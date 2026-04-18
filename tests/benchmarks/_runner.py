"""
Runners: long-running subprocess workers that speak the NDJSON protocol.

  parent → worker   {"path": "<abs path>", "request_id": "..."}\n
  worker → parent   one NDJSON record per input, then `{"event":"done"}` on EOF

Workers must also emit `{"event":"ready", ...}` on startup before the first
request is sent. The runner waits for that handshake before returning.
"""



import json
import os
import signal
import subprocess
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterator

HERE = Path(__file__).resolve().parent
REPO_ROOT = HERE.parent.parent


@dataclass
class RunnerConfig:
    name: str
    cmd: list[str]
    cwd: Path | None = None
    env: dict[str, str] | None = None
    batch_size: int = 50  # respawn after this many files
    per_file_timeout_s: float = 120.0


class Runner:
    """One long-running worker process. Restarts on every batch_size files."""

    def __init__(self, cfg: RunnerConfig) -> None:
        self.cfg = cfg
        self._proc: subprocess.Popen[str] | None = None
        self._files_since_start = 0
        self._ready_version: str | None = None

    # ------------------------------------------------------------------ lifecycle

    def _spawn(self) -> None:
        env = os.environ.copy()
        if self.cfg.env:
            env.update(self.cfg.env)
        self._proc = subprocess.Popen(
            self.cfg.cmd,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            cwd=str(self.cfg.cwd) if self.cfg.cwd else None,
            env=env,
            text=True,
            bufsize=1,  # line-buffered
        )
        assert self._proc.stdout is not None
        self._files_since_start = 0
        ready_line = self._read_line_with_timeout(10.0)
        if ready_line is None:
            self._terminate()
            raise RuntimeError(f"[{self.cfg.name}] worker never said ready")
        try:
            msg = json.loads(ready_line)
        except json.JSONDecodeError as exc:
            self._terminate()
            raise RuntimeError(f"[{self.cfg.name}] bad ready line: {ready_line!r}") from exc
        if msg.get("event") != "ready":
            self._terminate()
            raise RuntimeError(f"[{self.cfg.name}] expected ready, got {msg}")
        self._ready_version = msg.get("version")

    def _terminate(self) -> None:
        if self._proc is None:
            return
        try:
            if self._proc.poll() is None:
                self._proc.stdin.close() if self._proc.stdin else None  # type: ignore[func-returns-value]
                self._proc.send_signal(signal.SIGTERM)
                try:
                    self._proc.wait(timeout=2.0)
                except subprocess.TimeoutExpired:
                    self._proc.kill()
                    self._proc.wait(timeout=2.0)
        except Exception:  # noqa: BLE001
            try:
                self._proc.kill()
            except Exception:  # noqa: BLE001
                pass
        self._proc = None

    def stop(self) -> None:
        self._terminate()

    # ------------------------------------------------------------------ I/O

    def _read_line_with_timeout(self, timeout_s: float) -> str | None:
        """Blocking readline with a wall-clock timeout. Returns None on timeout."""
        assert self._proc is not None and self._proc.stdout is not None
        deadline = time.monotonic() + timeout_s
        # Cheap timeout: we poll os.read on the fd. This isn't perfect but works
        # fine for the harness — the worker always emits a line per file, so if
        # we don't see one within per_file_timeout_s the worker is stuck.
        import select

        fd = self._proc.stdout.fileno()
        buf = b""
        while True:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                return None
            r, _, _ = select.select([fd], [], [], remaining)
            if not r:
                return None
            chunk = os.read(fd, 8192)
            if not chunk:
                return None
            buf += chunk
            if b"\n" in buf:
                line, _, rest = buf.partition(b"\n")
                if rest:
                    # Stuff the unconsumed tail back for the next read — rare,
                    # since workers write exactly one line per event.
                    self._pending = rest  # type: ignore[attr-defined]
                    # But we don't have a real pushback. Simpler: the worker
                    # protocol guarantees exactly one line per request, so we
                    # shouldn't hit this path. Log and continue.
                return line.decode("utf-8", errors="replace").rstrip("\r")

    # ------------------------------------------------------------------ per-file

    def run(self, path: Path) -> dict[str, Any]:
        """Send one path, wait for one record, return parsed dict.

        Enforces per_file_timeout_s. On timeout, kills the worker so the next
        call re-spawns a fresh one. The returned dict always has status set —
        'ok' / 'error' / 'timeout' / 'oom'.
        """
        if self._proc is None or self._proc.poll() is not None:
            self._spawn()
        if self._files_since_start >= self.cfg.batch_size:
            self._terminate()
            self._spawn()

        assert self._proc is not None and self._proc.stdin is not None

        request = json.dumps({"path": str(path), "request_id": f"{id(path):x}"}) + "\n"
        try:
            self._proc.stdin.write(request)
            self._proc.stdin.flush()
        except BrokenPipeError:
            return _crash_record(path, self.cfg.name, self._ready_version,
                                 "worker pipe broken before sending request")

        line = self._read_line_with_timeout(self.cfg.per_file_timeout_s)
        self._files_since_start += 1

        if line is None:
            self._terminate()  # force respawn on next call
            return _timeout_record(path, self.cfg.name, self._ready_version,
                                   self.cfg.per_file_timeout_s)

        try:
            rec = json.loads(line)
        except json.JSONDecodeError:
            return _crash_record(path, self.cfg.name, self._ready_version,
                                 f"non-JSON line from worker: {line[:200]!r}")

        # If worker emitted an event-frame instead of a record, treat as error.
        if "event" in rec and "status" not in rec:
            return _crash_record(path, self.cfg.name, self._ready_version,
                                 f"worker event instead of record: {rec}")
        return rec


def _timeout_record(path: Path, parser: str, version: str | None, timeout_s: float) -> dict[str, Any]:
    return _blank_record(path, parser, version, "timeout", f"exceeded {timeout_s:.0f}s")


def _crash_record(path: Path, parser: str, version: str | None, msg: str) -> dict[str, Any]:
    return _blank_record(path, parser, version, "error", msg[:500])


def _blank_record(path: Path, parser: str, version: str | None, status: str, err: str) -> dict[str, Any]:
    try:
        size = path.stat().st_size
    except OSError:
        size = 0
    return {
        "file": str(path),
        "file_size_bytes": size,
        "parser": parser,
        "parser_version": version or "",
        "status": status,
        "error": err,
        "parse_time_ms": None,
        "peak_memory_mb": None,
        "sheets": None,
        "cells": None,
        "formulas": None,
        "formula_dependencies": None,
        "charts": None,
        "chart_types": None,
        "tables": None,
        "pivots": None,
        "merges": None,
        "cf_rules": None,
        "dv_rules": None,
        "named_ranges": None,
        "hyperlinks": None,
        "images": None,
        "comments": None,
        "sparklines": None,
        "chunks": None,
        "token_count": None,
        "schema_version": 1,
        "timestamp": "",
        "harness_commit": os.environ.get("HARNESS_COMMIT", ""),
    }


# ------------------------------------------------------------------ factories

def ks_runner(python_bin: str | None = None, timeout_s: float = 120.0) -> Runner:
    """Runner for the ks-xlsx-parser Python adapter."""
    py = python_bin or sys.executable
    return Runner(RunnerConfig(
        name="ks-xlsx-parser",
        cmd=[py, "-m", "tests.benchmarks.adapters.ks_adapter"],
        cwd=REPO_ROOT,
        per_file_timeout_s=timeout_s,
    ))


def hucre_runner(timeout_s: float = 120.0) -> Runner:
    """Runner for the hucre Node adapter."""
    node_dir = HERE / "hucre_node"
    adapter = node_dir / "hucre_adapter.mjs"
    if not adapter.exists():
        raise FileNotFoundError(f"hucre adapter missing: {adapter}")
    if not (node_dir / "node_modules" / "hucre").exists():
        raise FileNotFoundError(
            "hucre not installed. Run: "
            f"cd {node_dir} && pnpm install --frozen-lockfile"
        )
    return Runner(RunnerConfig(
        name="hucre",
        cmd=["node", "--max-old-space-size=4096", str(adapter)],
        cwd=node_dir,
        per_file_timeout_s=timeout_s,
    ))
