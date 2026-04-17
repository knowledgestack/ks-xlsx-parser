"""
testBench round-trip tests.

Parses every .xlsx under ``testBench/`` and asserts:

* ``parse_workbook()`` returns without raising.
* ``result.to_json()`` produces non-empty JSON (> 100 bytes).
* ``result.workbook`` has at least one sheet.

Failures are collected into ``metrics/testbench/failures.json`` so parser
regressions across the whole bench are easy to diff.

Runs under the ``testbench`` marker only (skipped by default). Invoke with:

    pytest tests/test_testbench_roundtrip.py -m testbench -q
    make testbench   # convenience wrapper
"""
from __future__ import annotations

import json
import os
import traceback
from pathlib import Path

import pytest

from xlsx_parser import parse_workbook

ROOT = Path(__file__).resolve().parent.parent
TESTBENCH_DIR = ROOT / "testBench"
METRICS_DIR = ROOT / "metrics" / "testbench"
FAILURES_PATH = METRICS_DIR / "failures.json"
FAILURES_JSONL = METRICS_DIR / "failures.jsonl"  # append-only, xdist-safe


def _collect_files() -> list[Path]:
    if not TESTBENCH_DIR.exists():
        return []
    return sorted(TESTBENCH_DIR.rglob("*.xlsx"))


ALL_FILES = _collect_files()

pytestmark = [pytest.mark.testbench, pytest.mark.timeout(60)]


def _record_failure(entry: dict) -> None:
    """Append one failure row to the JSONL log. Safe under xdist parallelism."""
    METRICS_DIR.mkdir(parents=True, exist_ok=True)
    entry["worker"] = os.environ.get("PYTEST_XDIST_WORKER", "main")
    with FAILURES_JSONL.open("a", encoding="utf-8") as f:
        f.write(json.dumps(entry) + "\n")


@pytest.fixture(scope="session", autouse=True)
def _reset_log():
    """Reset the append log at the start of the session (master worker only)."""
    # Under xdist, PYTEST_XDIST_WORKER is set for workers but not the master.
    # The master is responsible for cleanup before workers start writing.
    if os.environ.get("PYTEST_XDIST_WORKER") is None:
        METRICS_DIR.mkdir(parents=True, exist_ok=True)
        if FAILURES_JSONL.exists():
            FAILURES_JSONL.unlink()
    yield
    # After session, aggregate JSONL → JSON summary (master only)
    if os.environ.get("PYTEST_XDIST_WORKER") is None:
        failures: list[dict] = []
        if FAILURES_JSONL.exists():
            for line in FAILURES_JSONL.read_text().splitlines():
                if line.strip():
                    failures.append(json.loads(line))
        FAILURES_PATH.write_text(
            json.dumps(
                {"total": len(ALL_FILES), "failure_count": len(failures), "failures": failures},
                indent=2,
            )
        )


def _relpath(p: Path) -> str:
    return str(p.relative_to(ROOT))


@pytest.mark.parametrize("path", ALL_FILES, ids=lambda p: _relpath(p))
def test_parse_roundtrip(path: Path):
    """Each workbook must parse, serialize to JSON, and report ≥1 sheet."""
    try:
        result = parse_workbook(path=path)
    except Exception as exc:
        _record_failure({
            "file": _relpath(path),
            "stage": "parse",
            "error": f"{type(exc).__name__}: {exc}",
            "traceback": traceback.format_exc(limit=5),
        })
        raise

    assert result.workbook is not None, f"no workbook DTO for {path}"
    assert result.workbook.total_sheets >= 1, f"{path} reports zero sheets"

    try:
        js = result.to_json()
    except Exception as exc:
        _record_failure({
            "file": _relpath(path),
            "stage": "to_json",
            "error": f"{type(exc).__name__}: {exc}",
            "traceback": traceback.format_exc(limit=5),
        })
        raise

    assert isinstance(js, dict), f"to_json returned non-dict for {path}"
    assert "workbook" in js, f"to_json result missing 'workbook' key for {path}"
    try:
        encoded = json.dumps(js, default=str)
    except Exception as exc:
        _record_failure({
            "file": _relpath(path),
            "stage": "json_encode",
            "error": f"{type(exc).__name__}: {exc}",
            "traceback": traceback.format_exc(limit=5),
        })
        raise
    assert len(encoded) > 100, f"encoded JSON suspiciously short ({len(encoded)} chars) for {path}"


def test_testbench_has_files():
    """Guard against an empty testBench (e.g. missing dataset zip)."""
    assert ALL_FILES, (
        f"No .xlsx files found under {TESTBENCH_DIR}. "
        "Run `make testbench-build` or download the dataset zip from the GitHub release."
    )
