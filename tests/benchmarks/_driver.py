"""
Driver: enumerates the corpus, schedules work across parser runners, streams
records to CSV/JSONL, and produces summary + drift reports.

Deliberately simple: one file → all parsers in sequence → write → next file.
Randomized file order (seeded) prevents disk-cache from biasing whichever
parser goes second. Pre-warm is unnecessary: across 1053 files × 2 parsers
the OS page cache settles quickly either way; randomization does the job.
"""



import csv
import json
import platform
import random
import statistics
import subprocess
import sys
from collections import defaultdict
from dataclasses import asdict
from datetime import UTC, datetime
from pathlib import Path
from typing import Iterable

from ._runner import Runner
from ._schema import CSV_FIELDS, BenchmarkRecord, record_to_csv_row, validate_record


def enumerate_corpus(corpus_root: Path, subset: list[str] | None, sample: int | None,
                    seed: int) -> list[Path]:
    """Return absolute paths of .xlsx/.xlsm under corpus_root, sorted then
    optionally filtered by subset dir and randomly sampled."""
    all_files = [
        p for p in corpus_root.rglob("*")
        if p.is_file() and p.suffix.lower() in {".xlsx", ".xlsm"}
    ]
    if subset:
        filtered: list[Path] = []
        for p in all_files:
            rel = p.relative_to(corpus_root)
            top = rel.parts[0] if rel.parts else ""
            if top in subset:
                filtered.append(p)
        all_files = filtered
    all_files.sort()  # deterministic pre-shuffle

    if sample is not None and sample < len(all_files):
        rng = random.Random(seed)
        all_files = rng.sample(all_files, sample)

    return all_files


def run_benchmark(
    files: list[Path],
    runners: dict[str, Runner],
    out_dir: Path,
    seed: int = 1337,
    progress_every: int = 25,
) -> Path:
    """Execute the benchmark. Writes results.csv, failures.jsonl, raw.ndjson,
    summary.md, drift.md, manifest.json to `out_dir`. Returns `out_dir`."""
    out_dir.mkdir(parents=True, exist_ok=True)
    csv_path = out_dir / "results.csv"
    jsonl_path = out_dir / "raw.ndjson"
    fail_path = out_dir / "failures.jsonl"
    manifest_path = out_dir / "manifest.json"

    # Randomize parser order per file to kill cache-warm bias.
    rng = random.Random(seed)

    total = len(files) * len(runners)
    done = 0
    per_status: dict[str, int] = defaultdict(int)

    with csv_path.open("w", newline="") as csv_f, \
         jsonl_path.open("w") as jsonl_f, \
         fail_path.open("w") as fail_f:
        writer = csv.writer(csv_f)
        writer.writerow(CSV_FIELDS)

        for file in files:
            parser_order = list(runners.keys())
            rng.shuffle(parser_order)

            for name in parser_order:
                runner = runners[name]
                try:
                    raw = runner.run(file)
                except Exception as exc:  # noqa: BLE001
                    raw = {
                        "file": str(file),
                        "file_size_bytes": file.stat().st_size if file.exists() else 0,
                        "parser": name,
                        "parser_version": "",
                        "status": "error",
                        "error": f"runner threw: {exc}",
                        "parse_time_ms": None, "peak_memory_mb": None,
                        "sheets": None, "cells": None, "formulas": None,
                        "formula_dependencies": None, "charts": None,
                        "chart_types": None, "tables": None, "pivots": None,
                        "merges": None, "cf_rules": None, "dv_rules": None,
                        "named_ranges": None, "hyperlinks": None,
                        "images": None, "comments": None, "sparklines": None,
                        "chunks": None, "token_count": None,
                        "schema_version": 1, "timestamp": "",
                        "harness_commit": "",
                    }
                try:
                    validate_record(raw)
                except ValueError as exc:
                    raw["status"] = "error"
                    raw["error"] = f"schema validation failed: {exc}"

                # Stream: full record → NDJSON; flat → CSV; failures → jsonl
                jsonl_f.write(json.dumps(raw, separators=(",", ":"), default=str) + "\n")
                rec = BenchmarkRecord.from_dict(raw)
                writer.writerow(record_to_csv_row(rec))
                if rec.status != "ok":
                    fail_f.write(json.dumps(raw, separators=(",", ":"), default=str) + "\n")

                per_status[f"{name}:{rec.status}"] += 1
                done += 1

            if done % progress_every == 0:
                _progress(done, total, per_status)

    # Final progress line so users see the totals
    _progress(done, total, per_status)

    # Generate reports
    generate_summary(out_dir)
    generate_drift(out_dir)
    write_manifest(out_dir, files, list(runners.keys()))

    return out_dir


def _progress(done: int, total: int, per_status: dict[str, int]) -> None:
    parts = " ".join(f"{k}={v}" for k, v in sorted(per_status.items()))
    sys.stderr.write(f"\r[{done}/{total}] {parts}")
    sys.stderr.flush()
    if done == total:
        sys.stderr.write("\n")


# ----------------------------------------------------------------- reports

def generate_summary(out_dir: Path) -> None:
    raw_path = out_dir / "raw.ndjson"
    rows = [json.loads(line) for line in raw_path.read_text().splitlines() if line.strip()]
    by_parser: dict[str, list[dict]] = defaultdict(list)
    for r in rows:
        by_parser[r["parser"]].append(r)

    parsers = sorted(by_parser.keys())
    lines: list[str] = []
    lines.append("# Benchmark Summary\n")
    lines.append(f"- Generated: {datetime.now(UTC).isoformat()}")
    lines.append(f"- Parsers: {', '.join(parsers)}")
    lines.append(f"- Files attempted: {len(rows) // max(len(parsers), 1)}")
    lines.append("")

    # Status matrix
    lines.append("## Status")
    lines.append("")
    header = "| Parser | ok | error | timeout | oom | total |"
    lines.append(header)
    lines.append("|" + "|".join(["---"] * 6) + "|")
    for p in parsers:
        recs = by_parser[p]
        statuses = {"ok": 0, "error": 0, "timeout": 0, "oom": 0}
        for r in recs:
            s = r.get("status")
            if s in statuses:
                statuses[s] += 1
        lines.append(
            f"| {p} | {statuses['ok']} | {statuses['error']} | "
            f"{statuses['timeout']} | {statuses['oom']} | {len(recs)} |"
        )
    lines.append("")

    # Feature presence (capability matrix)
    lines.append("## Feature capability")
    lines.append("")
    lines.append("A ✅ means the parser populated a non-null value for at least one ok file.")
    lines.append("Null everywhere = feature not modelled by that parser.")
    lines.append("")
    features = [
        "formulas", "formula_dependencies", "charts", "chart_types", "tables",
        "pivots", "merges", "cf_rules", "dv_rules", "named_ranges",
        "hyperlinks", "images", "comments", "sparklines", "chunks",
        "token_count",
    ]
    lines.append("| Feature | " + " | ".join(parsers) + " |")
    lines.append("|" + "|".join(["---"] * (1 + len(parsers))) + "|")
    for f in features:
        row = [f]
        for p in parsers:
            has = any(r.get(f) is not None for r in by_parser[p] if r.get("status") == "ok")
            row.append("✅" if has else "—")
        lines.append("| " + " | ".join(row) + " |")
    lines.append("")

    # Aggregate counts
    lines.append("## Aggregate counts (status=ok only)")
    lines.append("")
    lines.append("| Feature | " + " | ".join(parsers) + " |")
    lines.append("|" + "|".join(["---"] * (1 + len(parsers))) + "|")
    count_fields = ["cells", "formulas", "charts", "tables", "pivots", "merges",
                    "cf_rules", "dv_rules", "named_ranges", "hyperlinks",
                    "images", "comments", "sparklines"]
    for f in count_fields:
        row = [f]
        for p in parsers:
            vals = [r.get(f) for r in by_parser[p]
                    if r.get("status") == "ok" and r.get(f) is not None]
            row.append(f"{sum(vals):,}" if vals else "—")
        lines.append("| " + " | ".join(row) + " |")
    lines.append("")

    # Perf
    lines.append("## Performance (status=ok only)")
    lines.append("")
    lines.append("| Parser | files | P50 ms | P95 ms | P99 ms | mean ms | total s | mean MB |")
    lines.append("|---|---|---|---|---|---|---|---|")
    for p in parsers:
        times = [r["parse_time_ms"] for r in by_parser[p]
                 if r.get("status") == "ok" and r.get("parse_time_ms") is not None]
        mems = [r["peak_memory_mb"] for r in by_parser[p]
                if r.get("status") == "ok" and r.get("peak_memory_mb") is not None]
        if not times:
            lines.append(f"| {p} | 0 | — | — | — | — | — | — |")
            continue
        times_sorted = sorted(times)
        p50 = times_sorted[len(times_sorted) // 2]
        p95 = times_sorted[int(len(times_sorted) * 0.95)] if len(times_sorted) > 1 else p50
        p99 = times_sorted[int(len(times_sorted) * 0.99)] if len(times_sorted) > 1 else p50
        mean_ms = statistics.mean(times)
        total_s = sum(times) / 1000.0
        mean_mb = statistics.mean(mems) if mems else 0.0
        lines.append(
            f"| {p} | {len(times)} | {p50:.1f} | {p95:.1f} | {p99:.1f} | "
            f"{mean_ms:.1f} | {total_s:.1f} | {mean_mb:.1f} |"
        )
    lines.append("")
    lines.append("> Memory is approximate (±30%). See `_mem.py` and the README for caveats.")
    lines.append("")

    # Per-sub-corpus breakdown
    lines.append("## Performance by sub-corpus")
    lines.append("")
    by_sub: dict[tuple[str, str], list[float]] = defaultdict(list)
    for r in rows:
        if r.get("status") != "ok" or r.get("parse_time_ms") is None:
            continue
        try:
            rel = Path(r["file"]).resolve()
            # Find segment after 'testBench/' or use file's parent name.
            parts = rel.parts
            if "testBench" in parts:
                idx = parts.index("testBench")
                sub = "/".join(parts[idx + 1: idx + 3]) if idx + 2 < len(parts) else parts[idx + 1]
            else:
                sub = rel.parent.name
        except Exception:  # noqa: BLE001
            sub = "?"
        by_sub[(r["parser"], sub)].append(r["parse_time_ms"])
    subs = sorted({k[1] for k in by_sub})
    lines.append("| Sub-corpus | " + " | ".join(f"{p} P50" for p in parsers) + " | "
                 + " | ".join(f"{p} P95" for p in parsers) + " |")
    lines.append("|" + "|".join(["---"] * (1 + 2 * len(parsers))) + "|")
    for sub in subs:
        cells_row = [sub]
        for p in parsers:
            tl = sorted(by_sub.get((p, sub), []))
            cells_row.append(f"{tl[len(tl)//2]:.1f}" if tl else "—")
        for p in parsers:
            tl = sorted(by_sub.get((p, sub), []))
            cells_row.append(f"{tl[int(len(tl)*0.95)]:.1f}" if len(tl) > 1 else ("—" if not tl else f"{tl[0]:.1f}"))
        lines.append("| " + " | ".join(cells_row) + " |")

    (out_dir / "summary.md").write_text("\n".join(lines) + "\n")


def generate_drift(out_dir: Path) -> None:
    raw_path = out_dir / "raw.ndjson"
    rows = [json.loads(line) for line in raw_path.read_text().splitlines() if line.strip()]

    # Index rows by file for pairwise comparison across parsers.
    by_file: dict[str, dict[str, dict]] = defaultdict(dict)
    for r in rows:
        by_file[r["file"]][r["parser"]] = r

    parsers = sorted({r["parser"] for r in rows})
    lines: list[str] = []
    lines.append("# Drift report\n")
    lines.append(f"- Generated: {datetime.now(UTC).isoformat()}")
    lines.append(f"- Files compared: {len(by_file)}")
    lines.append(f"- Parsers: {', '.join(parsers)}")
    lines.append("")
    lines.append("Drift = both parsers reported ok on this file AND both extracted a")
    lines.append("non-null value for the given feature, but the counts disagree.")
    lines.append("")

    count_fields = [
        ("cells", 0.01, 1),      # relative 1%, abs 1
        ("formulas", 0, 1),
        ("tables", 0, 1),
        ("merges", 0, 1),
        ("cf_rules", 0, 1),
        ("dv_rules", 0, 1),
        ("named_ranges", 0, 1),
        ("hyperlinks", 0, 1),
        ("images", 0, 1),
        ("comments", 0, 1),
    ]

    if len(parsers) != 2:
        lines.append("(drift report assumes exactly 2 parsers; skipping)")
        (out_dir / "drift.md").write_text("\n".join(lines) + "\n")
        return
    p_a, p_b = parsers

    for field_name, rel_th, abs_th in count_fields:
        lines.append(f"## {field_name}")
        lines.append("")
        lines.append(f"| file | {p_a} | {p_b} | diff |")
        lines.append("|---|---|---|---|")
        n_drift = 0
        for fname, entries in sorted(by_file.items()):
            a = entries.get(p_a)
            b = entries.get(p_b)
            if not a or not b:
                continue
            if a.get("status") != "ok" or b.get("status") != "ok":
                continue
            va = a.get(field_name)
            vb = b.get(field_name)
            if va is None or vb is None:
                continue
            diff = abs(va - vb)
            if diff < abs_th:
                continue
            if rel_th and max(va, vb) > 0 and diff / max(va, vb) < rel_th:
                continue
            n_drift += 1
            if n_drift <= 50:
                short = Path(fname).name
                lines.append(f"| {short} | {va} | {vb} | {va - vb:+d} |")
        if n_drift == 0:
            lines.append(f"| _no drift_ | | | |")
        elif n_drift > 50:
            lines.append(f"| … {n_drift - 50} more rows truncated | | | |")
        lines.append("")

    (out_dir / "drift.md").write_text("\n".join(lines) + "\n")


def write_manifest(out_dir: Path, files: list[Path], parsers: list[str]) -> None:
    try:
        git_sha = subprocess.check_output(
            ["git", "rev-parse", "HEAD"], cwd=str(out_dir.parent),
            stderr=subprocess.DEVNULL, text=True).strip()
    except Exception:  # noqa: BLE001
        git_sha = ""
    try:
        node_ver = subprocess.check_output(["node", "-v"], text=True).strip()
    except Exception:  # noqa: BLE001
        node_ver = ""
    manifest = {
        "timestamp": datetime.now(UTC).isoformat(),
        "file_count": len(files),
        "parsers": parsers,
        "git_sha": git_sha,
        "python_version": sys.version.split()[0],
        "platform": platform.platform(),
        "node_version": node_ver,
    }
    (out_dir / "manifest.json").write_text(
        json.dumps(manifest, indent=2, default=str) + "\n"
    )
