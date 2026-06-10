#!/usr/bin/env python3
"""Append the latest retrieval-benchmark run to a commit-over-commit history.

``eval_retrieval.py`` writes a timestamped ``summary.json`` per run. This
script picks the most recent one, flattens the headline metrics, tags it
with the current git commit, and appends one JSON line to
``tests/benchmarks/reports/history.jsonl``.

That history file is what makes "is recall improving over time?" answerable
— plot it, diff it in CI, or just ``tail`` it. Goal: text recall@5 > 0.90.

Usage:
    python scripts/append_bench_history.py
    python scripts/append_bench_history.py --reports-dir tests/benchmarks/reports/retrieval
"""
from __future__ import annotations

import argparse
import json
import subprocess
from datetime import UTC, datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
HISTORY = ROOT / "tests" / "benchmarks" / "reports" / "history.jsonl"


def git_commit() -> str:
    try:
        return subprocess.check_output(
            ["git", "rev-parse", "--short", "HEAD"], cwd=ROOT, text=True
        ).strip()
    except Exception:
        return "unknown"


def latest_summary(reports_dir: Path) -> Path:
    summaries = sorted(reports_dir.glob("*/summary.json"))
    if not summaries:
        raise SystemExit(
            f"no summary.json under {reports_dir} — run `make bench-retrieval` first"
        )
    return summaries[-1]


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--reports-dir", type=Path,
                    default=ROOT / "tests" / "benchmarks" / "reports" / "retrieval")
    ap.add_argument("--parser", default="excel-parser",
                    help="which parser's metrics to record")
    args = ap.parse_args(argv)

    summary_path = latest_summary(args.reports_dir)
    summary = json.loads(summary_path.read_text())
    metrics = summary.get(args.parser)
    if metrics is None:
        raise SystemExit(
            f"parser {args.parser!r} not in {summary_path}; "
            f"have: {list(summary)}"
        )

    row = {
        "timestamp": datetime.now(UTC).isoformat(),
        "commit": git_commit(),
        "parser": args.parser,
        "run": summary_path.parent.name,
        "instances": metrics.get("instances"),
        "in_scope_instances": metrics.get("in_scope_instances"),
        "out_of_scope_instances": metrics.get("out_of_scope_instances"),
        "recall_text@1": metrics.get("recall_text@1"),
        "recall_text@3": metrics.get("recall_text@3"),
        "recall_text@5": metrics.get("recall_text@5"),
        "recall_geometric@5": metrics.get("recall_geometric@5"),
        # In-scope numbers are the gate per the recall-90 roadmap.
        "recall_text@5_in_scope": metrics.get("recall_text@5_in_scope"),
        "recall_geometric@5_in_scope": metrics.get("recall_geometric@5_in_scope"),
        "table_fragmentation_rate": metrics.get("table_fragmentation_rate"),
        "mean_parse_ms": metrics.get("mean_parse_ms"),
        "errors": metrics.get("errors"),
        "failure_buckets": metrics.get("failure_buckets"),
    }

    HISTORY.parent.mkdir(parents=True, exist_ok=True)
    with HISTORY.open("a") as f:
        f.write(json.dumps(row, separators=(",", ":")) + "\n")

    print(f"appended to {HISTORY.relative_to(ROOT)}:")
    print(f"  commit {row['commit']}  recall_text@5={row['recall_text@5']}  "
          f"in_scope={row['recall_text@5_in_scope']}")

    # Show the trend if there's history to compare against.
    rows = [json.loads(ln) for ln in HISTORY.read_text().splitlines() if ln.strip()]
    if len(rows) >= 2:
        prev, cur = rows[-2], rows[-1]
        for k in ("recall_text@5", "recall_text@5_in_scope",
                  "recall_geometric@5", "recall_geometric@5_in_scope"):
            p, c = prev.get(k), cur.get(k)
            if isinstance(p, int | float) and isinstance(c, int | float):
                delta = c - p
                arrow = "▲" if delta > 0 else ("▼" if delta < 0 else "—")
                print(f"  {k}: {p:.4f} → {c:.4f}  {arrow} {delta:+.4f}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
