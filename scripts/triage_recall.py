#!/usr/bin/env python3
"""Triage retrieval-recall failures into a ranked, actionable worklist.

Reads a ``failures.ndjson`` produced by
``eval_retrieval.py --emit-failures`` and prints:

  1. A histogram of failure buckets (biggest = highest leverage).
  2. For each bucket, a few concrete example failures with the
     ground-truth answer and the top ranked chunks the parser produced.

The point: turn "recall@5 is 0.70" into "N misses are answer_absent_from_chunks
— the parser is dropping cells; here are 5 examples to reproduce."

Usage:
    python scripts/triage_recall.py tests/benchmarks/reports/retrieval/<stamp>/failures.ndjson
    python scripts/triage_recall.py <dir>           # finds the latest run
    python scripts/triage_recall.py <file> --bucket answer_absent_from_chunks --examples 10
"""
from __future__ import annotations

import argparse
import json
import sys
from collections import Counter
from pathlib import Path

# Ordered worst→least-actionable. Used to sort the worklist.
BUCKET_GUIDANCE = {
    "parse_error": "Parser raised on the file. Reproduce with parse_workbook(path) and fix the crash.",
    "no_chunks": "Parser produced zero chunks. Sheet/region detection collapsed — check chunking/segmenter.py.",
    "answer_absent_from_chunks": "Answer value is in NO chunk. EXTRACTION gap — the cell was dropped or garbled. Highest leverage.",
    "wrong_sheet": "Answer sheet was never chunked. Sheet enumeration bug — check workbook_parser.py sheet loop.",
    "geometric_no_overlap": "No chunk's A1 range overlaps ground truth. RANGE bookkeeping drift in merge/split.",
    "present_but_ranked_low": "A chunk DOES contain the answer but ranked >5. Not a parser bug — fix chunk granularity/embedding.",
    "unparseable_ground_truth": "Could not parse the dataset's answer_position. Benchmark-harness issue, not the parser.",
}


def find_failures_file(arg: Path) -> Path:
    if arg.is_file():
        return arg
    if arg.is_dir():
        direct = arg / "failures.ndjson"
        if direct.exists():
            return direct
        runs = sorted(p for p in arg.glob("*/failures.ndjson"))
        if runs:
            return runs[-1]
    sys.exit(f"ERROR: no failures.ndjson found at {arg}")


def load(path: Path) -> list[dict]:
    rows = []
    for line in path.read_text().splitlines():
        line = line.strip()
        if line:
            rows.append(json.loads(line))
    return rows


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("path", type=Path,
                    help="failures.ndjson, or a dir containing/parenting one")
    ap.add_argument("--bucket", help="only show examples for this bucket")
    ap.add_argument("--examples", type=int, default=3,
                    help="example failures to print per bucket (default 3)")
    args = ap.parse_args(argv)

    path = find_failures_file(args.path)
    rows = load(path)
    if not rows:
        print(f"{path} is empty — no failures recorded (or run was a no-op).")
        return 0

    print(f"# Recall failure triage — {path}")
    print(f"# {len(rows)} total failure rows\n")

    hist = Counter(r.get("failure_bucket") for r in rows)
    print("## Bucket histogram (ranked by count — fix the top one first)\n")
    width = max(len(b or "?") for b in hist)
    for bucket, count in hist.most_common():
        pct = 100.0 * count / len(rows)
        print(f"  {str(bucket):<{width}}  {count:5d}  ({pct:5.1f}%)")
        print(f"  {'':<{width}}         → {BUCKET_GUIDANCE.get(bucket, '')}")
    print()

    buckets = [args.bucket] if args.bucket else [b for b, _ in hist.most_common()]
    for bucket in buckets:
        examples = [r for r in rows if r.get("failure_bucket") == bucket][:args.examples]
        if not examples:
            continue
        print(f"## Examples — {bucket}\n")
        for r in examples:
            print(f"  instance {r.get('instance_id')}  ({r.get('parser')})")
            print(f"    Q: {(r.get('instruction') or '')[:160]}")
            print(f"    answer_position: {r.get('answer_position')}")
            print(f"    ground-truth values: {r.get('answer_values')}")
            print(f"    n_chunks={r.get('n_chunks')} "
                  f"rank_of_text_match={r.get('rank_of_text_match')}")
            if r.get("error"):
                print(f"    ERROR: {r['error']}")
            for c in (r.get("top_chunks") or [])[:4]:
                mark = "✓" if c.get("contains_answer") else " "
                snippet = (c.get("text") or "").replace("\n", " ")[:120]
                print(f"    [{mark}] #{c.get('rank')} {c.get('sheet')} "
                      f"{c.get('range')}  {snippet}")
            print()
    return 0


if __name__ == "__main__":
    sys.exit(main())
