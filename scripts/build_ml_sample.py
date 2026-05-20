#!/usr/bin/env python3
"""Build the stratified 50-sheet sample for ML iteration.

Run once after a clean ``enriched_failures.ndjson`` exists. Writes
``data/labels/sample_50_seed1337.json``.

Usage:
    python scripts/build_ml_sample.py
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from ks_xlsx_parser.ml.sample import build_sample, save_sample


def main() -> int:
    corpus = ROOT / "data" / "corpora" / "spreadsheetbench" / "all_data_912_v0.1"
    reports = ROOT / "tests" / "benchmarks" / "reports" / "retrieval"
    runs = sorted(reports.glob("*/enriched_failures.ndjson"))
    if not runs:
        print(
            "ERROR: no enriched_failures.ndjson under "
            f"{reports}. Run `make bench-track` + "
            "`python scripts/enrich_failures.py` first.",
            file=sys.stderr,
        )
        return 1
    enriched = runs[-1]
    dataset = corpus / "dataset.json"
    ids = build_sample(enriched, dataset, n=50, seed=1337)
    save_sample(ids, meta={
        "source_enriched": str(enriched.relative_to(ROOT)),
        "corpus": str(dataset.relative_to(ROOT)),
        "n": len(ids),
        "seed": 1337,
    })
    print(f"Built 50-sheet stratified sample → "
          f"{ROOT / 'data/labels/sample_50_seed1337.json'}")
    print(f"  ids: {ids[:10]}{' ...' if len(ids) > 10 else ''}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
