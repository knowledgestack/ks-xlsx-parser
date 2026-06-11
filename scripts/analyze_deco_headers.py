"""
Cohort + failure-bucket analysis of DECO header detection from an eval_deco run.

Splits GT-headered tables into single-row vs multi-row header cohorts and scores
each on the ISOLATED prediction (detector run on the GT region — measures the
header detector itself, independent of segmentation quality) and end-to-end.

Buckets every imperfect isolated prediction so failures can be pulled and
diagnosed against the actual workbook:

    SR-OVER   single-row GT, prediction is a strict superset (over-extension)
    SR-SHIFT  single-row GT, prediction wrong row(s) entirely
    SR-MISS   single-row GT, empty prediction
    MR-UNDER  multi-row GT, prediction is a non-empty strict subset
    MR-SHIFT  multi-row GT, prediction overlaps partially / disjoint
    MR-OVER   multi-row GT, prediction is a strict superset
    MR-MISS   multi-row GT, empty prediction

Usage:
    PYTHONPATH=src python scripts/analyze_deco_headers.py --run RUN_DIR \
        [--dump-buckets OUT.json] [--corpus data/corpora/deco/completed]
"""

from __future__ import annotations

import argparse
import json
from collections import Counter, defaultdict
from pathlib import Path


def _read_ndjson(path: Path) -> list[dict]:
    out = []
    with path.open() as fh:
        for line in fh:
            line = line.strip()
            if line:
                out.append(json.loads(line))
    return out


def _prf(pred: set[int], gt: set[int]) -> tuple[int, int, int]:
    return (len(pred & gt), len(pred), len(gt))


def _f1(tp: int, pred: int, gt: int) -> tuple[float, float, float]:
    p = tp / pred if pred else 0.0
    r = tp / gt if gt else 0.0
    f = 2 * p * r / (p + r) if (p + r) else 0.0
    return p, r, f


def bucket_of(pred: set[int], gt: set[int], multirow: bool) -> str | None:
    if pred == gt:
        return None
    pre = "MR" if multirow else "SR"
    if not pred:
        return f"{pre}-MISS"
    if pred > gt:
        return f"{pre}-OVER"
    if pred < gt:
        return f"{pre}-UNDER"
    return f"{pre}-SHIFT"


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--run", required=True, help="eval_deco output dir (with ks.ndjson)")
    ap.add_argument("--dump-buckets", default=None, help="write per-table failure records")
    ap.add_argument(
        "--max-per-bucket", type=int, default=0, help="cap exemplars per bucket in the dump (0 = no cap)"
    )
    args = ap.parse_args()

    recs = _read_ndjson(Path(args.run) / "ks.ndjson")

    # cohort accumulators: key -> [tp, pred, gt, exact, n]
    stats: dict[str, list[int]] = defaultdict(lambda: [0, 0, 0, 0, 0])
    buckets: Counter[str] = Counter()
    failures: list[dict] = []

    for rec in recs:
        for t in rec.get("tables", []):
            if not t["has_gt_header"]:
                continue
            gt = set(t["gt_hrows"])
            multirow = len(gt) > 1
            cohort = "multi" if multirow else "single"
            for mode, key in (("ks_hrows_iso", "iso"), ("ks_hrows_e2e", "e2e")):
                pred = set(t[mode])
                tp, np_, ng = _prf(pred, gt)
                s = stats[f"{cohort}/{key}"]
                s[0] += tp
                s[1] += np_
                s[2] += ng
                s[3] += pred == gt
                s[4] += 1
                a = stats[f"all/{key}"]
                a[0] += tp
                a[1] += np_
                a[2] += ng
                a[3] += pred == gt
                a[4] += 1
            pred_iso = set(t["ks_hrows_iso"])
            b = bucket_of(pred_iso, gt, multirow)
            if b:
                buckets[b] += 1
                failures.append(
                    {
                        "bucket": b,
                        "file": rec["file"],
                        "sheet": t["sheet"],
                        "gt_range": t["gt_range"],
                        "gt_hrows": t["gt_hrows"],
                        "ks_hrows_iso": t["ks_hrows_iso"],
                        "ks_hrows_e2e": t["ks_hrows_e2e"],
                    }
                )

    print(f"{'cohort':<14}{'P':>8}{'R':>8}{'F1':>8}{'exact':>10}{'n':>7}")
    for key in ["single/iso", "multi/iso", "all/iso", "single/e2e", "multi/e2e", "all/e2e"]:
        tp, np_, ng, ex, n = stats[key]
        p, r, f = _f1(tp, np_, ng)
        print(f"{key:<14}{p:>8.3f}{r:>8.3f}{f:>8.3f}{ex:>6}/{n:<4}{100 * ex / n if n else 0:>5.1f}%")

    print("\nfailure buckets (isolated):")
    for b, c in buckets.most_common():
        print(f"  {b:<10} {c}")

    if args.dump_buckets:
        if args.max_per_bucket:
            capped: list[dict] = []
            seen: Counter[str] = Counter()
            for f_ in failures:
                if seen[f_["bucket"]] < args.max_per_bucket:
                    capped.append(f_)
                    seen[f_["bucket"]] += 1
            failures = capped
        Path(args.dump_buckets).write_text(json.dumps(failures, indent=1))
        print(f"\n{len(failures)} failure records → {args.dump_buckets}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
