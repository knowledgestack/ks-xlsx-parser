"""Aggregate metrics emitted by enterprise/corpus tests."""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent
METRICS_DIR = ROOT / "metrics" / "corpus"


def track_metrics() -> dict:
    results = {}
    if not METRICS_DIR.exists():
        return {}

    for metric_file in METRICS_DIR.glob("*_scorecard.json"):
        with open(metric_file) as f:
            results[metric_file.stem] = json.load(f)

    if not results:
        return {}

    summary = {
        "timestamp": datetime.now().isoformat(),
        "files": len(results),
        "composite_avg": sum(v.get("composite", 0) for v in results.values())
        / max(len(results), 1),
    }

    summary_path = ROOT / "metrics" / "corpus_summary.json"
    summary_path.parent.mkdir(parents=True, exist_ok=True)
    with open(summary_path, "w") as f:
        json.dump(summary, f, indent=2)

    return summary


if __name__ == "__main__":
    summary = track_metrics()
    if summary:
        print(json.dumps(summary, indent=2))
    else:
        print("No metrics to summarize (run tests first).")
