"""
Re-aggregate a retrieval-bench `results.ndjson` into summary.json / summary.md.

Useful when:
  - `eval_retrieval.py` was interrupted (Ctrl-C, OOM, hang-watchdog) and
    didn't get to run its end-of-run aggregator
  - You want to inspect aggregates without re-running

Usage:
    python scripts/summarize_retrieval.py <results.ndjson>

Writes summary.json + summary.md next to the input file. Same metrics
the live aggregator emits — keep the two in sync if you change them.
"""

from __future__ import annotations

import json
import sys
from collections import defaultdict
from pathlib import Path


def aggregate(records: list[dict]) -> dict:
    by_parser: dict[str, list[dict]] = defaultdict(list)
    for r in records:
        by_parser[r["parser"]].append(r)

    out: dict[str, dict] = {}
    for parser, recs in by_parser.items():
        total = len(recs)
        errors = sum(1 for r in recs if r.get("error"))
        ok = total - errors

        def _recall_at(k: int, key: str) -> float:
            denom = 0
            hits = 0
            for r in recs:
                if r.get("error"):
                    continue
                rank = r.get(key)
                if rank is None:
                    denom += 1
                    continue
                denom += 1
                if rank <= k:
                    hits += 1
            return hits / denom if denom else 0.0

        frags = [r["chunks_overlapping_data"] for r in recs
                 if not r.get("error") and r.get("data_regions") == 1
                 and r.get("chunks_overlapping_data", 0) > 0]
        n_clean = sum(1 for f in frags if f == 1)
        n_frag = len(frags) - n_clean
        frag_rate = (n_frag / len(frags)) if frags else 0.0

        parse_times = [r["parse_ms"] for r in recs
                       if not r.get("error") and r.get("parse_ms") is not None]

        out[parser] = {
            "instances": total,
            "ok": ok,
            "errors": errors,
            "recall_geometric@1": _recall_at(1, "rank_of_first_overlap"),
            "recall_geometric@3": _recall_at(3, "rank_of_first_overlap"),
            "recall_geometric@5": _recall_at(5, "rank_of_first_overlap"),
            "recall_text@1": _recall_at(1, "rank_of_text_match"),
            "recall_text@3": _recall_at(3, "rank_of_text_match"),
            "recall_text@5": _recall_at(5, "rank_of_text_match"),
            "table_integrity_clean": n_clean,
            "table_integrity_fragmented": n_frag,
            "table_fragmentation_rate": round(frag_rate, 4),
            "mean_parse_ms": round(sum(parse_times) / len(parse_times), 2)
            if parse_times else None,
            "p50_parse_ms": round(sorted(parse_times)[len(parse_times) // 2], 2)
            if parse_times else None,
        }
    return out


def render_md(summary: dict, source: Path, partial: bool) -> str:
    parsers = sorted(summary.keys())
    lines = ["# Retrieval-recall benchmark (SpreadsheetBench)\n"]
    lines.append(f"- Source NDJSON: `{source}`")
    n_total = sum(s["instances"] for s in summary.values())
    n_per = n_total // max(len(parsers), 1)
    lines.append(f"- Records: {n_total} ({n_per} per parser){'  ⚠️ PARTIAL RUN — bench interrupted before completion' if partial else ''}")
    lines.append("- Embedding model: `BAAI/bge-small-en-v1.5`")
    lines.append("")
    lines.append("| Metric | " + " | ".join(parsers) + " |")
    lines.append("|---|" + "|".join(["---"] * len(parsers)) + "|")
    metrics = [
        ("recall_geometric@1", "Recall@1 (geometric)"),
        ("recall_geometric@3", "Recall@3 (geometric)"),
        ("recall_geometric@5", "Recall@5 (geometric)"),
        ("recall_text@1", "Recall@1 (text-match)"),
        ("recall_text@3", "Recall@3 (text-match)"),
        ("recall_text@5", "Recall@5 (text-match)"),
        ("table_fragmentation_rate", "Fragmentation rate"),
        ("mean_parse_ms", "Mean parse ms"),
        ("p50_parse_ms", "P50 parse ms"),
        ("errors", "Errors"),
    ]
    for key, label in metrics:
        row = [label]
        for p in parsers:
            v = summary[p].get(key)
            if v is None:
                row.append("—")
            elif isinstance(v, float):
                row.append(f"{v:.3f}")
            else:
                row.append(str(v))
        lines.append("| " + " | ".join(row) + " |")
    lines.append("")
    lines.append("**Geometric overlap** = chunk's reported A1 range overlaps the "
                 "ground-truth `data_position`. Requires the parser to surface "
                 "(sheet, range) per chunk — docling does not, so its geometric "
                 "recall is structurally 0.")
    lines.append("")
    lines.append("**Text-match** = the answer cell's actual string value appears "
                 "as a substring of the chunk's text, after numeric/date/boolean "
                 "normalization on both sides. Parser-agnostic; this is the "
                 "apples-to-apples retrieval comparison.")
    return "\n".join(lines) + "\n"


def main(argv: list[str]) -> int:
    if len(argv) != 2:
        sys.stderr.write("usage: python scripts/summarize_retrieval.py <results.ndjson>\n")
        return 2
    ndjson = Path(argv[1]).resolve()
    if not ndjson.exists():
        sys.stderr.write(f"file not found: {ndjson}\n")
        return 2

    records = [json.loads(line) for line in ndjson.read_text().splitlines() if line.strip()]
    summary = aggregate(records)

    out_dir = ndjson.parent
    (out_dir / "summary.json").write_text(json.dumps(summary, indent=2) + "\n")

    # If counts per parser are unequal, treat as partial.
    counts = [summary[p]["instances"] for p in summary]
    partial = len(set(counts)) > 1 if counts else False
    (out_dir / "summary.md").write_text(render_md(summary, ndjson, partial))

    sys.stderr.write(f"Wrote {out_dir / 'summary.json'}\n")
    sys.stderr.write(f"Wrote {out_dir / 'summary.md'}\n")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
