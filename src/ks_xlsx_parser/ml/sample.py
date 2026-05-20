"""Stratified 50-sheet sample for fast iteration on ML changes.

A full 912 bench pass takes ~70 min with the LLM enabled. While
iterating on a vertical (prompt design + integration code) you want
a much tighter loop. This sample is the agreed harness for that loop:

  * 50 instances, deterministic (`seed=1337`)
  * Stratified so each of these populations is over-represented vs.
    a random draw — the worst-case patterns are the ones that move
    the needle:
      - text_hit_geom_miss  (cluster-02 territory: range bookkeeping)
      - n_chunks_on_gt_sheet == 1 with large bbox (cluster-04 territory)
      - gt_cell_is_formula (cluster-01 / 03 territory)
      - instruction_requires_execution = False (must be in-scope —
        OOS instances are unscorable by retrieval and would just be
        noise in a sample meant to surface parser deltas)

The sample's instance ids are written into a JSON file under
``data/labels/sample_50_seed1337.json`` so it's reproducible across
runs, agent worktrees, and CI.
"""
from __future__ import annotations

import json
import random
from pathlib import Path
from typing import Any

SAMPLE_FILE = Path(__file__).resolve().parents[3] / "data" / "labels" / "sample_50_seed1337.json"


def build_sample(
    enriched_failures_path: Path,
    dataset_path: Path,
    n: int = 50,
    seed: int = 1337,
) -> list[str]:
    """Construct (or reproduce) the stratified sample of instance ids.

    Reads ``enriched_failures.ndjson`` to know each instance's bucket /
    flags, then samples deterministically. Run once before the agents
    fan out so they all measure against the same 50 sheets.
    """
    failures = [json.loads(line) for line in enriched_failures_path.read_text().splitlines()
                if line.strip()]
    dataset = json.loads(dataset_path.read_text())

    by_id = {str(d["id"]): d for d in dataset}
    enriched = {str(r["instance_id"]): r for r in failures}

    # Bucket each instance into one of four strata (priority order).
    strata: dict[str, list[str]] = {
        "text_hit_geom_miss": [],
        "whole_sheet_one_chunk": [],
        "formula_cell_answer": [],
        "in_scope_other": [],
    }
    seen: set[str] = set()
    for iid, r in enriched.items():
        flags = set(r.get("flags") or [])
        if "instruction_requires_execution" in flags:
            continue  # OOS — useless in a parser-quality sample
        if iid in seen:
            continue
        if r.get("bucket_combined") == "text_hit_geom_miss":
            strata["text_hit_geom_miss"].append(iid)
        elif (r.get("n_chunks_on_gt_sheet") == 1
              and r.get("n_chunks_total", 99) <= 2):
            strata["whole_sheet_one_chunk"].append(iid)
        elif "gt_cell_is_formula" in flags:
            strata["formula_cell_answer"].append(iid)
        else:
            strata["in_scope_other"].append(iid)
        seen.add(iid)

    # Round-robin draw, deterministic order via seeded shuffle inside each
    # stratum. Targets: 20 / 12 / 8 / 10 (sum = 50). Top up from
    # "in_scope_other" if a stratum is short.
    rng = random.Random(seed)
    for k in strata:
        rng.shuffle(strata[k])
    targets = {"text_hit_geom_miss": 20, "whole_sheet_one_chunk": 12,
               "formula_cell_answer": 8, "in_scope_other": 10}
    out: list[str] = []
    for k, n_target in targets.items():
        out.extend(strata[k][:n_target])
    # Pad to `n` from passing in-scope instances (i.e., dataset entries
    # not in the failure set, so the sample also has control cases).
    if len(out) < n:
        passing = [str(d["id"]) for d in dataset
                   if str(d["id"]) not in enriched and str(d["id"]) not in out]
        rng.shuffle(passing)
        out.extend(passing[: n - len(out)])
    return out[:n]


def load_sample() -> list[str]:
    """Read the persisted sample. Raises FileNotFoundError if not built yet."""
    return json.loads(SAMPLE_FILE.read_text())["instance_ids"]


def save_sample(instance_ids: list[str], meta: dict[str, Any] | None = None) -> None:
    SAMPLE_FILE.parent.mkdir(parents=True, exist_ok=True)
    SAMPLE_FILE.write_text(json.dumps({
        "instance_ids": instance_ids,
        "meta": meta or {},
    }, indent=2))
