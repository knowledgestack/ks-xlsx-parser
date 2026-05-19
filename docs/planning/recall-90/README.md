# Roadmap: ks-xlsx-parser retrieval recall → 0.90

This directory is the worklist for getting SpreadsheetBench recall on the
canonical ks-xlsx-parser benchmark from the v0.2.0 baseline up to:

| Metric                 | v0.2.0 (912) | 200-sample baseline | Target | Stretch |
|------------------------|--------------|---------------------|--------|---------|
| Text recall@5          | 0.704        | 0.635               | ≥ 0.90 | 0.95    |
| Geometric recall@5     | 0.369        | 0.310               | ≥ 0.70 | 0.85    |
| In-scope text recall@5 | unmeasured   | **0.59**            | ≥ 0.90 | —       |

Each `NN-*.md` in this folder is one **independent, parallel-safe TODO**.
Pick one, claim it (table below), execute, hand off the next agent.

## Read this first — the headline number is misleading

The benchmark dataset has 912 instances. **~63% of them require the
system to *compute and write* the answer** (e.g. "fill column C with
ROUND(B*5)") — the answer literally doesn't exist in the input
spreadsheet. A parser cannot retrieve what isn't there. We call this
class `instruction_requires_execution` and treat them as **out of scope
for parser work**. See [05-out-of-scope-execution-instances.md](./05-out-of-scope-execution-instances.md).

When the scoring is filtered to in-scope instances only, the real
ks-xlsx-parser recall is closer to **0.59**, and improving it past
0.90 means closing ~22 named failures out of ~30 actionable cases on the
200-instance seed=1337 sample. That's the actual work.

## Worklist

Status legend: 🆓 free to claim · 🔵 in progress · ✅ landed · ⛔ blocked

| # | File | Cluster | Instances | Primary slice | Status |
|---|---|---|---|---|---|
| 00 | [00-benchmark-spec-parser-bugs.md](./00-benchmark-spec-parser-bugs.md) | Benchmark harness fails to parse malformed `data_position`/`answer_position` (`Dashboard'!B8`, fullwidth `G12：J15`) — 8 instances scored as wrong even when the parser was correct. | 8 | `scripts/eval_retrieval.py` | 🆓 |
| 01 | [01-array-formula-rendering.md](./01-array-formula-rendering.md) | Array-formula cells surface as `<openpyxl.worksheet.formula.ArrayFormula object>` — value never lands in `render_text`. | 2 | `parsers/cell_parser.py`, `rendering/text_renderer.py` | 🆓 |
| 02 | [02-chunk-range-vs-text-mismatch.md](./02-chunk-range-vs-text-mismatch.md) | `text_hit_geom_miss` with chunks on the GT sheet — answer text is in *some* chunk's text but no chunk's claimed A1 range overlaps GT. Range bookkeeping drift during block merge/split. | 7 | `annotation/block_splitter.py`, `analysis/pattern_splitter.py`, `chunking/chunker.py` | 🆓 |
| 03 | [03-cell-drop-or-uncached-formula.md](./03-cell-drop-or-uncached-formula.md) | Chunk's range covers GT geometrically but its rendered text lacks the answer values — either cells dropped in render, or formula cells with no cached value rendered as the formula source. | up to 10 (some unscorable) | `parsers/cell_parser.py`, `parsers/sheet_parser.py`, `rendering/text_renderer.py` | 🆓 |
| 04 | [04-single-chunk-multi-region-sheet.md](./04-single-chunk-multi-region-sheet.md) | Sheets containing distinct logical regions emerge as one giant chunk; the answer text dilutes against the rest of the sheet at embedding time. | 2 (+ likely more at full corpus) | `chunking/chunker.py`, `analysis/light_block_detector.py`, `analysis/table_grouper.py` | 🆓 |
| 05 | [05-out-of-scope-execution-instances.md](./05-out-of-scope-execution-instances.md) | Informational + scoring filter. **Not a parser fix.** Surfaces benchmark instances where the parser fundamentally can't help; suggests filtering them from headline recall. | 127 / 200 | `scripts/eval_retrieval.py` (scoring), `scripts/enrich_failures.py` | 🆓 |

## How parallelism works here

Each TODO is scoped to a different file slice so agents won't collide:

```
slice A   benchmark harness         00, 05
slice B   parsers/cell_parser       01, 03
slice C   parsers/sheet_parser      03
slice D   rendering/text_renderer   01, 03
slice E   annotation/ + analysis/   02
slice F   chunking/                 04
```

TODOs 00 and 05 touch the same file (`eval_retrieval.py`) — **serialize**
those two. Everything else can run truly in parallel.

If you discover a cluster mid-flight that doesn't fit anywhere here,
add a new `NN-*.md` and append a row to the table above. Do NOT bury
findings inside an existing TODO.

## Working a TODO

Each cluster file has the same structure:

1. **What the cluster looks like** — a real failure example in full.
2. **Repro instance IDs** — the 200-sample seed=1337 instance IDs that
   match this cluster.
3. **Diagnostic signature** — exact `enriched_failures.ndjson` columns
   that identify a member.
4. **File scope** — what you can change; what you can't.
5. **Acceptance criteria** — measurable on the 200-sample seed=1337
   rerun. "The N instances listed in §2 all flip from miss → hit on
   `text_match_rank` AND no other previously-passing instance regresses."
6. **Failing test sketch** — one-line idea for a regression test.

What's **deliberately missing** from each file: the proposed fix and an
effort estimate. The agent that picks up the task does that diagnosis;
the doc just specifies success.

### The loop, end-to-end

```bash
# 1. Claim the task (edit README, change 🆓 → 🔵 with your name).
# 2. Read the cluster file + the named example instances.
# 3. Reproduce locally:
python scripts/eval_retrieval.py \
    --corpus data/corpora/spreadsheetbench/all_data_912_v0.1 \
    --parsers ks --sample 200 --seed 1337 --emit-failures
python scripts/enrich_failures.py tests/benchmarks/reports/retrieval
python scripts/triage_recall.py tests/benchmarks/reports/retrieval --examples 5

# 4. Write a regression test that fails today (see test sketch in cluster file).
# 5. Fix until the test passes.
# 6. Re-run the eval + enrichment. Confirm the cluster count dropped, no regressions.
# 7. Append to history; the delta should be visible:
python scripts/append_bench_history.py

# 8. PR the change. Title: `recall(NN): <cluster short name> — N→M`.
#    PR body links to the cluster file and pastes the before/after histogram.
# 9. Update README table: 🔵 → ✅ with PR link.
```

## Why the 200-sample seed=1337 is the gate, not the full 912

912 takes ~40 min on a laptop. PR validation needs to be fast. The
seed=1337 200-sample is deterministic and contains roughly proportional
representation of every cluster. Acceptance criteria are stated against
it. **Cross-check with the full corpus before declaring a release**:

```bash
python scripts/eval_retrieval.py --parsers ks --emit-failures
```

The weekly benchmark workflow (`.github/workflows/benchmark.yml`) runs
the full 912 on schedule + workflow_dispatch.

## Tracking progress

`tests/benchmarks/reports/history.jsonl` is appended to on each
`make bench-track` run, one row per commit. Tail it to see the trend:

```bash
tail -10 tests/benchmarks/reports/history.jsonl
```

Each row carries the failure_buckets histogram, so per-cluster progress
is in the data, not just the headline number.

## See also

- `docs/recall-investigation.md` — diagnosis framework + hypotheses
  (H1 chunk-size, H2 formula-rendering, H3 range-bookkeeping)
- `docs/benchmark-local-setup.md` — install → corpus → eval loop
- `scripts/triage_recall.py` — bucket histogram + exemplar dump
- `scripts/enrich_failures.py` — per-failure diagnostic columns
- `.claude/skills/recall-failure-triage/SKILL.md` — agent guide
