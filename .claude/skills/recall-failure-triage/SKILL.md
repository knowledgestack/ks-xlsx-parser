---
name: recall-failure-triage
description: Diagnose why ks-xlsx-parser's retrieval recall is below target by reading the failure-bucket histogram and exemplar failures, then propose ranked fixes. Use when investigating low recall@k on SpreadsheetBench, deciding what to fix next, or reviewing a benchmark run. Sister skill to excel-extraction-pipeline-improver.
---

# Recall Failure Triage — turn `recall@5 = X` into a ranked worklist

A single recall number tells you nothing about *what* to fix. Five
distinct failure modes produce identical recall drops, and the fix is
completely different per mode. This skill is the bridge from "recall is
low" to "open `parsers/cell_parser.py` line N, here is the bug".

Pair with `excel-extraction-pipeline-improver` — that skill implements
fixes; this one decides what to fix.

## When this skill fires

* User says "recall is low / dropped / not improving"
* New benchmark run produced a `failures.ndjson` or `summary.json`
* PR-time benchmark sample regressed on the CI step summary
* You're picking the next parser/chunking PR and need to choose what's
  highest leverage

## Inputs you should expect

| Input | Where it comes from |
|---|---|
| `summary.json` + `summary.md` | latest `tests/benchmarks/reports/retrieval/<stamp>/` |
| `failures.ndjson` | same dir, **only present if** the run used `--emit-failures` |
| `history.jsonl` | `tests/benchmarks/reports/history.jsonl` — recall over commits |
| current branch + recent commits | to know what changed since the last run |

## The decision procedure

### 0. Apply the in-scope filter — most "misses" are unfixable

Before counting buckets, run `scripts/enrich_failures.py` to mark
instances flagged `instruction_requires_execution` — those where the
benchmark's `answer_position` is empty in `input.xlsx` because the
question literally asks the system to *write* the answer. ~63% of
the SpreadsheetBench corpus is this class on the 200-sample. They are
NOT parser bugs and CANNOT be fixed by parser work. Exclude them.

```bash
python scripts/enrich_failures.py tests/benchmarks/reports/retrieval
jq -c 'select((.flags | contains(["instruction_requires_execution"])) | not)' \
    tests/benchmarks/reports/retrieval/*/enriched_failures.ndjson
```

### 1. Read the bucket histogram FIRST

```bash
python scripts/triage_recall.py tests/benchmarks/reports/retrieval
```

The output ranks buckets by count. **Spend your time on the largest
bucket only.** The next-largest is the next PR.

| Bucket | Root cause | Where to fix |
|---|---|---|
| `answer_absent_from_chunks` | The answer value never made it into ANY chunk. The cell was dropped or rendered as a formula expression instead of its computed value. | `parsers/cell_parser.py::_extract_cell_value` and `rendering/text_renderer.py::_cell_render_value`. Check `data_only` plumbing in `parsers/workbook_parser.py`. |
| `present_but_ranked_low` | A chunk DOES contain the answer but ranked >5. The chunk is too big/heterogeneous — embedding is diluted. | `chunking/chunker.py`. There is no token cap today; large blocks are emitted whole. Cap at ~512 tokens and row-split tall tables. |
| `wrong_sheet` | Answer sheet was never chunked. | `parsers/workbook_parser.py` sheet loop. Check hidden sheets, very-hidden sheets, and `wb.sheetnames` vs `wb.worksheets`. |
| `geometric_no_overlap` | Answer text matches but the chunk's A1 range doesn't overlap ground truth. Range drifts during merge/split. | `annotation/block_splitter.py` + `analysis/pattern_splitter.py`. Add invariant: `block.cell_range` ⊆ bbox of cells that actually rendered text. |
| `no_chunks` / `parse_error` | Upstream parser failure. | The exception. Reproduce with `parse_workbook(path)`. |

### 2. Read 5 example failures from the top bucket

```bash
python scripts/triage_recall.py tests/benchmarks/reports/retrieval \
  --bucket <largest> --examples 5
```

For each example, the script prints:
* the natural-language question
* `answer_position` (the ground-truth cell)
* the ground-truth string values that should appear in a chunk
* the top-8 ranked chunks with sheet, range, and whether each contains
  the answer

**Pattern-match.** What do the 5 examples have in common? E.g.
* All answer cells are formulas → H2 confirmed: render the cached value
* All chunks span the entire sheet → H1 confirmed: cap chunk size
* All answer sheets are hidden → wrong_sheet from hidden-sheet skip

### 3. State the hypothesis before changing code

In your reply, write the hypothesis as one sentence:
> "Bucket X dominates because [observable common pattern in examples].
> Expected fix: [specific file + change]. Expected lift on recall@5:
> [number]."

The expected lift = (bucket count) ÷ (scoreable instances). If a fix
is supposed to halve the biggest bucket but recall barely moves on the
next run, the hypothesis was wrong — go back to step 1, don't pile on.

### 4. Wire the fix into the existing pipeline

Pass the hypothesis to the `excel-extraction-pipeline-improver` skill
(or implement directly) with TDD: write a failing crossval/invariant
test that captures the bug from one example, fix, run `make test`, then
re-run `make bench-track` to confirm the bucket shrank.

### 5. Verify on the history

```bash
tail -5 tests/benchmarks/reports/history.jsonl
```

The script prints a row-over-row delta. Expected behaviour:
* The targeted bucket count drops
* `recall_text@5` rises by approximately `target_bucket_drop /
  scoreable_instances`
* `recall_text@1` improves proportionally OR more (better chunks rank
  higher too)
* Other buckets shouldn't grow — if they do, the fix had a side effect

## Anti-patterns

* **Don't chase the headline number.** "Recall went up 1 pp" without a
  bucket explanation is suspicious — it might be benchmark noise.
* **Don't fix two buckets at once.** You won't know which fix worked.
* **Don't tune chunk size as the first move.** It only helps
  `present_but_ranked_low`. If `answer_absent_from_chunks` is bigger,
  chunk tuning is wasted effort.
* **Don't trust geometric and text recall independently.** A docling-style
  parser with zero geometric overlap can still have high text recall
  (because it serializes everything to plain text). That's not the same
  thing as good — citation overlays need the geometric metric. Always
  report both.

## Quick reference — running it from scratch

```bash
make corpus-download                                    # one-time, large
make bench-track                                        # runs + triages
python scripts/triage_recall.py tests/benchmarks/reports/retrieval --examples 5
cat tests/benchmarks/reports/history.jsonl | tail -10   # commit-over-commit
```

CI runs a 60-instance sample on every PR (`.github/workflows/benchmark.yml`)
and posts the bucket histogram to the GitHub job summary. Weekly schedule
runs the full 912-instance corpus.

## Hand-off contract

When you finish a triage round, your final message MUST include:
1. The bucket histogram (counts + percentages).
2. The dominant bucket + one-sentence root cause hypothesis.
3. The 1–2 file paths where the fix lives.
4. The expected lift on `recall_text@5` if the fix lands.
5. A pointer to the example instance IDs you read.

Anything less and the next agent will have to re-do the analysis.
