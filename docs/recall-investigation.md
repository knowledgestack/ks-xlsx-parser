# Retrieval-recall investigation — getting ks-xlsx-parser to >0.90

## Where we are (v0.2.0 on SpreadsheetBench, 912 instances)

| Metric                 | ks-xlsx-parser | docling 2.93 |
|------------------------|----------------|--------------|
| Parse success          | 99.945%        | not run at scale |
| Recall@1 (text-match)  | 0.580          | 0.579 |
| Recall@3 (text-match)  | 0.697          | 0.670 |
| Recall@5 (text-match)  | **0.704**      | 0.686 |
| Recall@5 (geometric)   | 0.369          | 0.000 (no A1 anchors) |
| Mean parse time        | 251 ms         | 265 ms |

Recall@5 = 0.704 means **~30% of questions miss** with k=5. To reach 0.90
we need to roughly cut the miss rate to a third. A single recall number
hides which lever to pull, which is why this branch ships failure
bucketing (`scripts/eval_retrieval.py --emit-failures` +
`scripts/triage_recall.py`).

## The diagnosis framework — why the bucket histogram is the answer

Every recall@5 miss falls into one of these buckets. The fix is
completely different per bucket, and only one or two will dominate. The
job of the investigator is to read the histogram FIRST, then commit to
fixing the biggest one.

| Bucket | What it means | Where to look |
|---|---|---|
| `answer_absent_from_chunks` | Answer value is in NO chunk. Cell was dropped or garbled. | `parsers/cell_parser.py`, `rendering/text_renderer.py::_cell_render_value` |
| `present_but_ranked_low`    | A chunk DOES contain the answer but ranked >5. Chunk is too large/heterogeneous; the embedding is diluted. | `chunking/chunker.py` (no token cap), `analysis/table_assembler.py` (over-merging) |
| `wrong_sheet`               | Answer sheet was never chunked. Sheet enumeration missed it. | `parsers/workbook_parser.py` sheet loop |
| `geometric_no_overlap`      | No chunk's A1 range overlaps ground truth. Range bookkeeping drifts during merge/split. | `annotation/block_splitter.py`, `analysis/pattern_splitter.py` |
| `no_chunks` / `parse_error` | Upstream parser failure. | The parse exception — fix the crash. |

## First-run findings (200-sample seed=1337, May 2026)

Headline numbers refuted my **H1** ("present_but_ranked_low dominates")
and surfaced something the original analysis missed entirely. Of 157
failures in the seed=1337 sample:

* **127 (81%) are `instruction_requires_execution`** — the benchmark
  is asking the system to *compute and write* the answer. The
  `answer_position` cell range in `input.xlsx` is empty by design. A
  parser cannot retrieve what isn't there. These instances need to be
  filtered from the headline metric (TODO 05).
* **30 are truly actionable** parser failures, clustering into a small
  number of named buckets (benchmark-spec parser bugs in
  `scripts/eval_retrieval.py`; array-formula rendering; chunk range vs
  text mismatch; cell-drop / uncached-formula rendering; single-chunk
  dilution).
* Zero of the 30 are `present_but_ranked_low`. Chunk-size dilution
  may matter on the full 912-corpus but is NOT the dominant problem
  on the 200-sample.

The dominant **citation-grade killer** turned out to be
`text_hit_geom_miss` (84 of 200, mostly inside the 127 execution
instances; 9 are actionable parser bugs). The chunk text contains
the answer string, but the chunk's claimed A1 range doesn't overlap
ground truth — exactly **H3** ("range-bookkeeping drift").

When the in-scope filter is applied, real `recall_text@5` is
**0.59**, not 0.635. Closing the 30 actionable failures gets us
near 0.90 on the in-scope metric.

## A priori hypotheses (now confirmed/refuted; left as history)

### H1 — `present_but_ranked_low` is the biggest bucket

There is no per-chunk token cap in `chunking/chunker.py` (`CHARS_PER_TOKEN`
is only used to *report* `token_count`, never to split). On
SpreadsheetBench many input files are single-sheet ledgers where the
block-assembler collapses the whole sheet into one chunk. The
sentence-transformer query embedding then has to compete against ~2k
tokens of mostly irrelevant text; the relevant ~5 tokens get washed out.

If H1 is right, the histogram will show `present_but_ranked_low` ≫ the
others, and recall@1 (0.580) will be much worse than recall@5 (0.704)
— exactly what we observe (Δ = 12.4 pp, vs typical Δ ≈ 5–6 pp when
chunks are right-sized).

**Fix**: hard cap chunks at ~512 tokens by row-splitting tables and add
a "row group" sub-chunk for tall tables. This is a 1–2 day surgical
change in `chunking/chunker.py`.

### H2 — `answer_absent_from_chunks` dominates the geometric gap

`parsers/workbook_parser.py` loads both `data_only=False` (formula
expressions) and `data_only=True` (computed values). But what flows into
`render_text` is whichever `_cell_render_value` picks. If the cell is a
formula like `=SUM(B2:B10)`, `display_value` may be the *expression*
when the workbook was saved without cached values (LibreOffice and some
generated files do this). Those answer cells become unfindable by text
match even though the data IS in the spreadsheet.

**Diagnostic**: count failure rows where every `top_chunks[*].text`
matches the formula expression pattern (`=`, function name) but not the
expected numeric value. The bucket emits the top-8 chunks for inspection.

**Fix**: when the cached value is missing for a formula cell, evaluate it
with our own formula engine (`formula/formula_parser.py` already exists)
or use python-calamine's value-only pass as the source of truth for
render text — never the formula source.

### H3 — `geometric_no_overlap` is high because block ranges over-extend

Geometric recall@5 = 0.369 means **only ~37% of the time** does the
chunk a parser surfaces actually cover the ground-truth answer cell —
even when the text match works. The block-detection pipeline merges
sparse blocks (`analysis/light_block_detector.py`) and groups by
similarity (`analysis/table_grouper.py`). Each merge widens the
top-left/bottom-right anchors. If the anchors are widened past the
sheet's true content, downstream citation overlays in ks-backend will
highlight whitespace, and the geometric metric registers the chunk as
"not overlapping" because its claimed range is so large it's not useful.

**Fix**: after every merge/split, clip `cell_range` to the tight bounding
box of the cells that actually contributed text. Add an invariant test
that `block.cell_range` ⊆ `bounding_box(block.cells)`.

## How to confirm — the next benchmark run

1. `make corpus-download` (one-time, ~hundreds of MB).
2. `make bench-track` — runs the full benchmark, appends to
   `tests/benchmarks/reports/history.jsonl`, prints the bucket triage.
3. Read the histogram. Pick the biggest bucket. Open 3–5 example
   failures with `python scripts/triage_recall.py <reports-dir>
   --bucket <name> --examples 5`. Each row shows:
   * the natural-language question
   * the ground-truth answer cell + values
   * the top-8 ranked chunks we produced (sheet, A1 range, text snippet)
   * whether each chunk contains the answer
4. Pattern-match across 5 examples — what's the common parser behaviour?
   That tells you the fix.
5. Implement, re-run `make bench-track`. The script prints the delta
   row-over-row so improvement is visible immediately.

## How to use the Docker image (CI + reproducibility)

```bash
# Build once
docker build -f Dockerfile.bench -t ks-xlsx-parser-bench .

# Quick smoke (60 instances, < 2 min)
docker run --rm -e BENCH_SAMPLE=60 ks-xlsx-parser-bench

# Full corpus, persist reports + corpus cache
docker run --rm \
  -v "$PWD/tests/benchmarks/reports:/app/tests/benchmarks/reports" \
  -v "$PWD/data:/app/data" \
  ks-xlsx-parser-bench
```

The `Benchmark` GitHub workflow:
* Runs a 60-instance smoke on every PR that touches `src/` or the
  benchmark scripts.
* Runs the full 912-instance corpus weekly (Monday 06:00 UTC) and on
  manual dispatch.
* Uploads `tests/benchmarks/reports/*` as a build artifact and posts the
  recall summary to the job step summary.

## Goal & cadence

Target: **text recall@5 ≥ 0.90** by end of the current quarter.

Track in `tests/benchmarks/reports/history.jsonl` (commit-over-commit
row append). Refuse merges that drop recall@5 by ≥ 2 pp on the sample
run (planned gate; today the PR job is reporting-only).

## See also

* `scripts/eval_retrieval.py` — the benchmark itself.
* `scripts/triage_recall.py` — bucket histogram + example dump.
* `scripts/append_bench_history.py` — history.jsonl row writer.
* `.claude/skills/recall-failure-triage/SKILL.md` — agent guide.
* `Dockerfile.bench` — reproducible benchmark image.
