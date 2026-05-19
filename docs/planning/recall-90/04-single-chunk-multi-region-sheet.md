# 04 · Whole-sheet single chunk dilutes the embedding

**Status:** 🆓 free to claim
**Slice:** F (`chunking/chunker.py`, `chunking/segmenter.py`, `analysis/light_block_detector.py`, `analysis/table_grouper.py`)
**Independent of:** 00, 01, 02, 03

## What it looks like

Two clear instances on the 200-sample (`184-6`, `374-9`); the pattern
extends to many more on the full corpus where a sheet contains a
single large data table with a few headers and the chunker emits ONE
chunk covering A1:end. Note: TODO 02 also has instances where the
single-chunk-per-sheet pattern shows up (`53-12`, `189-9`,
`334-11`, `353-29`, `382-29`, `462-45`, `495-31`, `CF_3712`). Treat
those as **second-order beneficiaries**: when 02 lands first, the
chunk's range tightens and may already split the over-broad chunk.
When 04 lands first, the chunks get smaller and 02's range invariant
becomes easier to satisfy. Order is reviewer's choice.

Symptoms:

* `n_chunks_on_gt_sheet == 1` and the chunk spans most of the sheet.
* The answer values are present in the chunk's render_text BUT rank
  >5 (or rank=1 by luck — at recall@1 the model picks based on the
  bag-of-everything embedding). The recall@1 vs recall@5 gap of 12.4 pp
  in the v0.2.0 numbers is consistent with this.
* `summary.md` shows `table_fragmentation_rate ≈ 0` on these sheets
  (which sounds good but means we're under-splitting, not perfectly
  segmenting).

This is the **chunk-granularity** problem. The block-detection pipeline
fuses everything into one block; the chunker has no token cap and
emits the block verbatim.

## Diagnostic signature

```bash
jq -c 'select(.n_chunks_on_gt_sheet == 1 and .n_chunks_total <= 2
              and (.flags | contains(["instruction_requires_execution"]) | not))' \
    tests/benchmarks/reports/retrieval/*/enriched_failures.ndjson
```

Also: dump the per-chunk `token_count` (already computed but unused in
chunker.py). Any chunk over 800 tokens is over-broad on this corpus.

## File scope

You may touch:

* `src/ks_xlsx_parser/chunking/chunker.py` — introduce a chunk-size
  cap. Suggested target: ~512 tokens. When a block exceeds, split the
  block into row-groups (preserving headers in each group), with each
  group becoming its own chunk. `prev_chunk_id` / `next_chunk_id` is
  already wired for navigation.
* `src/ks_xlsx_parser/chunking/segmenter.py` — segmenter logic if any
  granularity decisions live there.
* `src/ks_xlsx_parser/analysis/light_block_detector.py` and
  `analysis/table_grouper.py` — investigate whether *over*-merging
  upstream is causing the one-block-per-sheet symptom. If the table
  grouper is collapsing 3 logical tables into 1, the upstream fix is
  better than splitting at chunk-emit time.

Do NOT touch the ranking/embedding step (`scripts/eval_retrieval.py`)
or any retrieval logic.

## Acceptance criteria

1. After landing, no chunk on the 200-sample seed=1337 run exceeds
   ~800 tokens (= ~3200 characters) by default. Allow override.
2. `table_fragmentation_rate` may rise (this is desired up to ~0.5 —
   purposeful row-group splits). It must NOT cause a regression in
   geometric@5 — each row-group chunk's claimed range must cover only
   its rows.
3. `recall_text@5` rises by ≥ 3 pp on the 200-sample.
4. `recall_text@1` rises by ≥ 4 pp (smaller chunks = sharper embeddings
   = top-1 sharpens).
5. The two named instances (`184-6`, `374-9`) flip to `both_hit`. Note
   that `184-6` is also subject to TODO 00's harness fix — order may
   matter.

## Failing test sketch

```python
# tests/test_chunk_size_cap.py
from ks_xlsx_parser.pipeline import parse_workbook

def test_no_chunk_exceeds_token_budget(tmp_path):
    # Generate a workbook with a 1000-row table. Without a cap the
    # chunker emits one giant chunk; with the cap it should emit several.
    import openpyxl
    p = tmp_path / "big.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["id", "name", "value"])
    for i in range(1, 1001):
        ws.append([i, f"name-{i}", i * 7])
    wb.save(p)

    chunks = parse_workbook(path=str(p)).chunks
    assert len(chunks) >= 2, "1000-row table should produce ≥2 chunks after the cap lands"
    # Each chunk's claimed range should be a contiguous row block on the same sheet.
    for c in chunks:
        assert c.sheet_name and c.top_left_cell and c.bottom_right_cell
    # No chunk should be larger than the budget (in render-text characters).
    for c in chunks:
        assert len(c.render_text or "") <= 4000, f"chunk too big: {len(c.render_text)} chars"
```

## Pitfalls

* **Don't strand headers.** The current chunker preserves table headers
  in render_text. When splitting a table into row groups, every group
  needs the header rows OR a header-summary block so an LLM can
  interpret column meanings.
* **Range bookkeeping** — when you split, the child chunks' A1 ranges
  MUST be tight to their row subset. Otherwise you create
  geometric_no_overlap failures (cluster 02 in reverse).
* **`token_count` is character-based** (`len(text) // 4`). Use the
  same metric for the cap to avoid a circular dependency on a tokenizer.
* **Don't split horizontally** (column subsets). Excel tables are
  usually wider than embeddings need but column-split chunks have no
  natural boundary and reading order breaks.

## Coordinating with TODO 02

If both 02 and 04 are in flight:
- 02 is *range-tightening* (semantically correct ranges).
- 04 is *granularity* (more chunks per sheet).
- The risk is they fight over `_block_to_chunk`. Decide which one
  emits the final A1 range and document the contract. Suggested:
  04 splits blocks into sub-blocks WITH tight ranges, then 02's
  invariant test passes for the sub-blocks.

## Repro

Use a synthetic 1000-row workbook (see test sketch) — much faster to
iterate on than the corpus instances. Once your local fixture passes,
run the 200-sample to confirm corpus-wide impact.
