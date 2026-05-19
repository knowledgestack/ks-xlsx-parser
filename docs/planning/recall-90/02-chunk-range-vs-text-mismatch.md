# 02 · `text_hit_geom_miss` — chunk contains answer text but range doesn't overlap GT

**Status:** 🆓 free to claim
**Slice:** E (`annotation/block_splitter.py`, `analysis/pattern_splitter.py`, `analysis/light_block_detector.py`)
**Independent of:** 00, 01, 03, 04

## What it looks like

Seven actionable instances on the 200-sample
(`61-4`, `353-29`, `382-29`, `80-42`, `334-11`, `495-31`, `462-45`)
where:

* `rank_of_text_match ≤ 5` — some chunk's `render_text` contains the
  answer value(s).
* `rank_of_first_overlap is None` — but no chunk's claimed A1 range
  overlaps the ground-truth range.
* `n_chunks_on_gt_sheet ≥ 1` — we DID emit chunks on the correct sheet.

So the parser knows what data exists on the sheet (text is there), but
the **range bookkeeping** on the chunk it emitted is wrong: the claimed
`top_left_cell` / `bottom_right_cell` excludes the answer cell even
though the chunk's text includes it.

This is the **citation-grade killer**. ks-backend's UI will highlight a
region that doesn't actually contain the answer the LLM cited — looks
buggy to the user.

Example — instance `61-4`:
```
ans = 'output'!A2:G15
chunks on 'output': 1 chunk
chunk text contains: the answer string (text_rank=1)
chunk's reported A1 range: NOT A2:G15
```

## Hypothesis (deliberately under-specified)

Block merge/split paths widen the *text* (by absorbing adjacent cells)
faster than they update `cell_range.bottom_right`, OR a pattern-split
narrows the `cell_range` to exclude the rows the splitter just lifted
into a child block while the rendered text still contains them. Find
the actual code path empirically.

## Diagnostic signature

```bash
jq -c 'select(.bucket_combined == "text_hit_geom_miss"
              and (.flags | contains(["instruction_requires_execution"]) | not)
              and .n_chunks_on_gt_sheet > 0)' \
    tests/benchmarks/reports/retrieval/*/enriched_failures.ndjson
```

For each match, the comparison that matters is `gt_range_bbox` vs each
chunk on the GT sheet's `top_left_cell`/`bottom_right_cell`. Print the
chunk's `render_text` and confirm it DOES contain the answer values
that openpyxl reads at `answer_position`.

## File scope

You may touch:

* `src/ks_xlsx_parser/models/block.py` — invariants on `BlockDTO.cell_range`.
* `src/ks_xlsx_parser/annotation/block_splitter.py`
* `src/ks_xlsx_parser/analysis/pattern_splitter.py`
* `src/ks_xlsx_parser/analysis/light_block_detector.py`
* `src/ks_xlsx_parser/chunking/chunker.py::_block_to_chunk` — at chunk
  creation time, you can defensively clip `cell_range` to the bounding
  box of cells that actually contributed text. That's a safety net even
  if the upstream bug isn't located.

Do NOT touch parsing (`parsers/*`) — the inputs are correct;
downstream block bookkeeping is wrong.

## Acceptance criteria

1. Add an invariant test (`tests/test_structural_invariants.py`
   pattern): for every chunk, the cells whose values appear in
   `render_text` MUST lie within
   `[chunk.top_left_cell, chunk.bottom_right_cell]`. Test should fail
   on `main` for at least 3 of the 7 cluster instances.
2. After the fix, all 7 instances flip from `text_hit_geom_miss` to
   `both_hit` (`rank_of_first_overlap ≤ 5`).
3. `geometric@5` on the 200-sample seed=1337 rises by ≥ 3 pp.
4. `recall_text@5` does NOT drop (tightening ranges shouldn't drop
   text — if it does, the chunker is also dropping text in step with
   the range clip, which is a separate bug).
5. `table_fragmentation_rate` does NOT rise — if you fixed range drift
   by splitting one chunk into many, fragmentation will spike. That
   would solve cluster 02 by creating cluster 04 — net zero. Either
   keep fragmentation flat or coordinate with whoever owns 04.

## Failing test sketch

```python
# tests/test_chunk_range_invariants.py
import re
from openpyxl.utils import column_index_from_string
from ks_xlsx_parser.pipeline import parse_workbook
from ks_xlsx_parser.models.common import col_letter_to_number

A1_RE = re.compile(r"^([A-Z]+)(\d+)$")

def _parse_a1(a1):
    m = A1_RE.match(a1)
    return (int(m.group(2)), col_letter_to_number(m.group(1)))

def test_chunk_range_covers_all_rendered_cells(tmp_path):
    # Use one of the cluster instances or a hand-built fixture:
    # the assertion is "for every chunk, the cell coordinates that appear
    # in render_text fall inside the claimed range box."
    chunks = parse_workbook(
        path="data/corpora/spreadsheetbench/all_data_912_v0.1/spreadsheet/61-4/1_61-4_input.xlsx"
    ).chunks
    for c in chunks:
        if not c.top_left_cell or not c.bottom_right_cell:
            continue
        r0, col0 = _parse_a1(c.top_left_cell)
        r1, col1 = _parse_a1(c.bottom_right_cell)
        # Pull A1 references that appear in the chunk header (e.g. "[Sheet1!A2:G15]")
        # and assert each is inside the claimed range. Adapt to whatever
        # renderer markers your current text actually uses.
        for ref in re.findall(r"\b([A-Z]+)(\d+)\b", c.render_text or ""):
            col = col_letter_to_number(ref[0])
            row = int(ref[1])
            assert r0 <= row <= r1, f"row {row} outside {c.top_left_cell}:{c.bottom_right_cell}"
            assert col0 <= col <= col1, f"col {ref[0]} outside {c.top_left_cell}:{c.bottom_right_cell}"
```

## Pitfalls

* The renderer prints A1 refs in its block header (e.g.
  `[Sheet1!A2:G15] (table)`) — those are the chunk's CLAIMED range, not
  a cell reference. The invariant test should look at rendered cell
  contents (table rows), not at the block header.
* Some text rendered into a chunk doesn't have an A1 reference (e.g.
  pure values inside a table grid). The invariant has to use the
  block's underlying `cells` collection, not regex on text.
* If the merge widens the chunk's range past adjacent empty rows, you
  may LOSE geometric overlap on those by clipping. That's fine here —
  the cluster is text_hit/geom_miss, so clipping to actual content is
  what's needed.

## Repro

```bash
python scripts/triage_recall.py tests/benchmarks/reports/retrieval \
    --bucket present_but_ranked_low --examples 0   # to ensure no overlap
# Then manually read enriched_failures.ndjson and pick instance 61-4
# or 353-29 — both are pure text_hit_geom_miss cases.
```
