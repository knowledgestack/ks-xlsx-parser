# 03 · Chunk's range covers GT but rendered text lacks the answer values

**Status:** 🆓 free to claim
**Slice:** B + C + D (`parsers/cell_parser.py`, `parsers/sheet_parser.py`, `rendering/text_renderer.py`)
**Independent of:** 00, 02, 04 — partial overlap with 01 on `text_renderer.py`; coordinate via README.

## What it looks like

Up to ten actionable instances on the 200-sample (`CF_3712`, `57033`,
`53-12`, `55794`, `45063`, `36842`, `384-4`, `262-17`, `189-9`, `48799`)
where:

* `rank_of_first_overlap ≤ 5` — chunk's claimed range overlaps GT.
* `rank_of_text_match is None` — but the chunk's rendered text does NOT
  contain the answer values.

The parser knows the block exists (range is right) but it dropped or
mis-rendered the cells whose values should appear in the chunk text.

**Two root causes nested inside this cluster — confirm by inspection:**

### Sub-cluster 3a: cell-drop within block

The block's `cells` collection is missing the answer cell, or the
renderer's grid-construction loop skips it (e.g. when the cell sits in
a row otherwise mostly empty, and a sparsity filter culls it).

Instance `CF_3712`: chunk on `Purchases` covers M3:M5 geometrically;
the answer header value (`gt_cell_raw = "Product "`) doesn't render.

### Sub-cluster 3b: uncached formula renders as `=A1+B1`, not the value

Cell stores a formula; the workbook was saved without computing it
(common for LibreOffice / programmatic generation). `data_only` load
returns `None` for that cell. The renderer falls back to the formula
source. Retrieval for the *value* (e.g. "1272") fails because the chunk
text contains "=A1+B1" instead.

Instance `48799` (`Sheet2!Z2` — a 24-deep nested IF/VLOOKUP). Same
shape: formula text in chunk, value missing.

> ⚠️ **Some of the named instances may be unscorable** — e.g. `189-9`
> and `53-12` have answer cells that are all `None` in `answer.xlsx`,
> meaning the benchmark expects the system to write nothing. In that
> case `answer_cell_values=[]` and `rank_text=None` is automatic,
> regardless of what the chunk renders. The first thing to do in this
> TODO is to filter the ten instances through openpyxl to find the
> truly unfixable ones; report that as part of the PR.

## Diagnostic signature

```bash
jq -c 'select(.bucket_combined == "text_miss_geom_hit"
              and (.flags | contains(["instruction_requires_execution"]) | not)
              and .n_chunks_on_gt_sheet >= 1)' \
    tests/benchmarks/reports/retrieval/*/enriched_failures.ndjson
```

Then run for each candidate:

```python
from openpyxl import load_workbook
wb = load_workbook(input_path, data_only=True)
ws = wb[answer_sheet]
# read answer.xlsx cells at answer_position — if all None, drop this instance.
```

Instances whose `answer.xlsx` GT cells are all `None` move to TODO 05
(out-of-scope; benchmark-scoring issue).

## File scope

You may touch:

* `src/ks_xlsx_parser/parsers/cell_parser.py` — formula vs cached-value
  selection. Add a code path: if cell has a formula AND no cached
  value in the `data_only` pass, mark as `formula_uncached` and emit
  both the formula text AND any reachable proxy (e.g. blank cell with
  formula reference, instead of literal formula in the grid).
* `src/ks_xlsx_parser/parsers/sheet_parser.py` — make sure cells with
  string values that look like formulas (`startswith('=')`) get
  classified as formula cells.
* `src/ks_xlsx_parser/rendering/text_renderer.py::_cell_render_value`
  — for formula cells, prefer cached value; if cached is None, render
  the formula source ONLY (not `None`, not the formula display).

Do NOT add a formula evaluator. That's a separate, larger effort.

## Acceptance criteria

1. After culling unscorable instances (those with all-None
   `answer.xlsx` cells), at least **5 of the remaining instances** flip
   from `text_miss_geom_hit` to `both_hit` (rank_of_text_match ≤ 5).
2. `recall_text@5` on the 200-sample seed=1337 rises by ≥ 2 pp.
3. The unscorable instances are documented in the PR description with
   a one-line "this instance has no answer values in answer.xlsx,
   moving to TODO 05".
4. A new test fixture in `tests/conftest.py` (a workbook with a
   `=SUM(A1:A3)` cell saved WITHOUT cached values) plus a regression
   test that asserts the chunk render_text contains `=SUM(A1:A3)`
   verbatim (not the openpyxl repr of the formula object).

## Failing test sketch

```python
# tests/test_formula_rendering_unfilled.py
import openpyxl
from ks_xlsx_parser.pipeline import parse_workbook

def test_uncached_formula_renders_formula_source(tmp_path):
    p = tmp_path / "uncached.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = 10
    ws["A2"] = 20
    ws["A3"] = 30
    ws["B1"] = "=SUM(A1:A3)"
    wb.save(p)  # NO data_only refresh — B1's cached value remains None.

    chunks = parse_workbook(path=str(p)).chunks
    assert chunks, "no chunks"
    text = chunks[0].render_text
    # Whichever is acceptable: the formula source verbatim, OR the
    # computed value via the formula engine. NOT a None/empty cell.
    assert ("=SUM(A1:A3)" in text) or ("60" in text), (
        "uncached formula cell dropped from render_text:\n" + text
    )
```

## Pitfalls

* `_cell_render_value` already has a branch for booleans, dates,
  integer-valued floats. Don't break that path while adding the formula
  fallback.
* If you add a formula-source string to render_text, make sure the
  retrieval text-match logic in `eval_retrieval.py::_normalize_value_for_match`
  still works on numeric values that DO have cached results — adding
  formula sources should AUGMENT, not replace, the cached-value path.
* Some "cell drop" cases are actually merged-cell handling: a value is
  on the top-left of a merged region but the renderer surfaces only an
  empty placeholder for the merged cells. Distinct from this cluster —
  if you find one, file it as a new TODO.
* Don't try to evaluate formulas — Excel's calc engine is a swamp and
  `formula/formula_parser.py` only parses, doesn't compute. The
  acceptable fallback is "render the formula source so an LLM can read
  it and reason."

## Repro

```bash
# Inspect one candidate end-to-end.
python << 'EOF'
from openpyxl import load_workbook
from ks_xlsx_parser.pipeline import parse_workbook
p = "data/corpora/spreadsheetbench/all_data_912_v0.1/spreadsheet/CF_3712/1_CF_3712_input.xlsx"
wb = load_workbook(p, data_only=True)
print("M3:M5 on", wb.sheetnames[0])
for r in range(3,6):
    print(" ", wb.active.cell(row=r,column=13).value)
r = parse_workbook(path=p)
for c in r.chunks:
    print(c.sheet_name, c.top_left_cell, c.bottom_right_cell)
    print((c.render_text or "")[:600])
EOF
```
