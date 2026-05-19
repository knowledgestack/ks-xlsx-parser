# 01 · Array-formula cells render as `<openpyxl.worksheet.formula.ArrayFormula object>`

**Status:** 🆓 free to claim
**Slice:** B + D (`parsers/cell_parser.py`, `rendering/text_renderer.py`)
**Independent of:** 00, 02, 04 — partially overlapping with 03 on
`text_renderer.py`; coordinate via the README table.

## What it looks like

Two instances (`43026`, `59969`) hit this exact pattern. The
ground-truth answer cell contains an Excel **array formula** (an
`{= ...}` formula entered with Ctrl-Shift-Enter that returns multiple
cells). openpyxl returns these as an `ArrayFormula` Python object, NOT
a string and NOT a number. Our cell extractor stuffs the repr of that
object into the cell's value, so the rendered chunk text contains
literal:

```
<openpyxl.worksheet.formula.ArrayFormula object at 0x10c0767b0>
```

This obviously kills both text-match (the answer values are nowhere in
the chunk) and any downstream LLM consumption.

Instance `43026` (`summary!D10`):

```
gt_cell_raw  = '<openpyxl.worksheet.formula.ArrayFormula object at 0x10c0767b0>'
gt_cell_formula = None    ← our heuristic missed it because raw_value isn't a str starting with '='
gt_cell_data_only = (cached value, if any)
```

## Diagnostic signature

```bash
jq -c 'select(.gt_cell_raw | tostring | contains("ArrayFormula"))' \
    tests/benchmarks/reports/retrieval/*/enriched_failures.ndjson
```

Likely a much larger cluster at full-corpus scale — most "tabular
report" workbooks use array formulas. Two on the 200-sample is the floor.

## File scope

You may touch:

* `src/ks_xlsx_parser/parsers/cell_parser.py` — wherever the raw value
  is extracted from `openpyxl.cell.Cell.value`. Detect `ArrayFormula`,
  pull the formula text out of it (`obj.text` on openpyxl) and treat it
  as a formula cell.
* `src/ks_xlsx_parser/rendering/text_renderer.py::_cell_render_value`
  — make sure array-formula cells render their *cached value* (from the
  `data_only` workbook pass) when one exists; fall back to the formula
  expression as a string only if no cached value.
* `src/ks_xlsx_parser/models/cell.py` — add an `is_array_formula: bool`
  field if useful for downstream consumers.

Do NOT add a new evaluator for array formulas — out of scope (and
covered indirectly by the cached-value path).

## Acceptance criteria

1. Build a tiny array-formula fixture inside `tests/conftest.py` (the
   existing programmatic-fixture pattern). Use `openpyxl`'s
   `ArrayFormula` constructor; populate the `data_only` workbook with a
   plausible cached value.
2. Add a test that asserts the chunk's `render_text` for that fixture
   contains the cached value AND does **not** contain the substring
   `ArrayFormula object`.
3. On the 200-sample seed=1337 rerun, instances `43026` and `59969`
   move out of the failure set on the text-match metric IF the answer
   values exist as cached data in the input (verify with openpyxl
   first; if they don't, the instance is `instruction_requires_execution`
   and you can't move it — note it in the PR).
4. No previously-passing crossval test regresses (`make test`).

## Failing test sketch

```python
# tests/test_array_formula_rendering.py
import openpyxl
from openpyxl.worksheet.formula import ArrayFormula
from ks_xlsx_parser.pipeline import parse_workbook

def test_array_formula_renders_cached_value(tmp_path):
    p = tmp_path / "af.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = 10
    ws["A2"] = 20
    ws["A3"] = 30
    # Array formula in B1:B3 summing the row; cache a plausible value.
    af = ArrayFormula("B1:B3", "=A1:A3*2")
    ws["B1"] = af
    # openpyxl writes the cached value separately when data_only=True is read;
    # for fixture purposes, write the file once and patch cached values via
    # a second openpyxl pass.
    wb.save(p)

    chunks = parse_workbook(path=str(p)).chunks
    assert chunks, "no chunks produced"
    text = chunks[0].render_text
    assert "ArrayFormula" not in text, "raw ArrayFormula object leaked"
    # Once the fix lands, expect cached values 20 / 40 / 60 to render.
```

## Pitfalls

* openpyxl's behaviour for `ArrayFormula` differs between read-only
  and normal mode. Make sure both paths in `workbook_parser.py` agree
  on how the cell is surfaced.
* If the input was saved by LibreOffice or generated programmatically
  without computing array results, `cached_value` will be `None`. In
  that case there's no useful retrieval signal — surfacing the formula
  expression itself is acceptable as a fallback, but it must be a
  string, not `repr()`.
* "Spilled" dynamic arrays (Excel 365's new `=A1:A3*2` without
  Ctrl-Shift-Enter) are a related but distinct case. Note in your PR
  if you handled both or only the classic array formula.

## Repro

```bash
python -c "
from openpyxl import load_workbook
wb = load_workbook('data/corpora/spreadsheetbench/all_data_912_v0.1/spreadsheet/43026/1_43026_input.xlsx')
ws = wb['summary']
print(type(ws['D10'].value), repr(ws['D10'].value))
"
```
