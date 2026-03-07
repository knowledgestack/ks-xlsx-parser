---
name: excel-stress-tester-builder
description: Creates Excel workbooks that exercise every content type in the Workbook Graph extraction spec. Use when building stress-test workbooks, measuring parser completeness, or regression-testing the xlsx_parser. Covers cells, formulas, tables, charts, shapes, conditional formatting, data validation, and edge cases.
---

# Excel Stress-Test Workbook Builder

Create one or more Excel workbooks that collectively exercise every content type required by the extraction spec, for parser completeness measurement and regression testing.

## Inputs

| Input | Type | Description |
|-------|------|-------------|
| `output_dir` | path | Where to write workbooks and manifests |
| `seed` | int | For deterministic generation |
| `scale` | enum | `tiny` \| `small` \| `medium` \| `large` — rows/cols, chart count |
| `formats` | list | `["xlsx","xlsm","xlsb"]` — include what tooling supports |
| `features` | dict | Optional overrides: pivots, power query, slicers, shapes, external links |

## Outputs

- **Workbooks**: `KS_ExcelStress_v1.xlsx`, `KS_ExcelStress_v1_reportlike.xlsx`, `KS_ExcelStress_v1_edges.xlsx`
- **Manifest JSON** per workbook: what features exist and where (sheet, range, object id)
- **Golden expectations** `expected_extraction_min.json`: must-extract entities for the "Practical minimum" set

## Coverage Checklist (Must Include)

### Workbook layer

- [ ] Properties: created/modified, custom properties
- [ ] Calculation mode + date system (1900 vs 1904 if possible)
- [ ] Workbook-level defined names (named ranges, named formulas like `=LET`, `=LAMBDA`)
- [ ] External link placeholder (formula referencing another workbook, may be broken)
- [ ] Connections/queries metadata (or placeholder sheet + OOXML manifest note if not embeddable)

### Sheet layer

- [ ] Visible + hidden + very hidden sheets
- [ ] Tab colors
- [ ] Freeze panes + split panes
- [ ] Row/col widths, hidden rows/cols
- [ ] Merged cells
- [ ] Print area + header/footer
- [ ] Sheet-scoped defined names

### Object layer

**Grid objects**

- [ ] Cells: numeric, text, bool, date, error, blank
- [ ] Number formats: currency, %, accounting, custom date, fraction
- [ ] Formulas: relative, absolute, 3D refs, structured refs, dynamic arrays (spill)
- [ ] Rich text runs in a cell
- [ ] Hyperlinks: external + internal
- [ ] Data validation: list from range, custom formula
- [ ] Conditional formatting: formula rule, color scale, icon set
- [ ] Tables (ListObjects): header, totals, calculated column
- [ ] PivotTable (if feasible): cache from table, row/col/value fields
- [ ] AutoFilter with criteria

**Canvas objects**

- [ ] Charts: line, column, pie — series referencing ranges and table columns
- [ ] Text boxes, shapes, arrows, callouts — overlaps, z-order, grouping
- [ ] Images with anchors
- [ ] Slicer/timeline (optional)

### Manifest (derived artifacts)

- [ ] Sheet purpose label: input / calc / dashboard / report / raw / lookup
- [ ] KPI cells and label cells
- [ ] Reading-order anchors for textboxes and blocks

## Construction Strategy

1. **Deterministic**: Given `seed`, always place the same objects in the same cells.
2. **Three workbook archetypes**:
   - **Tabular model**: tables, pivots, charts
   - **Report layout**: blocks, merged cells, text boxes, KPI callouts
   - **Edge-cases**: spills, circular refs, weird number formats, rich text, hidden sheets

## Failure Modes

If PivotTables, PowerQuery, or slicers cannot be generated with current libraries:

- **Do not** fake the file
- Add a `NotGenerated` section in the manifest
- Provide manual creation script/instructions to add in Excel and re-save

## Acceptance Criteria

- Every checklist item present and recorded in `manifest.json` (sheet, range/object id, expected fields)
- Workbook opens without corruption in Excel/LibreOffice
- `expected_extraction_min.json` usable as test oracle for the extraction pipeline

## Reference

See `DESIGN.md` for the parser architecture and `docs/PARSER_KNOWN_ISSUES.md` for known limitations.
