# Excel Workflow Parser — Test Plan

## Overview

This test plan covers the xlsx_parser system with 87 test cases across 6 test modules,
using programmatically generated fixture workbooks. Tests are organized by component
and verify both correctness and determinism.

## Test Modules

### 1. `test_models.py` — DTO & Utility Tests (13 tests)

| # | Test | Verifies |
|---|------|----------|
| 1 | `test_a1_simple` | CellCoord(1,1) → "A1" |
| 2 | `test_a1_double_letter` | CellCoord(1,27) → "AA1" |
| 3 | `test_a1_large_col` | CellCoord(100,26) → "Z100" |
| 4 | `test_a1_triple_letter` | CellCoord(1,703) → "AAA1" |
| 5 | `test_a1_range` | CellRange → "A1:C10" |
| 6 | `test_contains` | Range containment logic (inside and outside) |
| 7 | `test_row_col_count` | Row/col count computation |
| 8 | `test_letter_to_number` | A→1, Z→26, AA→27, AZ→52, BA→53 |
| 9 | `test_number_to_letter` | 1→A, 26→Z, 27→AA, 52→AZ |
| 10 | `test_roundtrip` | Letter↔number roundtrip for edge values |
| 11 | `test_same_input_same_hash` | Deterministic hash: same input = same output |
| 12 | `test_different_input_different_hash` | Different input = different hash |
| 13 | `test_hash_is_hex_string` | Hash is valid 16-char hex string (xxhash64) |

### 2. `test_formula_parser.py` — Formula Parsing & Dependencies (14 tests)

| # | Test | Verifies |
|---|------|----------|
| 14 | `test_simple_cell_ref` | Extract A1, B1 from "A1+B1" |
| 15 | `test_range_ref` | Extract A1:A10 from "SUM(A1:A10)" |
| 16 | `test_cross_sheet_ref` | Sheet2!A1, Sheet2!B1 extraction |
| 17 | `test_quoted_sheet_ref` | 'My Sheet'!A1 with quoted sheet name |
| 18 | `test_absolute_refs` | $A$1, $B$2:$C$3 → stripped absolute refs |
| 19 | `test_external_ref` | [Budget.xlsx]Sheet1!A1 external workbook ref |
| 20 | `test_structured_table_ref` | SalesData[Revenue] structured table ref |
| 21 | `test_complex_formula` | IF/SUM with mixed cross-sheet refs |
| 22 | `test_no_refs` | "1+2+3" yields zero references |
| 23 | `test_function_not_treated_as_table` | SUM not confused with table name |
| 24 | `test_upstream_traversal` | Graph forward traversal finds all deps |
| 25 | `test_downstream_traversal` | Graph backward traversal finds dependents |
| 26 | `test_circular_ref_detection` | A1→B1→A1 cycle detected and flagged |
| 27 | `test_cross_sheet_edge` | Cross-sheet dependency typed correctly |

### 3. `test_parsers.py` — Workbook/Sheet/Cell Parsing (24 tests)

| # | Test | Verifies |
|---|------|----------|
| 28 | `test_parse_cell_values` | Cell values extracted with correct types |
| 29 | `test_parse_formula` | Formula string extracted (without "=") |
| 30 | `test_parse_number_format` | Numeric values preserved as Python numbers |
| 31 | `test_workbook_hash_deterministic` | Same file = same hash on repeated parse |
| 32 | `test_cell_ids_populated` | All cells get cell_id and cell_hash after finalize |
| 33 | `test_merge_regions_detected` | Merged cell regions found (≥2 in fixture) |
| 34 | `test_merge_master_annotated` | Master cell has is_merged_master + extents |
| 35 | `test_merge_slave_annotated` | Slave cell has is_merged_slave flag |
| 36 | `test_cross_sheet_formulas` | Formula references other sheets ("Inputs!B1") |
| 37 | `test_dependency_graph_built` | Dependency edges created from formulas |
| 38 | `test_named_ranges_extracted` | "Price" and "Quantity" named ranges found |
| 39 | `test_table_detected` | Excel ListObject table found |
| 40 | `test_table_properties` | Table name, columns, column count correct |
| 41 | `test_table_range` | Table ref range = "A1:G5" |
| 42 | `test_rules_extracted` | Conditional formatting rule type + operator |
| 43 | `test_validation_extracted` | Data validation list type detected |
| 44 | `test_all_sheets_parsed` | 3-sheet workbook → 3 SheetDTOs |
| 45 | `test_hidden_sheet_detected` | Hidden sheet flagged in properties |
| 46 | `test_hidden_row_detected` | Hidden row 3 in hidden_rows set |
| 47 | `test_hidden_col_detected` | Hidden column B (col 2) in hidden_cols |
| 48 | `test_comments_extracted` | Comment text and author preserved |
| 49 | `test_sparse_cells_extracted` | Only 4 cells stored for sparse sheet |
| 50 | `test_used_range_spans_sparse` | Used range extends to row 1000 |
| 51 | `test_freeze_pane_extracted` | Freeze pane = "A2" |
| 52 | `test_font_color_extracted` | Bold font detected in styled cell |
| 53 | `test_fill_extracted` | Cell fill color preserved |
| 54 | `test_border_extracted` | Cell borders detected |
| 55 | `test_100_columns_parsed` | 500 cells (100 cols × 5 rows) parsed |
| 56 | `test_hyperlink_extracted` | Hyperlink URL preserved |

### 4. `test_charts.py` — Chart Extraction (6 tests)

| # | Test | Verifies |
|---|------|----------|
| 57 | `test_chart_detected` | Chart found in OOXML archive |
| 58 | `test_chart_type` | Bar chart type correctly identified |
| 59 | `test_chart_title` | "Monthly Revenue" title extracted |
| 60 | `test_chart_series` | At least 1 data series found |
| 61 | `test_chart_summary` | Summary text includes type + title |
| 62 | `test_chart_from_bytes` | Chart extraction from bytes (not file path) |

### 5. `test_segmentation.py` — Layout Segmentation (6 tests)

| # | Test | Verifies |
|---|------|----------|
| 63 | `test_simple_block_detected` | At least 1 block found for simple workbook |
| 64 | `test_table_block_from_listobject` | Excel table → TABLE block with table_name |
| 65 | `test_assumptions_block_classified` | Blank row separates assumptions from results (≥2 blocks) |
| 66 | `test_sparse_segmentation` | Distant cells form separate blocks |
| 67 | `test_blocks_have_bounding_box` | All blocks have cell_range + bounding_box |
| 68 | `test_block_hashes_deterministic` | Same input = same block content_hash |

### 6. `test_rendering.py` — HTML & Text Rendering (7 tests)

| # | Test | Verifies |
|---|------|----------|
| 69 | `test_basic_html_output` | Valid HTML table with data-sheet attribute |
| 70 | `test_merged_cell_rowspan_colspan` | Merged cells produce colspan attributes |
| 71 | `test_bold_rendered_as_style` | Bold font → "font-weight:bold" inline CSS |
| 72 | `test_data_ref_attributes` | data-ref="A1" on cell elements |
| 73 | `test_basic_text_output` | Pipe-delimited table format with sheet name |
| 74 | `test_formula_annotation` | Formula cells annotated with [=] marker |
| 75 | `test_text_includes_range` | Text includes sheet!range reference |

### 7. `test_pipeline.py` — End-to-End Pipeline (12 tests)

| # | Test | Verifies |
|---|------|----------|
| 76 | `test_simple_pipeline` | Full pipeline produces sheets, cells, chunks |
| 77 | `test_chunks_have_rendered_content` | Every chunk has HTML + text + token count |
| 78 | `test_chunks_have_source_uri` | Chunks have source_uri, chunk_id, content_hash |
| 79 | `test_chunk_navigation` | prev/next chunk IDs set for sequential traversal |
| 80 | `test_formula_workbook_pipeline` | 3 sheets, formulas, deps, named ranges extracted |
| 81 | `test_table_workbook_pipeline` | Table workbook → chunks with table block |
| 82 | `test_deterministic_output` | Two parses of same file → identical chunk IDs/hashes |
| 83 | `test_to_json` | ParseResult.to_json() is valid JSON-serializable dict |
| 84 | `test_parse_from_bytes` | Parsing from raw bytes (no file path) works |
| 85 | `test_serializer_records` | Serializer produces valid workbook/sheet/chunk/vector records |
| 86 | `test_multiple_blocks_detected` | Assumptions workbook → ≥2 separate blocks |
| 87 | `test_dependency_context_in_chunks` | At least one chunk has upstream dependency refs |

## Fixture Workbooks

All fixtures are generated programmatically in `tests/conftest.py`:

| Fixture | Description | Edge Cases Covered |
|---------|-------------|-------------------|
| `simple_workbook` | 4 cells, bold headers, formula, number format | Basic parsing |
| `merged_cells_workbook` | Horizontal merge, vertical merge, multi-cell merge | Merged cell handling |
| `formula_workbook` | 3 sheets with cross-sheet formulas + named ranges | Cross-sheet deps, named ranges |
| `table_workbook` | ListObject table with 7 cols, formulas, style | Table detection, structured refs |
| `chart_workbook` | Bar chart with 6 data points | OOXML chart parsing |
| `large_sparse_workbook` | Data at A1, Z100, CV1000 | Sparse sheet, used range |
| `conditional_format_workbook` | cellIs > 50 rule | Conditional formatting |
| `data_validation_workbook` | List dropdown validation | Data validation |
| `multi_sheet_workbook` | 3 sheets (1 hidden) with cross-refs | Hidden sheets, multi-sheet |
| `hidden_rows_cols_workbook` | Hidden row 3, hidden column B | Hidden rows/cols |
| `comment_workbook` | Cells with comments | Comment extraction |
| `freeze_panes_workbook` | Freeze at A2 | Freeze pane detection |
| `wide_workbook` | 100 columns × 5 rows | Wide sheet handling |
| `styled_workbook` | Borders, fills, fonts, alignment | Style extraction |
| `assumptions_workbook` | Assumptions + results with blank separator | Block segmentation, block typing |
| `hyperlink_workbook` | Cell with URL hyperlink | Hyperlink extraction |

## Golden Test Strategy

For snapshot/golden testing:
1. Parse a fixture workbook and serialize chunks to JSON
2. Hash the serialized JSON with xxhash64
3. Store the expected hash in a `.golden` file
4. On test run, compare computed hash to stored hash
5. If mismatch, fail with diff guidance

**Implementation**: Not yet implemented; planned for v0.2 once the DTO schema stabilizes.

## Performance Testing

Planned large-scale tests (not in current suite):
- 100-sheet workbook with 10K cells each → verify <30s parse time
- 1M-cell single sheet → verify <60s parse time and <2GB memory
- 10K-formula sheet → verify dependency graph builds in <5s

## Security Testing

- `.xlsm` file → verify macros flagged, never executed
- Workbook with external links → verify recorded but not resolved
- Malformed OOXML → verify partial parse with error collection
