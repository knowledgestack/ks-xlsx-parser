# Data Models

Every DTO is a [Pydantic v2](https://docs.pydantic.dev/) model. Fully
JSON-serialisable, validated on construction, shipping with `py.typed` so
your editor gives you autocomplete and type errors.

For the canonical machine-readable spec, see
[`docs/WORKBOOK_GRAPH_SPEC.md`](https://github.com/knowledgestack/ks-xlsx-parser/blob/main/docs/WORKBOOK_GRAPH_SPEC.md).

## High-level map

| Model | Role |
|-------|------|
| `ParseResult`       | Entry-point wrapper: workbook + chunks + serializer + verification. |
| `WorkbookDTO`       | Root: sheets, tables, charts, named ranges, dependency graph, errors. |
| `SheetDTO`          | Cells, merged regions, conditional formatting, data validation. |
| `CellDTO`           | Value, formula, style, coordinates, annotations. |
| `TableDTO`          | Excel ListObject: name, columns, range, style. |
| `ChartDTO`          | Chart metadata, series, axis labels, chart type. |
| `BlockDTO`          | Logical block (`HEADER` / `DATA` / `TABLE` / …) with bounding box + hash. |
| `ChunkDTO`          | LLM chunk: HTML + text rendering, token count, source URI, content hash. |
| `DependencyGraph`   | Directed graph of formula dependencies with traversal helpers. |
| `DependencyEdge`    | One edge in the graph. |
| `CellCoord`         | `(row, col)` pair with `.to_a1()` and `.from_a1()` helpers. |
| `CellRange`         | Inclusive top-left → bottom-right pair. |
| `TableStructure`    | Assembled table with header / data regions (post-segmentation). |
| `TreeNode`          | Hierarchical node from tree building. |
| `TemplateNode`      | Template node with degree-of-freedom annotations. |
| `ParseError`        | Structured error / warning attached to a workbook, sheet, or cell. |

## `ParseResult`

```python
class ParseResult:
    workbook: WorkbookDTO
    chunks: list[ChunkDTO]
    serializer: WorkbookSerializer

    def to_json(self) -> dict[str, Any]: ...
```

`to_json()` returns a nested dict, fully JSON-serialisable via
`json.dumps(result.to_json(), default=str)`.

## `WorkbookDTO`

Key fields:

| Field | Type | Description |
|---|---|---|
| `workbook_id` | `str` | Deterministic ID derived from `file_path` + `workbook_hash`. |
| `filename` | `str` | Display name. |
| `file_path` | `str \| None` | Original path if `parse_workbook(path=…)` was used. |
| `workbook_hash` | `str` | xxhash64 of the file bytes. |
| `sheets` | `list[SheetDTO]` | All sheets, in Excel tab order. |
| `tables` | `list[TableDTO]` | All Excel ListObjects across sheets. |
| `charts` | `list[ChartDTO]` | All charts across sheets. |
| `named_ranges` | `list[NamedRangeDTO]` | Workbook- and sheet-scoped. |
| `dependency_graph` | `DependencyGraph` | Global cell-level dep graph. |
| `kpi_catalog` | `list[KpiDTO]` | Workbook-level KPIs surfaced during annotation. |
| `errors` | `list[ParseError]` | Non-fatal warnings / errors. |
| `total_sheets` / `total_cells` / `total_formulas` | `int` | Aggregates. |
| `parse_duration_ms` | `float` | Wall time for the parse pipeline. |

## `CellDTO`

```python
class CellDTO(BaseModel):
    address: str          # "A1"
    row: int
    col: int
    value: Any            # str | int | float | bool | datetime | None
    formula: str | None   # "=SUM(A1:A5)"
    data_type: str        # "n", "s", "f", "d", "b", "e"
    number_format: str | None
    font: FontDTO | None
    fill: FillDTO | None
    border: BorderDTO | None
    alignment: AlignmentDTO | None
    hyperlink: str | None
    comment: str | None
    is_merged_master: bool
    is_merged_slave: bool
    merge_master: CellCoord | None
    merge_extent: int | None       # row span
    merge_col_extent: int | None   # col span
    is_empty: bool
```

## `ChunkDTO`

The object you'll spend most of your LLM-integration time with:

```python
class ChunkDTO(BaseModel):
    chunk_id: str                  # stable, deterministic across runs
    chunk_index: int
    sheet_name: str
    block_type: str                # "HEADER", "DATA", "TABLE", "CHART_ANCHOR", ...
    top_left_cell: str             # "A1"
    bottom_right_cell: str         # "F18"
    cell_range: CellRange | None
    key_cells: list[str]           # highlighted refs the block is "about"
    named_ranges: list[str]        # named ranges touching the block
    dependency_summary: DependencySummary
    render_html: str               # proper colspan/rowspan
    render_text: str               # pipe-delimited LLM-friendly
    token_count: int               # via tiktoken
    source_uri: str                # "file.xlsx#Sheet!A1:F18"
    content_hash: str              # xxhash64 of render_text
    prev_chunk_id: str | None
    next_chunk_id: str | None
    metadata: dict[str, Any]
```

## `DependencyGraph`

Relevant methods:

```python
graph.get_upstream(sheet: str, coord: CellCoord, max_depth: int = 2) -> list[DependencyEdge]
graph.get_downstream(sheet: str, coord: CellCoord, max_depth: int = 2) -> list[DependencyEdge]
graph.detect_circular_refs() -> set[str]     # {"Sheet!A1", "Sheet!B3", ...}
graph.edges_out_of(sheet: str, coord: CellCoord) -> list[DependencyEdge]
graph.edges_into(sheet: str, coord: CellCoord) -> list[DependencyEdge]
```

Each `DependencyEdge` carries `source_sheet`, `source_coord`,
`target_sheet`, `target_coord`, `target_ref_string`, and an `edge_type`
enum (`WITHIN_SHEET`, `CROSS_SHEET`, `TABLE_REF`, `EXTERNAL`,
`NAMED_RANGE`).

## `ParseError`

Non-fatal diagnostics live on `workbook.errors` (workbook-level) and
`sheet.errors` (sheet-level):

```python
class ParseError(BaseModel):
    severity: Severity           # INFO | WARNING | ERROR
    stage: str                   # "load", "parse", "annotate", "segment", ...
    message: str
    sheet_name: str | None
    cell_address: str | None
```

Anything that would traditionally raise an exception becomes a structured
record here, so a single pathological file can't crash a batch
processing pipeline.

## JSON shape

Full shape is recursive but roughly:

```jsonc
{
  "workbook": {
    "workbook_id": "...",
    "filename": "workbook.xlsx",
    "workbook_hash": "...",
    "sheets": [
      {
        "sheet_name": "Sheet1",
        "cells": { "A1": { /* CellDTO */ }, ... },
        "merged_regions": [ ... ],
        "conditional_format_rules": [ ... ],
        "data_validations": [ ... ],
        "properties": { ... }
      }
    ],
    "tables": [ ... ],
    "charts": [ ... ],
    "named_ranges": [ ... ],
    "dependency_edges": [ ... ],
    "kpi_catalog": [ ... ],
    "errors": [ ... ],
    "total_sheets": 3,
    "total_cells": 1240,
    "total_formulas": 87
  },
  "chunks": [ /* ChunkDTO[] */ ]
}
```

For the exact field-by-field breakdown including every optional field,
see [`docs/WORKBOOK_GRAPH_SPEC.md`](https://github.com/knowledgestack/ks-xlsx-parser/blob/main/docs/WORKBOOK_GRAPH_SPEC.md).
