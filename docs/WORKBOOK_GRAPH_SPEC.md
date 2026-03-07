# Workbook Graph Extraction Specification

> Canonical reference for what the XLSX parser extracts and how it is structured.
> Goal: an LLM can answer questions, reason about calculations, and navigate the
> workbook like a power user — without losing semantics.

---

## Architecture Overview

```
                          ┌──────────────────────┐
                          │    WorkbookDTO        │
                          │  (Workbook Layer)     │
                          └──────────┬───────────┘
                                     │
              ┌──────────────────────┼──────────────────────┐
              │                      │                      │
     ┌────────▼────────┐   ┌────────▼────────┐    ┌────────▼────────┐
     │   SheetDTO[0]   │   │   SheetDTO[1]   │    │   SheetDTO[N]   │
     │  (Sheet Layer)  │   │  (Sheet Layer)  │    │  (Sheet Layer)  │
     └───────┬─────────┘   └─────────────────┘    └─────────────────┘
             │
   ┌─────────┼──────────┬──────────┬──────────┐
   │         │          │          │          │
┌──▼──┐  ┌───▼───┐  ┌───▼───┐  ┌──▼──┐  ┌───▼───┐
│Cell │  │Table  │  │Chart  │  │Shape│  │Block  │
│DTOs │  │DTOs   │  │DTOs   │  │DTOs │  │DTOs   │
└─────┘  └───────┘  └───────┘  └─────┘  └───────┘
                    (Object Layer)
```

---

## 1. Workbook Layer

### WorkbookDTO

| Field                 | Type                    | Status       | Description                                     |
|-----------------------|-------------------------|--------------|--------------------------------------------------|
| `filename`            | `str`                   | Implemented  | Display filename                                 |
| `file_path`           | `str?`                  | Implemented  | Absolute path (if from disk)                     |
| `workbook_hash`       | `str`                   | Implemented  | xxhash64 of raw file bytes                       |
| `workbook_id`         | `str`                   | Implemented  | Deterministic ID (hash + filename)               |
| `sheets`              | `list[SheetDTO]`        | Implemented  | All worksheets                                   |
| `tables`              | `list[TableDTO]`        | Implemented  | Excel ListObject tables                          |
| `charts`              | `list[ChartDTO]`        | Implemented  | Extracted chart metadata + series                |
| `shapes`              | `list[ShapeDTO]`        | Implemented  | Images, text boxes, drawing objects               |
| `named_ranges`        | `list[NamedRangeDTO]`   | Implemented  | Defined names (workbook + sheet scope)           |
| `dependency_graph`    | `DependencyGraph`       | Implemented  | Formula dependency edges + cycle detection       |
| `external_links`      | `list[ExternalLink]`    | Implemented  | References to other workbooks                    |
| `properties`          | `WorkbookProperties`    | Implemented  | Creator, dates, calc settings                    |
| `pivot_tables`        | `list[PivotTableDTO]`   | **New**      | PivotTable structure (row/col/value fields)      |
| `sheet_summaries`     | `list[SheetSummaryDTO]` | **New**      | LLM-ready sheet purpose + key entities           |
| `kpi_catalog`         | `list[KpiDTO]`          | **New**      | Candidate KPI cells ranked by signal strength    |
| `table_structures`    | `list[TableStructure]`  | Implemented  | Header-body-footer assemblies (Stage 3)          |
| `tree_nodes`          | `list[TreeNode]`        | Implemented  | Hierarchical structure (Stage 7)                 |
| `template_nodes`      | `list[TemplateNode]`    | Implemented  | Degrees of freedom (Stage 8)                     |

### WorkbookProperties (Calculation Context)

| Field                    | Type               | Status       | Description                            |
|--------------------------|--------------------|--------------|----------------------------------------|
| `calc_mode`              | `str?`             | Implemented  | Raw calc mode string                   |
| `calculation_mode`       | `CalculationMode?` | **New**      | Typed enum (auto/manual/semiAutomatic) |
| `iterate_enabled`        | `bool`             | Implemented  | Iterative calculation enabled          |
| `iterate_count`          | `int?`             | Implemented  | Max iterations for circular refs       |
| `iterate_max_change`     | `float?`           | **New**      | Max delta for convergence              |
| `precision_as_displayed` | `bool`             | **New**      | Use displayed precision in calcs       |
| `date_system`            | `DateSystem`       | **New**      | 1900 (default) vs 1904 (Mac legacy)   |

### NamedRangeDTO

| Field              | Type           | Status       | Description                        |
|--------------------|----------------|--------------|------------------------------------|
| `name`             | `str`          | Implemented  | Human-readable name                |
| `ref_string`       | `str`          | Implemented  | Raw reference (e.g. Sheet1!$A$1)   |
| `scope_sheet`      | `str?`         | Implemented  | None = workbook scope              |
| `parsed_range`     | `CellRange?`   | Implemented  | Parsed cell range                  |
| `resolved_range`   | `CellRange?`   | **New**      | Fully resolved target range        |
| `usage_locations`  | `list[str]`    | **New**      | Cell refs that reference this name |
| `is_hidden`        | `bool`         | Implemented  | Hidden from user                   |
| `comment`          | `str?`         | Implemented  | Optional comment                   |

---

## 2. Sheet Layer

### SheetDTO

| Field                     | Type                        | Status       | Description                              |
|---------------------------|-----------------------------|--------------|------------------------------------------|
| `sheet_name`              | `str`                       | Implemented  | Tab name                                 |
| `sheet_index`             | `int`                       | Implemented  | 0-based position                         |
| `sheet_id`                | `str`                       | Implemented  | Deterministic ID                         |
| `cells`                   | `dict[str, CellDTO]`        | Implemented  | Sparse cell storage ("row,col" keys)     |
| `used_range`              | `CellRange?`                | Implemented  | Computed bounds of non-empty cells       |
| `row_heights`             | `dict[int, float]`          | Implemented  | Custom row heights (points)              |
| `col_widths`              | `dict[int, float]`          | Implemented  | Custom column widths (points)            |
| `hidden_rows`             | `set[int]`                  | Implemented  | Hidden row indices                       |
| `hidden_cols`             | `set[int]`                  | Implemented  | Hidden column indices                    |
| `merged_regions`          | `list[MergedRegion]`        | Implemented  | All merged cell ranges                   |
| `properties`              | `SheetProperties`           | Implemented  | Freeze panes, print area, visibility     |
| `conditional_format_rules`| `list[ConditionalFormatRule]`| Implemented  | CF rules with ranges and formulas        |
| `data_validations`        | `list[DataValidationRule]`  | Implemented  | Dropdown/validation rules                |
| `autofilter_range`        | `CellRange?`                | **New**      | Active autofilter range                  |
| `autofilter_criteria`     | `list[FilterCriteria]`      | **New**      | Column filter criteria                   |
| `sort_keys`               | `list[SortKey]`             | **New**      | Active sort state                        |

### SheetProperties

| Field                  | Type     | Status      | Description                        |
|------------------------|----------|-------------|------------------------------------|
| `is_hidden`            | `bool`   | Implemented | Sheet visibility (hidden/visible)  |
| `tab_color`            | `str?`   | Implemented | Tab color (hex)                    |
| `default_row_height`   | `float?` | Implemented | Default row height                 |
| `default_col_width`    | `float?` | Implemented | Default column width               |
| `freeze_pane`          | `str?`   | Implemented | Freeze pane split position         |
| `print_area`           | `str?`   | Implemented | Print area range                   |
| `auto_filter_range`    | `str?`   | Implemented | Autofilter range (A1 string)       |
| `sheet_protection`     | `bool`   | Implemented | Protection enabled                 |
| `right_to_left`        | `bool`   | Implemented | RTL reading order                  |

---

## 3. Object Layer

### CellDTO

| Field                  | Type              | Status       | Description                              |
|------------------------|-------------------|--------------|------------------------------------------|
| `coord`                | `CellCoord`       | Implemented  | Row and column (1-indexed)               |
| `sheet_name`           | `str`             | Implemented  | Parent sheet name                        |
| `raw_value`            | `Any`             | Implemented  | Python-native value                      |
| `display_value`        | `str?`            | Implemented  | Formatted string as shown in Excel       |
| `data_type`            | `str?`            | Implemented  | s/n/d/b/f/e type code                    |
| `formula`              | `str?`            | Implemented  | Raw formula (without leading =)          |
| `formula_value`        | `Any`             | Implemented  | Computed value from data_only pass       |
| `formula_r1c1`         | `str?`            | **New**      | R1C1-style formula                       |
| `formula_references`   | `list[str]`       | **New**      | Resolved cell/range refs from formula    |
| `rich_text_runs`       | `list[RichTextRun]`| **New**     | Mixed formatting runs within cell        |
| `spill_range`          | `CellRange?`      | **New**      | Dynamic array spill range                |
| `style`                | `CellStyle?`      | Implemented  | Font/fill/border/alignment snapshot      |
| `is_merged_master`     | `bool`            | Implemented  | Master of a merged region                |
| `is_merged_slave`      | `bool`            | Implemented  | Covered by a merge                       |
| `merge_master`         | `CellCoord?`      | Implemented  | Master cell coordinate (for slaves)      |
| `merge_extent`         | `int?`            | Implemented  | Row span (for masters)                   |
| `merge_col_extent`     | `int?`            | Implemented  | Column span (for masters)                |
| `comment_text`         | `str?`            | Implemented  | Comment body text                        |
| `comment_author`       | `str?`            | Implemented  | Comment author                           |
| `hyperlink`            | `str?`            | Implemented  | URL or internal target                   |
| `annotation`           | `CellAnnotation?` | Implemented | Stage 1 role (data/label)                |
| `cell_id`              | `str`             | Implemented  | Deterministic ID                         |
| `cell_hash`            | `str`             | Implemented  | Content hash (value+formula)             |

### RichTextRun

```json
{
  "text": "bold part",
  "bold": true,
  "italic": false,
  "color": "FF0000",
  "font_name": "Arial",
  "font_size": 12.0
}
```

### TableDTO (ListObject)

| Field              | Type               | Status      | Description                          |
|--------------------|--------------------|-------------|--------------------------------------|
| `table_name`       | `str`              | Implemented | Internal name                        |
| `display_name`     | `str`              | Implemented | User-visible name                    |
| `sheet_name`       | `str`              | Implemented | Parent sheet                         |
| `ref_range`        | `CellRange`        | Implemented | Full table range                     |
| `header_row_range` | `CellRange?`       | Implemented | Header row range                     |
| `data_range`       | `CellRange?`       | Implemented | Body data range                      |
| `total_row_range`  | `CellRange?`       | Implemented | Totals row range                     |
| `columns`          | `list[TableColumn]` | Implemented | Column names + indices              |
| `style_name`       | `str?`             | Implemented | Table style name                     |
| `has_header_row`   | `bool`             | Implemented | Whether header row exists            |
| `has_total_row`    | `bool`             | Implemented | Whether total row exists             |
| `table_id`         | `str`              | Implemented | Deterministic ID                     |

### ChartDTO

| Field            | Type               | Status      | Description                         |
|------------------|--------------------|-------------|-------------------------------------|
| `chart_type`     | `ChartType`        | Implemented | bar/line/pie/scatter/etc.           |
| `title`          | `str?`             | Implemented | Chart title text                    |
| `sheet_name`     | `str`              | Implemented | Parent sheet                        |
| `series`         | `list[ChartSeries]`| Implemented | Data series with refs               |
| `axes`           | `list[ChartAxis]`  | Implemented | Axis titles and types               |
| `anchor`         | `ChartAnchor?`     | Implemented | Position on sheet                   |
| `chart_id`       | `str`              | Implemented | Deterministic ID                    |

### ShapeDTO (Canvas Objects)

| Field              | Type           | Status       | Description                         |
|--------------------|----------------|--------------|-------------------------------------|
| `shape_type`       | `str`          | Implemented  | image/textBox/rectangle/etc.        |
| `sheet_name`       | `str`          | Implemented  | Parent sheet                        |
| `alt_text`         | `str?`         | Implemented  | Alt text for images                 |
| `text_content`     | `str?`         | Implemented  | Text box content                    |
| `image_ref`        | `str?`         | Implemented  | Path in OOXML package               |
| `anchor`           | `ShapeAnchor?` | Implemented  | Cell-based position                 |
| `width_emu`        | `int?`         | Implemented  | Width in EMUs                       |
| `height_emu`       | `int?`         | Implemented  | Height in EMUs                      |
| `z_index`          | `int?`         | **New**      | Drawing order (higher = on top)     |
| `reading_order`    | `int?`         | **New**      | Top-to-bottom, left-to-right order  |
| `group_id`         | `str?`         | **New**      | Parent group shape ID               |
| `rotation`         | `float?`       | **New**      | Rotation in degrees                 |

### PivotTableDTO

| Field                | Type                  | Status  | Description                       |
|----------------------|-----------------------|---------|-----------------------------------|
| `name`               | `str`                 | **New** | PivotTable name                   |
| `sheet_name`         | `str`                 | **New** | Parent sheet                      |
| `location`           | `str?`                | **New** | Output range                      |
| `cache_source_type`  | `str`                 | **New** | range/external/consolidation      |
| `cache_source_ref`   | `str?`                | **New** | Source data reference              |
| `row_fields`         | `list[PivotField]`    | **New** | Row area fields                   |
| `col_fields`         | `list[PivotField]`    | **New** | Column area fields                |
| `filter_fields`      | `list[PivotField]`    | **New** | Filter/page area fields           |
| `value_fields`       | `list[PivotValueField]`| **New** | Measure definitions              |
| `layout_type`        | `PivotLayoutType`     | **New** | compact/tabular/outline           |
| `slicer_connections` | `list[str]`           | **New** | Connected slicer names            |

---

## 4. Formula Dependency Graph

### DependencyGraph

| Field              | Type                    | Status      | Description                        |
|--------------------|-------------------------|-------------|------------------------------------|
| `edges`            | `list[DependencyEdgeDTO]`| Implemented | All dependency edges              |
| `topological_order`| `list[str]`             | Implemented | Evaluation order (if acyclic)      |
| `circular_groups`  | `list[list[str]]`       | Implemented | Circular reference groups          |

### DependencyEdgeDTO

| Field              | Type          | Status      | Description                          |
|--------------------|---------------|-------------|--------------------------------------|
| `source_sheet`     | `str`         | Implemented | Sheet containing the formula         |
| `source_coord`     | `CellCoord?`  | Implemented | Cell with the formula                |
| `target_sheet`     | `str?`        | Implemented | Sheet being referenced               |
| `target_coord`     | `CellCoord?`  | Implemented | Cell being referenced                |
| `target_range`     | `CellRange?`  | Implemented | Range being referenced               |
| `edge_type`        | `EdgeType`    | Implemented | cell_to_cell/cross_sheet/etc.        |

```
    A1 ──depends_on──► B1  (same sheet, cell_to_cell)
    A1 ──depends_on──► Sheet2!C5  (cross_sheet)
    A1 ──depends_on──► D1:D10  (cell_to_range)
```

---

## 5. LLM-Ready Derived Artifacts

### SheetSummaryDTO

Auto-detected purpose and key facts per sheet:

```json
{
  "sheet_name": "Revenue",
  "purpose": "raw_data",
  "purpose_confidence": 0.8,
  "total_cells": 500,
  "formula_count": 10,
  "formula_density": 0.02,
  "has_data_validation": false,
  "has_charts": false,
  "key_tables": ["RevenueData"],
  "key_output_cells": ["Revenue!G2"],
  "key_entities": ["Product", "Region", "Q1", "Q2", "Revenue"],
  "summary_text": "Sheet \"Revenue\" (Raw Data): 500 cells, 10 formulas. Tables: RevenueData."
}
```

**Purpose Detection Heuristics:**

| Purpose       | Signal                                                          |
|---------------|-----------------------------------------------------------------|
| `dashboard`   | Has charts + low formula density                                |
| `raw_data`    | Low formula density + many cells + tabular structure            |
| `lookup`      | High cross-sheet in-degree (other sheets reference this one)    |
| `calculation` | High formula density + few charts                               |
| `input`       | Has data validation + low formula density                       |
| `report`      | Has print area + formatted headers                              |
| `template`    | Has protection + data validation + named ranges                 |
| `config`      | Small sheet + named ranges                                      |

### KpiDTO

Candidate KPI cells ranked by signal strength:

```json
{
  "label": "Year 1 Net",
  "cell_ref": "Model!B10",
  "value_display": "829,000",
  "sheet_name": "Model",
  "in_degree": 3,
  "drivers": ["Model!B8", "Model!B9"]
}
```

**KPI Detection Signals:**
- Currency/percentage number format (+2/+1)
- Bold font with formula (+2)
- High in-degree in dependency graph (+3/+4)
- Referenced by charts (+2)
- Large font size (+1)

### Entity Index

Business entities extracted from headers, table columns, and named ranges:

```json
{
  "entities": [
    {
      "name": "Revenue",
      "category": "measure",
      "locations": [
        {"sheet_name": "P&L", "range_a1": "B2", "source": "header"},
        {"sheet_name": "Summary", "range_a1": "A1:G50", "source": "table_column"}
      ]
    }
  ]
}
```

### Reading-Order Linearization

Slide-like text per sheet for LLM consumption:

```
## Sheet: Financial Summary
### [A1] Title: "Q4 2024 Financial Report"
[A2] Revenue
[A3] Cost
### [A1:G5] Table: "SalesData" (4 rows)
Columns: Product, Region, Q1, Q2, Q3, Q4, Total
### Chart: "Revenue by Region" (bar, 4 series)
[A20] Note: "Excludes discontinued operations"
```

---

## 6. Chunking Strategy (RAG)

### Principles

1. **BBox-aware** — chunks respect spatial layout; a table is never split mid-row
2. **Anchor-preserving** — every chunk has a `cell_range` back to its source location
3. **Token-budgeted** — chunks target a configurable token limit (default: 512 tokens)
4. **Context-carrying** — each chunk includes sheet name, block type, and parent headers

### Chunk Structure (ChunkDTO)

| Field              | Type          | Description                               |
|--------------------|---------------|-------------------------------------------|
| `chunk_id`         | `str`         | Deterministic ID                          |
| `sheet_name`       | `str`         | Source sheet                              |
| `cell_range`       | `CellRange`   | Bounding range of the chunk               |
| `block_type`       | `BlockType`   | table/header/text_block/etc.              |
| `html_content`     | `str`         | HTML rendering (with rowspan/colspan)     |
| `text_content`     | `str`         | Plain text rendering                      |
| `token_count`      | `int`         | Estimated token count                     |
| `content_hash`     | `str`         | xxhash64 of content                       |
| `dependencies`     | `list[str]`   | Chunk IDs this chunk depends on           |
| `metadata`         | `dict`        | Arbitrary metadata (headers, labels, etc.)|

### Chunk Ordering

Chunks are ordered within each sheet by their bounding box:
1. Top-to-bottom (by `cell_range.top_left.row`)
2. Left-to-right (by `cell_range.top_left.col`)

---

## 7. Merge Recovery (OOXML Fallback)

When openpyxl reports an empty master cell in a merged region, the parser
falls back to raw OOXML XML parsing:

1. Opens the `.xlsx` as a ZIP archive
2. Parses `xl/worksheets/sheet{N}.xml` with `lxml`
3. Loads `xl/sharedStrings.xml` for string values
4. Scans `<c>` elements within the merge range
5. Promotes the first non-None value found to the master cell

This resolves the `merge_empty_master` issue documented in
`PARSER_KNOWN_ISSUES.md`.

---

## 8. Practical Minimum (80% Coverage)

For a first integration, extract this subset:

- [x] Cells: value + display_text + formula + number_format + hyperlink
- [x] Tables + headers + columns
- [x] Named ranges/formulas (workbook + sheet scope)
- [x] Dependency graph edges + cycle detection
- [x] Charts: series refs + title + type
- [x] Sheet geometry: row heights, col widths, hidden rows/cols
- [x] Merged cells: master/slave annotations + recovery
- [x] Comments/notes
- [x] Canvas text boxes (text + bbox)
- [x] Rich text runs
- [x] Formula references
- [x] Calculation context
- [x] Autofilter state
- [x] Sheet summaries + KPI catalog + entity index

---

## 9. Pipeline Stages

```
Stage 0: Sheet Chunking ─────► Adaptive gap + style boundary detection
Stage 1: Cell Annotation ────► Feature-based scoring (header/data/label)
Stage 2: Solid Block ID ─────► Annotation-based contiguous blocks
Stage 3: Table Assembly ─────► Associate labels with data regions
Stage 4: Light Block Detect ─► Associate sparse/isolated blocks
Stage 5: Table Grouping ─────► Cluster structurally similar tables
Stage 6: Pattern Splitting ──► Detect repeating label/template patterns
Stage 7: Tree Building ──────► Build recursive hierarchy
Stage 8: Template Extraction ► Identify degrees of freedom
Stage 9: Template Comparison ► Cross-workbook comparison (multi-wb)
Stage 10: Model Export ──────► Generate Python importer classes
```

---

## 10. Hashing Strategy

All content-addressable IDs use **xxhash64** (via `xxhash` Python package):

- **Workbook ID**: `xxh64(workbook_hash + filename)`
- **Sheet ID**: `xxh64(workbook_hash + sheet_name + sheet_index)`
- **Cell ID**: `sheet_name|row|col` (human-readable)
- **Cell Hash**: `xxh64(sheet_name + row + col + value + formula)`
- **Chunk Hash**: `xxh64(content)` — for change detection
- **Table ID**: `xxh64(workbook_hash + sheet_name + table_name)`

Properties:
- Deterministic across runs (same input = same hash)
- Order-independent where appropriate
- Collision-resistant (64-bit, ~2^32 items before birthday collision)
