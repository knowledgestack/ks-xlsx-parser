# Excel Workflow Parser — Technical Design Document

## 1. Architecture Overview

### Purpose
Parse `.xlsx` workbooks into structured, loss-minimizing representations for a RAG (Retrieval-Augmented Generation) + auditability system ("Knowledge Stack"). Every extracted element is traceable back to exact sheet coordinates and a stable workbook version hash.

### Assumptions
- **Input**: `.xlsx` files only (OOXML format). `.xls` (legacy BIFF) is out of scope; conversion via external tools is assumed.
- **No Excel automation**: No COM/OLE/Excel app dependency. Pure Python using `openpyxl` + direct OOXML XML parsing.
- **No macro execution**: `.xlsm` files are accepted but macros are flagged and never executed.
- **Deterministic output**: Given identical input bytes, the parser produces byte-identical output DTOs and hashes.
- **Streaming-friendly**: Large workbooks (100+ sheets, 1M+ cells) are processed with bounded memory via lazy parsing and sheet-level streaming.
- **Thread safety**: Sheet-level parsing is independent and can be parallelized.

### High-Level Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                     XLSX Workflow Parser                      │
├─────────────────────────────────────────────────────────────┤
│                                                              │
│  ┌──────┐   ┌───────┐   ┌───────────┐   ┌────────┐         │
│  │ Load │──▶│ Parse │──▶│ Normalize │──▶│Segment │         │
│  └──────┘   └───────┘   └───────────┘   └────────┘         │
│                                              │               │
│                                              ▼               │
│                                         ┌────────┐          │
│                                         │ Render │          │
│                                         └────────┘          │
│                                              │               │
│                                              ▼               │
│                                         ┌────────┐          │
│                                         │ Store  │          │
│                                         └────────┘          │
│                                                              │
├─────────────────────────────────────────────────────────────┤
│  Cross-cutting: Logging · Hashing · Error Collection         │
└─────────────────────────────────────────────────────────────┘
```

## 2. Pipeline Stages

### Stage 1: Load
- Accept file path or byte stream.
- Compute workbook-level content hash (xxhash64 of raw bytes) → `workbook_hash`.
- Open with `openpyxl.load_workbook(data_only=False)` for formulas.
- Open a second pass with `data_only=True` for computed values.
- For chart/image extraction: open as ZIP and parse OOXML parts directly.

### Stage 2: Parse
- **Per-sheet**: Extract cells (value, formula, format), merges, dimensions.
- **Workbook-level**: Extract defined names, external links, workbook properties.
- **Tables**: Extract ListObject table definitions via openpyxl.
- **Charts**: Extract from `/xl/charts/*.xml` via direct OOXML parsing.
- **Images/Shapes**: Extract from `/xl/drawings/*.xml` and `/xl/media/*`.
- **Comments**: Extract from comment parts.
- **Conditional Formatting**: Extract rules per sheet.

### Stage 3: Normalize
- Resolve merged cells (assign value to top-left, mark others as merged-into).
- Compute pixel bounding boxes using row heights and column widths.
- Normalize number formats to a canonical representation.
- Detect the "used range" per sheet (skip truly empty areas).
- Build formula dependency graph.

### Stage 4: Segment
- Identify logical blocks: tables, calculation regions, text headers, chart anchors.
- Use heuristics: blank row/col gaps, style continuity, border boundaries, merge groups.
- Assign each block a type and bounding coordinates.

### Stage 5: Render
- Per block: produce HTML (with merged cell spans, formatting) and plain-text/markdown.
- Charts: produce text summary + optional SVG placeholder.
- Compute token counts for RAG embedding.

### Stage 6: Store
- Serialize DTOs to JSON.
- Map to Postgres tables + vector store entries.
- Compute chunk hashes for dedup and versioning.

## 3. Library Choices & Limitations

### openpyxl
**Can do**: Cell values, formulas, styles, merges, defined names, tables (ListObject), conditional formatting, data validations, comments, sheet properties, workbook properties, column widths, row heights, images (partial).

**Cannot do well**:
- Chart data extraction (openpyxl has chart *creation* support but limited *reading* of existing charts)
- Pivot table internals
- Sparklines
- Threaded comments (only legacy comments)
- External link resolution
- Some advanced conditional formatting (data bars, icon sets—present but sometimes incomplete)

**Workarounds**:
- Charts: Parse `/xl/charts/chart*.xml` directly from the ZIP archive using `lxml`.
- Drawings/shapes: Parse `/xl/drawings/drawing*.xml` and cross-reference with `/_rels/` relationship files.
- Pivot tables: Detect presence via `/xl/pivotTables/` and `/xl/pivotCache/`; extract range metadata.

### Other libraries
- `lxml`: Fast XML parsing for OOXML parts.
- `xxhash`: Fast deterministic hashing (xxhash64).
- `tiktoken`: Token count estimation for RAG chunking.
- `pydantic`: DTO validation and serialization.

## 4. Performance Strategy

- **Used range detection**: Compute actual used range by scanning for non-empty cells; ignore openpyxl's sometimes-inflated `sheet.dimensions`.
- **Lazy style loading**: Cache shared style objects; don't re-parse identical styles.
- **Streaming for large sheets**: Use `openpyxl.load_workbook(read_only=True)` for initial cell scanning of very large sheets (>100K cells), then switch to full mode for style-heavy sheets.
- **Parallel sheet parsing**: Each sheet is independent after workbook-level metadata is loaded. Use `concurrent.futures.ProcessPoolExecutor` for CPU-bound parsing.
- **Sparse sheet handling**: Store only non-empty cells in a dict keyed by (row, col). Skip entirely empty rows/columns in segmentation.
- **Shared string caching**: openpyxl handles this internally; no additional work needed.

## 5. Hashing Strategy

### Workbook Hash
`xxhash64(file_bytes)` — stable across platforms, deterministic.

### Cell Hash
`xxhash64(f"{sheet_name}|{row}|{col}|{raw_value}|{formula}")` — captures identity and content.

### Block Hash
`xxhash64(sorted([cell_hash for cell in block_cells]))` — order-independent within block, stable.

### Chunk Hash
`xxhash64(f"{workbook_hash}|{sheet_name}|{block_type}|{top_left}|{bottom_right}|{block_hash}")` — globally unique, stable.

## 6. Security

- **No macro execution**: VBA modules are detected and flagged but never executed.
- **No external link resolution**: External references are recorded as-is.
- **Redaction hooks**: Every cell value passes through an optional `RedactionFilter` before rendering. Default is pass-through; can be extended for PII detection.
- **Input validation**: File size limits, sheet count limits, and cell count limits are configurable.
- **ZIP bomb protection**: Limit decompressed size when reading OOXML parts.

## 7. Logging

Structured logging with fields:
- `workbook_id`: hash of workbook
- `sheet_name`: current sheet
- `block_id`: current block being processed
- `stage`: pipeline stage (load, parse, normalize, segment, render, store)
- `level`: DEBUG/INFO/WARN/ERROR

Partial parse + error collection: errors are accumulated in a list on the DTO rather than raising exceptions, allowing maximum data extraction even from malformed workbooks.

## 8. Storage Mapping (Postgres + Vector Store)

### Postgres Tables
- `workbooks`: id, file_hash, filename, metadata_json, created_at
- `sheets`: id, workbook_id, sheet_name, index, properties_json
- `blocks`: id, sheet_id, block_type, top_left, bottom_right, content_hash, render_html, render_text, metadata_json
- `cells`: id, block_id, sheet_id, row, col, raw_value, display_value, formula, style_json
- `dependencies`: id, source_cell_id, target_cell_id, edge_type
- `charts`: id, sheet_id, chart_type, title, series_json, position_json, summary_text
- `named_ranges`: id, workbook_id, name, scope_sheet_id, ref_string

### Vector Store
- One embedding per block, using `render_text` as input.
- Metadata: chunk_id, sheet_name, block_type, coordinates, workbook_hash.
- Chunk IDs are deterministic hashes for dedup.

## 9. Limitations & Future Work

- **Pivot tables**: Detection and range extraction only; full pivot cache parsing is deferred.
- **Sparklines**: Not extracted in v1.
- **VBA macros**: Flagged but not analyzed.
- **External links**: Recorded but not resolved.
- **Threaded comments**: Only legacy comments in v1 (openpyxl limitation).
- **.xls files**: Out of scope; recommend conversion via LibreOffice CLI.
- **Locale-dependent formats**: Number formats are stored as-is; locale interpretation is deferred to the rendering layer.
- **Embedded OLE objects**: Detected but not parsed.
