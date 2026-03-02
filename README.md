# XLSXParser

Production-grade Excel parser built for RAG (Retrieval-Augmented Generation) systems with full auditability. Parses `.xlsx` workbooks into structured, loss-minimizing representations while preserving cell values, formulas, formatting, tables, charts, layout, and full dependency graphs with citation support.

## Table of Contents

- [Features](#features)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [API Reference](#api-reference)
- [Pipeline Architecture](#pipeline-architecture)
- [Web API](#web-api)
- [Data Models](#data-models)
- [Storage & Serialization](#storage--serialization)
- [Testing](#testing)
- [Project Structure](#project-structure)
- [Limitations](#limitations)
- [License](#license)

## Features

### Core Parsing
- **Cell extraction** -- values, formulas, number formats, data types, hyperlinks, and comments
- **Style preservation** -- fonts, fills, borders, alignment, and conditional formatting
- **Merged cells** -- full master/slave relationship detection with correct colspan/rowspan
- **Hidden elements** -- detection of hidden rows, columns, and sheets
- **Named ranges** -- workbook-scoped and sheet-scoped definitions
- **Data validation** -- extraction of dropdown lists and cell constraints

### Formula & Dependency Analysis
- Parse Excel formulas and extract all cell/range references
- Cross-sheet references (`Sheet2!A1`, `'My Sheet'!B2`)
- Structured table references (`SalesData[Revenue]`)
- External workbook references (`[Budget.xlsx]Sheet1!A1`)
- Directed dependency graphs with upstream/downstream traversal
- Circular reference detection

### Table & Structure Detection
- Excel ListObject table extraction with column definitions
- Auto-detection of table boundaries, headers, and data regions
- Layout segmentation into logical blocks via adaptive gap analysis and style boundaries
- Multi-table sheet support -- vertical, horizontal, and mixed layouts

### Chart Extraction
- Direct OOXML parsing for bar, line, pie, and scatter charts
- Extraction of chart titles, series data, axis labels, and category references
- Text summaries for RAG ingestion

### RAG-Optimized Output
- Token-counted chunks (via `tiktoken`) for LLM context window management
- HTML rendering with proper colspan/rowspan for merged cells
- Pipe-delimited plain text rendering for text-based RAG
- Source URIs with exact sheet coordinates for citation and traceability
- Content-addressable hashing (xxhash64) for deduplication and change detection

### Security
- No macro execution -- VBA modules are flagged but never run
- No external link resolution
- Input validation with file size and cell count limits
- ZIP bomb protection

## Installation

Requires Python 3.10+.

```bash
# Core library
pip install xlsx-parser

# With FastAPI web server
pip install xlsx-parser[api]

# With development/test tools
pip install xlsx-parser[dev]
```

### From Source

```bash
git clone https://github.com/arnav2/XLSXParser.git
cd XLSXParser
pip install -e ".[dev]"
```

### Dependencies

| Package | Purpose |
|---------|---------|
| `openpyxl>=3.1.0` | Excel file reading and cell extraction |
| `pydantic>=2.0` | Data validation and serialization |
| `lxml>=4.9.0` | Fast OOXML/XML parsing |
| `xxhash>=3.0.0` | Deterministic content hashing |
| `tiktoken>=0.5.0` | Token counting for RAG chunking |

## Quick Start

### Parse a Workbook

```python
from xlsx_parser import parse_workbook

result = parse_workbook(path="workbook.xlsx")

# Workbook metadata
print(f"Sheets: {result.workbook.total_sheets}")
print(f"Cells: {result.workbook.total_cells}")
print(f"Formulas: {result.workbook.total_formulas}")
print(f"Parse time: {result.workbook.parse_duration_ms:.0f}ms")
```

### Access RAG Chunks

```python
for chunk in result.chunks:
    print(f"Source: {chunk.source_uri}")
    print(f"Type: {chunk.block_type}")
    print(f"Tokens: {chunk.token_count}")
    print(f"Text:\n{chunk.render_text[:200]}")
```

### Inspect Formulas & Dependencies

```python
# Find all formula cells
for sheet in result.workbook.sheets:
    for cell in sheet.cells.values():
        if cell.formula:
            print(f"  {sheet.sheet_name}!{cell.address}: ={cell.formula}")

# Traverse dependency graph
from xlsx_parser.models import CellCoord

upstream = result.workbook.dependency_graph.get_upstream(
    "Sheet1", CellCoord(row=10, col=3), max_depth=3
)
for edge in upstream:
    print(f"  {edge.source_sheet}!{edge.source_coord.to_a1()} -> {edge.target_ref_string}")
```

### Parse from Bytes

```python
with open("workbook.xlsx", "rb") as f:
    content = f.read()

result = parse_workbook(content=content, filename="workbook.xlsx")
```

### Serialize to JSON

```python
json_output = result.to_json()

# Or get database-ready records
serializer = result.serializer
workbook_record = serializer.to_workbook_record()
sheet_records = serializer.to_sheet_records()
chunk_records = serializer.to_chunk_records()
vector_entries = serializer.to_vector_store_entries()
```

## API Reference

### `parse_workbook()`

Parse a single Excel workbook through stages 0-8 of the pipeline.

```python
def parse_workbook(
    path: str | Path | None = None,
    content: bytes | None = None,
    filename: str | None = None,
    max_cells_per_sheet: int = 2_000_000,
) -> ParseResult
```

**Parameters:**
- `path` -- Path to a `.xlsx` file
- `content` -- Raw file bytes (alternative to `path`)
- `filename` -- Display name when using `content`
- `max_cells_per_sheet` -- Safety limit per sheet (default: 2M)

**Returns:** `ParseResult` with `.workbook`, `.chunks`, and `.serializer`

### `compare_workbooks()`

Stage 9: Compare templates across multiple workbooks to find structural similarities and degrees of freedom.

```python
def compare_workbooks(
    paths: list[str | Path],
    dof_threshold: int = 50,
) -> GeneralizedTemplate
```

### `export_importer()`

Stage 10: Generate a reusable Python importer class from a generalized template.

```python
def export_importer(
    template: GeneralizedTemplate,
    output_path: str | Path,
    class_name: str = "GeneratedImporter",
) -> Path
```

### Multi-Workbook Workflow (Stages 9-10)

```python
from xlsx_parser import compare_workbooks, export_importer

# Compare multiple workbooks to find a common template
template = compare_workbooks([
    "report_q1.xlsx",
    "report_q2.xlsx",
    "report_q3.xlsx",
])

# Generate a Python importer for the template
export_importer(template, "generated_importer.py", class_name="QuarterlyReportImporter")
```

## Pipeline Architecture

XLSXParser implements the **Excellent Algorithm**, an 11-stage pipeline for deep structural analysis of Excel workbooks.

### Single-Document Stages (0-8)

| Stage | Name | Description |
|-------|------|-------------|
| 0 | Sheet Chunking | Adaptive gap analysis + style boundary detection via `LayoutSegmenter` |
| 1 | Cell Annotation | Feature-based scoring to classify cells (header, data, label, etc.) |
| 2 | Solid Block ID | Annotation-based splitting into contiguous blocks |
| 3 | Table Assembly | Associate labels with data regions to form table structures |
| 4 | Light Block Detection | Associate sparse/isolated blocks with nearby tables |
| 5 | Table Grouping | Cluster structurally similar tables |
| 6 | Pattern Splitting | Detect repeating label/template patterns |
| 7 | Tree Building | Build recursive hierarchy from blocks and structures |
| 8 | Template Extraction | Identify degrees of freedom across the structure |

### Multi-Document Stages (9-10)

| Stage | Name | Description |
|-------|------|-------------|
| 9 | Template Comparison | Compare templates across workbooks, resolve conflicts |
| 10 | Model Export | Generate Python importer classes from generalized templates |

### Final Processing

- **Rendering** -- HTML and plain text output with proper merged cell handling
- **Chunking** -- Token-counted chunks with source URIs for RAG systems
- **Serialization** -- Database-ready records for Postgres and vector stores

## Web API

XLSXParser includes a built-in FastAPI web application with a drag-and-drop UI.

```bash
# Install with API dependencies
pip install xlsx-parser[api]

# Start the server
uvicorn xlsx_parser.api:app --reload
```

Open `http://localhost:8000` to access the upload UI, or POST files directly:

```bash
curl -X POST http://localhost:8000/parse \
  -F "file=@workbook.xlsx"
```

The response includes:
- `parse_result` -- Full structured JSON output (workbook metadata + chunks)
- `verification_markdown` -- Pipeline stage verification report
- `verification` -- Structured verification data

## Data Models

All models use Pydantic v2 for validation and serialization.

| Model | Description |
|-------|-------------|
| `WorkbookDTO` | Root object: sheets, tables, charts, named ranges, dependency graph, errors |
| `SheetDTO` | Sheet with cells, merged regions, conditional formatting, data validation |
| `CellDTO` | Cell value, formula, style, coordinates, annotations |
| `TableDTO` | Excel ListObject table with name, columns, range, and style |
| `ChartDTO` | Chart metadata, series data, axis labels, chart type |
| `BlockDTO` | Logical block (HEADER, DATA, TABLE, etc.) with bounding box and hash |
| `ChunkDTO` | RAG chunk with HTML/text rendering, token count, source URI, content hash |
| `DependencyGraph` | Directed graph of formula dependencies with traversal methods |
| `TableStructure` | Assembled table structure with header/data regions |
| `TreeNode` | Hierarchical node from Stage 7 tree building |
| `TemplateNode` | Template node with degree-of-freedom annotations |

## Storage & Serialization

The `WorkbookSerializer` produces database-ready records:

```python
serializer = result.serializer

# Postgres-style records
workbook_record = serializer.to_workbook_record()   # Single workbook row
sheet_records = serializer.to_sheet_records()         # One row per sheet
chunk_records = serializer.to_chunk_records()         # One row per chunk

# Vector store entries (for embedding)
vector_entries = serializer.to_vector_store_entries()  # Text + metadata per chunk
```

### Deterministic Hashing

All elements are hashed with xxhash64 for content-addressable storage:

- **Workbook hash** -- derived from raw file bytes
- **Cell hash** -- sheet name + coordinates + value + formula
- **Block hash** -- sorted cell hashes (order-independent)
- **Chunk hash** -- globally unique and stable across repeated parses

## Testing

```bash
# Run all tests (excluding corpus tests)
pytest tests/ -v

# Run with coverage
pytest tests/ --cov=xlsx_parser --cov-report=term-missing

# Run specific test modules
pytest tests/test_parsers.py -v
pytest tests/test_pipeline.py -v
pytest tests/test_formula_parser.py -v

# Run cross-validation tests (requires python-calamine)
pytest tests/ -m crossval -v

# Run structural invariant tests
pytest tests/ -m invariant -v

# Run all tests including corpus
pytest tests/ -m "" -v
```

### Test Suite

The test suite includes 87 tests across 15 modules:

- **Unit tests** -- models, formula parsing, cell parsing
- **Integration tests** -- full pipeline, rendering, segmentation
- **Cross-validation** -- results verified against `python-calamine` as an independent reader
- **Real-world datasets** -- Iris, Titanic, Boston Housing, Wine Quality, Superstore, and more
- **Structural invariants** -- hash determinism, chunk completeness, coverage guarantees
- **Stress tests** -- 25 levels of increasingly complex workbooks

### Test Markers

| Marker | Description |
|--------|-------------|
| `crossval` | Cross-validation tests against calamine |
| `invariant` | Structural invariant tests |
| `corpus` | External corpus tests (skipped by default) |
| `slow` | Tests taking >10 seconds |

## Project Structure

```
XLSXParser/
├── pyproject.toml                          # Build config and dependencies
├── DESIGN.md                               # Technical design document
├── TEST_PLAN.md                            # Test strategy and coverage plan
├── docs/
│   └── PARSER_KNOWN_ISSUES.md              # Known limitations
├── examples/
│   ├── demo.py                             # Usage examples
│   ├── generate_examples.py                # Generate sample workbooks
│   ├── *.xlsx                              # Sample workbooks
│   └── stress_test/                        # Stress test fixtures and runner
├── src/xlsx_parser/
│   ├── __init__.py                         # Public API exports
│   ├── api.py                              # FastAPI web application
│   ├── pipeline.py                         # Main entry point (Stages 0-10)
│   ├── models/                             # Pydantic data models (DTOs)
│   ├── parsers/                            # Excel reading and extraction
│   ├── formula/                            # Formula parsing and dependency graphs
│   ├── charts/                             # Chart extraction via OOXML
│   ├── chunking/                           # Layout segmentation and RAG chunking
│   ├── rendering/                          # HTML and text output
│   ├── annotation/                         # Cell annotation and block splitting (Stages 1-2)
│   ├── analysis/                           # Table assembly through template extraction (Stages 3-8)
│   ├── comparison/                         # Multi-workbook template comparison (Stage 9)
│   ├── export/                             # Python code generation (Stage 10)
│   ├── storage/                            # Database serialization
│   ├── verification/                       # Pipeline stage verification
│   └── utils/                              # Logging configuration
└── tests/
    ├── conftest.py                         # 30+ programmatic test fixtures
    ├── helpers/                            # Test utilities (calamine reader, invariant checker)
    ├── fixtures/                           # Test data (real-world datasets)
    └── test_*.py                           # 15 test modules
```

## Limitations

- **`.xls` not supported** -- only `.xlsx` and `.xlsm` formats (OOXML); convert legacy files externally
- **Pivot tables** -- detected but not fully parsed
- **Sparklines** -- not extracted
- **VBA macros** -- flagged but never executed or analyzed
- **External links** -- recorded but not resolved
- **Threaded comments** -- only legacy comments are supported (openpyxl limitation)
- **Embedded OLE objects** -- detected but not extracted
- **Locale-dependent number formats** -- not interpreted

See [docs/PARSER_KNOWN_ISSUES.md](docs/PARSER_KNOWN_ISSUES.md) for additional edge cases.

## License

MIT
