# XLSXParser

**A production-grade Excel parser for RAG, data-pipeline, and audit systems — now open source.**

[![PyPI version](https://img.shields.io/pypi/v/ks-xlsx-parser.svg)](https://pypi.org/project/ks-xlsx-parser/)
[![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Tests](https://img.shields.io/badge/tests-1053%20workbooks-success.svg)](#the-testbench-dataset)

XLSXParser turns any `.xlsx` file into a fully-typed, loss-minimising
representation — cells, formulas, tables, charts, styles, merged regions,
conditional formatting, dependency graphs, named ranges, and RAG-ready chunks
with source URIs for citation. It is built for systems that need the *whole*
workbook, not just a dataframe: knowledge bases, document understanding
pipelines, financial auditing tools, and LLM agents that reason about
spreadsheets.

---

## ⭐ If this is useful to you

This project is free, open source (MIT), and maintained by a small team.
The single most helpful thing you can do is **[star the repo on GitHub](https://github.com/arnav2/XLSXParser)** —
it's how we justify spending more time on it. 👍

**Get involved:**

- 💬 [Discussions](https://github.com/arnav2/XLSXParser/discussions) — ask questions, share what you built, or float an idea
- 🐞 [Issues](https://github.com/arnav2/XLSXParser/issues/new/choose) — report a bug, request a feature, or file a parser edge case
- 🎯 [Show & Tell](https://github.com/arnav2/XLSXParser/discussions/new?category=show-and-tell) — tell us about your production use
- 🔐 [Security](https://github.com/arnav2/XLSXParser/security/advisories/new) — report a vulnerability privately
- 🙌 [Contribute](CONTRIBUTING.md) — every PR is reviewed, and good-first-issues are labeled

Not sure where to start? Run `make testbench`, find a file that breaks, and
open a [Parser edge case](https://github.com/arnav2/XLSXParser/issues/new?template=parser_edge_case.yml).
That's the fastest path to a merged PR.

---

## Table of Contents

- [Why another Excel parser?](#why-another-excel-parser)
- [Architecture](#architecture)
- [Installation](#installation)
- [Quick start](#quick-start)
- [API reference](#api-reference)
- [Web API](#web-api)
- [The testBench dataset](#the-testbench-dataset)
- [Data models](#data-models)
- [Limitations](#limitations)
- [Contributing](#contributing)
- [License](#license)

---

## Why another Excel parser?

Most Excel libraries answer one of two questions well: *"read a rectangle of
values"* (pandas, openpyxl) or *"run Excel headless"* (xlwings, LibreOffice).
XLSXParser answers a third one: **"give me a structured, inspectable,
loss-minimising graph of a workbook that an LLM or an auditor can reason
about."**

Concretely, one call to `parse_workbook()` gives you:

| Output | What it's for |
|--------|---------------|
| Typed cell graph (values, formulas, styles, coordinates) | Faithful round-trip into JSON / DB / vector store |
| Formula AST + directed dependency graph | Impact analysis, lineage, circular-reference detection |
| Detected tables, merged regions, layout blocks | Structure preservation without losing multi-table sheets |
| Chart extractions (bar / line / pie / scatter / area / radar / bubble) | Text summaries for RAG |
| Token-counted render chunks (HTML + pipe-text) | Drop straight into an embedding pipeline |
| Citation-ready source URIs (`sheet!A1:B10`) | Point an LLM back at the exact cells it cited |
| Deterministic content hashes (xxhash64) | Deduplication, change detection |

Everything is deterministic, everything is tested, and everything is
open source.

---

## Architecture

```
                          ┌────────────────────────────┐
   .xlsx bytes ────────▶  │  parsers/ (OOXML drivers)  │
                          │  openpyxl + raw lxml       │
                          └─────────────┬──────────────┘
                                        │
                                        ▼
                          ┌────────────────────────────┐
                          │  models/ (Pydantic DTOs)   │
                          │  WorkbookDTO, SheetDTO,    │
                          │  CellDTO, TableDTO,        │
                          │  ChartDTO, BlockDTO, ...   │
                          └─────────────┬──────────────┘
                                        │
             ┌──────────────────────────┼────────────────────────────┐
             ▼                          ▼                            ▼
   ┌─────────────────┐        ┌──────────────────┐        ┌────────────────────┐
   │ formula/        │        │ analysis/        │        │ charts/            │
   │ lexer + parser, │        │ dependency graph │        │ OOXML chart extr.  │
   │ token refs,     │        │ impact + cycles  │        │ series + axes      │
   │ cross-sheet,    │        └──────────────────┘        └────────────────────┘
   │ table refs,     │                  │                           │
   │ array CSE       │                  ▼                           │
   └────────┬────────┘        ┌──────────────────┐                  │
            │                 │ annotation/      │◀─────────────────┘
            │                 │ semantic roles,  │
            │                 │ KPIs, block type │
            │                 └────────┬─────────┘
            │                          │
            ▼                          ▼
   ┌─────────────────┐        ┌──────────────────┐
   │ chunking/       │        │ rendering/       │
   │ segmenter that  │        │ HTML + pipe-text │
   │ splits sheets   │        │ preserving       │
   │ into logical    │        │ colspan/rowspan  │
   │ blocks          │        └────────┬─────────┘
   └────────┬────────┘                 │
            │                          │
            └──────────────┬───────────┘
                           ▼
                 ┌──────────────────┐          ┌────────────────────┐
                 │ storage/         │          │ verification/      │
                 │ DB records,      │          │ stage-by-stage     │
                 │ vector entries,  │          │ assertions + diff  │
                 │ to_json(), to_db │          │ reports            │
                 └──────────────────┘          └────────────────────┘
                           │
                           ▼
                 ┌──────────────────┐          ┌────────────────────┐
                 │ comparison/      │          │ export/            │
                 │ cross-workbook   │          │ generate Python    │
                 │ templates + DOF  │          │ importer classes   │
                 └──────────────────┘          └────────────────────┘
```

### Pipeline stages (`pipeline.py`)

1. **Parse** — `parsers/` pulls OOXML through openpyxl + targeted lxml for the
   parts openpyxl loses (chart refs, dynamic arrays, some validation edge
   cases). Output is a typed `WorkbookDTO`.
2. **Analyse** — `formula/` tokenises every expression; `analysis/` assembles
   them into a directed dependency graph, detects cycles, and resolves
   cross-sheet / table / external references.
3. **Annotate** — `annotation/` tags blocks with semantic roles (`HEADER`,
   `DATA`, `TOTAL`, `KPI`, …) and extracts workbook-level KPIs.
4. **Segment** — `chunking/` splits each sheet into logical blocks using
   adaptive gap analysis + style boundaries (handles vertical, horizontal,
   and mixed multi-table layouts).
5. **Render** — `rendering/` emits HTML (with faithful colspan/rowspan) and
   pipe-delimited text per block, with token counts.
6. **Serialize** — `storage/` produces JSON, DB-ready records, and
   vector-store entries addressable by source URI.
7. **Verify** — `verification/` runs stage-level assertions so regressions
   show up as structured diffs, not silent failures.
8. **Compare / Export** (optional) — `comparison/` aligns multiple workbooks
   of the same template and `export/` turns that alignment into a reusable
   Python importer class.

### Public API surface

```python
from xlsx_parser import (
    parse_workbook,      # 1 file  → ParseResult
    compare_workbooks,   # N files → GeneralizedTemplate
    export_importer,     # template → generated Python class
    ParseResult,
    StageVerifier,       # run individual stages for debugging
    VerificationReport,
    ExcellentStage,
    __version__,
)
```

The package is type-annotated end-to-end (`py.typed` is shipped).

---

## Installation

Requires Python 3.10+.

```bash
# Core library
pip install ks-xlsx-parser

# With FastAPI web server
pip install ks-xlsx-parser[api]

# With development/test tools
pip install ks-xlsx-parser[dev]
```

### From source

```bash
git clone https://github.com/arnav2/XLSXParser.git
cd XLSXParser
make install           # pip install -e ".[dev,api]"
make test              # run the default test suite
make testbench-build   # generate the 1000-file stress corpus
make testbench         # round-trip every workbook through the parser
```

### Dependencies

| Package | Purpose |
|---------|---------|
| `openpyxl>=3.1.0` | Excel file reading and cell extraction |
| `pydantic>=2.0`   | Data validation and serialization |
| `lxml>=4.9.0`     | Fast OOXML/XML parsing |
| `xxhash>=3.0.0`   | Deterministic content hashing |
| `tiktoken>=0.5.0` | Token counting for RAG chunking |

---

## Quick start

### Parse a workbook

```python
from xlsx_parser import parse_workbook

result = parse_workbook(path="workbook.xlsx")

print(f"Sheets:   {result.workbook.total_sheets}")
print(f"Cells:    {result.workbook.total_cells}")
print(f"Formulas: {result.workbook.total_formulas}")
print(f"Parsed in {result.workbook.parse_duration_ms:.0f} ms")
```

### Iterate RAG chunks

```python
for chunk in result.chunks:
    print(f"[{chunk.block_type}] {chunk.source_uri} ({chunk.token_count} tokens)")
    print(chunk.render_text[:200])
```

### Walk the formula dependency graph

```python
from xlsx_parser.models import CellCoord

for edge in result.workbook.dependency_graph.get_upstream(
    "Sheet1", CellCoord(row=10, col=3), max_depth=3
):
    print(f"{edge.source_sheet}!{edge.source_coord.to_a1()} → {edge.target_ref_string}")
```

### Serialise for a DB or vector store

```python
as_dict  = result.to_json()                       # fully JSON-compatible dict
records  = result.serializer.to_workbook_record() # DB row
sheets   = result.serializer.to_sheet_records()
chunks   = result.serializer.to_chunk_records()
vectors  = result.serializer.to_vector_store_entries()
```

### Parse from bytes

```python
with open("workbook.xlsx", "rb") as f:
    content = f.read()
result = parse_workbook(content=content, filename="workbook.xlsx")
```

---

## API reference

### `parse_workbook()`

```python
def parse_workbook(
    path: str | Path | None = None,
    content: bytes | None = None,
    filename: str | None = None,
    max_cells_per_sheet: int = 2_000_000,
) -> ParseResult: ...
```

Returns a `ParseResult` with `.workbook`, `.chunks`, and `.serializer`.

### `compare_workbooks()`

Align multiple workbooks of the same template to find structural similarities
and degrees of freedom.

```python
from xlsx_parser import compare_workbooks

template = compare_workbooks(["q1.xlsx", "q2.xlsx", "q3.xlsx"], dof_threshold=50)
```

### `export_importer()`

Generate a reusable Python importer class from a generalised template.

```python
from xlsx_parser import export_importer

export_importer(template, "quarterly_importer.py", class_name="QuarterlyReportImporter")
```

---

## Web API

XLSXParser ships with an optional FastAPI application (drag-and-drop UI
included).

```bash
pip install ks-xlsx-parser[api]
xlsx-parser-api                          # starts on http://localhost:8080
# or:
uvicorn xlsx_parser.api:app --reload --port 8080
```

Open the UI in a browser, or POST a file:

```bash
curl -X POST http://localhost:8080/parse -F "file=@workbook.xlsx"
```

The response includes the full parse result plus a verification report.

---

## The testBench dataset

A companion **1053-workbook stress corpus** is shipped under
[`testBench/`](testBench/):

| Group | Files | What it covers |
|-------|------:|----------------|
| `real_world/`           | 8    | Real anonymised workbooks (financial, engineering, project tracking) |
| `enterprise/`           | 4    | Deterministic enterprise templates |
| `github_datasets/`      | 10   | Public datasets (iris, titanic, superstore, …) |
| `stress/curated/`       | 26   | 26 progressive stress levels authored by hand |
| `stress/merges/`        | 5    | Pathological merge patterns |
| `generated/matrix/`     | 297  | One feature per file across 18 categories |
| `generated/combo/`      | 400  | Deterministic feature cocktails (5 densities × 80 seeds) |
| `generated/adversarial/`| 300  | 1M-row sheets, 250-sheet workbooks, unicode bombs, circular refs, 32 k-char cells, deep formula chains |

The `generated/` tree is produced deterministically by
[`scripts/build_testbench.py`](scripts/build_testbench.py). Every parser
regression becomes a new entry in `metrics/testbench/failures.json`, so the
whole bench is a fast, diffable acceptance gate.

```bash
make testbench-build   # regenerate testBench/generated/ (~1 minute)
make testbench         # parse every file, record failures
make testbench-zip     # package as dist/testBench-vX.Y.Z.zip for GitHub release
```

The zipped dataset is attached to every release as
`testBench-v<version>.zip`. Pull it from the Releases page if you don't want
to clone the full repo.

---

## Data models

All DTOs are Pydantic v2.

| Model | Description |
|-------|-------------|
| `WorkbookDTO`     | Root: sheets, tables, charts, named ranges, dependency graph, errors |
| `SheetDTO`        | Cells, merged regions, conditional formatting, data validation |
| `CellDTO`         | Value, formula, style, coordinates, annotations |
| `TableDTO`        | Excel ListObject table with columns, range, style |
| `ChartDTO`        | Chart metadata, series data, axis labels, chart type |
| `BlockDTO`        | Logical block (`HEADER` / `DATA` / `TABLE` / …) with bounding box + hash |
| `ChunkDTO`        | RAG chunk: HTML + text rendering, token count, source URI, content hash |
| `DependencyGraph` | Directed graph of formula dependencies with traversal helpers |
| `TableStructure`  | Assembled table with header / data regions |
| `TreeNode`        | Hierarchical node from tree building |
| `TemplateNode`    | Template node with degree-of-freedom annotations |

---

## Limitations

- **`.xls` not supported** — only `.xlsx` and `.xlsm` (OOXML). Convert legacy files externally.
- **Pivot tables** — detected but not fully parsed.
- **Sparklines** — not extracted.
- **VBA macros** — flagged but never executed or analysed.
- **External links** — recorded but not resolved.
- **Threaded comments** — only legacy comments are supported (openpyxl limitation).
- **Embedded OLE objects** — detected but not extracted.
- **Locale-dependent number formats** — not interpreted.

See [`docs/PARSER_KNOWN_ISSUES.md`](docs/PARSER_KNOWN_ISSUES.md) for edge cases.

---

## Contributing

We love contributions. Three paths, in order of speed-to-merge:

1. **Report a testBench failure** — run `make testbench`, find a file that
   breaks, attach it to a
   [Parser edge case issue](https://github.com/arnav2/XLSXParser/issues/new?template=parser_edge_case.yml).
2. **Add a new adversarial workbook** — contribute a builder to
   `scripts/build_testbench.py`. Any file that makes the parser crash or
   lose information is welcome.
3. **Fix a flagged issue** — see [`docs/PARSER_KNOWN_ISSUES.md`](docs/PARSER_KNOWN_ISSUES.md).

Full dev loop, PR checklist, and code style in [`CONTRIBUTING.md`](CONTRIBUTING.md).
See the [Code of Conduct](CODE_OF_CONDUCT.md) and
[Security policy](SECURITY.md) before posting.

If you don't have time to contribute but the project helped you, please
**[star the repo](https://github.com/arnav2/XLSXParser)**. That's the main
signal that keeps this maintained.

---

## License

[MIT](LICENSE). Use it, fork it, ship it.
