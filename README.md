# Make XLSX LLM Ready

**`ks-xlsx-parser` is the missing ETL step between your spreadsheets and your LLM.**

[![PyPI version](https://img.shields.io/pypi/v/ks-xlsx-parser.svg)](https://pypi.org/project/ks-xlsx-parser/)
[![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Tests](https://img.shields.io/badge/testBench-1054%2F1054-success.svg)](#the-testbench-dataset)
[![CI](https://github.com/knowledgestack/ks-xlsx-parser/actions/workflows/ci.yml/badge.svg)](https://github.com/knowledgestack/ks-xlsx-parser/actions/workflows/ci.yml)

> `.xlsx` → structured, typed, citation-ready JSON that an LLM can actually reason about.
> Cells, formulas, merged regions, tables, charts, conditional formatting,
> dependency graphs, and RAG-ready chunks — deterministic, fully tested, MIT.

Spreadsheets are still the #1 unstructured data source in the enterprise.
Feeding a `.xlsx` directly to an LLM loses structure (rows, formulas, merges),
loses provenance (which cell said what), and blows through context windows.
`ks-xlsx-parser` turns an Excel workbook into a token-counted, source-addressable
graph that drops straight into LangChain, LangGraph, CrewAI, or any MCP-aware
agent.

---

## ⭐ If this helps you

This project is free, open source (MIT), and part of the [Knowledge Stack](https://github.com/knowledgestack)
ecosystem. The single most helpful thing you can do is **[star the repo](https://github.com/knowledgestack/ks-xlsx-parser)** —
that's how we justify spending more time on it. 👍

**Get involved:**

- 💬 [Discussions](https://github.com/knowledgestack/ks-xlsx-parser/discussions) — ask questions, share what you built, or float an idea
- 🐞 [Issues](https://github.com/knowledgestack/ks-xlsx-parser/issues/new/choose) — report a bug, request a feature, or file a parser edge case
- 🎯 [Show & Tell](https://github.com/knowledgestack/ks-xlsx-parser/discussions/new?category=show-and-tell) — tell us about your production use
- 🔐 [Security](https://github.com/knowledgestack/ks-xlsx-parser/security/advisories/new) — report a vulnerability privately
- 🙌 [Contribute](CONTRIBUTING.md) — every PR is reviewed, `good-first-issue` labels live on Issues

Not sure where to start? Run `make testbench`, find a file that breaks, open a
[Parser edge case](https://github.com/knowledgestack/ks-xlsx-parser/issues/new?template=parser_edge_case.yml).
That's the fastest path to a merged PR.

---

## 30-second demo

```bash
pip install ks-xlsx-parser
```

```python
from ks_xlsx_parser import parse_workbook

result = parse_workbook(path="q4_forecast.xlsx")

# LLM-ready chunks with citation URIs
for chunk in result.chunks:
    print(chunk.source_uri)          # q4_forecast.xlsx#Revenue!A1:F18
    print(chunk.token_count)         # 412
    print(chunk.render_text[:200])   # Pipe-delimited Markdown-ish text
    print(chunk.render_html[:200])   # HTML with proper colspan/rowspan

# Or dump the whole workbook graph
import json
json.dump(result.to_json(), open("workbook.json", "w"), default=str)
```

That's it. Every chunk has:
- `source_uri` — cite back to exact cells
- `render_text` / `render_html` — LLM-consumable bodies
- `token_count` — cap your context window properly
- `dependency_summary` — upstream/downstream formulas
- content hash — dedupe across versions

---

## Table of Contents

- [Why a dedicated XLSX parser for LLMs?](#why-a-dedicated-xlsx-parser-for-llms)
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

## Why a dedicated XLSX parser for LLMs?

Most Excel libraries answer one of two questions well: *"read a rectangle of
values"* (pandas, openpyxl) or *"run Excel headless"* (xlwings, LibreOffice).
`ks-xlsx-parser` answers a third one: **"give me a structured, inspectable,
loss-minimising graph that an LLM or auditor can reason about."**

| Output | Why an LLM cares |
|--------|------------------|
| Typed cell graph (values, formulas, styles, coordinates) | Round-trips to JSON/DB/vector store without losing formulas or data types |
| Formula AST + directed dependency graph | Answer "what drives Q4 revenue?" via upstream traversal |
| Detected tables, merged regions, layout blocks | Multi-table sheets no longer collapse into one giant CSV |
| Chart extractions (bar / line / pie / scatter / area / radar / bubble) | Text summaries the model can read |
| Token-counted render chunks (HTML + pipe-text) | Plug straight into an embedding pipeline without blowing context |
| Citation-ready source URIs (`sheet!A1:B10`) | The LLM can cite the exact cell it's talking about |
| Deterministic content hashes (xxhash64) | Dedupe across versions, detect change between uploads |

Everything is deterministic, everything is tested on a 1054-workbook stress
corpus, and everything is open source.

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
                 │ to_json(),       │          │ reports            │
                 │ LLM chunks       │          └────────────────────┘
                 └──────────────────┘
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
6. **Serialise** — `storage/` produces JSON, DB-ready records, and vector-store
   entries addressable by source URI.
7. **Verify** — `verification/` runs stage-level assertions so regressions
   show up as structured diffs, not silent failures.
8. **Compare / Export** (optional) — `comparison/` aligns multiple workbooks
   of the same template and `export/` turns that alignment into a reusable
   Python importer class.

### Public API surface

```python
from ks_xlsx_parser import (
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

The package is fully type-annotated (`py.typed` is shipped).

> **Note**: the importable module is `xlsx_parser`; the PyPI package is
> `ks-xlsx-parser`. Both names above work — `ks_xlsx_parser` is re-exported
> as a convenience alias matching the package name.

---

## Installation

Requires Python 3.10+.

```bash
# Core library
pip install ks-xlsx-parser

# With the FastAPI web server
pip install ks-xlsx-parser[api]

# Dev / test tools
pip install ks-xlsx-parser[dev]
```

### From source

```bash
git clone https://github.com/knowledgestack/ks-xlsx-parser.git
cd ks-xlsx-parser
make install           # pip install -e ".[dev,api]"
make test              # default suite
make testbench-build   # generate the 1000-file stress corpus
make testbench         # round-trip every workbook through the parser
```

### Dependencies

| Package | Purpose |
|---------|---------|
| `openpyxl>=3.1.0` | Excel file reading and cell extraction |
| `pydantic>=2.0`   | Data validation and serialisation |
| `lxml>=4.9.0`     | Fast OOXML/XML parsing |
| `xxhash>=3.0.0`   | Deterministic content hashing |
| `tiktoken>=0.5.0` | Token counting for LLM context management |

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

### LLM chunks with citations

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
as_dict = result.to_json()                             # fully JSON-compatible dict
records = result.serializer.to_workbook_record()       # DB row
sheets = result.serializer.to_sheet_records()
chunks = result.serializer.to_chunk_records()
vectors = result.serializer.to_vector_store_entries()  # ready for Qdrant / pgvector / Weaviate
```

### Parse from bytes (typical server path)

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

Align multiple workbooks that share a template to find structural similarities
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

`ks-xlsx-parser` ships with an optional FastAPI application with a drag-and-drop UI.

```bash
pip install ks-xlsx-parser[api]
xlsx-parser-api                          # starts on http://localhost:8080
# or:
uvicorn xlsx_parser.api:app --reload --port 8080
```

POST a file:

```bash
curl -X POST http://localhost:8080/parse -F "file=@workbook.xlsx"
```

The response includes the full parse result plus a verification report.

---

## The testBench dataset

A companion **1054-workbook stress corpus** is shipped under
[`testBench/`](testBench/):

| Group | Files | What it covers |
|-------|------:|----------------|
| `real_world/`            | 8    | Real anonymised workbooks (financial, engineering, project tracking) |
| `enterprise/`            | 4    | Deterministic enterprise templates |
| `github_datasets/`       | 10   | Public datasets (iris, titanic, superstore, …) |
| `stress/curated/`        | 26   | 26 progressive stress levels authored by hand |
| `stress/merges/`         | 5    | Pathological merge patterns |
| `generated/matrix/`      | 297  | One feature per file across 18 categories |
| `generated/combo/`       | 400  | Deterministic feature cocktails (5 densities × 80 seeds) |
| `generated/adversarial/` | 300  | Unicode bombs, circular refs, 32k-char cells, deep formula chains, sparse 1M-row sheets, 250-sheet workbooks |

The `generated/` tree is produced deterministically by
[`scripts/build_testbench.py`](scripts/build_testbench.py). Every parser
regression becomes a new entry in `metrics/testbench/failures.json`, so the
whole bench is a fast, diffable acceptance gate.

```bash
make testbench-build   # regenerate testBench/generated/ (~1 minute)
make testbench         # 1054/1054 in ~70 seconds
make testbench-zip     # package as dist/testBench-vX.Y.Z.zip for a GitHub release
```

The zipped dataset is attached to every release. Pull it from the
[Releases page](https://github.com/knowledgestack/ks-xlsx-parser/releases)
if you don't want to clone the full repo.

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
| `ChunkDTO`        | LLM chunk: HTML + text rendering, token count, source URI, content hash |
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
   [Parser edge case issue](https://github.com/knowledgestack/ks-xlsx-parser/issues/new?template=parser_edge_case.yml).
2. **Add a new adversarial workbook** — contribute a builder to
   `scripts/build_testbench.py`. Any file that makes the parser crash or
   lose information is welcome.
3. **Fix a flagged issue** — see [`docs/PARSER_KNOWN_ISSUES.md`](docs/PARSER_KNOWN_ISSUES.md).

Full dev loop, PR checklist, and code style in [`CONTRIBUTING.md`](CONTRIBUTING.md).
See the [Code of Conduct](CODE_OF_CONDUCT.md) and
[Security policy](SECURITY.md) before posting.

If you don't have time to contribute but the project helped you, please
**[star the repo](https://github.com/knowledgestack/ks-xlsx-parser)**. That's
the main signal that keeps this maintained.

---

## Ecosystem

`ks-xlsx-parser` is part of the [Knowledge Stack](https://github.com/knowledgestack)
open-source family:

- [**ks-cookbook**](https://github.com/knowledgestack/ks-cookbook) — 32
  production-style flagship agents + recipes for LangChain, LangGraph, CrewAI,
  Temporal, and the OpenAI Agents SDK.
- **ks-xlsx-parser** (this repo) — turn `.xlsx` into LLM-ready JSON.

---

## License

[MIT](LICENSE). Use it, fork it, ship it.
