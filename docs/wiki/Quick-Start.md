# Quick Start

Everything you need to go from `pip install` to production-shape output in
five short snippets. Each one is runnable standalone against a real
`.xlsx` file.

## Install

```bash
pip install ks-xlsx-parser                 # core library
pip install "ks-xlsx-parser[api]"          # + FastAPI web server
pip install "ks-xlsx-parser[dev]"          # + test tooling
```

Python 3.10+, tested on Ubuntu and macOS.

## 1. Parse a workbook

```python
from ks_xlsx_parser import parse_workbook

result = parse_workbook(path="workbook.xlsx")

print(f"Sheets:   {result.workbook.total_sheets}")
print(f"Cells:    {result.workbook.total_cells}")
print(f"Formulas: {result.workbook.total_formulas}")
print(f"Parsed in {result.workbook.parse_duration_ms:.0f} ms")
```

`result` is a `ParseResult` with three properties:

- `result.workbook` — the full typed workbook graph (`WorkbookDTO`).
- `result.chunks` — LLM-ready chunks with citations and token counts.
- `result.serializer` — DB / vector-store record exporters.

## 2. Iterate LLM chunks with citations

```python
for chunk in result.chunks:
    print(f"[{chunk.block_type}] {chunk.source_uri} ({chunk.token_count} tokens)")
    print(chunk.render_text[:200])
    # chunk.render_html is also available for agents that render Markdown+HTML.
```

Each chunk carries:

- `source_uri` — `file.xlsx#Sheet!A1:F18`, ready to cite.
- `render_text` — pipe-delimited, LLM-friendly representation.
- `render_html` — HTML with faithful colspan/rowspan, for UI embedding.
- `token_count` — computed with `tiktoken`.
- `dependency_summary` — upstream/downstream formula refs within the block.
- `content_hash` — deterministic xxhash64, for dedup and change detection.

## 3. Walk the formula dependency graph

```python
from ks_xlsx_parser.models import CellCoord

upstream_edges = result.workbook.dependency_graph.get_upstream(
    sheet="Sheet1",
    coord=CellCoord(row=10, col=3),
    max_depth=3,
)
for edge in upstream_edges:
    print(f"{edge.source_sheet}!{edge.source_coord.to_a1()} → {edge.target_ref_string}")
```

`get_downstream()` works the same way. Both support `max_depth` and return
typed `DependencyEdge` objects with `edge_type` markers
(`WITHIN_SHEET`, `CROSS_SHEET`, `TABLE_REF`, `EXTERNAL`).

Cycle detection:

```python
circular = result.workbook.dependency_graph.detect_circular_refs()
# → set of "Sheet!A1" strings that participate in a cycle
```

## 4. Serialise for a DB or vector store

```python
import json

# Full JSON-compatible dict (cells, formulas, chunks, errors, verification)
as_dict = result.to_json()
with open("workbook.json", "w") as f:
    json.dump(as_dict, f, default=str)

# DB-ready records (one row per level of the hierarchy)
ser = result.serializer
workbook_row = ser.to_workbook_record()
sheet_rows = ser.to_sheet_records()
chunk_rows = ser.to_chunk_records()

# Vector-store entries (id + text + metadata)
vectors = ser.to_vector_store_entries()
# → ready to upsert into Qdrant / pgvector / Weaviate / Pinecone
```

## 5. Parse from bytes (typical server path)

```python
from ks_xlsx_parser import parse_workbook

with open("workbook.xlsx", "rb") as f:
    content = f.read()

result = parse_workbook(
    content=content,
    filename="workbook.xlsx",       # shown in source URIs / logs
    max_cells_per_sheet=2_000_000,  # safety limit; default shown
)
```

Passing `content` instead of `path` is the usual pattern for FastAPI / Flask
upload handlers, S3 object processors, or anywhere the `.xlsx` doesn't have
a filesystem home.

## Using it as an LLM tool

The two methods you'll most often wrap as agent tools:

```python
def load_spreadsheet(path: str) -> list[dict]:
    """Load a workbook and return LLM-ready chunks."""
    result = parse_workbook(path=path)
    return [
        {
            "source_uri": c.source_uri,
            "text": c.render_text,
            "tokens": c.token_count,
            "block_type": c.block_type,
        }
        for c in result.chunks
    ]


def cite_cell(path: str, sheet: str, a1: str) -> dict:
    """Fetch one cell with its full context (value, formula, upstream deps)."""
    from ks_xlsx_parser.models import CellCoord
    from openpyxl.utils import coordinate_to_tuple

    row, col = coordinate_to_tuple(a1)
    result = parse_workbook(path=path)
    cell = next(
        (c for s in result.workbook.sheets
         for c in s.cells.values()
         if s.sheet_name == sheet and c.row == row and c.col == col),
        None,
    )
    if cell is None:
        return {"error": f"{sheet}!{a1} not found"}
    deps = result.workbook.dependency_graph.get_upstream(
        sheet, CellCoord(row=row, col=col), max_depth=2
    )
    return {
        "source_uri": f"{path}#{sheet}!{a1}",
        "value": cell.value,
        "formula": cell.formula,
        "upstream": [e.target_ref_string for e in deps],
    }
```

Both are trivially wrappable with LangChain `@tool`, LangGraph `ToolNode`, or
the OpenAI Agents SDK `@function_tool`.

## Next

- [**API Reference**](API-Reference) — full signatures and options.
- [**Web API**](Web-API) — if you'd rather call the parser over HTTP.
- [**Data Models**](Data-Models) — the Pydantic DTOs you'll be reading.
