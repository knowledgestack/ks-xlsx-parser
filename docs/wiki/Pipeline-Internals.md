# Pipeline Internals

The 8 stages that turn `.xlsx` bytes into LLM-ready chunks, in order, with
pointers to the code that implements each one. Read this if you're
extending the parser or hunting a regression.

## Stage map

```
.xlsx bytes
    │
    ▼
1. Parse         ── src/xlsx_parser/parsers/            openpyxl + lxml → WorkbookDTO
2. Analyse       ── src/xlsx_parser/formula/            tokenise, resolve refs
                   src/xlsx_parser/analysis/            build dependency graph
3. Annotate      ── src/xlsx_parser/annotation/         semantic roles, KPIs
4. Segment       ── src/xlsx_parser/chunking/           sheets → logical blocks
5. Render        ── src/xlsx_parser/rendering/          HTML + pipe-text
6. Serialise     ── src/xlsx_parser/storage/            to_json, DB rows, vectors
7. Verify        ── src/xlsx_parser/verification/       stage-level assertions
8. Compare/Export── src/xlsx_parser/comparison/         multi-workbook templates
                   src/xlsx_parser/export/              generated importer classes
```

The entry point is `src/xlsx_parser/pipeline.py`. Each stage is an
independent module you can unit-test in isolation.

## 1. Parse

`parsers/workbook_parser.py` loads the workbook twice through openpyxl —
once with `data_only=False` (formulas) and once with `data_only=True`
(computed values) — then hands each sheet to `SheetParser.parse()`.

Openpyxl loses a few things on load (some chart references, dynamic
array formulas, a few data-validation edge cases, values in
empty-master merged regions), which we recover by opening the `.xlsx`
as a ZIP and parsing the raw OOXML XML with `lxml` (see
`SheetParser._recover_empty_merge_masters()` for the canonical example).

**Perf note**: we iterate `ws._cells` (openpyxl's stored-cell dict)
rather than `ws.iter_rows()`, because the latter walks the full
bounding box — a single `XFD1048576` cell otherwise forces a ~17 B
empty-cell walk. See
[`CHANGELOG.md`](https://github.com/knowledgestack/ks-xlsx-parser/blob/main/CHANGELOG.md#performance).

## 2. Analyse

`formula/lexer.py` + `formula/parser.py` tokenise every formula and
resolve references (cell / range / cross-sheet / table / external).
`analysis/dependency_builder.py` consumes the token stream and emits
`DependencyEdge` objects, which go into `DependencyGraph`.

Circular-reference detection is O(V+E) DFS with memoisation at the
edge level. It's cached per workbook inside `ChunkBuilder` — running it
per chunk is how Walbridge Coatings used to take 307 s.

## 3. Annotate

`annotation/` tags blocks with semantic roles. The main output is
`BlockDTO.block_type` ∈ `{HEADER, DATA, TABLE, TOTAL, KPI,
CHART_ANCHOR, NOTES, ...}` and a workbook-level `kpi_catalog` surfacing
named KPIs. This is the easiest stage to extend — most domain-specific
customisation lives here.

## 4. Segment

`chunking/segmenter.py` splits each sheet into logical blocks using
adaptive gap analysis (blank rows/columns, sudden style boundaries,
explicit ListObjects, named ranges). Handles vertical, horizontal, and
mixed multi-table sheets.

`chunking/chunker.py` then turns each block into a `ChunkDTO` by calling
rendering + dependency-summary + token-counting + hashing. This is where
the LLM-ready output is assembled.

## 5. Render

- `rendering/html_renderer.py` — HTML with faithful colspan/rowspan so a
  browser renders the chunk like Excel would.
- `rendering/text_renderer.py` — pipe-delimited text. Headers are
  promoted; merged masters get their value repeated across the slaves
  in text form. Designed for LLM prompt assembly.

## 6. Serialise

`storage/serializer.py` exposes:

- `to_workbook_record()` — one DB row per workbook.
- `to_sheet_records()` — one per sheet.
- `to_chunk_records()` — one per chunk.
- `to_vector_store_entries()` — `id` + `text` + `metadata` triples.

`pipeline.ParseResult.to_json()` returns the full nested dict;
`json.dumps(..., default=str)` makes it JSON-safe.

## 7. Verify

`verification/stage_verifier.py` runs the same 11-stage Excellent
algorithm as an opt-in audit — load, parse, merge-resolve, formula,
graph, annotate, segment, render, chunk, serialise, kpi. Each stage
reports a `StageResult` with `ok`, `duration_ms`, and diagnostics.

Call `StageVerifier(path=...).run()` when you're debugging why a
specific file produces unexpected output. The FastAPI Web API returns
this alongside the parse result so users can diff behaviours.

## 8. Compare / Export (multi-workbook)

`comparison/` aligns two or more workbooks of the same template and
computes `GeneralizedTemplate` — what's fixed vs what's a
degree-of-freedom. `export/importer_generator.py` then turns that tree
into a generated Python class with one `import_one(path)` method.

Use case: you get quarterly reports from 50 subsidiaries, all loosely
the same shape but with subtle variations. Instead of writing one
importer by hand, you `compare_workbooks([...])` on a sample and the
parser writes the importer for you.

## Where to hook in

| You want to… | Edit this |
|---|---|
| Add a new chart type | `charts/chart_parser.py` |
| Support a new formula function (affects dependency traversal) | `formula/known_functions.py` |
| Tag blocks with a new semantic role | `annotation/semantic_tagger.py` |
| Split sheets differently | `chunking/segmenter.py` |
| Change how HTML / text is rendered | `rendering/*_renderer.py` |
| Add a new serialisation target (e.g. Arrow) | `storage/serializer.py` |
| Add a verification stage | `verification/stage_verifier.py` |
| Add a new DTO field | `models/*.py` (+ serializer + renderer) |

When in doubt, write the test first — the
[`testBench/`](https://github.com/knowledgestack/ks-xlsx-parser/tree/main/testBench)
corpus is the fastest signal that a pipeline change didn't regress
anything else.
