# Make XLSX LLM Ready

*How we built a citation-grade Excel parser for agents — and open-sourced it.*

**Subtitle (alt):** *Every enterprise RAG pipeline eventually hits a spreadsheet. Here's the ETL step that turns `.xlsx` into a graph your LLM can cite.*

**Tags:** `llm`, `rag`, `python`, `open-source`, `excel`, `agents`, `ai-engineering`

**Cover image:** the hero screenshot from `assets/hero-highlight.png` (financial model on the left, parser output on the right).

---

## TL;DR

- We open-sourced [**`ks-xlsx-parser`**](https://github.com/knowledgestack/ks-xlsx-parser), an MIT-licensed library that turns `.xlsx` workbooks into **citation-ready JSON** for LLM agents and RAG pipelines.
- Every output chunk carries a `source_uri` like `file.xlsx#Sheet!A1:F18`, a token count, an HTML + pipe-text rendering, and a deterministic content hash.
- It preserves what every other parser drops on the floor: **formulas, merges, charts, conditional formatting, data validation, and a directed dependency graph** with cycle detection.
- Ships with a **1054-workbook stress corpus** that runs in CI on every commit (**1054/1054 passing, ~70 s**).
- `pip install ks-xlsx-parser`.

---

## 1. Every enterprise agent eventually meets a spreadsheet

You build an agent. It does chain-of-thought, tool-calls, streaming — the whole song and dance. Then you deploy it against a real enterprise customer and the first message is:

> *"Can it read our quarterly forecast? The Excel one."*

And your options collapse:

1. **Dump the `.xlsx` as text.** You lose structure, provenance, and context-window discipline in one shot. The LLM sees a wall of characters and starts confabulating.
2. **Convert to CSV.** You throw away every formula, every merged header, every chart, and every multi-table sheet in the workbook. The CFO notices.
3. **Pandas / openpyxl.** Great if you already know which sheet and which rectangle of cells to look at. Not great if the agent is supposed to *figure that out*.
4. **Docling / markdown extractors.** Designed for documents, not for spreadsheets. They'll give you text; they won't give you a graph.
5. **Headless Excel.** Starts at $0 and ends in COM automation nightmares.

None of those give the LLM a **graph it can cite**. And citation is the load-bearing requirement: once compliance/finance/legal is in the loop, "the answer is somewhere in `Q4_forecast.xlsx`" is not an acceptable output.

That's the gap we built `ks-xlsx-parser` to fill.

---

## 2. What does "LLM-ready" actually mean?

We picked seven concrete outputs per parse. These are the ones that matter once you're wiring spreadsheets into an agent:

| Output | Why the agent cares |
|---|---|
| **Typed cell graph** (values, formulas, styles, coordinates) | Round-trips to JSON / DB / vector store without dropping formulas or types |
| **Formula AST + directed dependency graph** | Answer "what drives Q4 revenue?" with one upstream walk |
| **Tables, merges, layout blocks** | Multi-table sheets stop collapsing into one giant CSV |
| **Chart extractions** (7 types) | Text summaries the model can read and cite |
| **Token-counted render chunks** (HTML + pipe-text) | Plug into an embedding pipeline without blowing context |
| **Citation URIs** (`sheet!A1:B10`) | The LLM can point back at the exact cell |
| **Deterministic content hashes** (xxhash64) | Dedupe across versions, detect change between uploads |

That's the graph. Once you have it, "load this workbook" and "cite this cell" are two one-liner agent tools:

```python
def load_spreadsheet(path: str) -> list[dict]:
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
```

Drop that into LangChain, LangGraph, CrewAI, or the OpenAI Agents SDK as-is.

---

## 3. The 8-stage pipeline

The whole library is a linear pipeline. Each stage is an independent module you can unit-test in isolation.

```
.xlsx bytes
   │
   ▼
1. Parse         openpyxl + targeted lxml for what openpyxl loses
2. Analyse       tokenise formulas, build a directed dependency graph
3. Annotate      semantic roles (HEADER / DATA / TOTAL / KPI), KPI catalog
4. Segment       adaptive gap analysis → logical blocks
5. Render        HTML (colspan/rowspan) + pipe-text, token counts
6. Serialise     JSON, DB rows, vector-store entries
7. Verify        stage-level assertions so regressions surface as diffs
8. Compare       align N workbooks of the same template (optional)
```

The one decision that most shapes the downstream story is **where citation happens**. We resisted the temptation to bolt "add a URL" on at serialisation time. Instead, every block carries its bounding range end-to-end, and the chunk serialiser just *formats* it. That means the dependency summary on a chunk references the same coordinates as the source URI, and a downstream diff between two parses is structurally comparable.

Small rule, big payoff.

---

## 4. Two performance wins worth sharing

Prepping the library for the public release, we hit two bottlenecks that are interesting on their own merits. Both are openpyxl-adjacent and probably relevant if you're building in the same space.

### Win #1 — cached circular-ref detection (307 s → 4.6 s on a real 21k-cell workbook)

`detect_circular_refs()` on the dependency graph is O(V+E) with DFS + memoisation. Fine. But our chunk builder was calling it **once per chunk** inside `_build_dependency_summary()`, because every chunk's `has_circular` flag needed the global cycle set.

On a small workbook: invisible. On a 13-sheet, 21k-cell real-world financial model (Walbridge Coatings, now our favourite regression fixture): **115 chunks × ~2.6 s each = 307 s of CPU.** The chunker was dominating the parse.

The fix is almost embarrassing:

```python
class ChunkBuilder:
    def __init__(self, workbook):
        self._workbook = workbook
        self._dep_graph = workbook.dependency_graph
        self._circular_refs_cache: set[str] | None = None

    def _circular_refs(self) -> set[str]:
        if self._circular_refs_cache is None:
            self._circular_refs_cache = self._dep_graph.detect_circular_refs()
        return self._circular_refs_cache
```

Cache at the builder level (workbook-scoped lifetime). Replace the per-chunk call. **307 s → 4.6 s.** 66×.

The moral: performance regressions that only show up on real-world data are the hardest to catch with synthetic tests. Having a real financial model in our `testBench/real_world/` fixture set was what surfaced it.

### Win #2 — stored-cell iteration for sparse sheets (60 s timeout → 135 ms)

Openpyxl's `ws.iter_rows()` walks the full bounding box of a worksheet. For a sheet whose only non-empty cells are `A1` and `XFD1048576` (a real adversarial fixture in our stress corpus), that's **~17 billion empty-cell visits**. Even if we skip them in 30 ns each, we're looking at days.

Openpyxl keeps a private `ws._cells` dict that maps `(row, col) → Cell`, containing only actually-stored cells. Using it turns the iteration from O(max_row × max_col) into O(stored_cells):

```python
stored_cells = getattr(self._ws, "_cells", None)
if isinstance(stored_cells, dict):
    # Merged cells aren't in _cells; materialise them separately.
    merge_keys = {(mr, mc) for (mr, mc), (_m, _rs, _cs) in merge_masters.items()}
    cell_iter = [
        self._ws.cell(row=r, column=c)
        for (r, c) in sorted(set(stored_cells.keys()) | merge_keys)
    ]
else:
    cell_iter = (cell for row in self._ws.iter_rows() for cell in row)
```

That's a private API, which is usually a red flag. But `_cells` has been stable across openpyxl versions for years, and the alternative is *hours* of runtime on edge-case inputs. We guard with `isinstance(stored_cells, dict)` so the fallback kicks in if openpyxl ever changes the shape.

**60 s pytest-timeout → 135 ms.**

---

## 5. The testBench

Writing parser code without a stress corpus is writing parser bugs. We ship 1054 workbooks under `testBench/` and round-trip every one on CI.

- **`real_world/`** (8) — anonymised real workbooks shipped as demos.
- **`enterprise/`** (4) — deterministic enterprise templates.
- **`github_datasets/`** (10) — iris, titanic, superstore, apple stock, world happiness…
- **`stress/curated/`** (26) — 26 hand-authored progressive stress levels.
- **`stress/merges/`** (5) — pathological merge patterns.
- **`generated/matrix/`** (297) — **one feature per file** across 18 categories (formulas, merges, named ranges, CF, data validation, tables, charts, styles, dates, errors, hidden rows/cols, hyperlinks, comments, rich text, freeze panes, edge addresses, sheet names, 3D refs).
- **`generated/combo/`** (400) — randomised feature cocktails at 5 densities × 80 seeds.
- **`generated/adversarial/`** (300) — files engineered to break parsers: unicode bombs, 32k-char cells, deep formula chains, 1M-row sparse sheets, 250-sheet workbooks, broken refs.

The `generated/` tree is built deterministically by a single Python script. Any time the parser regresses on any of those 1000 files, `metrics/testbench/failures.json` grows an entry with the stage, error type, and a 5-line traceback. That's the acceptance gate.

You can pull the full corpus without cloning — every release attaches a `testBench-vX.Y.Z.zip`.

---

## 6. One example, end-to-end

Here's what parsing a financial model looks like in practice.

**Input:** `q4_forecast.xlsx`, 13 sheets, 21k cells, multiple tables per sheet, charts, conditional formatting, named ranges.

```python
from ks_xlsx_parser import parse_workbook

result = parse_workbook(path="q4_forecast.xlsx")

print(f"Sheets:   {result.workbook.total_sheets}")
print(f"Cells:    {result.workbook.total_cells}")
print(f"Formulas: {result.workbook.total_formulas}")
```

**Iterate chunks with citations:**

```python
for chunk in result.chunks:
    print(f"[{chunk.block_type}] {chunk.source_uri} — {chunk.token_count} tok")
    print(chunk.render_text[:200])
```

Example output:

```
[TABLE]  q4_forecast.xlsx#Assumptions!A1:D11 — 287 tok
Parameter | Value   | Unit | Notes
Revenue Growth Rate | 8.0% | % | Year-over-year
COGS Margin         | 45.0% | % | % of revenue
...

[DATA]   q4_forecast.xlsx#Revenue!A3:F12 — 612 tok
...
```

**Walk the dependency graph:**

```python
from ks_xlsx_parser.models import CellCoord

upstream = result.workbook.dependency_graph.get_upstream(
    "Revenue", CellCoord(row=10, col=3), max_depth=3
)
for edge in upstream:
    print(f"  {edge.source_sheet}!{edge.source_coord.to_a1()} → {edge.target_ref_string}")
```

That's a five-line "explain what drives this cell" feature.

**Dump for a vector store:**

```python
vectors = result.serializer.to_vector_store_entries()
# List[{"id": str, "text": str, "metadata": {...}}]
# Drop straight into Qdrant / pgvector / Weaviate / Pinecone.
```

---

## 7. Safety notes

Spreadsheets are a great attack surface. We're explicit:

- **No macro execution.** VBA is flagged and never run. We never open the workbook through COM.
- **No external-link resolution.** We record what the workbook claims, we don't follow anything.
- **ZIP-bomb protection.** Incoming bytes are size-checked before openpyxl sees them.
- **Cell-count ceiling.** Per-sheet `max_cells_per_sheet` (default 2M) truncates rather than OOMs.

You can safely point `ks-xlsx-parser` at untrusted uploads. We do.

---

## 8. Where this fits in the ecosystem

`ks-xlsx-parser` is the first library in the [**Knowledge Stack**](https://github.com/knowledgestack) open-source family — document intelligence for agents, so engineering teams can focus on agents and we handle the messy parts of enterprise data.

- [**ks-cookbook**](https://github.com/knowledgestack/ks-cookbook) — 32 production-style flagship agents + recipes for LangChain, LangGraph, CrewAI, Temporal, and the OpenAI Agents SDK.
- [**ks-xlsx-parser**](https://github.com/knowledgestack/ks-xlsx-parser) — this library.
- Next up: PDF, DOCX, PPTX parsers with the same citation model, plus an MCP server so Claude Desktop / Cursor / Windsurf / Zed can call the parsers without glue code.

---

## 9. How to help

If you got this far, please:

1. ⭐ **[Star the repo](https://github.com/knowledgestack/ks-xlsx-parser).** It's the single biggest signal that keeps maintainers paid.
2. 💬 **[Join the Discord](https://discord.gg/4uaGhJcx).** We hang out there. Ask questions, float ideas, show off what you built.
3. 🧪 **Run `make testbench` and send us a workbook that breaks it.** Every edge-case report becomes a fixture in the next release. There's even a [Parser edge case issue template](https://github.com/knowledgestack/ks-xlsx-parser/issues/new?template=parser_edge_case.yml) specifically for this.

We'll ship more parsers. We'll ship an MCP server. We'll ship a native agent runtime that knows how to ground its citations. But none of that matters if nobody's telling us which `.xlsx` files break the parser first.

Drop by Discord. Tell us what you're building.

---

*`ks-xlsx-parser` is MIT-licensed. Use it, fork it, ship it. If you build something on top of it, we'd love a [Show & Tell](https://github.com/knowledgestack/ks-xlsx-parser/discussions/new?category=show-and-tell) — or a shoutout in Discord.*

**`pip install ks-xlsx-parser`** · [GitHub](https://github.com/knowledgestack/ks-xlsx-parser) · [Discord](https://discord.gg/4uaGhJcx) · [Knowledge Stack](https://github.com/knowledgestack)
