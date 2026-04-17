# ks-xlsx-parser v0.1.1 — Make XLSX LLM Ready 🚀

**First public release** of `ks-xlsx-parser`, an open-source (MIT) ETL layer
that turns `.xlsx` workbooks into structured, citation-ready JSON your agents
and RAG pipelines can actually reason about.

Previously shipped to production as part of the [Knowledge Stack](https://github.com/knowledgestack)
ecosystem. Now open for the rest of the world.

## Highlights

- 🧠 **LLM-ready chunks with citations** — every `ChunkDTO` carries a
  `source_uri` like `q4_forecast.xlsx#Revenue!A1:F18`, a `token_count`,
  HTML + pipe-text renderings, and a deterministic xxhash64 content hash.
- 📊 **Full workbook graph** — cells, formulas, merged regions, tables
  (ListObjects), charts (bar / line / pie / scatter / area / radar / bubble),
  conditional formatting (all Excel rule types), data validation, named
  ranges, and a directed dependency graph with cycle detection.
- 🧪 **1054-workbook testBench** — one-feature-per-file matrix (297) + combo
  cocktails (400) + adversarial files (300) + 57 real-world and
  curated-stress workbooks. Round-trip gate in CI, **1054/1054 passing in
  ~70 s**. Ship fixtures in the
  [`testBench-v0.1.1.zip`](https://github.com/knowledgestack/ks-xlsx-parser/releases/tag/v0.1.1)
  asset attached to this release.
- ⚡ **Parser perf fixes** — real-world workbooks that used to hang now
  finish in under a second.
  - Cached `detect_circular_refs()` per workbook: Walbridge Coatings
    **307 s → 4.6 s (66×)**.
  - Sparse-cell iteration: files with two non-empty cells at `A1` and
    `XFD1048576` drop from 60 s timeout → **135 ms**.
- 🧰 **Framework-agnostic** — drops straight into
  [LangChain](https://www.langchain.com/),
  [LangGraph](https://langchain-ai.github.io/langgraph/),
  [CrewAI](https://www.crewai.com/),
  [OpenAI Agents SDK](https://github.com/openai/openai-agents-python), or any
  [MCP](https://modelcontextprotocol.io/)-aware client.
- 🔐 **Security-first** — no macro execution, no external-link resolution,
  ZIP-bomb protection, input size limits, private vulnerability disclosure
  flow (see `SECURITY.md`).

## 30-second demo

```bash
pip install ks-xlsx-parser
```

```python
from ks_xlsx_parser import parse_workbook

result = parse_workbook(path="q4_forecast.xlsx")

for chunk in result.chunks:
    print(chunk.source_uri)          # q4_forecast.xlsx#Revenue!A1:F18
    print(chunk.token_count)         # 412
    print(chunk.render_text[:200])   # pipe-delimited, LLM-friendly
```

## Install

```bash
pip install ks-xlsx-parser           # core library
pip install ks-xlsx-parser[api]      # + FastAPI web server
pip install ks-xlsx-parser[dev]      # + test tooling
```

Python 3.10+, tested on Ubuntu and macOS.

## Artifacts attached

- `ks_xlsx_parser-0.1.1-py3-none-any.whl` — wheel, published to
  [PyPI](https://pypi.org/project/ks-xlsx-parser/)
- `ks_xlsx_parser-0.1.1.tar.gz` — sdist
- `testBench-v0.1.1.zip` — 1053-workbook stress corpus (17 MB). Drop into
  any parser for a stiff regression test.

## Community

- 💬 **Discord**: <https://discord.gg/4uaGhJcx>
- 🗣 **Discussions**: <https://github.com/knowledgestack/ks-xlsx-parser/discussions>
- 🐞 **Issues**: <https://github.com/knowledgestack/ks-xlsx-parser/issues>
- ⭐ **Star the repo**: <https://github.com/knowledgestack/ks-xlsx-parser>
- 🧰 **Knowledge Stack org**: <https://github.com/knowledgestack>

## What's next

- PDF / DOCX / PPTX parsers with the same citation model.
- MCP server wrapping the parser so Claude Desktop, Cursor, Windsurf, and
  Zed can call it without any glue code.
- Threaded-comment support (openpyxl limitation — planned via raw OOXML).
- Pivot-table parsing.
- Incremental re-parse when only part of a workbook changes.

Bug reports, edge-case workbooks, and PRs welcome — especially `.xlsx`
files that break the parser. See
[`CONTRIBUTING.md`](https://github.com/knowledgestack/ks-xlsx-parser/blob/main/CONTRIBUTING.md).

**Thanks to every team that filed an edge case during the private beta.**

— The Knowledge Stack team
