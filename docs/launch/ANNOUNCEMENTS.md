# Launch announcements — v0.1.1

Copy-paste these. Tweak the tone to match the channel.

---

## 🎮 Discord — `#announcements`

> **🚀 ks-xlsx-parser is now open source!**
>
> We just shipped **ks-xlsx-parser v0.1.1** — the Knowledge Stack ETL layer
> that turns `.xlsx` into LLM-ready JSON with proper citations, dependency
> graphs, and per-chunk token counts.
>
> **What you get:**
> • Cells, formulas, merges, charts, tables, conditional formatting, data
>   validation — all preserved, all citable back to a source cell.
> • RAG-ready chunks with `source_uri` + `render_text` + `token_count` that
>   drop straight into LangChain, LangGraph, CrewAI, the OpenAI Agents SDK,
>   or any MCP client.
> • A 1054-workbook stress corpus (`testBench/`) that we round-trip on every
>   CI run — 1054/1054 passing in ~70s.
>
> **Install:** `pip install ks-xlsx-parser`
> **Repo:** <https://github.com/knowledgestack/ks-xlsx-parser>
> **Release:** <https://github.com/knowledgestack/ks-xlsx-parser/releases/tag/v0.1.1>
>
> ⭐ Star the repo if this saves you time, and drop your edge-case workbooks
> in <#edge-cases> (or just DM us). Every `.xlsx` that breaks the parser
> becomes a new fixture in the next release.
>
> More parsers (PDF / DOCX / PPTX) + an MCP server are next. Follow the org
> to get pinged: <https://github.com/knowledgestack>

---

## 🎮 Discord — `#general` (shorter)

> We just open-sourced **ks-xlsx-parser** 🎉
> Turn `.xlsx` into LLM-ready JSON with citations + dependency graphs.
> `pip install ks-xlsx-parser` · <https://github.com/knowledgestack/ks-xlsx-parser>
> Break it, star it, hang out here 🙌

---

## 🐦 Twitter / X

> Spreadsheets are still the #1 unstructured data source in the enterprise.
>
> We just open-sourced **ks-xlsx-parser** — turns `.xlsx` into
> citation-ready JSON your agents can actually reason about.
>
> 1054 stress-test workbooks. 100% pass rate. MIT.
>
> `pip install ks-xlsx-parser`
> 🔗 github.com/knowledgestack/ks-xlsx-parser

Follow-up tweet (thread):

> Why a dedicated parser? pandas/openpyxl give you a dataframe. Docling
> gives you prose. Neither gives you the *graph* an LLM needs: formulas,
> merges, charts, dependency edges, and citation URIs per chunk.
>
> That's the gap ks-xlsx-parser fills.

---

## 💼 LinkedIn

> We just open-sourced ks-xlsx-parser — the Knowledge Stack ETL layer for
> turning Excel workbooks into LLM-ready, citation-grounded JSON.
>
> Built for teams shipping agents and RAG pipelines that ingest real
> enterprise data: finance models, legal schedules, insurance claims,
> construction takeoffs. The messy `.xlsx` files your competitors can't
> handle.
>
> Highlights:
> → Preserves formulas, merges, charts, CF, data validation
> → Per-chunk source URIs so the LLM can cite the exact cell
> → Directed dependency graph with cycle detection
> → 1054-workbook stress corpus that we round-trip on every CI run
>
> MIT licensed. Part of the Knowledge Stack ecosystem (https://github.com/knowledgestack).
>
> `pip install ks-xlsx-parser`
>
> ⭐ Star: https://github.com/knowledgestack/ks-xlsx-parser
> 💬 Discord: https://discord.gg/4uaGhJcx

---

## 🧡 Hacker News — "Show HN"

**Title:** `Show HN: ks-xlsx-parser – turn .xlsx into citation-ready JSON for LLMs`

**Body:**

> Hey HN — we open-sourced the Excel parser we've been using in production
> behind our agent platform. It turns `.xlsx` workbooks into a structured,
> loss-minimising graph that an LLM or auditor can reason about: typed
> cells, formulas with a directed dependency graph, merged regions,
> tables, charts, conditional formatting, data validation, named ranges,
> and RAG-ready chunks with per-cell source URIs for citation.
>
> Every chunk carries a `source_uri` (`file.xlsx#Sheet!A1:F18`), a token
> count (via tiktoken), and both an HTML and a pipe-delimited text
> rendering. It drops straight into LangChain / LangGraph / CrewAI / the
> OpenAI Agents SDK, or any MCP-aware client.
>
> The library itself is Python 3.10+ and MIT, built on openpyxl + targeted
> lxml for the parts openpyxl loses. We ship a 1054-workbook stress corpus
> (`testBench/`) alongside the library — feature-per-file matrix (297) +
> combinatoric cocktails (400) + adversarial files (300: unicode bombs,
> 32k-char cells, deep formula chains, sparse 1M-row sheets, 250-sheet
> workbooks). CI round-trips every file; 1054/1054 currently pass in ~70s.
>
> Two performance fixes we shipped while preparing for the public release
> might be interesting on their own:
>
> 1. Cached `detect_circular_refs()` per workbook — we were re-running it
>    per block. A real 21k-cell financial model went from 307s to 4.6s.
>
> 2. Sparse cell iteration via openpyxl's `_cells` dict. A workbook whose
>    only non-empty cells are `A1` and `XFD1048576` was iterating ~17B
>    empty cells before. Now 135ms.
>
> Install: `pip install ks-xlsx-parser`
>
> Repo: https://github.com/knowledgestack/ks-xlsx-parser
> Discord: https://discord.gg/4uaGhJcx
>
> Would love bug reports — especially `.xlsx` files that break it.

---

## 🧵 Reddit (r/MachineLearning, r/LangChain, r/Python)

**Title:** `[P] Open-sourced ks-xlsx-parser: turn .xlsx into citation-ready JSON for LLMs`

**Body:**

> Just open-sourced the Excel parser I've been using in production for my
> agent platform. It's the missing ETL step between `.xlsx` and an LLM —
> preserves formulas, merges, charts, conditional formatting, and data
> validation, and emits RAG-ready chunks with per-cell source URIs for
> citation.
>
> Highlights:
>
> - LLM-ready chunks with `source_uri`, token count, HTML + pipe-text
>   renderings, xxhash64 content hash for dedup.
> - Directed formula dependency graph with upstream / downstream traversal
>   and cycle detection.
> - 1054-workbook stress corpus (`testBench/`) round-tripped on every CI
>   run. 1054/1054 pass in ~70s.
> - Python 3.10+, MIT, no macro execution, no external-link resolution.
>
> `pip install ks-xlsx-parser`
>
> Repo: https://github.com/knowledgestack/ks-xlsx-parser
>
> Would genuinely love `.xlsx` files that break it — every edge-case
> report becomes a new fixture in the next release. We have a
> [Parser edge case](https://github.com/knowledgestack/ks-xlsx-parser/issues/new?template=parser_edge_case.yml)
> issue template specifically for that.
>
> (Part of the wider Knowledge Stack open-source family —
> https://github.com/knowledgestack — more parsers coming for PDF / DOCX /
> PPTX, plus an MCP server.)

---

## 📢 Dev.to / Medium / blog post outline

**Title:** `Make XLSX LLM Ready: why we built ks-xlsx-parser`

**Outline** (use as prompt to yourself, then expand):

1. **The problem** — why dropping a `.xlsx` into an LLM loses structure,
   provenance, and context-window budget.
2. **What citations mean for enterprise RAG** — every number an agent
   cites must point back to a coordinate; without that, compliance /
   finance / legal won't ship.
3. **The graph model** — cells, formulas, merges, tables, charts, CF, DV
   as typed Pydantic DTOs. Show a 10-line snippet of `to_json()`.
4. **The chunking story** — layout segmentation into logical blocks,
   pipe-text + HTML rendering, token count, source URI.
5. **Dependency traversal** — one code sample answering "what drives Q4
   revenue?" by walking `get_upstream(...)`.
6. **Stress testing on 1054 workbooks** — how we use the generator,
   what kinds of edge cases catch regressions.
7. **Two perf wins** — the 66× circular-ref cache and sparse-cell
   iteration. Enough detail to be useful if someone's hitting the same
   openpyxl bottleneck.
8. **What's next** — PDF/DOCX/PPTX parsers, MCP server, pivot tables.
9. **Call to action** — star the repo, join Discord, send us your
   weirdest spreadsheet.
