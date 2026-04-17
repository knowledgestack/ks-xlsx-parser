# ks-xlsx-parser Wiki

Welcome! This wiki holds the implementation detail we'd rather keep out of
the front-page README so it stays scannable. The code-heavy stuff lives here.

## Start here

- **[Quick Start](Quick-Start)** — 5 end-to-end snippets that cover ~90 %
  of real-world usage: parse, iterate chunks, walk the dependency graph,
  serialise for a DB/vector store, parse from bytes.
- **[API Reference](API-Reference)** — full signatures and examples for
  `parse_workbook`, `compare_workbooks`, `export_importer`,
  `StageVerifier`.
- **[Web API](Web-API)** — running the bundled FastAPI app and calling
  `POST /parse` from `curl` / Python / TypeScript.
- **[Data Models](Data-Models)** — the Pydantic DTOs you'll be reading in
  JSON output, field by field.
- **[Pipeline Internals](Pipeline-Internals)** — how the 8 stages fit
  together, and where to hook in if you want to extend the parser.

## Related docs in the main repo

- [`README.md`](https://github.com/knowledgestack/ks-xlsx-parser#readme) —
  hero page, architecture diagram, comparison table, community links.
- [`docs/WORKBOOK_GRAPH_SPEC.md`](https://github.com/knowledgestack/ks-xlsx-parser/blob/main/docs/WORKBOOK_GRAPH_SPEC.md) —
  the canonical specification for the extraction output.
- [`docs/PARSER_KNOWN_ISSUES.md`](https://github.com/knowledgestack/ks-xlsx-parser/blob/main/docs/PARSER_KNOWN_ISSUES.md) —
  known edge cases and how we handle them.
- [`docs/corpora.md`](https://github.com/knowledgestack/ks-xlsx-parser/blob/main/docs/corpora.md) —
  the testBench stress corpus and public-corpus benchmarks.
- [`CONTRIBUTING.md`](https://github.com/knowledgestack/ks-xlsx-parser/blob/main/CONTRIBUTING.md) —
  dev loop, PR checklist, community channels.
- [`CHANGELOG.md`](https://github.com/knowledgestack/ks-xlsx-parser/blob/main/CHANGELOG.md) —
  release history.

## Community

- 💬 [Discord](https://discord.gg/4uaGhJcx) — fastest way to get a real
  answer from a human.
- 🗣 [GitHub Discussions](https://github.com/knowledgestack/ks-xlsx-parser/discussions) —
  async Q&A and RFCs.
- 🐞 [Issues](https://github.com/knowledgestack/ks-xlsx-parser/issues/new/choose) —
  bugs, feature requests, parser edge cases.

Something in the wiki out of date or confusing? Open a PR against
[`docs/wiki/`](https://github.com/knowledgestack/ks-xlsx-parser/tree/main/docs/wiki)
— the wiki is rebuilt from that directory on every release.
