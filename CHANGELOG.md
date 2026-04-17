# Changelog

All notable changes to **ks-xlsx-parser** are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

Each release lives under a version heading linked to its GitHub compare view
at the bottom. Subsections use a fixed set of labels so the log is skimmable:

- **Added** — new features
- **Changed** — changes in existing behaviour
- **Deprecated** — soon-to-be removed features (keep at least one release ahead)
- **Removed** — removed features
- **Fixed** — bug fixes
- **Security** — vulnerability fixes (link to the GHSA advisory)
- **Performance** — noteworthy perf wins, with numbers
- **Docs** — user-facing documentation changes
- **Internal** — refactors, test infra, tooling (only when it affects contributors)

Breaking changes get a `⚠️ BREAKING:` prefix and are called out at the top of
the release. Keep entries in the imperative ("add X"), one line each, linking
issues or PRs in parentheses (`#123`).

<!--
Template for a new release (copy this block, fill in, move Unreleased items in):

## [X.Y.Z] — YYYY-MM-DD

### Added
- <entry> (#PR)

### Changed
- <entry> (#PR)

### Fixed
- <entry> (#PR)

### Performance
- <entry — include before/after numbers> (#PR)

### Docs
- <entry> (#PR)
-->

## [Unreleased]

Nothing yet. Open a PR and add your entry under the appropriate heading.

## [0.1.1] — 2026-04-17

**First public release.** MIT-licensed, open-sourced under the
[Knowledge Stack](https://github.com/knowledgestack) ecosystem. Detailed
announcement: [`docs/launch/RELEASE_NOTES_v0.1.1.md`](docs/launch/RELEASE_NOTES_v0.1.1.md).

### Added
- Public Python package **`ks-xlsx-parser`** on PyPI; import as
  `xlsx_parser` or the alias `ks_xlsx_parser`.
- `parse_workbook()` returning a `ParseResult` with `.workbook`,
  `.chunks`, and `.serializer` — full workbook graph (cells, formulas,
  merges, tables, charts, CF, DV, named ranges, dependency edges).
- `compare_workbooks()` + `export_importer()` for multi-workbook template
  alignment and Python-importer generation.
- `StageVerifier` / `VerificationReport` / `ExcellentStage` for pipeline
  stage-level assertions.
- RAG-ready `ChunkDTO` with `source_uri`, `render_text`, `render_html`,
  `token_count`, `dependency_summary`, and xxhash64 content hash.
- **`testBench/`** — 1053-workbook stress corpus (real_world 8 + enterprise 4
  + github_datasets 10 + stress/curated 26 + stress/merges 5 + generated
  1000). Ships as `testBench-v0.1.1.zip` release asset.
- `scripts/build_testbench.py` — deterministic generator (matrix: 297,
  combo: 400, adversarial: 300).
- `tests/test_testbench_roundtrip.py` — parallel round-trip gate;
  1054/1054 passing in ~70 s.
- FastAPI web server (`xlsx-parser-api`) in the `[api]` extra.
- GitHub Actions: `ci.yml` (test matrix on py3.10/3.11/3.12 × ubuntu/macos
  + dedicated testBench job) and `release.yml` (wheel + sdist + testBench
  zip, PyPI Trusted Publishing).
- Community infra: `CODE_OF_CONDUCT.md`, `SECURITY.md`, issue / PR /
  discussion templates, `FUNDING.yml`, pre-commit config.

### Performance
- Chunk builder caches `detect_circular_refs()` per workbook instead of
  re-running it per block. Real 21k-cell financial model (Walbridge):
  **307 s → 4.6 s (66×)**.
- Sheet parser iterates openpyxl's `_cells` dict instead of `iter_rows()`
  over the full bounding box. Workbooks with extreme sparse addresses
  (e.g. `A1` + `XFD1048576`): **60 s timeout → 135 ms**.

### Fixed
- Conditional-formatting rules (`top10`, `uniqueValues`, `duplicateValues`,
  `containsText`, `aboveAverage`, `belowAverage`) no longer reference a
  non-existent `dxfId=0` in generated fixtures, so openpyxl can load them
  back without an `IndexError`.
- `test_formula_cached_values_match` now applies a 15 % threshold for
  workbooks with known openpyxl `data_only` caching gaps (Walbridge),
  5 % everywhere else. See
  [`docs/PARSER_KNOWN_ISSUES.md`](docs/PARSER_KNOWN_ISSUES.md).

### Docs
- New README positioned as *"Make XLSX LLM Ready"* with architecture
  diagram, comparison table vs pandas/openpyxl/Docling, vertical-use-case
  section, Knowledge Stack ecosystem links, and prominent Discord + ⭐
  call-to-actions.
- [`CONTRIBUTING.md`](CONTRIBUTING.md) rewritten with three first-PR
  paths and Discord as the primary community channel.
- [`docs/MAINTAINERS.md`](docs/MAINTAINERS.md) — branch-protection
  playbook, label script, Discussions categories, PyPI Trusted
  Publishing setup, release checklist.
- [`testBench/README.md`](testBench/README.md) — dataset layout, manifest
  schema, licensing.
- [`docs/launch/`](docs/launch/) — v0.1.1 release notes +
  Discord / Twitter / LinkedIn / HN / Reddit / blog announcement drafts.

### Internal
- Consolidated 53 checked-in `.xlsx` fixtures under a single `testBench/`
  tree; updated every path reference in tests, scripts, and demos.
- Removed internal-only tooling: Ralph loop scripts, Cursor / Serena
  agent configs, iteration logs, Knowledge-Stack-internal framing in
  DESIGN.md.
- Rebranded from `arnav2/XLSXParser` to `knowledgestack/ks-xlsx-parser`;
  transferred the repo into the `knowledgestack` org and made it public.
- `uv.lock` regenerated after dropping the `[ralph]` extra and adding
  `pytest-timeout` / `ruff` / `mypy` to `[dev]`.

## [0.1.0] — 2026-02-25 (private beta)

Private-beta release used inside the Knowledge Stack ecosystem. Not
published to PyPI. Superseded by 0.1.1.

<!-- Compare links -->
[Unreleased]: https://github.com/knowledgestack/ks-xlsx-parser/compare/v0.1.1...HEAD
[0.1.1]: https://github.com/knowledgestack/ks-xlsx-parser/releases/tag/v0.1.1
