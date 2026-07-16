# Changelog

All notable changes to **excel-parser** are documented here.

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

## [0.2.1] — 2026-05-19

### ⚠️ BREAKING (Fixed — see also #excel-parser channel report)
- Repository layout flattened on `src/` was leaking 13 generic top-level
  packages (`models`, `utils`, `parsers`, …) into installed wheels and
  silently dropping `pipeline.py` and `api.py` (setuptools `packages.find`
  only finds *packages*, not top-level modules). Users hitting
  `from excel_parser.pipeline import ...` on 0.2.0 from PyPI got
  `ModuleNotFoundError`. **All modules now live under
  `src/excel_parser/`**; the wheel's `top_level.txt` contains only
  `excel_parser`. Imports inside the package switched from
  `from pipeline import` to `from excel_parser.pipeline import`.
  Downstream code that imported the leaked generics
  (`from models import …`) MUST migrate to `from excel_parser.models …`.

### Added
- `scripts/verify_wheel.py` — builds the wheel, installs it in a fresh
  venv, and asserts the public import surface resolves. Wired into a
  new `wheel-check` job in `.github/workflows/ci.yml` and a `Verify wheel`
  step in `release.yml`. Regression guard for the packaging bug above.
- `scripts/triage_recall.py` + `scripts/append_bench_history.py` — turn
  `failures.ndjson` into a ranked bucket histogram with exemplar
  failures, and append each benchmark run to
  `tests/benchmarks/reports/history.jsonl` so recall is tracked
  commit-over-commit. Goal: text recall@5 > 0.90.
- `eval_retrieval.py --emit-failures` — dumps top-8 ranked chunks per
  miss with a `failure_bucket` (answer_absent_from_chunks /
  present_but_ranked_low / wrong_sheet / geometric_no_overlap / …) for
  triage. Summary JSON gains a `failure_buckets` histogram.
- `Dockerfile.bench` + `.github/workflows/benchmark.yml` — reproducible
  benchmark image; PR sample run (60 instances), weekly full corpus run.
- `make install-dev` alias and `make wheel-check` / `make bench-track`
  / `make docker-bench` targets.
- New `bench` optional-dependency group (`sentence-transformers`,
  `numpy`) — only the benchmark needs these.
- `docs/recall-investigation.md` documenting the diagnosis framework and
  three named hypotheses (chunk-size dilution, formula-expression
  rendering, range-bookkeeping drift).
- `.claude/skills/recall-failure-triage/SKILL.md` — agent skill that
  consumes the bucket output and proposes ranked fixes.

### Changed
- Dropped `PYTHONPATH=src` from Makefile benchmark targets — the
  package is now properly installable so callers don't need it.
- `pyproject.toml`: `packages.find` constrained to `excel_parser*`,
  `py.typed` declared as package data, `excel-parser-api` console script
  updated to `excel_parser.api:main`.

### ⚠️ BREAKING
- Retired the in-tree `testBench/` corpus. The 1054-workbook stress dataset
  and `make testbench*` targets are gone — benchmarks now run against the
  public SpreadsheetBench v0.1 corpus, downloaded on demand to `data/corpora/`
  (gitignored). See `docs/corpora.md`.

### Removed
- `testBench/` directory and all bundled real-world / generated workbooks.
- `make testbench-build`, `make testbench`, `make testbench-zip` targets.
- `testbench` job in `.github/workflows/ci.yml`.
- `testBench-vX.Y.Z.zip` release asset from the release workflow.
- `tests/test_testbench_roundtrip.py`, `tests/test_enterprise_scoring.py`,
  `tests/test_real_world_datasets.py`, `tests/test_cross_validation.py`.
- `scripts/build_testbench.py`, `scripts/generate_enterprise_fixtures.py`.
- `static_xlsx` pytest fixture (the test bench it iterated is gone).

### Changed
- README, wiki, examples, and contributor docs now point at SpreadsheetBench
  (`make bench-robust` / `make bench-retrieval`) as the canonical benchmark.
- `examples/demo.py` + `examples/generate_examples.py` now write/read fixtures
  under `examples/fixtures/` instead of the (removed) `testBench/real_world/`.


## [0.2.0] — 2026-05-11

**Benchmark + retrievability release.** Adds a head-to-head benchmark against
[Docling](https://github.com/DS4SD/docling) on the [SpreadsheetBench](https://github.com/RUCKBReasoning/SpreadsheetBench)
corpus (912 instances, 5,458 xlsx files) and fixes three rendering bugs that
were silently torpedoing RAG retrieval. excel-parser parses **99.945%** of
SpreadsheetBench and **ties Docling at recall@1 / wins at recall@3 (+2.7 pp)
and recall@5 (+1.8 pp)**, plus 36.9% citation-grade geometric recall (Docling
0%, structurally — no A1 anchors).

### Added
- `tests/benchmarks/adapters/docling_adapter.py` — Docling adapter speaking the
  same NDJSON-worker protocol as `ks_adapter.py` (#TBD).
- `tests/benchmarks/_runner.py`: `docling_runner` factory wired into
  `vs_hucre.py`'s `--parsers` dispatch.
- `scripts/eval_retrieval.py` — retrieval-recall benchmark over
  SpreadsheetBench's `(instruction, data_position, answer_position)` triples.
  Uses `sentence-transformers` (default `BAAI/bge-small-en-v1.5`) and computes
  geometric overlap + numeric/date/boolean-normalized text-match recall@k.
  Persistent docling subprocess with hard-kill timeout — PyTorch's table-rec
  loop holds the GIL through C-land so in-process timeouts don't work.
- `scripts/summarize_retrieval.py` — re-aggregate a `results.ndjson` into
  `summary.json` / `summary.md` if a long run is interrupted.
- `scripts/download_corpora.sh`: fetches SpreadsheetBench v0.1 (~96 MB tar.gz)
  into `data/corpora/spreadsheetbench/` (gitignored).
- `tests/benchmarks/README.md` — adapter design notes + benchmark how-to.
- `tests/benchmarks/reports/COMPARISON.md` — head-to-head report incl.
  methodology, capability matrix, caveats.
- `Makefile`: `bench`, `bench-robust`, `bench-retrieval` targets.

### Fixed
- `src/rendering/text_renderer.py`: numeric cells now render the raw value
  (`1272`) instead of Excel's display-formatted string (`1,272.00`). The
  display format defeated substring-match retrieval for the most common RAG
  query shape ("what was the value in 2020?" → user types `1272`).
- `src/rendering/text_renderer.py`: the `[=]` formula marker no longer
  spuriously inflates a cell past its column width, which used to trigger
  a sci-notation fallback (`1.272000e+03`) on perfectly small values.
  Column widths now computed using the same rendering pipeline data rows
  will use, so the long-value path only triggers on genuinely-too-wide
  values.
- `src/rendering/text_renderer.py`: dates render as ISO `YYYY-MM-DD` and drop
  the spurious `00:00:00` time component on midnight datetimes.
- `src/rendering/text_renderer.py`: embedded newlines inside header cells
  (e.g. `"租金\n天数"`) collapse to spaces so they don't tear apart the
  Markdown grid (regression fixed for `租赁收入计提表.xlsx`-class layouts).
- `src/chunking/segmenter.py`: removed `_detect_style_boundaries`. The
  function split a coherent table into 5 fragments at fill-color band
  boundaries (year-banding, alternating-row shading), shedding header
  context from data rows. The connected-components + gap detection
  already handles real boundaries; fill banding is not a semantic one.
- `src/parsers/cell_parser.py`: `GradientFill` cells no longer crash the
  sheet parser. Accessing `.patternType` on a `GradientFill` (vs the
  expected `PatternFill`) raised `AttributeError`, which propagated up and
  killed every cell on the sheet. We don't model gradients but we no
  longer drop the sheet because of them (caught by SpreadsheetBench
  instance `118-8`, 8 sheets / 1,244 cells previously lost).

### Changed
- `tests/benchmarks/_schema.py`: `formulas` is now nullable on `status=ok`
  records. Parsers that don't model formulas (Docling, Marker) can now
  emit valid `BenchmarkRecord`s without tripping schema validation. The
  schema's load-bearing `None` vs `0` distinction is preserved: `None` =
  "feature not modeled by this parser", `0` = "modeled and observed zero".

### Removed
- `scripts/compare_docling.py` — superseded by the unified `tests/benchmarks/`
  framework + `eval_retrieval.py`. The old script's `ScoreCard` composite
  score was structurally biased (formula-preservation gave Docling a 0 by
  definition while contributing 20/100 points; header-propagation used
  different proxies for each parser); replaced by parser-agnostic
  text-match and geometric recall metrics.

### Performance
- excel-parser is now ~5% faster on average parse time on SpreadsheetBench
  than Docling (251 ms vs 265 ms mean), while producing a richer output
  (formulas, dependency graph, charts, named ranges, etc.).

### Docs
- `tests/benchmarks/README.md` — new — methodology + adapter design.
- `tests/benchmarks/reports/COMPARISON.md` — new — head-to-head report.
- README — new "Benchmark — excel-parser vs Docling on SpreadsheetBench"
  section near the top with the headline table.

### Internal
- `tests/test_rendering.py`: updated `test_numeric_cells_use_scientific_notation_not_truncation`
  to assert the new raw-numeric rendering (test renamed
  `test_numeric_cells_render_raw_not_display_formatted`).
- `.gitignore`: `data/corpora/` (downloaded benchmark corpora; can run to
  several GB).
- `Makefile`: `bench`, `bench-robust`, `bench-retrieval` targets.

## [0.1.1] — 2026-04-17

**First public release.** MIT-licensed, open-sourced under the
[Knowledge Stack](https://github.com/knowledgestack) ecosystem. Detailed
announcement: [`docs/launch/RELEASE_NOTES_v0.1.1.md`](docs/launch/RELEASE_NOTES_v0.1.1.md).

### Added
- Public Python package **`excel-parser`** on PyPI; import as
  `excel_parser` or the alias `excel_parser`.
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
- FastAPI web server (`excel-parser-api`) in the `[api]` extra.
- GitHub Actions: `ci.yml` (test matrix on py3.10/3.11/3.12 × ubuntu/macos
  + dedicated testBench job) and `release.yml` (wheel + sdist + testBench
  zip, PyPI Trusted Publishing).
- Community infra: `CODE_OF_CONDUCT.md`, `SECURITY.md`, issue / PR /
  discussion templates, `FUNDING.yml`, pre-commit config.

### Performance
- Chunk builder caches `detect_circular_refs()` per workbook instead of
  re-running it per block. Real 21k-cell financial model:
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
  workbooks with known openpyxl `data_only` caching gaps, 5 % everywhere
  else. See [`docs/PARSER_KNOWN_ISSUES.md`](docs/PARSER_KNOWN_ISSUES.md).

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
- Rebranded from `arnav2/XLSXParser` to `knowledgestack/excel-parser`;
  transferred the repo into the `knowledgestack` org and made it public.
- `uv.lock` regenerated after dropping the `[ralph]` extra and adding
  `pytest-timeout` / `ruff` / `mypy` to `[dev]`.

## [0.1.0] — 2026-02-25 (private beta)

Private-beta release used inside the Knowledge Stack ecosystem. Not
published to PyPI. Superseded by 0.1.1.

<!-- Compare links -->
[Unreleased]: https://github.com/knowledgestack/excel-parser/compare/v0.1.1...HEAD
[0.1.1]: https://github.com/knowledgestack/excel-parser/releases/tag/v0.1.1
