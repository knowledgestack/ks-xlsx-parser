"""
Local-only benchmark harness. Not part of the public test suite.

Runs `ks-xlsx-parser` head-to-head against external parsers (currently `hucre`,
a TypeScript zero-dependency spreadsheet I/O library) across the `testBench/`
corpus and produces per-file perf + feature-coverage records.

Not committed by default — reports and node_modules are git-ignored. Invoke
via `python -m tests.benchmarks.vs_hucre --corpus testBench`.

Pitfalls this harness is designed to avoid (read before editing):

1. Don't count features by regex over adapter output — call parser APIs.
2. Don't conflate round-trip-preserved objects with extracted data. Report
   counts but `null` the detail fields (e.g., hucre charts: count yes,
   chart_types no).
3. Don't buffer all results in memory — stream NDJSON → CSV line by line.
4. Don't time Node from Python; each worker times itself.
5. Don't compare raw wall-clock without a per-1k-cell normalization.
6. Don't let disk cache favor the second parser — pre-warm buffers, pass
   bytes to both adapters.
7. Don't report `0` when a parser doesn't model a feature — use `None`. The
   summary/drift generators treat `None` and `0` differently.
8. Don't fail a file because one parser crashed — record both independently.
9. Don't run under pytest; standalone `python -m`.
10. Don't commit reports. `.gitignore` covers them.
11. Pin hucre exact (`0.3.0`), `--frozen-lockfile`.
"""

__all__: list[str] = []
