# Corpus & Benchmarks

excel-parser benchmarks against public corpora that are downloaded on demand —
nothing large is committed to the repo.

## Primary corpus — SpreadsheetBench v0.1

912 instruction × xlsx tasks (5,458 unique workbooks) covering financial models,
project trackers, HR records, scientific data, and a long tail of small-business
spreadsheets. Each task ships with an `instruction`, a `data_position`, and
(usually) an `answer_position`, which gives us ground truth for retrieval recall.

```bash
make corpus-download    # fetch SpreadsheetBench + a few smaller corpora under data/corpora/
make bench-robust       # parse-success rate + structural counts vs Docling (~20 min)
make bench-retrieval    # top-k retrieval recall + table fragmentation rate vs Docling (~40 min)
```

Reports land in `tests/benchmarks/reports/<timestamp>_<git-sha>/`. The headline
numbers and methodology live in
[`tests/benchmarks/reports/COMPARISON.md`](../tests/benchmarks/reports/COMPARISON.md).

## Other public corpora — opt-in robustness

`scripts/download_corpora.sh` also fetches a handful of smaller xlsx corpora
(EUSES, Enron `.xlsx` subset, SheetJS / openpyxl samples) under
`data/corpora/`. These are useful for spot-checking specific failure modes.

```bash
python -m pytest -m corpus -v    # opt-in robustness run against external corpora
```
