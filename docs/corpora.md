# Corpus & Benchmarks

The XLSXParser test bench is split into two tiers.

## 1. `testBench/` — checked into the repo

A 1053-workbook corpus shipped with every clone, exercising the full extraction
spec. Round-tripped on every CI run. See [`testBench/README.md`](../testBench/README.md)
for the layout.

```bash
make testbench-build   # regenerate the 1000-file `generated/` subtree
make testbench         # parse every workbook, record failures to metrics/testbench/
make testbench-zip     # package as a GitHub release asset
```

## 2. External public corpora — downloaded on demand

Heavier public datasets (EUSES, Enron `.xlsx` subset, SheetJS/openpyxl samples)
stay out of git and download under `tests/fixtures/corpus/`.

```bash
make corpus-download                    # fetch external corpora
python -m pytest -m corpus -v           # opt-in robustness run
```

## Enterprise scorecard (runs by default)

```bash
python -m pytest tests/test_enterprise_scoring.py -v
```

Four small deterministic fixtures under `testBench/enterprise/` are regenerated
if missing by `scripts/generate_enterprise_fixtures.py`. Per-file scorecards
are written to `metrics/corpus/`; git ignores the `metrics/` tree so CI can
upload the artifacts without polluting history.
