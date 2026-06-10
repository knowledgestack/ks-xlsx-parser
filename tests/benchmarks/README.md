# xlsx parser benchmarks

Two benchmarks, both reproducible:

| Benchmark | What it measures | Corpus | Cost |
|---|---|---|---|
| `vs_hucre.py` (structural) | Parse-success rate + structural counts (cells, formulas, tables, merges, etc.) across many files | `data/corpora/spreadsheetbench/` (5,458 real-world) | Cheap — 1–20 min |
| `scripts/eval_retrieval.py` (chunk quality) | Recall@k for retrieving the relevant chunk given a natural-language instruction, + table-integrity fragmentation rate | SpreadsheetBench `dataset.json` (912 instruction + position pairs) | Medium — 10 min on 100 instances |

## 1. Structural benchmark — `vs_hucre.py`

Long-running NDJSON-protocol workers, per-file timeout, batch respawn, randomized parser order. Adapters live under `adapters/`; runners in `_runner.py`. Add a new parser by:

1. Write `adapters/<name>_adapter.py` that speaks the protocol (see `ks_adapter.py` and `docling_adapter.py` as references).
2. Add a runner factory in `_runner.py`.
3. Wire it into `vs_hucre.py`'s `--parsers` handling.

Supported parsers today: `ks` (excel-parser), `hucre` (TypeScript, requires `pnpm install` under `hucre_node/`), `docling` (IBM Docling — `uv pip install docling`).

```bash
# Quick smoke (50 random files from SpreadsheetBench)
PYTHONPATH=src uv run python -m tests.benchmarks.vs_hucre \
    --corpus data/corpora/spreadsheetbench --sample 50 --parsers ks

# Robustness on full SpreadsheetBench (5,458 files, ~20 min)
PYTHONPATH=src uv run python -m tests.benchmarks.vs_hucre \
    --corpus data/corpora/spreadsheetbench --parsers ks \
    --per-file-timeout 120 --out tests/benchmarks/reports/spreadsheetbench

# ks vs docling on a sample
PYTHONPATH=src uv run python -m tests.benchmarks.vs_hucre \
    --corpus data/corpora/spreadsheetbench --sample 100 \
    --parsers ks,docling
```

Outputs (per run, timestamped subdir):
- `results.csv` — one row per (file, parser), schema-validated
- `raw.ndjson` — full records with `extra` fields preserved
- `failures.jsonl` — `status != "ok"` rows
- `summary.md` — status matrix, capability matrix, aggregate counts, perf percentiles, per-sub-corpus breakdown
- `drift.md` — pairwise feature-count disagreement between two parsers
- `manifest.json` — git sha, Python/Node version, timestamp

### Schema notes (`_schema.py`)

Null vs zero is **load-bearing**: `None` means the parser doesn't model a feature, `0` means it does and observed zero. Drift and capability reports treat them differently. Adapters must respect this.

## 2. Chunk-quality benchmark — `scripts/eval_retrieval.py`

Uses SpreadsheetBench's `dataset.json` (912 instances) — each has an `instruction`, a `data_position`, and an `answer_position`. We:

1. Parse the input `.xlsx` with each parser → list of `(text, sheet, A1 range)` chunks.
2. Embed all chunks + instruction with sentence-transformers (BGE-small by default).
3. Rank chunks by cosine similarity to the instruction.
4. Check whether the chunk containing `data_position` (falling back to `answer_position` for the 561 instances that omit `data_position`) is in top-k.

**Two recall metrics**:
- **Geometric overlap** — chunk's `A1` range overlaps the ground-truth region. Requires parsers to surface (sheet, range) per chunk; docling doesn't, so its geometric recall is structurally zero. Useful for measuring ks against its own past output.
- **Text-match** — the answer cell's actual string value appears as a substring of the chunk's text. Parser-agnostic — this is the apples-to-apples retrieval metric for cross-parser comparison.

**Fragmentation rate** — for single-region instances, how many chunks does the data region span? `1` = clean; `>1` = the table was fragmented across chunks (the 租赁收入计提表-class regression).

```bash
# Sample 100 instances, both parsers
PYTHONPATH=src uv run python scripts/eval_retrieval.py \
    --sample 100 --parsers ks,docling

# Full corpus (912 instances) — single-parser baseline
PYTHONPATH=src uv run python scripts/eval_retrieval.py --parsers ks
```

Outputs: `results.ndjson` (per-instance per-parser), `summary.json` (machine-readable aggregate), `summary.md` (human-readable table).

## 3. Structural table/header benchmark — `scripts/eval_deco.py`

Scores **table-boundary** and **header-row** detection against ground truth, which
SpreadsheetBench cannot (it only has answer positions, not table/header regions).
Uses [DECO](https://github.com/ddenron/deco_dataset) (Dresden Enron COrpus, ICDAR
2019): 852 real `.xlsx` whose hidden `Range_Annotations_Data` sheet marks, per
sheet, every table region and its header/data/derived row ranges.

For each GT table we score:
- **ks** — table-region IoU vs GT, fragmentation, and header-row precision/recall/F1
  via the shipped `find_header_span`.
- **docling** — tables-per-sheet vs GT count (its only measurable axis; it emits no
  A1 coordinates, so it can't be scored on IoU or headers).

Run one parser per process (memory isolation), then a report pass:

```bash
PYTHONPATH=src uv run python scripts/eval_deco.py --parser ks      --out tests/benchmarks/reports/deco/run
PYTHONPATH=src uv run python scripts/eval_deco.py --parser docling --out tests/benchmarks/reports/deco/run
PYTHONPATH=src uv run python scripts/eval_deco.py --report         --out tests/benchmarks/reports/deco/run
```

Outputs: `ks.ndjson` / `docling.ndjson` (per-table records), `summary.json`,
`summary.md`. A per-file timeout (`--timeout`, default 60s) bounds pathological
large sheets; reports give both micro (per-table) and macro (per-file) aggregates.

## Corpora

Run `scripts/download_corpora.sh` once to populate `data/corpora/` (gitignored). Currently fetches:
- SpreadsheetBench v0.1 (912 task instances × ~6 xlsx each = 5,458 files)
- DECO (852 annotated `.xlsx` with table + header ground truth, under `deco/completed/`)
- EUSES (mostly .xls)
- Enron spreadsheets (mostly .xls)
- A handful of SheetJS / openpyxl sample xlsx

## Caveats

- **Marker is intentionally absent.** Marker's xlsx pipeline goes xlsx → HTML → PDF (WeasyPrint) → markdown via PDF layout-recognition models. On a CPU-only machine it took >30 min on a single 1k-row workbook; not viable for a 5,458-file corpus. The structural framework supports adding a marker adapter (see `adapters/docling_adapter.py` as a template) — the speed wall is the obstacle, not the integration.
- **Memory measurement** in `_mem.py` is RSS-delta and approximate (±30%). For trustworthy memory numbers run one parser per process and look at peak RSS reported by the OS.
- **The retrieval embedding model** matters. BGE-small-en-v1.5 is the default for speed; switch with `--model BAAI/bge-large-en-v1.5` for better recall at ~10× the compute.
