# Running the retrieval benchmark locally

This is the loop you'll use whenever you want to know "is my parser
change actually moving recall?" — the same pipeline CI runs, just on
your laptop. Goal: text recall@5 > 0.90 (currently ~0.70).

The full SpreadsheetBench corpus is 912 instances and takes ~30–45 min
on a Mac. For an iteration loop you'll mostly use `--sample 60` (≈ 3 min
after the first embedding-model load).

## TL;DR

```bash
# One-time
make install-dev              # installs the parser + dev deps
pip install -e ".[bench]"     # sentence-transformers + numpy (~500 MB)
make corpus-download          # downloads SpreadsheetBench → data/corpora/
                              # (91 MB tarball, 2,726 .xlsx after extract)

# Each time you want to score
python scripts/eval_retrieval.py \
  --corpus data/corpora/spreadsheetbench/all_data_912_v0.1 \
  --parsers ks \
  --sample 60 \
  --emit-failures
python scripts/triage_recall.py tests/benchmarks/reports/retrieval
python scripts/append_bench_history.py
```

Output: `tests/benchmarks/reports/retrieval/<UTC stamp>/` with
`results.ndjson`, `failures.ndjson`, `summary.json`, `summary.md`.

## Step-by-step

### 1. Install bench deps

The benchmark uses `sentence-transformers` (≈ 500 MB with torch) for
embeddings. They're a separate optional group so the parser package
itself stays lean:

```bash
pip install -e ".[bench]"
```

First run also downloads the embedding model (`BAAI/bge-small-en-v1.5`,
≈ 130 MB) into `~/.cache/huggingface/`.

### 2. Download the corpus

```bash
make corpus-download
```

This is the same `scripts/download_corpora.sh` CI uses. It pulls
SpreadsheetBench v0.1 plus a few legacy XLSX samples; only
SpreadsheetBench matters for retrieval scoring.

Layout you should end up with:

```
data/corpora/spreadsheetbench/
  all_data_912_v0.1/
    dataset.json                       # 912 (question, answer_position) tuples
    spreadsheet/
      <instance-id>/
        1_<instance-id>_input.xlsx     # test case 1 input
        1_<instance-id>_answer.xlsx    # test case 1 ground-truth output
        2_..., 3_...                   # additional test cases per instance
```

`data/` is gitignored — never commit corpus files.

### 3. Run the benchmark

A typical iteration cycle uses a sample for speed:

```bash
python scripts/eval_retrieval.py \
  --corpus data/corpora/spreadsheetbench/all_data_912_v0.1 \
  --parsers ks \
  --sample 60 \
  --emit-failures
```

Flags worth knowing:

| Flag                  | What it does                                                |
|-----------------------|-------------------------------------------------------------|
| `--parsers ks,docling`| Score one or both parsers. Docling is heavy; skip unless comparing. |
| `--sample N`          | Random N-instance subset (seeded). Omit for the full 912.   |
| `--seed 1337`         | Random seed for `--sample`. Stays stable across runs.       |
| `--emit-failures`     | Also write `failures.ndjson` with top-8 chunks per miss.    |
| `--test-case 1`       | Which of the (usually 3) test cases per instance to score.  |
| `--per-parser-timeout`| Wall-clock seconds before a hung parse is killed. Default 60. |

For a full run, drop `--sample` and add `--per-parser-timeout 120`:

```bash
python scripts/eval_retrieval.py \
  --corpus data/corpora/spreadsheetbench/all_data_912_v0.1 \
  --parsers ks \
  --emit-failures \
  --per-parser-timeout 120
```

`make bench-retrieval` is the same thing with the canonical defaults.

### 4. Read the triage report

```bash
python scripts/triage_recall.py tests/benchmarks/reports/retrieval
```

This auto-finds the most recent run and prints:

* **Bucket histogram** ranked by count. The top bucket is the
  highest-leverage thing to fix next.
* **3 example failures** per bucket showing the question, the
  ground-truth answer cell + values, and the top-8 chunks the parser
  produced (with a ✓ next to chunks that contain the answer).

Drill into one bucket:

```bash
python scripts/triage_recall.py tests/benchmarks/reports/retrieval \
  --bucket answer_absent_from_chunks --examples 10
```

Five buckets and what they mean:

| Bucket | Root cause | Fix lives in |
|---|---|---|
| `answer_absent_from_chunks` | Answer value in NO chunk. Cell dropped or rendered as formula. | `parsers/cell_parser.py`, `rendering/text_renderer.py` |
| `present_but_ranked_low` | A chunk DOES contain the answer but ranked >5. Chunk too big/heterogeneous. | `chunking/chunker.py` |
| `wrong_sheet` | Answer sheet never chunked. | `parsers/workbook_parser.py` |
| `geometric_no_overlap` | Text matches but the chunk's A1 range doesn't overlap GT. | `annotation/block_splitter.py`, `analysis/pattern_splitter.py` |
| `no_chunks` / `parse_error` | Upstream parser failure. | The exception. |

See `docs/recall-investigation.md` for the named hypotheses behind each
bucket and `.claude/skills/recall-failure-triage/SKILL.md` for the
agent-driven loop.

### 5. Append to history.jsonl

```bash
python scripts/append_bench_history.py
```

Appends one row per benchmark run to
`tests/benchmarks/reports/history.jsonl` tagged with the current git
commit, and prints the row-over-row delta on the headline metrics:

```
appended to tests/benchmarks/reports/history.jsonl:
  commit 421783f  recall_text@5=0.704  recall_text@1=0.580
  recall_text@5: 0.6800 → 0.7040  ▲ +0.0240
```

That's how "is recall improving?" gets answered. Goal: `recall_text@5 > 0.90`.

`make bench-track` chains eval + history-append + triage in one go.

## Docker path (matches CI exactly)

When you want to make sure local results aren't drifting from CI:

```bash
docker build -f Dockerfile.bench -t ks-xlsx-parser-bench .

# Quick sanity (60 instances, ~3 min after image load):
docker run --rm \
  -e BENCH_SAMPLE=60 \
  -v "$PWD/tests/benchmarks/reports:/app/tests/benchmarks/reports" \
  -v "$PWD/data:/app/data" \
  ks-xlsx-parser-bench

# Full corpus:
docker run --rm \
  -v "$PWD/tests/benchmarks/reports:/app/tests/benchmarks/reports" \
  -v "$PWD/data:/app/data" \
  ks-xlsx-parser-bench
```

The image pre-warms the embedding model at build time so the first
`docker run` doesn't pay the 130 MB download.

Environment knobs (also work in CI dispatch):

| Env var          | Default | What it does                              |
|------------------|---------|-------------------------------------------|
| `BENCH_SAMPLE`   | `0`     | Sample N instances (0 = full 912)         |
| `BENCH_PARSERS`  | `ks`    | Comma list (e.g. `ks,docling`)            |
| `BENCH_TIMEOUT`  | `120`   | Per-file parse timeout in seconds         |

## Adding a new failure bucket

If you find a recall failure mode that doesn't fit any of the existing
six buckets, add it instead of stuffing it into `answer_absent_from_chunks`:

1. Append the new bucket name to `FAILURE_BUCKETS` in
   `scripts/eval_retrieval.py`.
2. Update `classify_text_failure` so it can return the new name. Keep
   the predicate cheap — it runs once per scored instance.
3. Add the bucket + a one-line root cause + fix location to the table
   in `docs/recall-investigation.md` and the SKILL file.
4. Re-run `make bench-track`; confirm the histogram shows the new
   bucket and counts make sense.

## Troubleshooting

* `ModuleNotFoundError: No module named 'sentence_transformers'`
  — you skipped `pip install -e ".[bench]"`.
* `dataset.json not found in ...` — your `--corpus` is pointing at
  `data/corpora/spreadsheetbench`, not the nested
  `data/corpora/spreadsheetbench/all_data_912_v0.1`. The benchmark
  expects the leaf directory that contains `dataset.json`.
* `FileNotFoundError: 1_<id>_input.xlsx` — the corpus tarball didn't
  fully extract. Delete `data/corpora/spreadsheetbench/` and re-run
  `make corpus-download`.
* `recall_text@5 = 0.0` on a sample of 5 — small samples have huge
  variance because the benchmark seeds. Bump to `--sample 60` minimum
  before trusting the number; use `--sample 0` (full corpus) for a
  decision-grade comparison.
* MPS / CUDA torch errors on first sentence-transformers import —
  re-install with `pip install --upgrade torch torchvision`. The
  benchmark runs fine on CPU.

## See also

* `scripts/eval_retrieval.py` — the benchmark itself.
* `scripts/triage_recall.py` — the bucket histogram + exemplar dump.
* `scripts/append_bench_history.py` — history.jsonl row writer.
* `Dockerfile.bench` — reproducible benchmark image.
* `docs/recall-investigation.md` — diagnosis framework & hypotheses.
* `.claude/skills/recall-failure-triage/SKILL.md` — agent guide.
