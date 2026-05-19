#!/usr/bin/env bash
# Entrypoint for the benchmark Docker image (Dockerfile.bench).
#
# Ensures the SpreadsheetBench corpus is present, runs the retrieval-recall
# benchmark for ks-xlsx-parser, appends the result to history.jsonl, and
# prints a failure-bucket triage so accuracy can be tracked over time.
#
# Env vars:
#   BENCH_SAMPLE   parse only N random instances (0 / unset = full 912)
#   BENCH_PARSERS  comma list passed to eval_retrieval.py (default: ks)
#   BENCH_TIMEOUT  per-file parser timeout in seconds (default: 120)
set -euo pipefail

cd "$(dirname "$0")/.."

CORPUS_DIR="data/corpora/spreadsheetbench"
SAMPLE="${BENCH_SAMPLE:-0}"
PARSERS="${BENCH_PARSERS:-ks}"
TIMEOUT="${BENCH_TIMEOUT:-120}"

if [ ! -d "$CORPUS_DIR" ]; then
  echo "→ Downloading SpreadsheetBench corpus ..."
  mkdir -p "$CORPUS_DIR"
  curl -L --fail --retry 3 --connect-timeout 20 \
    -o /tmp/sb.tar.gz \
    "https://raw.githubusercontent.com/RUCKBReasoning/SpreadsheetBench/main/data/spreadsheetbench_912_v0.1.tar.gz"
  tar -xzf /tmp/sb.tar.gz -C "$CORPUS_DIR"
  rm -f /tmp/sb.tar.gz
fi

# eval_retrieval.py expects the dataset.json + spreadsheet/ dir. Find it.
CORPUS_ARG="$CORPUS_DIR"
if [ -d "$CORPUS_DIR/all_data_912_v0.1" ]; then
  CORPUS_ARG="$CORPUS_DIR/all_data_912_v0.1"
fi

SAMPLE_ARG=()
if [ "$SAMPLE" != "0" ]; then
  SAMPLE_ARG=(--sample "$SAMPLE")
  echo "→ Sampling $SAMPLE instances"
fi

echo "→ Running retrieval benchmark (parsers=$PARSERS) ..."
python scripts/eval_retrieval.py \
  --corpus "$CORPUS_ARG" \
  --parsers "$PARSERS" \
  --emit-failures \
  --per-parser-timeout "$TIMEOUT" \
  --out tests/benchmarks/reports/retrieval \
  "${SAMPLE_ARG[@]}"

echo "→ Appending to history.jsonl ..."
python scripts/append_bench_history.py

echo
python scripts/triage_recall.py tests/benchmarks/reports/retrieval
