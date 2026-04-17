#!/usr/bin/env bash
set -euo pipefail

# Download public XLSX corpora into data/corpora/.
# Idempotent: skips datasets that already exist.

ROOT="$(cd -- "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
CORPUS_DIR="$ROOT/data/corpora"
TMP_DIR="$(mktemp -d)"

cleanup() { rm -rf "$TMP_DIR"; }
trap cleanup EXIT

need_tool() {
  if ! command -v "$1" >/dev/null 2>&1; then
    echo "Missing required tool: $1" >&2
    exit 1
  fi
}

need_tool curl
need_tool unzip

mkdir -p "$CORPUS_DIR"

fetch_zip() {
  local name="$1"
  local url="$2"
  local dest="$CORPUS_DIR/$name"

  if [ -d "$dest" ]; then
    echo "✓ $name already present, skipping"
    return
  fi

  echo "→ Downloading $name ..."
  local zip_path="$TMP_DIR/$name.zip"
  curl -L --fail --retry 3 --connect-timeout 20 -o "$zip_path" "$url"

  local unzip_dir="$TMP_DIR/$name"
  mkdir -p "$unzip_dir"
  unzip -q "$zip_path" -d "$unzip_dir"

  mkdir -p "$dest"
  find "$unzip_dir" -type f -iname "*.xlsx" -print0 | while IFS= read -r -d '' f; do
    cp "$f" "$dest/$(basename "$f")"
  done

  local count
  count="$(find "$dest" -type f -name '*.xlsx' | wc -l | tr -d ' ')"
  if [ "$count" = "0" ]; then
    echo "⚠️  No .xlsx files found for $name (archive may be .xls-heavy)."
  else
    echo "✓ $name: $count xlsx files"
  fi
}

fetch_single() {
  local name="$1"
  local url="$2"
  local dest_dir="$CORPUS_DIR/github_samples"
  mkdir -p "$dest_dir"

  local dest="$dest_dir/$name"
  if [ -f "$dest" ]; then
    echo "✓ $name already present, skipping"
    return
  fi

  echo "→ Downloading $name ..."
  curl -L --fail --retry 3 --connect-timeout 20 -o "$dest" "$url"
  echo "✓ $name"
}

# EUSES (mostly .xls, but keep any .xlsx present)
fetch_zip "euses" "https://zenodo.org/records/581673/files/EUSES.zip"

# Enron (mostly .xls; will warn if no .xlsx)
fetch_zip "enron" "https://github.com/SheetJS/enron_xls/archive/refs/heads/master.zip"

# SheetJS + openpyxl sample files (individual XLSX)
fetch_single "comments_stress.xlsx" "https://raw.githubusercontent.com/SheetJS/sheetjs/master/test_files/comments_stress_test.xlsx"
fetch_single "merge_cells.xlsx" "https://raw.githubusercontent.com/SheetJS/sheetjs/master/test_files/merge_cells.xlsx"
fetch_single "number_format.xlsx" "https://raw.githubusercontent.com/SheetJS/sheetjs/master/test_files/number_format.xlsx"
fetch_single "genuine_empty.xlsx" "https://raw.githubusercontent.com/openpyxl/openpyxl/master/openpyxl/tests/data/genuine/empty.xlsx"

echo "Done. Corpora in $CORPUS_DIR"
