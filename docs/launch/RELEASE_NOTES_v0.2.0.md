# excel-parser v0.2.0 — Benchmark + Retrievability 📊

**Headline:** excel-parser now has a head-to-head benchmark against [Docling](https://github.com/DS4SD/docling) on the [SpreadsheetBench](https://github.com/RUCKBReasoning/SpreadsheetBench) corpus (912 task instances, 5,458 xlsx files). ks **parses 99.945%** of the corpus and **ties Docling at recall@1 / wins at recall@3 (+2.7 pp) and recall@5 (+1.8 pp)** on apples-to-apples retrieval, with **36.9% citation-grade geometric recall** that Docling structurally cannot achieve.

Plus three quiet RAG-breaking rendering bugs in 0.1.1 are gone.

## What's new

### 🏁 SpreadsheetBench benchmark — `make bench`

A reproducible, parser-agnostic benchmark over real-world workbooks scraped from ExcelHome / Mr.Excel / r/excel:

| Metric | **excel-parser** | Docling 2.93 | Δ |
|---|---:|---:|---:|
| Parse success (5,458 files) | **99.945%** | not run at scale | — |
| Recall@1 (text-match) | 0.580 | 0.579 | **+0.1 pp (tied)** |
| Recall@3 (text-match) | **0.697** | 0.670 | **+2.7 pp** |
| Recall@5 (text-match) | **0.704** | 0.686 | **+1.8 pp** |
| Recall@5 (geometric, A1 anchor overlap) | **0.369** | 0.000 | Docling has no per-chunk anchors |
| Mean parse time per file | **251 ms** | 265 ms | ks ~5% faster |

**Why "geometric" recall matters for RAG:** ks emits a `sheet!A1:Z99` range with every chunk. A retrieval system that surfaces the chunk can render a citation that points at the exact source cells. Docling produces markdown without per-chunk anchors, so it can't satisfy this metric at all. This is the difference between "the answer was in *the workbook*" and "the answer was in *cell C7 of the Revenue sheet*."

Marker is intentionally absent — its xlsx → HTML → PDF → layout-model pipeline clocks >30 min per workbook on CPU. The harness supports adding a Marker adapter (`tests/benchmarks/adapters/docling_adapter.py` as a template); the speed wall is the obstacle.

Full methodology, capability matrix, and caveats: [`tests/benchmarks/reports/COMPARISON.md`](https://github.com/knowledgestack/excel-parser/blob/main/tests/benchmarks/reports/COMPARISON.md).

### 🔧 Three rendering bugs that were silently torpedoing retrieval

1. **Comma-formatted numbers.** `1272` rendered as `"1,272.00"` (Excel's display format). A user query `"1272"` substring-missed. Now: numeric cells render the raw value.
2. **Spurious sci-notation.** The `[=]` formula marker inflated a cell past column width, tripping a long-value fallback that rendered `1272` as `"1.272000e+03"`. Now: column widths computed using the same rendering pipeline data rows will use.
3. **Embedded newlines in headers** (common in CJK workbooks like `"租金\n天数"`) tore apart the Markdown table grid. Now: collapsed to spaces.

These three together accounted for the entire retrieval-recall gap we initially measured against Docling.

### 🧹 Segmenter — no more banded-table fragmentation

Removed `_detect_style_boundaries` from `chunking/segmenter.py`. The function split a coherent table into 5 fragments at fill-color band boundaries (year-banding, alternating-row shading), shedding header context from data rows. The connected-components + gap detection already handles real boundaries; fill banding is not a semantic one.

### 🛡️ GradientFill safety

Cells using `GradientFill` (rare but real — caught by SpreadsheetBench instance `118-8`, 8 sheets / 1,244 cells previously lost) used to crash the sheet parser. Now: defensively skipped, sheet keeps parsing.

### 🐳 Productionization

- `Makefile`: `make bench`, `make bench-robust`, `make bench-retrieval`
- `scripts/download_corpora.sh` now fetches SpreadsheetBench v0.1
- `scripts/summarize_retrieval.py` — re-aggregate a partial `results.ndjson` if a long run gets interrupted
- New benchmark framework supports adding parsers (Marker, hucre, others) via the NDJSON-worker protocol; see `tests/benchmarks/README.md`

## Reproduce

```bash
pip install -U excel-parser==0.2.0      # or
git clone https://github.com/knowledgestack/excel-parser
cd excel-parser
make corpus-download                       # one-time, ~100 MB
make bench                                 # ~30 min for both benchmarks
open tests/benchmarks/reports/COMPARISON.md
```

## Upgrading from 0.1.1

No breaking API changes. The only behavioral change is that **`render_text` on numeric cells now contains the raw value instead of the Excel-display-formatted string** (e.g. `1272` instead of `1,272.00`). If you were relying on display formatting in retrieval keys or downstream regex parsing, switch to the cell's `display_value` field on the `ChunkDTO`. For everything else, drop-in.

Full changelog: [`CHANGELOG.md`](https://github.com/knowledgestack/excel-parser/blob/main/CHANGELOG.md#020--2026-05-11).

## Thanks

To the [SpreadsheetBench](https://github.com/RUCKBReasoning/SpreadsheetBench) team at Renmin University for publishing a clean, real-world xlsx corpus with structured ground truth — none of this comparison would have been possible without it.
