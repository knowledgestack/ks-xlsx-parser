# SpreadsheetBench — `excel-parser` vs Docling

Head-to-head retrieval-recall on [SpreadsheetBench v0.1](https://github.com/RUCKBReasoning/SpreadsheetBench)
(912 real-world instruction × xlsx instances). **Both parsers were scored in the
same run on the same harness**, so the comparison is apples-to-apples.

- Embedding model: `BAAI/bge-small-en-v1.5`
- Parsers: `excel-parser` (this project) vs `docling==2.93.0`
- Reproduce: `make corpus-download && uv run --extra bench python scripts/eval_retrieval.py --parsers ks,docling`

## Headline (text-match recall, all scoreable instances)

| Metric | excel-parser | Docling 2.93 | Δ |
|---|---:|---:|---|
| Recall@1 (text-match) | 0.693 | **0.708** | Docling **+1.5 pp** |
| Recall@3 (text-match) | **0.848** | 0.820 | excel-parser **+2.8 pp** |
| Recall@5 (text-match) | **0.859** | 0.840 | excel-parser **+1.9 pp** |
| **Geometric Recall@5** (chunk `sheet!A1` overlaps answer) | **0.889** | 0.000 | excel-parser (citation-grade) |
| Mean parse time / file | 349 ms | **238 ms** | Docling |
| Median (p50) parse time / file | **11 ms** | 13 ms | excel-parser |
| Parser errors / 912 | 0 | 0 | tie |
| Table fragmentation rate | 0.221 | 0.000* | — |

\* Docling emits markdown tables with no `sheet!range`, so the fragmentation
metric (chunks-per-answer-region) is structurally 0/undefined for it, not a win.

## In-scope (retrieval-satisfiable instances only)

"In-scope" excludes the 245/912 instances whose answer must be *computed and
written* (the input has no value to retrieve — e.g. "modify this formula"); a
retrieval parser cannot satisfy those by design. Applied equally to both parsers.

| Metric (in-scope, n=667) | excel-parser | Docling 2.93 |
|---|---:|---:|
| Recall@1 (text-match) | 0.772 | **0.786** |
| Recall@3 (text-match) | **0.909** | 0.886 |
| Recall@5 (text-match) | **0.919** | 0.904 |
| Geometric Recall@5 | **0.960** | 0.000 |

## How to read this

- **Text-match retrieval quality is roughly even.** Docling wins @1; excel-parser
  wins @3 and @5. Neither parser dominates the answer-string-retrieval task.
- **Geometric recall is the real differentiator and it is a capability gap, not
  a quality gap.** `excel-parser` attaches a `sheet!A1:Z99` range to every chunk,
  so a retrieved chunk can cite the exact source cells (`Revenue!C7`). Docling
  outputs markdown without coordinates, so it is structurally 0.000 — it can say
  *what* the answer is but not *where* it is.
- **Parse latency is mixed.** Medians are close (~11–13 ms); on the mean Docling
  is faster (238 vs 349 ms) because `excel-parser` row-windows very large tables
  (rendering big sheets more than once). Both have 0 hard errors across 912.

## Harness corrections (honesty note)

The numbers above use a harness corrected vs. earlier releases. Both fixes are
parser-independent (they change *measurement*, not either parser) and are covered
by unit tests in `tests/test_eval_retrieval_classify.py`:

1. **Geometric overlap with an unspecified sheet.** ~62% of instances omit the
   sheet name (single-sheet workbooks). The old harness compared a real chunk
   sheet name against `""` and rejected correct matches; it now matches on
   geometry when the ground-truth sheet is unspecified. This raised
   `excel-parser`'s geometric recall from the previously-reported **0.369** — the
   parser always pointed to the right cells; the metric under-counted. Docling
   stays **0.000** (no coordinates to match).
2. **Unscoreable text instances excluded.** Instances whose answer cell is empty
   or an uncached formula in `answer.xlsx` have no ground-truth string to match
   and were being counted as text misses. They are now excluded from the text
   denominator for **both** parsers equally (consistent with the harness's
   existing `had_answer_values` bucket filter).

## Excluded: Marker

`Marker`'s xlsx → HTML → PDF → layout-recognition pipeline runs >30 min/workbook
on CPU. The framework can add a Marker adapter when a GPU is available; see
`tests/benchmarks/adapters/docling_adapter.py` as a template.
