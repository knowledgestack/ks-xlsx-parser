# Retrieval-recall benchmark (SpreadsheetBench)

- Corpus: `data/corpora/spreadsheetbench/all_data_912_v0.1`
- Instances scored: 50
- Embedding model: `BAAI/bge-small-en-v1.5`

| Metric | ks-xlsx-parser |
|---|---|
| Recall@1 (geometric, all) | 0.220 |
| Recall@3 (geometric, all) | 0.280 |
| Recall@5 (geometric, all) | 0.280 |
| Recall@1 (text-match, all) | 0.340 |
| Recall@3 (text-match, all) | 0.400 |
| Recall@5 (text-match, all) | 0.400 |
| Recall@5 (geometric, in-scope) ** | 0.280 |
| Recall@5 (text-match, in-scope) ** | 0.400 |
| In-scope instances | 50 |
| Out-of-scope (execution-required) | 0 |
| Fragmentation rate | 0.000 |
| Mean parse ms | 74.800 |
| P50 parse ms | 9.260 |
| Errors | 0 |

**Geometric overlap** = chunk's reported A1 range overlaps the ground-truth `answer_position`. Requires the parser to surface (sheet, range) per chunk — docling does not, so its geometric recall is structurally 0.

** **In-scope** excludes instances where the input spreadsheet has nothing at `answer_position` (the question asks the system to *compute and write* the answer; a retrieval parser cannot help). The recall-90 roadmap gates on the in-scope number. See `docs/planning/recall-90/05-out-of-scope-execution-instances.md`.

## Failure buckets (text-match recall@5 misses)

Why each miss happened — the biggest bucket is the highest-
leverage fix. `answer_absent_from_chunks` → fix extraction; 
`present_but_ranked_low` → fix chunking/embedding.

| Bucket | ks-xlsx-parser |
|---|---|
| answer_absent_from_chunks | 12 |
| present_but_ranked_low | 0 |
| wrong_sheet | 0 |
| geometric_no_overlap | 0 |
| no_chunks | 0 |
| parse_error | 0 |
| unparseable_ground_truth | 0 |

**Text-match** = the answer cell's actual string value appears as a substring of the chunk's text. Parser-agnostic; this is the apples-to-apples retrieval comparison.

