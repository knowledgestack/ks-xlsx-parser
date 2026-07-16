# DECO structural benchmark — ks vs docling

Corpus: DECO `completed/` · GT tables scored: **1480** (across 852 files)

ks parse: 0 files timed out, 0 errored (excluded from metrics below).

## Table-boundary detection (ks — needs A1 localisation)

| metric | value |
|---|---|
| mean best-IoU vs GT table | 0.511 |
| detected @ IoU≥0.5 | 723/1480 (48.9%) |
| detected @ IoU≥0.3 | 899/1480 (60.7%) |
| mean best-IoU, macro by file | 0.517 |
| ks regions overlapping one GT table (fragmentation) | mean 8.79, median 2 |
| GT tables split across >1 ks region | 916/1480 (61.9%) |

## Header-row detection (ks — shipped `find_header_span`)

GT tables with a header: **1480** · of which multi-row: **534** (36.1%)

| metric | precision | recall | F1 |
|---|---|---|---|
| end-to-end (header on ks's own region) | 0.817 | 0.370 | 0.509 |
| isolated (header on GT region) | 0.936 | 0.617 | 0.744 |
| end-to-end, macro F1 by file | | | 0.506 |

- ks emitted a non-empty header for **979/1480** (66.1%) GT-headered tables (end-to-end).
- exact header-row match (isolated): **1060/1480** (71.6%).

## Tables-per-sheet (docling — its only measurable axis)

docling emits no A1 coordinates, so it can't be scored on IoU or header rows. The one axis it exposes is how many table objects it produces per sheet vs the GT count.

_No docling records (run `--parser docling` first)._

## Header cohorts — single-row vs multi-row GT headers

Isolated = detector run directly on the GT region (detector quality, independent
of segmentation). Produced by `scripts/analyze_deco_headers.py --run RUN`.

| cohort | precision | recall | F1 | exact match |
|---|---|---|---|---|
| single-row (n=946), isolated | 0.896 | 0.906 | 0.901 | 819/946 (86.6%) |
| multi-row (n=534), isolated | 0.972 | 0.489 | 0.650 | 241/534 (45.1%) |
| all (n=1480), isolated | 0.936 | 0.617 | 0.744 | 1060/1480 (71.6%) |
| single-row, end-to-end | 0.735 | 0.563 | 0.638 | 475/946 (50.2%) |
| multi-row, end-to-end | 0.907 | 0.283 | 0.432 | 129/534 (24.2%) |
| all, end-to-end | 0.817 | 0.370 | 0.509 | 604/1480 (40.8%) |

### vs the previous detector (run `full_fix2`, isolated)

| cohort | metric | before → after |
|---|---|---|
| single-row | exact | 79.5% → **86.6%** |
| single-row | precision | 0.800 → **0.896** |
| multi-row | F1 | 0.495 → **0.650** |
| multi-row | exact | 23.8% → **45.1%** |
| multi-row | recall | 0.335 → **0.489** |
| all | F1 | 0.628 → **0.744** |
| all | exact | 59.4% → **71.6%** |

Remaining failure mass (isolated): MR-UNDER 260 (≈50 of which are DECO's
transposed/attributes-as-rows GT convention that a row-band detector cannot
express), SR-SHIFT 53, SR-OVER 38, SR-MISS 36, MR-SHIFT 12, MR-MISS 12,
MR-OVER 9. Diagnose any bucket with `scripts/diag_header_failure.py`.

Known follow-up: the renderers/chunker consume only `span.top` for column
naming and window headers, so rows 2..bottom of a multi-row band render as
data in part 1 and are absent from later windowed parts. With multi-row
detection now ~2× more frequent, rendering the full band is the next win.
