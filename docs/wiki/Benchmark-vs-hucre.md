# `ks-xlsx-parser` vs [`hucre`](https://github.com/productdevbook/hucre)

An honest, reproducible head-to-head against [**`hucre`**](https://github.com/productdevbook/hucre) —
an excellent zero-dependency TypeScript spreadsheet I/O engine by
[**@productdevbook**](https://github.com/productdevbook). Hucre reads **and
writes** xlsx/csv/ods, runs in Node/Deno/Bun/browsers/Cloudflare Workers, and
ships in ~18 KB gzipped. It's a different *category* of tool than
`ks-xlsx-parser` — they're an I/O engine, we're a semantic extractor — but
since xlsx reading overlaps, it's worth putting both on the same corpus and
publishing what we find. We built the comparison as much to learn from hucre
as to measure ourselves.

---

## TL;DR

- **hucre is faster on raw throughput**: ~3× at P50 in our fast mode, ~25–100× at P95 on data-heavy files.
- **We extract more**: formula dependency graph, chart type/series, pivots, RAG chunks with token counts + citation URIs, content hashes. Hucre extracts **sparklines** and round-trips charts — we don't.
- **We agree on every feature both parsers extract** to exact parity (tables, merges, CF rules, DV rules, hyperlinks, comments) or near-exact (formulas: 0.05% drift).
- Accuracy is the primary constraint of `ks-xlsx-parser`: **1631-test pytest suite**, cross-validated against [`calamine`](https://github.com/tafia/calamine), zero regressions required on every perf change.

Pick hucre for edge-runtime / browser / CF-Worker I/O.
Pick `ks-xlsx-parser` for Python LLM / RAG / auditing pipelines.

---

## Performance — 1053-workbook testBench corpus

Same machine, same run, same OS page cache. `parse_workbook(mode="fast")`
is the apples-to-apples configuration for hucre's read-only path (it skips
LLM-specific chunking + template/tree extraction but still extracts every
metadata feature hucre extracts).

| metric | `hucre` 0.3.0 | `ks-xlsx-parser` **full** | `ks-xlsx-parser` **fast** |
|---|---:|---:|---:|
| P50 parse time | **1.3 ms** | 5.0 ms | **3.9 ms** |
| P95 parse time | **3.5 ms** | 368 ms | 206 ms |
| P99 parse time | **30.2 ms** | 469 ms | 246 ms |
| mean parse time | **2.7 ms** | 73.9 ms | 39.5 ms |
| total wall-clock | **2.8 s** | 77.8 s | 41.6 s |
| Walbridge Coatings<br>(17.6k formulas, worst real-world file) | **139 ms** | 1413 ms | 686 ms |

### Ratio to hucre

| mode | P50 ratio | P95 ratio | mean ratio |
|---|:---:|:---:|:---:|
| full | 3.8× slower | 105× slower | 27× slower |
| **fast** | **3.0× slower** | 60× slower | 15× slower |

Hucre's per-file speed is genuinely remarkable — hand-rolled SAX parsing of
OOXML in TypeScript, zero allocations in the hot loop. If raw read
throughput is your bottleneck, use it.

---

## Where `hucre` wins

| | `hucre` | `ks-xlsx-parser` |
|---|:---:|:---:|
| **Writes** xlsx/csv/ods (round-trip) | ✅ | ❌ read-only |
| **CSV / ODS / HTML** input | ✅ | ❌ xlsx / xlsm only |
| **Sparkline** extraction | ✅ | ❌ not modelled |
| **Chart round-trip preservation** (open → modify → save) | ✅ | ❌ read-only |
| **Edge runtime** (Cloudflare Workers / Deno / browser) | ✅ | ❌ Python-only |
| **Bundle size** | ~18 KB, zero deps | ~500 KB incl. deps |
| **Streaming row iterator** API | ✅ `streamXlsxRows` | ❌ full-workbook parse |
| **CSP-compliant, no eval** | ✅ | N/A (Python) |
| **Raw parse throughput** | ✅ 3-100× faster | ❌ |

---

## Where `ks-xlsx-parser` wins

| | `ks-xlsx-parser` | `hucre` |
|---|:---:|:---:|
| **Formula dependency graph** (topological, cycle detection via Tarjan's SCC) | ✅ | ❌ formula stored as string only |
| **Chart type + series extraction** (7 types: bar, line, pie, scatter, area, radar, bubble) | ✅ | ❌ round-trip preservation only |
| **Pivot table structure** (cache source, row/col/filter fields, slicer connections) | ✅ | ❌ listed as "No" |
| **RAG chunking** with configurable token budget | ✅ | ❌ no LLM positioning |
| **Source URIs** for citations (`file.xlsx#Sheet!A1:F18`) | ✅ | ❌ |
| **Sheet-purpose classification** (raw_data / dashboard / calc / …) | ✅ | ❌ |
| **KPI ranking** by formula connectivity + entity index | ✅ | ❌ |
| **Deterministic content hashes** (xxhash64 per cell / block / chunk) | ✅ | ❌ |
| **Adversarial-corpus robustness** | ✅ 1053/1053 parsed | ⚠️ 2 timeouts on pathological address-space files |
| **Stress corpus** (1053 workbooks checked into repo + CI round-trip) | ✅ | ❌ |

---

## Extraction-count agreement (1053 workbooks)

On every feature **both** parsers extract, the drift is zero or near-zero:

| feature | `hucre` | `ks-xlsx-parser` | drift |
|---|---:|---:|:---:|
| formulas | 46,411 | 46,433 | 0.05% |
| tables | 523 | 523 | **0** |
| merges | 10,488 | 10,488 | **0** |
| conditional-format rules | 70 | 70 | **0** |
| data validations | 503 | 503 | **0** |
| hyperlinks | 511 | 511 | **0** |
| comments | 486 | 486 | **0** |
| named ranges | 822 | 809 | 1.6% (tracked) |

The 22-formula disagreement is dominated by one workbook
(`real_world/Walbridge Coatings 8.9.23.xlsx`) where we parse 16 formulas
that hucre misses — we surface this in the drift report, not hide it.

The cell-count difference on adversarial merge-heavy files (we emit ~50%
more rows) is a **methodology difference**: `ks-xlsx-parser` counts every
addressable cell in a merged region; hucre counts the master cell only.
Both are defensible; document in the drift report generated by the
benchmark harness.

---

## Our accuracy commitment

Every perf change in `ks-xlsx-parser` has to pass, in order:

1. The **1631-test pytest suite** (unit + integration + corpus-slice)
2. **Cross-validation** against [`calamine`](https://github.com/tafia/calamine) — the Rust reference parser — on a golden fixture set
3. **Zero regressions** on the 1053-file testBench across eight sub-corpora (`real_world/`, `enterprise/`, `github_datasets/`, `stress/curated/`, `stress/merges/`, `generated/matrix/`, `generated/combo/`, `generated/adversarial/`)
4. **Feature-count stability** vs. the hucre benchmark above

That's the order. If a perf change breaks any gate, we don't ship it.
Every number on this page came from a run that passed all four gates.

If you're building RAG / agent / auditing pipelines where a silently
dropped formula or a misread merge is a user-visible bug, that order
matters. If you're shipping an I/O library for edge runtimes,
[**use hucre**](https://github.com/productdevbook/hucre) — it's the right
tool.

---

## Reproducing these numbers

The benchmark harness lives at [`tests/benchmarks/`](https://github.com/knowledgestack/ks-xlsx-parser/tree/main/tests/benchmarks).
Full details in [`tests/benchmarks/README`](https://github.com/knowledgestack/ks-xlsx-parser/tree/main/tests/benchmarks)
but the short version:

```bash
# From the repo root, in the ks-xlsx-parser venv
cd tests/benchmarks/hucre_node && pnpm install --frozen-lockfile
cd ../../..

# Full mode (default)
python -m tests.benchmarks.vs_hucre --corpus testBench --out tests/benchmarks/reports

# Fast mode
KS_PARSE_MODE=fast python -m tests.benchmarks.vs_hucre \
    --corpus testBench --out tests/benchmarks/reports
```

Outputs (under `tests/benchmarks/reports/<timestamp>_<git-sha>/`):

- `results.csv` — one row per `(file, parser)` pair
- `raw.ndjson` — full per-row records (nullable fields preserved)
- `failures.jsonl` — status != ok rows
- `summary.md` — aggregate counts, capability matrix, perf percentiles
- `drift.md` — per-feature disagreement between parsers
- `manifest.json` — run metadata (git sha, node / python versions, host, timestamp, CLI args)

The harness:

- Pins hucre exact (`0.3.0`, `--frozen-lockfile`) so numbers are reproducible
- Randomises `(file, parser)` ordering per seed to kill OS-page-cache bias
- Each parser times itself in-process; Python driver doesn't measure the other
- Per-file 60s timeout, 4 GB memory ceiling, worker respawn per 50-file batch
- Uses `null` (not `0`) for features a parser doesn't model — the summary generator distinguishes them

---

## Credit

This comparison wouldn't exist without [`hucre`](https://github.com/productdevbook/hucre)
and its author [**@productdevbook**](https://github.com/productdevbook).
Their work on a zero-dep TypeScript parser pushed us to actually measure
our perf floor and invest in the Rust fast-path, the Tarjan's SCC swap,
and `parse_workbook(mode='fast')`.

If you need a fast, tiny, edge-runtime xlsx / csv / ods library with
write support — that's them, not us.
