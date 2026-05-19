# 05 · Out-of-scope: benchmark instances that require execution, not retrieval

**Status:** 🆓 free to claim (scoring/harness work — **NOT a parser fix**)
**Slice:** A (`scripts/eval_retrieval.py`, `scripts/enrich_failures.py`)
**Independent of:** 01, 02, 03, 04 — partial overlap with 00 on
`eval_retrieval.py`; coordinate via README.

## Why this exists

127 of 200 instances in the seed=1337 sample (≈ 63%) have
`instruction_requires_execution = True` — the `answer_position` in the
INPUT spreadsheet is empty because the question literally asks the
system to *compute and write* a value there.

Example instructions from this class:

* "Round the value 367.5 to the nearest $5 and put the result in C2."
* "Sum each row of B:F and fill column G."
* "Fix the broken formulas in D2:D613."

A retrieval-grade parser cannot satisfy these instances. The answer
doesn't exist in `input.xlsx` to be retrieved. Including these
instances in the headline `recall_text@5` number drags it down ~25
points without representing any parser quality signal.

**This TODO is not a parser improvement.** It's a benchmark-scoring
correction so that subsequent parser work has a clean signal.

## What the headline numbers actually mean today

On the 200-sample seed=1337:

|                         | All 200 | In-scope (73) |
|-------------------------|---------|---------------|
| Text recall@5           | 0.635   | 0.59          |
| Geometric recall@5      | 0.310   | (compute it)  |
| `instruction_requires_execution` | 127 of 200 | 0 (excluded) |

The "in-scope" column is the metric the planning roadmap targets. Today
it's hidden inside the eval output; this TODO surfaces it as a
first-class metric.

## What to ship

1. **A scoring filter switch** in `scripts/eval_retrieval.py`.
   - New optional `--exclude-execution-instances` flag (or compute both
     filtered and unfiltered metrics every run).
   - Definition of "execution instance": at run time, peek at
     `input.xlsx[answer_sheet][answer_position]`. If the union of
     non-empty cells in that range is empty, mark the instance as
     `instruction_requires_execution` and report it in a separate
     bucket. **Do this once per instance, not per parser** — the
     classification is parser-independent.
2. **Two recall numbers** in `summary.json` / `summary.md`:
   - `recall_text@5_all` (current behaviour — count every instance).
   - `recall_text@5_in_scope` (denominator excludes
     execution instances).
3. **History tracking**: include both numbers in
   `scripts/append_bench_history.py` so the trend chart shows the
   in-scope number — that's the target for the 0.90 goal.
4. **README + recall-investigation.md update** so the in-scope number
   is named and called out as the gate. The 0.90 target is on
   *in-scope* recall, not all-instances recall.
5. **Out-of-scope summary file** alongside the run reports:
   `tests/benchmarks/reports/retrieval/<stamp>/out_of_scope.txt` listing
   the instance IDs that were filtered, so the filter's behaviour is
   auditable. If somebody disagrees with the classification, they can
   diff the list.

## File scope

You may touch:

* `scripts/eval_retrieval.py` — add classification step BEFORE
  `score_instance` (so unfiltered metric still works).
* `scripts/enrich_failures.py` — make `instruction_requires_execution`
  the canonical source of truth (it already detects this).
* `scripts/append_bench_history.py` — record both numbers per run.
* `scripts/triage_recall.py` — by default exclude out-of-scope rows.
* `docs/planning/recall-90/README.md` — update target row.
* `docs/recall-investigation.md` — call out the in-scope number.

Do NOT touch any `src/ks_xlsx_parser/*` code for this cluster.

## Acceptance criteria

1. On the same 200-sample seed=1337 run, the new
   `recall_text@5_in_scope` number is emitted alongside the existing
   one and printed in `summary.md`.
2. The classifier reproducibly identifies ≥ 120 of the 127 currently-
   flagged instances (some borderline cases — a single non-empty
   header in the answer range — are debatable).
3. `history.jsonl` rows after the change have both fields populated.
4. A unit test under `tests/test_eval_retrieval_classify.py` covers
   the classifier with hand-built fixtures: empty range → out-of-scope,
   range with one value → in-scope, range with only string headers
   (no numeric data) → in-scope, range with formula cells that have
   cached values → in-scope.

## Why this isn't TODO 00

* TODO 00 fixes the harness's *spec parser* — wrong sheet names get
  resolved.
* TODO 05 fixes the harness's *scoring* — instances the parser can't
  affect get a separate bucket.

Both touch `eval_retrieval.py` but in different functions. If they're
worked simultaneously, coordinate over `parse_position_spec` —
TODO 00 owns the regex; TODO 05 owns the classifier.

## Pitfalls

* Don't over-exclude. An instance where `answer_position` covers
  `A2:G15` and only `A1` is non-empty (a single header row above the
  empty area) is still execution-required, but use the right cell
  range — read the actual `answer_position`, not `data_position`.
* The classifier needs the openpyxl read, which costs ~50ms per
  instance. Cache it or do it once at script start, not inside the
  ranking loop.
* Don't delete the unfiltered metric. Both numbers are useful: the
  unfiltered one tells you how a downstream consumer experiences
  ks-xlsx-parser on real spreadsheet questions (most of which DO
  require execution); the in-scope one tells you whether the parser
  is doing its job.

## Sample of what gets filtered

Random sampling of `instruction_requires_execution = True` instances:

| Instance | Instruction (first 60 chars) | Answer range |
|---|---|---|
| ...      | "Calculate the sum of column B for products where..."   | `Sheet1!E2:E10` |
| ...      | "For each row in column A, look up the corresponding..."| `Sheet1!C2:C50` |
| ...      | "Pivot the data in A1:D100 by month..."                 | `Sheet2!A1:E13` |

Run the enricher to see the full list.
