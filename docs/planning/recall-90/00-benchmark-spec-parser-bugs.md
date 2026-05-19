# 00 · Benchmark harness fails to parse malformed `answer_position` strings

**Status:** 🆓 free to claim
**Slice:** A (benchmark harness — `scripts/eval_retrieval.py`)
**Independent of:** all other TODOs in this folder

## What it looks like

Eight instances in the 200-sample (seed=1337) score as recall@5 misses
even though the parser produced perfectly reasonable chunks. The
benchmark harness's `parse_position_spec()` returns no regions because
the `answer_position` string in `dataset.json` doesn't match
`SHEET_RANGE_RE`. The score code then sees `data_regions=0`, marks the
instance as "no scorable target", and the recall denominator counts it
as a miss. Net effect: the parser is blamed for benchmark-spec parsing
bugs.

Concrete malformed strings observed:

| Instance | `answer_position` string | What's wrong |
|---|---|---|
| `50442`  | `RESULTS 1'!G17`                                  | unmatched closing quote, opening quote missing |
| `49490`  | `Sheet1'!H3:H58`                                  | same |
| `49036`  | `Dashboard'!B8`                                   | same |
| `55427`  | `Compiled and located schools da'!B2:B1461`       | same + long sheet name |
| `48975`  | `Output'!B11:B17`                                 | same |
| `37456`  | `G12：J15`                                        | fullwidth Chinese colon `：` instead of `:` |
| `184-6`  | `'RAWDATA!'A1:P6,'OUTPUT!'A1:P6'`                 | quote between sheet and `!`, comma-joined |
| `164-22` | `'YEAR1!'A1:G1478,'YEAR2!'A1:G1480'`              | same shape as 184-6 |

(`374-9` may also belong here — investigate.)

## Diagnostic signature

In `enriched_failures.ndjson`, members of this cluster have:

* `gt_sheet == null` (spec parser couldn't extract the sheet name), OR
* `gt_sheet` wrapped in apostrophes (`"'Sheet1'"`), AND
* the workbook *does* contain a sheet with the bare name.

```bash
python scripts/enrich_failures.py tests/benchmarks/reports/retrieval
jq -c 'select(.gt_sheet == null or (.gt_sheet | startswith("'")))' \
    tests/benchmarks/reports/retrieval/*/enriched_failures.ndjson
```

## File scope

You may touch:

* `scripts/eval_retrieval.py` — `SHEET_RANGE_RE`, `parse_position_spec`,
  `_normalize_value_for_match`. This is where the bugs live.
* `scripts/enrich_failures.py` — must stay consistent with the harness;
  re-run after editing.
* Add new tests under `tests/test_eval_retrieval.py` (or similar — file
  doesn't exist yet; create it).

Do **NOT** touch any `src/ks_xlsx_parser/*` code for this cluster. If
the urge to "make the parser tolerate these patterns too" arises,
that's a separate TODO and a category error here — these are dataset
typos, not parser inputs.

## Acceptance criteria

On `python scripts/eval_retrieval.py --corpus data/corpora/spreadsheetbench/all_data_912_v0.1 --parsers ks --sample 200 --seed 1337 --emit-failures`:

1. All 8 instances listed above stop appearing as recall@5 misses
   (text or geometric) on the diagnostic ndjson.
2. `geometric@5` rises by ≥ 4 pp (each fixed instance flips from miss
   to hit since the chunks already overlap the correct region).
3. No previously-passing instance becomes a miss. Diff the
   `results.ndjson` before and after — `pass → fail` count must be 0.
4. `python scripts/enrich_failures.py` shows zero rows with
   `gt_sheet == null` from the eight listed instances.

## Failing test sketch

```python
# tests/test_eval_retrieval_spec_parser.py
import pytest
from scripts.eval_retrieval import parse_position_spec

@pytest.mark.parametrize("spec, expected_sheet", [
    ("Dashboard'!B8", "Dashboard"),
    ("Sheet1'!H3:H58", "Sheet1"),
    ("'RAWDATA!'A1:P6", "RAWDATA"),
    ("G12：J15", None),  # this one keeps default_sheet; sheet unaltered
])
def test_unmatched_quotes_resolve_sheet(spec, expected_sheet):
    regions = parse_position_spec(spec, default_sheet="DEFAULT")
    assert regions, f"no regions for {spec!r}"
    if expected_sheet is not None:
        assert regions[0][0] == expected_sheet

def test_fullwidth_colon_in_range():
    # Excel-China dataset entries occasionally use the fullwidth colon.
    regions = parse_position_spec("G12：J15", default_sheet="S")
    assert regions and regions[0][1] == (12, 7, 15, 10)
```

These should fail on `main`. They pass when the fix lands.

## Pitfalls

* `'YEAR1!'A1:G1478,'YEAR2!'A1:G1480'` is multi-region. Make sure the
  comma-split path still returns BOTH regions after the fix.
* `_normalize_value_for_match` is not involved here — don't refactor it.
* The current regex `SHEET_RANGE_RE` uses a backreference for matching
  quotes; tightening the closing quote to "optional, but only if opening
  present" is fragile. Consider a two-pass approach: strip stray
  apostrophes, then re-attempt the strict regex.

## Repro fixtures

The eight instances live under
`data/corpora/spreadsheetbench/all_data_912_v0.1/spreadsheet/<id>/`.
Quote one as a unit-test data point — do NOT vendor the corpus into
the repo. If a snapshot of the malformed strings is wanted, hand-copy
them into a Python fixture (they're ≤ 100 chars each).
