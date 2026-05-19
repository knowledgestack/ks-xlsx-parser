# ks-xlsx-parser v0.2.1 — Packaging hotfix 📦

**Headline:** v0.2.0 shipped a broken wheel. `from ks_xlsx_parser.pipeline import parse_workbook` (and every public entry point that depends on it) raised `ModuleNotFoundError` for everyone who installed from PyPI. **v0.2.1 fixes it.**

If you're on 0.2.0:

```bash
pip install --upgrade ks-xlsx-parser   # or: uv pip install --upgrade ks-xlsx-parser
python -c "from ks_xlsx_parser.pipeline import parse_workbook; print('ok')"
```

## What was broken

The source tree was laid out flat under `src/` — `pipeline.py` and `api.py` as top-level *modules*, with 13 sibling *packages* (`models`, `utils`, `parsers`, `analysis`, …) next to them. The setuptools build picked up the packages but **silently dropped the top-level modules**, because `[tool.setuptools.packages.find]` only finds packages, never modules. The published wheel's `top_level.txt`:

```
analysis annotation charts chunking comparison export formula
ks_xlsx_parser models parsers rendering storage utils verification
```

— 14 top-level entries, 13 of them generic. Two consequences:

1. `from ks_xlsx_parser.pipeline import ...` failed because `pipeline.py` was never copied into the wheel.
2. Anyone with an unrelated `models` / `utils` / `parsers` package in `site-packages` had it shadowed by ks-xlsx-parser's internal ones.

CI couldn't catch this because the test matrix used an editable install — `src/` lives on `sys.path` and the wheel-packaging defect is invisible.

## What changed in 0.2.1

* **Proper nested package.** Everything now lives under `src/ks_xlsx_parser/`. The wheel's `top_level.txt` contains only `ks_xlsx_parser`. Imports inside the package switched from `from pipeline import …` to relative / fully-qualified `ks_xlsx_parser.pipeline`.
* **`scripts/verify_wheel.py`** — builds the wheel, installs it in a clean venv, asserts the public import surface resolves and there's no namespace pollution. Wired into both `ci.yml` (new `wheel-check` job — required) and `release.yml` (a `Verify wheel` step that runs before PyPI publish). The class of bug that produced 0.2.0 cannot recur.
* **CI overhaul** — separate `test` / `wheel-check` / `lint` / `typecheck` jobs; `uv`-backed installs (faster); `make install-dev` alias.
* **Docker benchmark + accuracy tracking** — `Dockerfile.bench` reproduces the SpreadsheetBench retrieval benchmark; a new `.github/workflows/benchmark.yml` runs a 60-instance sample on every PR touching the parser and the full 912-corpus weekly. `scripts/append_bench_history.py` keeps a commit-over-commit history file. (Goal: text recall@5 > 0.90.)
* **Recall failure triage** — `eval_retrieval.py --emit-failures` + `scripts/triage_recall.py` turn "recall@5 = X" into a bucket histogram with exemplar failures. `docs/recall-investigation.md` documents the diagnosis framework.

## ⚠️ Breaking — if you depended on the leaked top-level packages

Anyone whose downstream code did:

```python
from models import WorkbookDTO          # ← used to "work" because of the leak
from parsers.workbook_parser import …
```

must update to:

```python
from ks_xlsx_parser.models import WorkbookDTO
from ks_xlsx_parser.parsers.workbook_parser import …
```

If you only used the documented public surface (`from ks_xlsx_parser import parse_workbook` and `ks_xlsx_parser.pipeline.parse_workbook`), nothing changes — those names still resolve, and now they actually *load*.

## Upgrade

```bash
pip install --upgrade ks-xlsx-parser
# or
uv pip install --upgrade ks-xlsx-parser
```

Then:

```python
from ks_xlsx_parser import parse_workbook
result = parse_workbook(path="report.xlsx")
print(result.workbook.total_cells)
for chunk in result.chunks:
    print(chunk.source_uri, chunk.render_text[:120])
```

## Verification used to ship this release

```
1041 passed, 11 deselected           # full test suite
verifying ks_xlsx_parser-0.2.1-py3-none-any.whl
wheel contents OK (61 entries, top-level: ks_xlsx_parser)
clean-venv import OK
wheel verification PASSED
```

`scripts/verify_wheel.py` ran as a release-workflow step before the PyPI publish, with the wheel that's now on PyPI.

## Thanks

Frank for the bug report — the channel screenshot was the right level of detail to reproduce in five minutes.

## See also

* [CHANGELOG.md](https://github.com/knowledgestack/ks-xlsx-parser/blob/main/CHANGELOG.md) — full diff log.
* [`docs/recall-investigation.md`](https://github.com/knowledgestack/ks-xlsx-parser/blob/main/docs/recall-investigation.md) — diagnosis framework for the recall→0.90 roadmap.
* [`docs/benchmark-local-setup.md`](https://github.com/knowledgestack/ks-xlsx-parser/blob/main/docs/benchmark-local-setup.md) — reproduce the benchmark on your laptop.
