# excel_core

Rust + PyO3 fast-path for `excel-parser`. Wraps the `calamine` Rust crate
and exposes cell values **and formulas** to Python.

This exists because `python-calamine` (0.6.2) returns cached values only —
no formula strings. The underlying Rust `calamine` crate does expose them.

**This crate is removable.** See [REMOVAL.md](REMOVAL.md) for the upgrade
path when upstream Python bindings gain formula support.

## Build

From the repo root, inside the venv:

```
cd rust/excel_core
maturin develop --release
```

`--release` is important — debug builds are ~10× slower and benchmark
numbers will be misleading.

## Python API

Exactly one function:

```python
import excel_core

sheets: list[SheetData] = excel_core.read_workbook("workbook.xlsx")
# SheetData: {name: str, cells: list[(row, col, value, formula, dtype)]}
```

All consumers in `excel-parser` go through `src/parsers/calamine_core.py`,
never through this module directly.
