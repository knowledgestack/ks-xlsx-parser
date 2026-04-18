# Removing `ks_xlsx_core`

This Rust crate exists for one reason: **python-calamine (the Python binding)
does not expose formula strings as of 0.6.2**, only cached values via
`to_python()` / `iter_rows()`. Our Rust crate wraps the underlying Rust
`calamine` crate (which *does* expose formulas via `worksheet_formula`) and
re-exports what we need through PyO3.

## When to remove it

Check python-calamine release notes at
<https://github.com/dimastbk/python-calamine/releases>. Once a release exposes
formula strings per cell (the issue to track:
<https://github.com/dimastbk/python-calamine/issues> — search "formula"),
this crate is redundant.

## How to remove it

The isolation point is a single Python adapter module. When the dependency
flips, these are the only changes required:

1. **Delete this directory**: `rm -rf rust/ks_xlsx_core`.
2. **Drop the maturin build from CI** (wherever it's wired — today it's
   invoked as `maturin develop` inside `rust/ks_xlsx_core`).
3. **Rewrite the adapter** at `src/parsers/calamine_core.py` to call
   `python_calamine` directly for both values and formulas. The public
   surface of that module must not change — the rest of the parser only
   depends on `read_sheet_cells(path_or_bytes, sheet_name) -> SheetCells`.
4. **Remove the optional-dep entry** `ks-xlsx-core` from `pyproject.toml`.
5. **Run** `pytest tests/` to confirm nothing regresses.

No other code in `src/` imports `ks_xlsx_core` directly — that's by design.
Grep for `ks_xlsx_core` before removing to confirm the only references are
inside `src/parsers/calamine_core.py` and the Rust crate itself.

## Why not just use `maturin` on pypi's `python-calamine` source?

Because we'd need to fork it. This crate is intentionally thin: it
re-exports only what upstream is missing, with no divergent behaviour.
When upstream catches up, delete and move on.
