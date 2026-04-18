//! ks_xlsx_core — PyO3 bindings exposing `calamine`'s cell values **and**
//! formula strings to Python.
//!
//! Rationale lives in REMOVAL.md. Scope is intentionally minimal: one entry
//! point, `read_workbook(path) -> list[SheetData]`, returning primitive
//! Python objects (tuples of (row, col, value, formula, dtype)). Anything
//! richer (styles, CF, DV, charts) stays on the Python side via openpyxl so
//! this crate has a narrow, deletable surface.
//!
//! Coordinates returned are **1-indexed** to match openpyxl, avoiding any
//! conversion on the Python side.

mod formula;

use calamine::{open_workbook_auto, Data, Reader, Sheets};
use pyo3::exceptions::PyIOError;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList, PyTuple};

/// Convert a `calamine::Data` cell value into a Python object.
fn data_to_py(py: Python<'_>, d: &Data) -> PyObject {
    match d {
        Data::Empty => py.None(),
        Data::String(s) => s.to_object(py),
        Data::Float(f) => f.to_object(py),
        Data::Int(i) => i.to_object(py),
        Data::Bool(b) => b.to_object(py),
        Data::DateTime(dt) => {
            // Emit the numeric serial; Python side has the full locale-aware
            // conversion machinery (date system 1900 vs 1904).
            dt.as_f64().to_object(py)
        }
        Data::DateTimeIso(s) => s.to_object(py),
        Data::DurationIso(s) => s.to_object(py),
        Data::Error(e) => format!("#{:?}", e).to_object(py),
    }
}

/// One-character type tag mirroring openpyxl's `data_type`:
/// `s` string, `n` numeric, `b` bool, `d` date, `e` error, `f` formula, `-` empty.
/// Formulas are tagged `f` regardless of their cached-value type — matches
/// openpyxl's convention and is what `cell_parser.py` expects.
fn data_type_tag(d: &Data, has_formula: bool) -> &'static str {
    if has_formula {
        return "f";
    }
    match d {
        Data::Empty => "-",
        Data::String(_) => "s",
        Data::Float(_) | Data::Int(_) => "n",
        Data::Bool(_) => "b",
        Data::DateTime(_) | Data::DateTimeIso(_) | Data::DurationIso(_) => "d",
        Data::Error(_) => "e",
    }
}

#[pyfunction]
#[pyo3(signature = (path))]
fn read_workbook(py: Python<'_>, path: &str) -> PyResult<PyObject> {
    let mut wb: Sheets<_> = open_workbook_auto(path)
        .map_err(|e| PyIOError::new_err(format!("calamine open failed: {e}")))?;

    let sheet_names: Vec<String> = wb.sheet_names().to_vec();
    let result = PyList::empty_bound(py);

    for name in sheet_names {
        // Values (cached formula results + literals).
        let values = match wb.worksheet_range(&name) {
            Ok(r) => r,
            Err(_) => continue, // skip sheets calamine can't read
        };
        // Formulas (strings without leading '='; empty where no formula).
        let formulas = wb.worksheet_formula(&name).ok();

        let cells = PyList::empty_bound(py);
        let (v_start_row, v_start_col) = values.start().unwrap_or((0, 0));
        let (f_start_row, f_start_col) = formulas
            .as_ref()
            .and_then(|f| f.start())
            .unwrap_or((0, 0));

        for (r, c, v) in values.cells() {
            // Absolute 1-indexed coordinates for openpyxl parity.
            let abs_row = (v_start_row + r as u32) + 1;
            let abs_col = (v_start_col + c as u32) + 1;

            let formula_opt: Option<String> = formulas.as_ref().and_then(|f| {
                // Formula grid may have a different origin — translate to
                // the value grid's frame before indexing.
                let fr = (v_start_row as i64 + r as i64) - f_start_row as i64;
                let fc = (v_start_col as i64 + c as i64) - f_start_col as i64;
                if fr < 0 || fc < 0 {
                    return None;
                }
                f.get_value((fr as u32, fc as u32)).and_then(|s| {
                    if s.is_empty() {
                        None
                    } else {
                        Some(s.clone())
                    }
                })
            });

            let has_formula = formula_opt.is_some();
            // Skip genuinely empty cells (no value AND no formula). This
            // mirrors openpyxl's sparse storage behaviour.
            if matches!(v, Data::Empty) && !has_formula {
                continue;
            }

            let tag = data_type_tag(v, has_formula);
            let value_obj = data_to_py(py, v);
            let formula_obj = match formula_opt {
                Some(s) => s.to_object(py),
                None => py.None(),
            };

            let tup = PyTuple::new_bound(
                py,
                [
                    abs_row.to_object(py),
                    abs_col.to_object(py),
                    value_obj,
                    formula_obj,
                    tag.to_object(py),
                ],
            );
            cells.append(tup)?;
        }

        // Merged cells: calamine's Range doesn't expose them directly on
        // xlsx; we rely on openpyxl for merges today. Future: lift the
        // merged_cells list out of xl/worksheets/sheet_N.xml here.
        let sheet_obj = PyDict::new_bound(py);
        sheet_obj.set_item("name", &name)?;
        sheet_obj.set_item("cells", cells)?;
        result.append(sheet_obj)?;
    }

    Ok(result.into())
}

#[pymodule]
fn ks_xlsx_core(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(read_workbook, m)?)?;
    m.add_function(wrap_pyfunction!(formula::scan_formula, m)?)?;
    m.add("__version__", env!("CARGO_PKG_VERSION"))?;
    Ok(())
}
