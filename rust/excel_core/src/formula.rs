//! Hand-rolled Excel formula reference scanner.
//!
//! Returns tuples of primitives; all DTO construction stays on the Python
//! side so behavior is preserved byte-for-byte with
//! `src/formula/formula_parser.py`. We only replace the three regex scans.
//!
//! Emitted tuple layout (per reference):
//!   (kind: &str, sheet: Option<String>, col1: Option<u32>, row1: Option<u32>,
//!    col2: Option<u32>, row2: Option<u32>,
//!    external_workbook: Option<String>, table_name: Option<String>,
//!    ref_string: String)
//!
//! `kind` is one of: "cell", "range", "external_cell", "external_range",
//! "structured".
//!
//! Dedup is preserved by emitting each distinct `ref_string` only once,
//! matching the Python implementation's `seen: set[str]` behavior.

use pyo3::prelude::*;
use pyo3::types::{PyList, PyTuple};
use std::collections::HashSet;

/// Convert A..ZZZ column letters to a 1-indexed column number.
fn col_letters_to_num(letters: &[u8]) -> u32 {
    let mut n: u32 = 0;
    for &c in letters {
        let v = (c as u32).saturating_sub(b'A' as u32) + 1;
        n = n.saturating_mul(26).saturating_add(v);
    }
    n
}

/// Recognizer helpers — indexing directly on bytes keeps the hot loop tight.

#[inline]
fn is_ascii_upper(b: u8) -> bool {
    b.is_ascii_uppercase()
}

#[inline]
fn is_ascii_alpha_lower_upper(b: u8) -> bool {
    b.is_ascii_alphabetic()
}

#[inline]
fn is_ident_start(b: u8) -> bool {
    b.is_ascii_alphabetic() || b == b'_'
}

#[inline]
fn is_ident_cont(b: u8) -> bool {
    b.is_ascii_alphanumeric() || b == b'_' || b == b'.'
}

#[inline]
fn is_digit(b: u8) -> bool {
    b.is_ascii_digit()
}

/// Excel built-ins we refuse to classify as structured table refs.
fn is_function_name(name: &str) -> bool {
    matches!(
        name.to_ascii_uppercase().as_str(),
        "SUM"
            | "AVERAGE"
            | "COUNT"
            | "COUNTA"
            | "COUNTIF"
            | "COUNTIFS"
            | "SUMIF"
            | "SUMIFS"
            | "VLOOKUP"
            | "HLOOKUP"
            | "INDEX"
            | "MATCH"
            | "IF"
            | "IFERROR"
            | "IFNA"
            | "IFS"
            | "SWITCH"
            | "LEFT"
            | "RIGHT"
            | "MID"
            | "LEN"
            | "TRIM"
            | "UPPER"
            | "LOWER"
            | "CONCATENATE"
            | "CONCAT"
            | "TEXTJOIN"
            | "TEXT"
            | "VALUE"
            | "MAX"
            | "MIN"
            | "LARGE"
            | "SMALL"
            | "RANK"
            | "ABS"
            | "ROUND"
            | "ROUNDUP"
            | "ROUNDDOWN"
            | "CEILING"
            | "FLOOR"
            | "MOD"
            | "POWER"
            | "SQRT"
            | "LOG"
            | "LN"
            | "EXP"
            | "AND"
            | "OR"
            | "NOT"
            | "TRUE"
            | "FALSE"
            | "DATE"
            | "TODAY"
            | "NOW"
            | "YEAR"
            | "MONTH"
            | "DAY"
            | "HOUR"
            | "MINUTE"
            | "SECOND"
            | "DATEVALUE"
            | "TIMEVALUE"
            | "OFFSET"
            | "INDIRECT"
            | "ROW"
            | "COLUMN"
            | "ROWS"
            | "COLUMNS"
            | "TRANSPOSE"
            | "SORT"
            | "FILTER"
            | "UNIQUE"
            | "SEQUENCE"
            | "XLOOKUP"
            | "XMATCH"
            | "LET"
            | "LAMBDA"
            | "MAP"
            | "REDUCE"
            | "CHOOSE"
            | "HYPERLINK"
            | "TYPE"
            | "ISBLANK"
            | "ISERROR"
            | "ISNA"
            | "ISNUMBER"
            | "ISTEXT"
            | "ISLOGICAL"
            | "ISREF"
            | "SUMPRODUCT"
            | "MMULT"
            | "MINVERSE"
            | "NPV"
            | "IRR"
            | "PMT"
            | "PV"
            | "FV"
            | "RATE"
            | "NPER"
            | "STDEV"
            | "VAR"
            | "MEDIAN"
            | "MODE"
            | "PERCENTILE"
            | "SUBSTITUTE"
            | "REPLACE"
            | "FIND"
            | "SEARCH"
            | "REPT"
            | "NUMBERVALUE"
            | "FIXED"
            | "DOLLAR"
            | "ADDRESS"
            | "CELL"
            | "INFO"
            | "N"
            | "NA"
            | "ARRAYTOTEXT"
            | "VALUETOTEXT"
            | "PI"
            | "RAND"
            | "RANDBETWEEN"
    )
}

/// Consume an A1-style cell ref starting at `start`. Returns `(end, col, row)`
/// or `None` if the next chars don't form a valid cell ref.
fn consume_a1(bytes: &[u8], start: usize) -> Option<(usize, u32, u32)> {
    let mut i = start;
    if i < bytes.len() && bytes[i] == b'$' {
        i += 1;
    }
    let col_start = i;
    let mut col_len = 0;
    while i < bytes.len() && is_ascii_upper(bytes[i]) && col_len < 3 {
        i += 1;
        col_len += 1;
    }
    if col_len == 0 {
        return None;
    }
    let col_end = i;
    if i < bytes.len() && bytes[i] == b'$' {
        i += 1;
    }
    let row_start = i;
    let mut row_len = 0;
    while i < bytes.len() && is_digit(bytes[i]) && row_len < 7 {
        i += 1;
        row_len += 1;
    }
    if row_len == 0 {
        return None;
    }
    let col = col_letters_to_num(&bytes[col_start..col_end]);
    let row: u32 = std::str::from_utf8(&bytes[row_start..i])
        .ok()?
        .parse()
        .ok()?;
    Some((i, col, row))
}

/// Consume an optional sheet prefix ending at `!`. Returns the sheet name and
/// the post-`!` index, or `None` if no sheet prefix is present here.
fn consume_sheet_prefix(bytes: &[u8], start: usize) -> Option<(usize, String)> {
    // Quoted sheet: 'name'!
    if start < bytes.len() && bytes[start] == b'\'' {
        let mut i = start + 1;
        let name_start = i;
        while i < bytes.len() && bytes[i] != b'\'' {
            i += 1;
        }
        if i >= bytes.len() || bytes[i] != b'\'' {
            return None;
        }
        let name_end = i;
        i += 1; // past closing quote
        if i >= bytes.len() || bytes[i] != b'!' {
            return None;
        }
        let name = std::str::from_utf8(&bytes[name_start..name_end])
            .ok()?
            .to_string();
        return Some((i + 1, name));
    }
    // Unquoted sheet: name!
    let mut i = start;
    if i >= bytes.len() || !is_ident_start(bytes[i]) {
        return None;
    }
    let name_start = i;
    while i < bytes.len() && (is_ident_cont(bytes[i])) {
        i += 1;
    }
    if i >= bytes.len() || bytes[i] != b'!' {
        return None;
    }
    let name_end = i;
    // Must be followed by something that looks like a ref: [A-Z$] or another
    // sheet start. Guard against matching "FUNCNAME!" which shouldn't happen
    // in real formulas but defensive.
    let name = std::str::from_utf8(&bytes[name_start..name_end])
        .ok()?
        .to_string();
    Some((i + 1, name))
}

/// Scan a formula and emit reference tuples. Matches the behaviour of
/// `src/formula/formula_parser.py::parse` modulo the ordering of dedup
/// (which doesn't matter — callers process the returned list as a set).
#[pyfunction]
pub fn scan_formula(py: Python<'_>, formula: &str) -> PyResult<PyObject> {
    let bytes = formula.as_bytes();
    let n = bytes.len();
    let out = PyList::empty_bound(py);
    let mut seen: HashSet<String> = HashSet::new();

    // ---------- Pass 1: external refs `[book]Sheet!Ref` ------------------
    let mut i = 0;
    while i < n {
        if bytes[i] != b'[' {
            i += 1;
            continue;
        }
        // Find matching ']'
        let br_start = i + 1;
        let mut j = br_start;
        while j < n && bytes[j] != b']' {
            j += 1;
        }
        if j >= n {
            break;
        }
        let workbook = std::str::from_utf8(&bytes[br_start..j])
            .unwrap_or("")
            .to_string();
        let mut k = j + 1;
        // Sheet prefix required
        let (after_sheet, sheet_name) = match consume_sheet_prefix(bytes, k) {
            Some(x) => x,
            None => {
                i = k;
                continue;
            }
        };
        k = after_sheet;
        let (ref_end, col1, row1) = match consume_a1(bytes, k) {
            Some(x) => x,
            None => {
                i = k;
                continue;
            }
        };
        // Optional :end
        let mut end = ref_end;
        let mut col2: Option<u32> = None;
        let mut row2: Option<u32> = None;
        if end < n && bytes[end] == b':' {
            if let Some((e, c2, r2)) = consume_a1(bytes, end + 1) {
                end = e;
                col2 = Some(c2);
                row2 = Some(r2);
            }
        }
        let ref_str = std::str::from_utf8(&bytes[i..end]).unwrap_or("").to_string();
        if seen.insert(ref_str.clone()) {
            let kind = if col2.is_some() { "external_range" } else { "external_cell" };
            let tup = PyTuple::new_bound(
                py,
                [
                    kind.to_object(py),
                    sheet_name.to_object(py),
                    col1.to_object(py),
                    row1.to_object(py),
                    col2.to_object(py),
                    row2.to_object(py),
                    workbook.to_object(py),
                    py.None(),
                    ref_str.to_object(py),
                ],
            );
            out.append(tup)?;
        }
        i = end;
    }

    // ---------- Pass 2: structured table refs `Name[...]` ----------------
    let mut i = 0;
    while i < n {
        if !is_ident_start(bytes[i]) {
            i += 1;
            continue;
        }
        let name_start = i;
        while i < n && is_ident_cont(bytes[i]) {
            i += 1;
        }
        let name_end = i;
        // Must be followed by `[`
        if i >= n || bytes[i] != b'[' {
            continue;
        }
        let name = std::str::from_utf8(&bytes[name_start..name_end])
            .unwrap_or("")
            .to_string();
        if is_function_name(&name) {
            // Skip - that's a function call, not a structured ref
            continue;
        }
        // Find matching `]`, respecting nested brackets (Table[[#Headers],[Col]])
        let open = i;
        let mut depth = 1;
        i += 1;
        while i < n && depth > 0 {
            if bytes[i] == b'[' {
                depth += 1;
            } else if bytes[i] == b']' {
                depth -= 1;
            }
            i += 1;
        }
        if depth != 0 {
            break;
        }
        let ref_str = std::str::from_utf8(&bytes[name_start..i])
            .unwrap_or("")
            .to_string();
        if seen.insert(ref_str.clone()) {
            let tup = PyTuple::new_bound(
                py,
                [
                    "structured".to_object(py),
                    py.None(),
                    py.None(),
                    py.None(),
                    py.None(),
                    py.None(),
                    py.None(),
                    name.to_object(py),
                    ref_str.to_object(py),
                ],
            );
            out.append(tup)?;
        }
    }

    // ---------- Pass 3: plain A1 refs (with optional sheet prefix) -------
    let mut i = 0;
    while i < n {
        // Skip over what we already ate as external refs to avoid partial
        // matches like "Sheet1!A1" inside `[book]Sheet1!A1`. The Python
        // version checks `formula[pos - 1] == "]"`; we do the same.
        let save_i = i;
        let sheet_opt;
        let ref_text_start = i;
        let mut after_sheet = i;
        if let Some((k, name)) = consume_sheet_prefix(bytes, i) {
            sheet_opt = Some(name);
            after_sheet = k;
        } else {
            sheet_opt = None;
        }
        let (ref_end, col1, row1) = match consume_a1(bytes, after_sheet) {
            Some(x) => x,
            None => {
                i = save_i + 1;
                continue;
            }
        };
        // Guard: don't re-match a ref that was preceded by `]` (already
        // captured in the external-ref pass).
        if ref_text_start > 0 && bytes[ref_text_start - 1] == b']' {
            i = ref_end;
            continue;
        }
        let mut end = ref_end;
        let mut col2: Option<u32> = None;
        let mut row2: Option<u32> = None;
        if end < n && bytes[end] == b':' {
            // Optional sheet prefix inside range end (rare)
            let range_end_sheet_after =
                consume_sheet_prefix(bytes, end + 1).map(|(k, _)| k).unwrap_or(end + 1);
            if let Some((e, c2, r2)) = consume_a1(bytes, range_end_sheet_after) {
                end = e;
                col2 = Some(c2);
                row2 = Some(r2);
            }
        }
        let ref_str = std::str::from_utf8(&bytes[ref_text_start..end])
            .unwrap_or("")
            .to_string();
        if seen.insert(ref_str.clone()) {
            let kind = if col2.is_some() { "range" } else { "cell" };
            let tup = PyTuple::new_bound(
                py,
                [
                    kind.to_object(py),
                    match sheet_opt {
                        Some(s) => s.to_object(py),
                        None => py.None(),
                    },
                    col1.to_object(py),
                    row1.to_object(py),
                    col2.to_object(py),
                    row2.to_object(py),
                    py.None(),
                    py.None(),
                    ref_str.to_object(py),
                ],
            );
            out.append(tup)?;
        }
        i = end.max(save_i + 1);
    }

    Ok(out.into())
}

#[allow(unused_imports)]
use pyo3::wrap_pyfunction;
