# Parser Known Issues & Limitations

This document lists Excel files that fail to parse correctly or exhibit known limitations.
Use this when working with the XLSX parser to understand edge cases.

---

## Files That Fail (Hard Errors)

None. All stress-tested files parse without hard errors.

---

## Resolved Issues

### merge_empty_master — RESOLVED

**Status**: Fixed via raw OOXML recovery fallback.

**Original problem**: Value was in non-master cell before merge; openpyxl
`MergedCell` has no value, content was lost.

**Fix**: When a merged region's master cell has no value, the parser opens
the `.xlsx` as a ZIP, parses the sheet XML with `lxml`, and recovers values
from any `<c>` element within the merge range. The recovered value is
promoted to the master cell.

**Previously affected files**:
- `stress_level_20.xlsx` — EmptyMaster!A1
- `stress_level_21.xlsx` — EmptyMaster!A1
- `stress_level_22.xlsx` — EmptyMaster!A1
- `stress_level_23.xlsx` — EmptyMaster!A1
- `stress_level_24.xlsx` — EmptyMaster!A1
- `stress_level_25.xlsx` — EmptyMaster!A1
- `merge_stress_empty_master.xlsx` — EmptyMaster!A1

---

## Documented Limitations (No Hard Fail)

### `Walbridge Coatings 8.9.23.xlsx` — formula cached-value drift

**Symptom**: ~11% of formula cells in this real-world workbook produce a
different cached value than calamine reads. Hard failures are zero; parsing
and serialization succeed end-to-end.

**Root cause**: openpyxl's `data_only=True` reader does not always surface the
most recently written cached value for complex dynamic-array or volatile
formulas when the calc chain references across multiple sheets. This is an
openpyxl limitation, not an ks-xlsx-parser bug; calamine reads from the raw XML
and catches the newer values.

**Current mitigation**: `tests/test_cross_validation.py::test_formula_cached_values_match`
uses a 15% threshold for files in a `known_loose_files` set and the default
5% threshold for everything else.

**Potential fixes** (tracked):
1. Read cached values directly from the OOXML XML instead of via openpyxl (like
   we already do for empty merge masters).
2. Re-evaluate formulas ourselves via a sandboxed evaluator.

---

## Quick Reference

| Issue ID             | Status   | Description                                                                       |
|----------------------|----------|-----------------------------------------------------------------------------------|
| `merge_empty_master` | RESOLVED | Value was in right cell, merge created left — now recovered via OOXML XML fallback |
