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

None at this time.

---

## Quick Reference

| Issue ID             | Status   | Description                                                                       |
|----------------------|----------|-----------------------------------------------------------------------------------|
| `merge_empty_master` | RESOLVED | Value was in right cell, merge created left — now recovered via OOXML XML fallback |
