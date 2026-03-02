# Parser Known Issues & Limitations

This document lists Excel files that fail to parse correctly or exhibit known limitations.
Use this when working with the XLSX parser to understand edge cases.

---

## Files That Fail (Hard Errors)

None. All stress-tested files parse without hard errors.

---

## Documented Limitations (No Hard Fail)

These files exhibit known openpyxl/parser quirks. The pipeline completes but output may be incomplete.

### merge_empty_master

**Why**: Value was in non-master cell before merge; openpyxl MergedCell has no value, content is lost.

**Affected files**:
- `stress_level_20.xlsx` — EmptyMaster!A1
- `stress_level_21.xlsx` — EmptyMaster!A1
- `stress_level_22.xlsx` — EmptyMaster!A1
- `stress_level_23.xlsx` — EmptyMaster!A1
- `stress_level_24.xlsx` — EmptyMaster!A1
- `stress_level_25.xlsx` — EmptyMaster!A1
- `merge_stress_empty_master.xlsx` — EmptyMaster!A1

---

## Quick Reference

| Issue ID | Description |
|----------|-------------|
| `merge_empty_master` | Value was in right cell, merge created left→content lost in openpyxl |
