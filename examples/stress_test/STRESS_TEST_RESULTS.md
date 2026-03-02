# Parser Stress Test Results

## Summary

- Stress levels tested: 0..25 (+ 5 merge-stress files)
- Passed: 31
- Failed: 0

## Per-Level Results

| Level | Pass | Cells | Chunks | Tables | Charts | Errors |
|-------|------|-------|--------|--------|--------|--------|
| 0 | ✓ | 2 | 1 | 0 | 0 | 0 |
| 1 | ✓ | 6 | 2 | 0 | 0 | 0 |
| 2 | ✓ | 21 | 3 | 1 | 0 | 0 |
| 3 | ✓ | 21 | 4 | 1 | 1 | 0 |
| 4 | ✓ | 22 | 4 | 1 | 1 | 0 |
| 5 | ✓ | 26 | 5 | 1 | 1 | 0 |
| 6 | ✓ | 32 | 7 | 1 | 1 | 0 |
| 7 | ✓ | 140 | 10 | 2 | 2 | 0 |
| 8 | ✓ | 140 | 10 | 2 | 2 | 0 |
| 9 | ✓ | 143 | 13 | 2 | 2 | 0 |
| 10 | ✓ | 145 | 15 | 2 | 2 | 0 |
| 11 | ✓ | 155 | 16 | 2 | 2 | 0 |
| 12 | ✓ | 157 | 18 | 2 | 3 | 0 |
| 13 | ✓ | 159 | 19 | 2 | 3 | 0 |
| 14 | ✓ | 160 | 20 | 2 | 3 | 0 |
| 15 | ✓ | 160 | 20 | 2 | 3 | 0 |
| 16 | ✓ | 166 | 21 | 2 | 3 | 0 |
| 17 | ✓ | 166 | 21 | 2 | 3 | 0 |
| 18 | ✓ | 167 | 22 | 2 | 3 | 0 |
| 19 | ✓ | 317 | 23 | 2 | 3 | 0 |
| 20 | ✓ | 319 | 24 | 2 | 3 | 0 |
| 21 | ✓ | 334 | 25 | 3 | 3 | 0 |
| 22 | ✓ | 734 | 26 | 3 | 3 | 0 |
| 23 | ✓ | 934 | 27 | 3 | 3 | 0 |
| 24 | ✓ | 984 | 37 | 3 | 3 | 0 |
| 25 | ✓ | 985 | 38 | 3 | 3 | 0 |
| 0 | ✓ | 300 | 1 | 0 | 0 | 0 |
| 1 | ✓ | 900 | 1 | 0 | 0 | 0 |
| 2 | ✓ | 2 | 1 | 0 | 0 | 0 |
| 3 | ✓ | 33 | 1 | 1 | 0 | 0 |
| 4 | ✓ | 320 | 1 | 0 | 0 | 0 |

## Failures & Issues

## Documented Issues (Known Limitations)

### merge_empty_master
- stress_level_20.xlsx (EmptyMaster!A1)
- stress_level_21.xlsx (EmptyMaster!A1)
- stress_level_22.xlsx (EmptyMaster!A1)
- stress_level_23.xlsx (EmptyMaster!A1)
- stress_level_24.xlsx (EmptyMaster!A1)
- stress_level_25.xlsx (EmptyMaster!A1)
- merge_stress_empty_master.xlsx (EmptyMaster!A1)

## Identified Pipeline Problems

- No hard failures in tested range.
