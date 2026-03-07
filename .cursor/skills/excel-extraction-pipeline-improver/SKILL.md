---
name: excel-extraction-pipeline-improver
description: Iteratively improves the xlsx_parser extraction pipeline using feedback, test failures, and golden outputs. Use when fixing extraction gaps, addressing parser feedback, or improving coverage on stress workbooks. Follows TDD with instrumentation.
---

# Feedback-Driven Extraction Pipeline Iteration

Given extraction output, test failures, and user feedback, iteratively improve the Excel ingestion pipeline until it passes golden tests and reaches high coverage.

## Inputs

| Input | Description |
|-------|-------------|
| `repo_path` | Path to the xlsx_parser repository |
| `stress_workbooks_dir` | Directory containing stress test workbooks |
| `golden_expected_dir` | Golden expectations (`expected_extraction_min.json`) |
| `current_output_dir` | Latest extraction output |
| `feedback` | Human notes, failure logs, diffs |
| `time_budget_minutes` | Default 120 |

## Operating Procedure (TDD + Instrumentation)

### 1. Baseline

- Run unit and integration tests
- Run extraction on stress workbooks
- Produce coverage report with:
  - % cells with value
  - % cells with formula
  - % tables detected
  - % charts detected
  - % objects with anchors/bboxes
  - % conditional formats captured

### 2. Triage

Classify each failure:

| Class | Meaning |
|-------|---------|
| **Parsing gap** | Not extracted |
| **Schema gap** | Extracted but not represented in DTO |
| **Normalization gap** | Extracted but wrong type / wrong display_text |
| **Linking gap** | Refs not resolved |
| **Performance regression** | Slower than before |

### 3. Fix One Class at a Time

- Add/adjust DTO fields only when needed
- **Add a focused test (golden assertion) before writing code**
- Implement extraction improvements
- Add debug logs (behind a flag) for formula parsing, table mapping, object anchors/bboxes

### 4. Regression

- Re-run extraction on all stress workbooks
- Update coverage report
- Update golden outputs only when:
  - Change is clearly correct
  - Accompanied by rationale in `CHANGELOG.md` or test comment

## Must-Have Behaviors

- **Lossless references**: raw formula string + resolved dependency list (ranges/names)
- **Display semantics**: display_text + number_format
- **Layout semantics**: merged cells, row/col sizing, anchors/bboxes for objects
- **Deterministic output ordering** for stable diffs

## Deliverables per Iteration

A PR-style commit (or patch) with:

- Tests added
- Code changes
- Updated coverage report
- Before/after diff summary

## Acceptance Criteria

- Passes all golden tests for `expected_extraction_min.json`
- No new regressions on prior workbooks
- Coverage increases monotonically for targeted categories
- Extraction output stable across runs

## Reference

See `DESIGN.md` for pipeline stages and `src/xlsx_parser/pipeline.py` for orchestration.
