# XLSXParser — Agent Instructions

## Project

Excel workflow parser for RAG + auditability. Pure Python, openpyxl + lxml. See `DESIGN.md` for architecture.

## Commands

- **Tests**: `python -m pytest tests/ -v --tb=short -W ignore::UserWarning -m 'not corpus'`
- **Parse a file**: `python -c "from xlsx_parser.pipeline import parse_workbook; print(parse_workbook('path/to/file.xlsx').to_json())"`

## Ralph Loop (Iterative Parser Improvement)

When the user asks to **improve the parser**, **fix all tests**, **run the Ralph loop**, or **iteratively improve the repo**, follow this workflow:

1. **Apply** the `excel-extraction-pipeline-improver` and `excel-stress-tester-builder` skills (from `.cursor/skills/`).
2. **Run** pytest. If any tests fail, fix them using TDD (add a focused test first, then implement).
3. **Re-run** pytest. Repeat until **all tests pass**.
4. **Stop** when pytest exits 0. Summarize what you fixed. Do not continue after tests pass.

**Guidelines:**
- Fix one failure class at a time: parsing gap, schema gap, normalization gap, linking gap.
- Preserve lossless references, display semantics, layout semantics.
- Add tests before implementation. See `docs/PARSER_KNOWN_ISSUES.md` for known edge cases.

## Skills

- `excel-stress-tester-builder` — Build stress-test workbooks that cover the extraction spec
- `excel-extraction-pipeline-improver` — TDD-based pipeline fixes from feedback
