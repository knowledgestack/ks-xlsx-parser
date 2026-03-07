# Ralph Loop — Fix All Failing Tests

You are improving the xlsx_parser Excel extraction pipeline. Apply the **excel-extraction-pipeline-improver** and **excel-stress-tester-builder** skills from `.claude/skills/`.

## Task

1. Run the test suite: `python -m pytest tests/ -v --tb=short -W ignore::UserWarning -m 'not corpus'`
2. If any tests fail, fix them following TDD: add a focused test first, then implement.
3. Re-run tests. Repeat until **all tests pass**.
4. **When pytest exits with code 0, you are DONE.** Summarize what you fixed and stop. Do not continue.

## Critical

- Stop as soon as tests pass. There are no more issues.
- Fix one failure class at a time (parsing gap, schema gap, normalization gap, etc.).
- Preserve lossless references, display semantics, and layout semantics.
- Work in this repo (xlsx_parser). Use Bash to run commands, Read/Edit to modify files.
