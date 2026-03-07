# Contributing to XLSX Parser

Thank you for your interest in contributing.

## Development Setup

```bash
git clone https://github.com/your-username/xlsx-parser.git
cd xlsx-parser
pip install -e ".[dev]"
```

## Running Tests

```bash
pytest
```

Skip slow/corpus tests (default):

```bash
pytest -m "not corpus"
```

## Code Style

- Use type hints where practical.
- Follow existing patterns in the codebase.
- Run `pytest` before submitting a PR.

## Adding Tests

- Place tests in `tests/`.
- Use fixtures from `tests/conftest.py` for programmatic workbooks.
- For cross-validation against calamine, use the `crossval` marker.

## Reporting Issues

Include:

- Python version
- Minimal `.xlsx` file that reproduces the issue (if applicable)
- Expected vs actual behavior
- Output of `xlsx_parser.__version__`

## Pull Requests

1. Fork the repository.
2. Create a branch from `main`.
3. Add tests for new behavior.
4. Ensure all tests pass.
5. Open a PR with a clear description.
