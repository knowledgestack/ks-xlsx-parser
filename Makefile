.PHONY: help install test test-ci lint format typecheck clean corpus-download bench-robust bench-retrieval bench

PYTHON ?= python
PKG_VERSION := $(shell $(PYTHON) -c "import tomllib, pathlib; print(tomllib.loads(pathlib.Path('pyproject.toml').read_text())['project']['version'])")

help:
	@echo "ks-xlsx-parser — common targets"
	@echo ""
	@echo "  make install         Install package and dev deps (editable)"
	@echo "  make test            Run the default test suite"
	@echo "  make test-ci         Run the suite with verbose output for CI"
	@echo ""
	@echo "  make lint            Ruff lint"
	@echo "  make format          Ruff format"
	@echo "  make typecheck       mypy"
	@echo ""
	@echo "  make corpus-download Fetch SpreadsheetBench for benchmark runs"
	@echo ""
	@echo "  make bench-robust    Robustness on SpreadsheetBench (ks vs docling, ~20 min)"
	@echo "  make bench-retrieval Retrieval recall on SpreadsheetBench (ks vs docling, ~40 min)"
	@echo "  make bench           Run both benchmarks back-to-back"

install:
	$(PYTHON) -m pip install -e ".[dev,api]"

test:
	$(PYTHON) -m pytest tests/ -v --tb=short -W ignore::UserWarning

test-ci:
	$(PYTHON) -m pytest tests/ -v --tb=short -W ignore::UserWarning --junitxml=reports/junit.xml

lint:
	$(PYTHON) -m ruff check src/ tests/ scripts/

format:
	$(PYTHON) -m ruff format src/ tests/ scripts/

typecheck:
	$(PYTHON) -m mypy src/xlsx_parser

clean:
	rm -rf build/ dist/ *.egg-info src/*.egg-info .pytest_cache .ruff_cache .mypy_cache
	find . -type d -name __pycache__ -prune -exec rm -rf {} +

corpus-download:
	./scripts/download_corpora.sh

bench-robust:
	@test -d data/corpora/spreadsheetbench || (echo "Corpus missing. Run 'make corpus-download' first." && exit 1)
	PYTHONPATH=src $(PYTHON) -m tests.benchmarks.vs_hucre \
		--corpus data/corpora/spreadsheetbench --parsers ks,docling \
		--per-file-timeout 120 \
		--out tests/benchmarks/reports/spreadsheetbench

bench-retrieval:
	@test -d data/corpora/spreadsheetbench || (echo "Corpus missing. Run 'make corpus-download' first." && exit 1)
	PYTHONPATH=src $(PYTHON) scripts/eval_retrieval.py --parsers ks,docling

bench: bench-robust bench-retrieval
