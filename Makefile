.PHONY: help install install-dev test test-ci lint format typecheck wheel-check clean \
	corpus-download bench-robust bench-retrieval bench bench-track docker-bench

PYTHON ?= python
PKG_VERSION := $(shell $(PYTHON) -c "import tomllib, pathlib; print(tomllib.loads(pathlib.Path('pyproject.toml').read_text())['project']['version'])")

help:
	@echo "ks-xlsx-parser — common targets"
	@echo ""
	@echo "  make install         Install package and dev deps (editable)"
	@echo "  make install-dev     Alias for install (matches ks-backend)"
	@echo "  make test            Run the default test suite"
	@echo "  make test-ci         Run the suite with verbose output for CI"
	@echo ""
	@echo "  make lint            Ruff lint"
	@echo "  make format          Ruff format"
	@echo "  make typecheck       mypy"
	@echo "  make wheel-check     Build wheel + verify it imports in a clean venv"
	@echo ""
	@echo "  make corpus-download Fetch SpreadsheetBench for benchmark runs"
	@echo ""
	@echo "  make bench-robust    Robustness on SpreadsheetBench (ks vs docling, ~20 min)"
	@echo "  make bench-retrieval Retrieval recall on SpreadsheetBench (ks vs docling, ~40 min)"
	@echo "  make bench           Run both benchmarks back-to-back"
	@echo "  make bench-track     Run retrieval bench + append metrics to history"
	@echo "  make docker-bench    Build + run the benchmark Docker image"

install:
	$(PYTHON) -m pip install -e ".[dev,api]"

# Alias — junior devs pattern-match off ks-backend's `make install-dev`.
install-dev: install

test:
	$(PYTHON) -m pytest tests/ -v --tb=short -W ignore::UserWarning

test-ci:
	$(PYTHON) -m pytest tests/ -v --tb=short -W ignore::UserWarning --junitxml=reports/junit.xml

lint:
	$(PYTHON) -m ruff check src/ tests/ scripts/

format:
	$(PYTHON) -m ruff format src/ tests/ scripts/

typecheck:
	$(PYTHON) -m mypy src/ks_xlsx_parser

# Build the wheel and prove it imports outside the editable source tree.
# This is the regression guard for the v0.2.0 packaging bug (pipeline.py
# missing from the wheel because it was a top-level module, not a package).
wheel-check:
	rm -rf dist build
	$(PYTHON) -m build --wheel
	$(PYTHON) scripts/verify_wheel.py

clean:
	rm -rf build/ dist/ *.egg-info src/*.egg-info .pytest_cache .ruff_cache .mypy_cache
	find . -type d -name __pycache__ -prune -exec rm -rf {} +

corpus-download:
	./scripts/download_corpora.sh

bench-robust:
	@test -d data/corpora/spreadsheetbench || (echo "Corpus missing. Run 'make corpus-download' first." && exit 1)
	$(PYTHON) -m tests.benchmarks.vs_hucre \
		--corpus data/corpora/spreadsheetbench --parsers ks,docling \
		--per-file-timeout 120 \
		--out tests/benchmarks/reports/spreadsheetbench

bench-retrieval:
	@test -d data/corpora/spreadsheetbench || (echo "Corpus missing. Run 'make corpus-download' first." && exit 1)
	$(PYTHON) scripts/eval_retrieval.py --parsers ks,docling

bench: bench-robust bench-retrieval

# Run the retrieval benchmark and append a row to history.jsonl so
# accuracy can be tracked commit-over-commit. Goal: text recall@5 > 0.90.
bench-track:
	@test -d data/corpora/spreadsheetbench || (echo "Corpus missing. Run 'make corpus-download' first." && exit 1)
	$(PYTHON) scripts/eval_retrieval.py --parsers ks --emit-failures \
		--out tests/benchmarks/reports/retrieval
	$(PYTHON) scripts/append_bench_history.py
	$(PYTHON) scripts/triage_recall.py tests/benchmarks/reports/retrieval

docker-bench:
	docker build -f Dockerfile.bench -t ks-xlsx-parser-bench .
	docker run --rm -v "$(PWD)/tests/benchmarks/reports:/app/tests/benchmarks/reports" ks-xlsx-parser-bench
