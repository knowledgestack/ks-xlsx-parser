.PHONY: help install test test-ci testbench testbench-build testbench-zip lint format typecheck clean corpus-download

PYTHON ?= python
PKG_VERSION := $(shell $(PYTHON) -c "import tomllib, pathlib; print(tomllib.loads(pathlib.Path('pyproject.toml').read_text())['project']['version'])")
TESTBENCH_ZIP := dist/testBench-v$(PKG_VERSION).zip

help:
	@echo "ks-xlsx-parser — common targets"
	@echo ""
	@echo "  make install         Install package and dev deps (editable)"
	@echo "  make test            Run the default test suite (skips corpus + testbench)"
	@echo "  make test-ci         Run the suite with verbose output for CI"
	@echo ""
	@echo "  make testbench-build Generate the 1000-file testBench dataset"
	@echo "  make testbench       Run parser round-trip across the full testBench"
	@echo "  make testbench-zip   Package testBench into $(TESTBENCH_ZIP) for GitHub release"
	@echo ""
	@echo "  make lint            Ruff lint"
	@echo "  make format          Ruff format"
	@echo "  make typecheck       mypy"
	@echo ""
	@echo "  make corpus-download Fetch public XLSX corpora for extended robustness"

install:
	$(PYTHON) -m pip install -e ".[dev,api]"

test:
	$(PYTHON) -m pytest tests/ -v --tb=short -W ignore::UserWarning

test-ci:
	$(PYTHON) -m pytest tests/ -v --tb=short -W ignore::UserWarning --junitxml=reports/junit.xml

testbench-build:
	$(PYTHON) scripts/build_testbench.py --clean

testbench:
	@test -d testBench/generated || (echo "testBench/generated missing. Run 'make testbench-build' first." && exit 1)
	$(PYTHON) -m pytest tests/test_testbench_roundtrip.py -m testbench --tb=short -W ignore::UserWarning

testbench-zip: testbench-build
	@mkdir -p dist
	@echo "→ packaging testBench into $(TESTBENCH_ZIP)"
	@rm -f $(TESTBENCH_ZIP)
	@cd . && zip -qr $(TESTBENCH_ZIP) testBench \
		-x "testBench/**/__pycache__/*" \
		-x "testBench/**/.DS_Store"
	@ls -lh $(TESTBENCH_ZIP)
	@echo "→ attach with: gh release create v$(PKG_VERSION) $(TESTBENCH_ZIP) --generate-notes"

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
