.PHONY: test ralph-loop ralph-loop-cli help

help:
	@echo "Targets:"
	@echo "  test             - Run pytest (default: skip corpus)"
	@echo "  ralph-loop-cli   - Run Ralph loop via Claude CLI (recommended)"
	@echo "  ralph-loop       - Run Ralph loop via Python + Anthropic API"
	@echo "  install          - pip install -e .[dev]"

test:
	python -m pytest tests/ -v --tb=short -W ignore::UserWarning -m 'not corpus'

ralph-loop-cli:
	@chmod +x scripts/ralph_loop_claude_cli.sh 2>/dev/null || true
	./scripts/ralph_loop_claude_cli.sh --duration-minutes 120

ralph-loop:
	@if [ -z "$$ANTHROPIC_API_KEY" ]; then \
		echo "Error: ANTHROPIC_API_KEY not set"; \
		echo "  export ANTHROPIC_API_KEY=sk-ant-..."; \
		exit 1; \
	fi
	pip install -e ".[dev,ralph]" -q 2>/dev/null || true
	python scripts/ralph_loop_improve.py --duration-minutes 120

install:
	pip install -e ".[dev]"
