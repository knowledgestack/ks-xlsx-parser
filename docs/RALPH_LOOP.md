# RALPH Loop — Iterative Repository Improvement

The Ralph loop runs an AI agent to iteratively fix the xlsx_parser until **all tests pass** or a time/iteration limit is reached. It stops **immediately when there are no more issues**.

## Quick Start (Claude CLI — recommended)

```bash
# Install and authenticate Claude CLI
# https://claude.com/download

claude auth login

# Run for up to 2 hours (stops early if tests pass)
./scripts/ralph_loop_claude_cli.sh

# Or with make
make ralph-loop-cli
```

## Alternative: Python + Anthropic API

```bash
pip install -e ".[dev,ralph]"
export ANTHROPIC_API_KEY="sk-ant-..."
python scripts/ralph_loop_improve.py
# Or: make ralph-loop
```

## Options

| Option | Default | Description |
|--------|---------|-------------|
| `--duration-minutes` | 120 | Max runtime (2 hours) |
| `--max-iterations` | 50 | Max agent iteration count |
| `ANTHROPIC_API_KEY` | required (Python) | API key for Claude |
| `RALPH_MODEL` | claude-sonnet-4-5 | Claude model ID (Python only) |
| Claude CLI | — | Must be installed and authenticated |

## Stop Conditions

The loop terminates when **any** of:

1. **All tests pass** (pytest exit code 0) → success, exit 0
2. **Time budget exceeded** → exit 1
3. **Max iterations reached** → exit 1

**Important**: As soon as `run_tests` returns exit code 0, the agent stops. No further iterations.

## Skills Used

The agent uses two project skills:

1. **excel-stress-tester-builder** — Build stress workbooks covering extraction spec
2. **excel-extraction-pipeline-improver** — TDD-based pipeline fixes from feedback

## Loop Flow (RALPH)

- **R**ecap: Summarize current failures and opportunities
- **A**ct: Smallest code/test change to fix one failure
- **L**og: Record what changed (changelog, coverage delta)
- **P**rove: Re-run tests and extraction; diff golden outputs
- **H**andoff: If time remains and issues remain, pick next item; else produce final report

## Iteration Logs

Each iteration can write a note to `docs/iterations/` (optional). The script creates the directory on run.

## Running Unattended

**Claude CLI** (recommended):

```bash
nohup ./scripts/ralph_loop_claude_cli.sh --duration-minutes 120 > ralph.log 2>&1 &
tail -f ralph.log
```

**Python/API**:

```bash
nohup python scripts/ralph_loop_improve.py --duration-minutes 120 > ralph.log 2>&1 &
tail -f ralph.log
```

## Skills Location

Skills are in `.claude/skills/` for Claude CLI (auto-discovered) and `.cursor/skills/` for Cursor.
