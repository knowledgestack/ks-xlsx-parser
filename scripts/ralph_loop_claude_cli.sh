#!/usr/bin/env bash
#
# Ralph Loop for Claude CLI
# Runs `claude -p` in a loop until all tests pass or time/iteration limit.
# Stops immediately when there are no more issues.
#
# Usage: ./scripts/ralph_loop_claude_cli.sh [--duration-minutes 120] [--max-iterations 50]
# Or: make ralph-loop-cli
#
# Requires: Claude CLI installed and authenticated (claude auth login)

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
PROMPT_FILE="$SCRIPT_DIR/ralph_prompt.md"
TEST_LOG="$(mktemp)"
trap 'rm -f "$TEST_LOG"' EXIT

DURATION_MINUTES=120
MAX_ITERATIONS=50

# Parse args
while [[ $# -gt 0 ]]; do
  case $1 in
    --duration-minutes) DURATION_MINUTES="$2"; shift 2 ;;
    --max-iterations)   MAX_ITERATIONS="$2";   shift 2 ;;
    *) echo "Unknown option: $1"; exit 1 ;;
  esac
done

if ! command -v claude &>/dev/null; then
  echo "Error: Claude CLI not found. Install from https://claude.com/download"
  exit 1
fi

if ! claude auth status &>/dev/null; then
  echo "Error: Claude CLI not authenticated. Run: claude auth login"
  exit 1
fi

END_TIME=$(($(date +%s) + DURATION_MINUTES * 60))
ITER=0
cd "$REPO_ROOT"

echo "=== Ralph Loop (Claude CLI) ==="
echo "Duration: ${DURATION_MINUTES} min | Max iterations: ${MAX_ITERATIONS}"
echo "Stops when: pytest exits 0 (no more issues)"
echo ""

while [[ $(date +%s) -lt $END_TIME ]] && [[ $ITER -lt $MAX_ITERATIONS ]]; do
  ITER=$((ITER + 1))
  echo "--- Ralph iteration $ITER ---"

  if [[ $ITER -eq 1 ]]; then
    # First iteration: initial prompt (skills auto-loaded from .claude/skills/)
    claude -p "$(cat "$PROMPT_FILE")" \
      --allowedTools "Bash,Read,Edit,Grep,Glob" \
      --output-format text 2>&1 || true
  else
    # Subsequent: continue with feedback
    FEEDBACK="Tests still failing. Fix them. Last pytest output:

$(cat "$TEST_LOG")

Run pytest again after fixing. Stop when exit code 0."
    claude -c -p "$FEEDBACK" \
      --allowedTools "Bash,Read,Edit,Grep,Glob" \
      --output-format text 2>&1 || true
  fi

  # Verify: run pytest ourselves
  echo ""
  echo "Verifying tests..."
  if python -m pytest tests/ -v --tb=short -W ignore::UserWarning -m 'not corpus' 2>&1 | tee "$TEST_LOG"; then
    echo ""
    echo "=========================================="
    echo "SUCCESS: All tests pass after $ITER iteration(s)."
    echo "=========================================="
    exit 0
  fi

  echo "Tests failed. Continuing next iteration..."
  echo ""
done

echo "=========================================="
echo "Stopped: time or iteration limit reached ($ITER iterations)"
echo "Last test output:"
cat "$TEST_LOG" | tail -100
echo "=========================================="
exit 1
