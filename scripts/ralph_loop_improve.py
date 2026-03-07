#!/usr/bin/env python3
"""
RALPH Loop: Iteratively improve the xlsx_parser repository.

Runs an AI agent in a loop until:
  - All tests pass (no more issues) → SUCCESS, exit 0
  - Time budget (default 2 hours) exceeded → exit 1
  - Max iterations (default 50) reached → exit 1

Uses the excel-stress-tester-builder and excel-extraction-pipeline-improver skills.
Requires: pip install anthropic  (or: pip install -e ".[dev,ralph]")

Usage:
  python scripts/ralph_loop_improve.py [--duration-minutes 120] [--max-iterations 50]
  # Or: make ralph-loop
"""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
import time
from pathlib import Path

try:
    import anthropic
except ImportError:
    print(
        "ERROR: anthropic package required. Run: pip install anthropic",
        file=sys.stderr,
    )
    sys.exit(1)

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
REPO_ROOT = Path(__file__).resolve().parent.parent
SKILLS_DIR = REPO_ROOT / ".cursor" / "skills"
ITERATIONS_DIR = REPO_ROOT / "docs" / "iterations"


def load_skill(name: str) -> str:
    """Load skill content from .cursor/skills/<name>/SKILL.md."""
    path = SKILLS_DIR / name / "SKILL.md"
    if path.exists():
        return path.read_text()
    return f"[Skill {name} not found at {path}]"


def load_both_skills() -> str:
    stress = load_skill("excel-stress-tester-builder")
    improver = load_skill("excel-extraction-pipeline-improver")
    return f"""
## Skill 1: Excel Stress-Test Workbook Builder
{stress}

---
## Skill 2: Excel Extraction Pipeline Improver
{improver}
"""


# ---------------------------------------------------------------------------
# Tools (execute and return results as strings for the model)
# ---------------------------------------------------------------------------
def run_tests() -> str:
    """Run pytest. Returns stdout+stderr and exit code summary."""
    result = subprocess.run(
        ["python", "-m", "pytest", "tests/", "-v", "--tb=short", "-W", "ignore::UserWarning"],
        cwd=REPO_ROOT,
        capture_output=True,
        text=True,
        timeout=300,
    )
    out = result.stdout + result.stderr
    return f"Exit code: {result.returncode}\n\n{out}"


def run_command(cmd: str, cwd: str | None = None) -> str:
    """Run a shell command. cwd defaults to repo root."""
    try:
        result = subprocess.run(
            cmd,
            shell=True,
            cwd=cwd or str(REPO_ROOT),
            capture_output=True,
            text=True,
            timeout=120,
        )
        return f"Exit code: {result.returncode}\nStdout:\n{result.stdout}\nStderr:\n{result.stderr}"
    except subprocess.TimeoutExpired:
        return "Error: Command timed out after 120s"
    except Exception as e:
        return f"Error: {e}"


def read_file(path: str) -> str:
    """Read a file from the repo."""
    p = Path(path)
    if not p.is_absolute():
        p = REPO_ROOT / p
    if not p.exists():
        return f"Error: File not found: {p}"
    try:
        return p.read_text()
    except Exception as e:
        return f"Error reading file: {e}"


def write_file(path: str, content: str) -> str:
    """Write content to a file in the repo."""
    p = Path(path)
    if not p.is_absolute():
        p = REPO_ROOT / p
    try:
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(content)
        return f"Wrote {len(content)} chars to {p.relative_to(REPO_ROOT)}"
    except Exception as e:
        return f"Error writing file: {e}"


def list_dir(path: str = ".") -> str:
    """List directory contents."""
    p = REPO_ROOT / path if path != "." else REPO_ROOT
    if not p.is_dir():
        return f"Error: Not a directory: {p}"
    items = sorted(p.iterdir())
    return "\n".join(
        f"  {'(dir) ' if x.is_dir() else ''}{x.name}"
        for x in items
        if not x.name.startswith(".")
    )


TOOLS = [
    {
        "name": "run_tests",
        "description": "Run pytest. Use this to verify no test failures. If exit code 0, the loop can stop.",
        "input_schema": {"type": "object", "properties": {}},
    },
    {
        "name": "run_command",
        "description": "Run a shell command (e.g., python -m pytest, pip install).",
        "input_schema": {
            "type": "object",
            "properties": {"cmd": {"type": "string"}, "cwd": {"type": "string"}},
            "required": ["cmd"],
        },
    },
    {
        "name": "read_file",
        "description": "Read a file from the repository.",
        "input_schema": {
            "type": "object",
            "properties": {"path": {"type": "string"}},
            "required": ["path"],
        },
    },
    {
        "name": "write_file",
        "description": "Write or overwrite a file in the repository.",
        "input_schema": {
            "type": "object",
            "properties": {"path": {"type": "string"}, "content": {"type": "string"}},
            "required": ["path", "content"],
        },
    },
    {
        "name": "list_dir",
        "description": "List directory contents.",
        "input_schema": {
            "type": "object",
            "properties": {"path": {"type": "string"}},
        },
    },
]


def execute_tool(name: str, args: dict) -> str:
    if name == "run_tests":
        return run_tests()
    if name == "run_command":
        return run_command(args.get("cmd", ""), args.get("cwd"))
    if name == "read_file":
        return read_file(args.get("path", ""))
    if name == "write_file":
        return write_file(args.get("path", ""), args.get("content", ""))
    if name == "list_dir":
        return list_dir(args.get("path", "."))
    return f"Unknown tool: {name}"


# ---------------------------------------------------------------------------
# Ralph Loop
# ---------------------------------------------------------------------------
INITIAL_PROMPT = """You are improving the xlsx_parser Excel extraction pipeline. Apply the excel-extraction-pipeline-improver and excel-stress-tester-builder skills.

Your task: Fix any failing tests and extraction gaps. Work iteratively:
1. Run tests to see current state
2. Fix one issue at a time (TDD: add test first, then implement)
3. Re-run tests
4. Repeat until all tests pass

When all tests pass (run_tests exit code 0), you are DONE. Summarize what you fixed and stop.
Do not continue iterating once tests pass."""


def run_ralph_loop(
    duration_minutes: int = 120,
    max_iterations: int = 50,
) -> tuple[bool, str]:
    """
    Run the Ralph loop. Returns (success, final_message).
    Stops immediately when run_tests returns exit code 0 (no more issues).
    """
    if not os.environ.get("ANTHROPIC_API_KEY"):
        return False, "ANTHROPIC_API_KEY environment variable not set"

    skills_text = load_both_skills()
    system = f"""You have two skills loaded. Follow their procedures.

{skills_text}

CRITICAL: When run_tests returns exit code 0, there are NO MORE ISSUES. Summarize and stop.
"""

    client = anthropic.Anthropic()
    end_time = time.time() + duration_minutes * 60
    iterations = 0
    messages: list[anthropic.types.MessageParam] = [
        {"role": "user", "content": INITIAL_PROMPT}
    ]
    last_test_result = ""

    while time.time() < end_time and iterations < max_iterations:
        iterations += 1
        print(f"\n--- Ralph iteration {iterations} ---", flush=True)

        # Inner loop: handle tool_use until model sends end_turn
        while True:
            response = client.messages.create(
                model=os.environ.get("RALPH_MODEL", "claude-sonnet-4-5"),
                max_tokens=16384,
                system=system,
                messages=messages,
                tools=TOOLS,
            )

            # Check for tool use
            tool_use_blocks = [b for b in response.content if b["type"] == "tool_use"]
            if not tool_use_blocks:
                # end_turn - model finished this round
                break

            # Execute tools and build tool_result content
            tool_results: list[dict] = []
            for block in tool_use_blocks:
                name = block["name"]
                tid = block["id"]
                inp = block.get("input", {})
                result = execute_tool(name, inp)
                tool_results.append({"type": "tool_result", "tool_use_id": tid, "content": result})
                if name == "run_tests":
                    last_test_result = result
                    if "Exit code: 0" in result:
                        return True, f"All tests pass after {iterations} iterations."
                print(f"  Tool: {name} -> {result[:200]}...", flush=True)

            messages.append({"role": "assistant", "content": response.content})
            messages.append({"role": "user", "content": tool_results})

        # Model ended turn. Verify: run tests ourselves.
        last_test_result = run_tests()
        if "Exit code: 0" in last_test_result:
            return True, f"All tests pass after {iterations} iterations."

        # Inject feedback for next iteration
        feedback = f"""Tests still failing. Fix the issues and try again.

Last test output:
{last_test_result[:2000]}
"""
        messages.append({"role": "assistant", "content": response.content})
        messages.append({"role": "user", "content": feedback})

    return False, (
        f"Stopped after {iterations} iterations (time or limit). "
        f"Last test result:\n{last_test_result[:1000]}"
    )


def main() -> None:
    parser = argparse.ArgumentParser(description="Ralph loop: improve xlsx_parser until tests pass")
    parser.add_argument("--duration-minutes", type=int, default=120, help="Max runtime in minutes")
    parser.add_argument("--max-iterations", type=int, default=50, help="Max agent iterations")
    args = parser.parse_args()

    ITERATIONS_DIR.mkdir(parents=True, exist_ok=True)
    success, msg = run_ralph_loop(
        duration_minutes=args.duration_minutes,
        max_iterations=args.max_iterations,
    )
    print("\n" + "=" * 60)
    print(msg)
    print("=" * 60)
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
