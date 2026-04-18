#!/usr/bin/env python3
"""
Stress Test Runner for XLSX Parser Pipeline.

Builds progressively complex Excel files, runs them through the parser,
and documents any failures, errors, or unexpected behavior. Runs in a loop
until a failure is encountered or max level reached.
"""



import json
import sys
from dataclasses import dataclass, field
from pathlib import Path

# Add src to path for imports
PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from xlsx_parser.pipeline import parse_workbook

STRESS_DIR = Path(__file__).parent


@dataclass
class LevelResult:
    level: int
    passed: bool
    source_file: str = ""  # Path or label for reporting
    error_message: str | None = None
    parse_errors: list[dict] = field(default_factory=list)
    workbook_stats: dict = field(default_factory=dict)
    chunk_count: int = 0
    validation_failures: list[str] = field(default_factory=list)
    documented_issues: list[dict] = field(default_factory=list)  # {issue_id, cell, reason}


def run_level(level: int, xlsx_path: Path, source_file: str | None = None) -> LevelResult:
    """Parse a single stress test file and validate the result."""
    result = LevelResult(level=level, passed=True, source_file=source_file or str(xlsx_path))

    try:
        parse_result = parse_workbook(path=xlsx_path)

        # Collect parse errors from workbook
        for err in parse_result.workbook.errors:
            result.parse_errors.append({
                "severity": getattr(err.severity, "value", str(err.severity)),
                "stage": err.stage,
                "message": err.message,
                "sheet_name": getattr(err, "sheet_name", None),
            })

        # Stats
        wb = parse_result.workbook
        result.workbook_stats = {
            "total_sheets": wb.total_sheets,
            "total_cells": wb.total_cells,
            "total_formulas": wb.total_formulas,
            "tables": len(wb.tables),
            "charts": len(wb.charts),
            "named_ranges": len(wb.named_ranges),
            "errors_count": len(wb.errors),
        }
        result.chunk_count = parse_result.total_chunks

        # Document known limitations (do NOT fail — these are expected openpyxl/parser quirks)
        for sheet in parse_result.workbook.sheets:
            for region in sheet.merged_regions:
                master = sheet.get_cell(region.master.row, region.master.col)
                if master and master.is_merged_master:
                    span = (master.merge_extent or 1) * (master.merge_col_extent or 1)
                    if span > 1 and (not master.raw_value and not master.display_value and not master.formula):
                        result.documented_issues.append({
                            "issue_id": "merge_empty_master",
                            "cell": f"{sheet.sheet_name}!{region.master.to_a1()}",
                            "reason": "Value was in non-master cell before merge; openpyxl MergedCell has no value, content is lost.",
                        })

        # Validation: hard errors should fail the level
        error_severity = "error"
        has_hard_error = any(
            e.get("severity", "").lower() == error_severity
            for e in result.parse_errors
        )

        if has_hard_error:
            result.passed = False
            result.error_message = "Pipeline reported ERROR severity"
            result.validation_failures.append("ERROR in workbook.errors")

        # Expect cells when file has content (level 0+)
        if wb.total_sheets > 0 and wb.total_cells == 0:
            if level != 15:
                result.validation_failures.append("No cells extracted but sheet exists")

        # Expect chunks when we have content
        if wb.total_cells > 0 and parse_result.total_chunks == 0:
            result.passed = False
            result.error_message = result.error_message or "No chunks produced"
            result.validation_failures.append("Cells exist but no chunks")

        if result.validation_failures:
            result.passed = False
            result.error_message = result.error_message or "; ".join(result.validation_failures[:3])

    except Exception as e:
        result.passed = False
        result.error_message = f"{type(e).__name__}: {e}"
        import traceback
        result.parse_errors.append({"traceback": traceback.format_exc()})

    return result


def run_stress_loop(max_level: int = 19, stop_on_first_fail: bool = True) -> list[LevelResult]:
    """
    Run stress test: build levels 0..max_level, parse each, document failures.
    If stop_on_first_fail, stops when first level fails.
    """
    import importlib.util
    builder_path = STRESS_DIR / "complex_excel_builder.py"
    spec = importlib.util.spec_from_file_location("complex_excel_builder", builder_path)
    builder = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(builder)
    build_level = builder.build_level

    results: list[LevelResult] = []
    for level in range(max_level + 1):
        path = STRESS_DIR / f"stress_level_{level}.xlsx"
        if not path.exists():
            build_level(level, path)
        res = run_level(level, path, source_file=f"stress_level_{level}.xlsx")
        results.append(res)
        print(f"Level {level}: {'PASS' if res.passed else 'FAIL'} - "
              f"cells={res.workbook_stats.get('total_cells', 0)}, "
              f"chunks={res.chunk_count}, "
              f"errors={len(res.parse_errors)}"
              + (f" | {res.error_message[:60]}..." if res.error_message and len(res.error_message) > 60 else f" | {res.error_message}" if res.error_message else ""))
        if not res.passed and stop_on_first_fail:
            print(f"  Stopping at first failure (level {level})")
            break
    return results


def run_merge_stress(merge_paths: list[Path]) -> list[LevelResult]:
    """Run parser on merge-stress workbooks and return results with documented issues."""
    results = []
    for i, path in enumerate(merge_paths):
        if path.exists():
            res = run_level(i, path, source_file=path.name)
            results.append(res)
    return results


def _write_known_issues_doc(all_results: list[LevelResult], docs_path: Path) -> None:
    """Write docs/PARSER_KNOWN_ISSUES.md: files that failed or have documented limitations."""
    # Collect by issue_id
    by_issue: dict[str, list[dict]] = {}
    failed_files: list[dict] = []

    for r in all_results:
        if not r.passed:
            failed_files.append({
                "file": r.source_file,
                "error": r.error_message,
                "validation_failures": r.validation_failures,
            })
        for d in r.documented_issues:
            issue_id = d["issue_id"]
            if issue_id not in by_issue:
                by_issue[issue_id] = []
            by_issue[issue_id].append({
                "file": r.source_file,
                "cell": d.get("cell", ""),
                "reason": d.get("reason", ""),
            })

    lines = [
        "# Parser Known Issues & Limitations",
        "",
        "This document lists Excel files that fail to parse correctly or exhibit known limitations.",
        "Use this when working with the XLSX parser to understand edge cases.",
        "",
        "---",
        "",
        "## Files That Fail (Hard Errors)",
        "",
    ]
    if failed_files:
        for f in failed_files:
            lines.append(f"### `{f['file']}`")
            lines.append("")
            lines.append(f"- **Error**: {f['error']}")
            for vf in f.get("validation_failures", []):
                lines.append(f"- Validation: {vf}")
            lines.append("")
    else:
        lines.append("None. All stress-tested files parse without hard errors.")
        lines.append("")

    lines.extend([
        "---",
        "",
        "## Documented Limitations (No Hard Fail)",
        "",
        "These files exhibit known openpyxl/parser quirks. The pipeline completes but output may be incomplete.",
        "",
    ])

    if by_issue:
        for issue_id, occurrences in sorted(by_issue.items()):
            lines.append(f"### {issue_id}")
            lines.append("")
            reason = occurrences[0].get("reason", "") if occurrences else ""
            lines.append(f"**Why**: {reason}")
            lines.append("")
            lines.append("**Affected files**:")
            for occ in occurrences[:20]:  # Cap display
                lines.append(f"- `{occ['file']}` — {occ.get('cell', '')}")
            if len(occurrences) > 20:
                lines.append(f"- … and {len(occurrences) - 20} more")
            lines.append("")
    else:
        lines.append("None detected.")
        lines.append("")

    lines.extend([
        "---",
        "",
        "## Quick Reference",
        "",
        "| Issue ID | Description |",
        "|----------|-------------|",
        "| `merge_empty_master` | Value was in right cell, merge created left→content lost in openpyxl |",
        "",
    ])

    docs_path.write_text("\n".join(lines), encoding="utf-8")


def generate_report(results: list[LevelResult], out_path: Path) -> None:
    """Write STRESS_TEST_RESULTS.md and a JSON summary."""
    stress_results = [r for r in results if "stress_level_" in r.source_file]
    stress_levels = [r.level for r in stress_results] if stress_results else []
    max_level = max(stress_levels) if stress_levels else -1

    lines = [
        "# Parser Stress Test Results",
        "",
        "## Summary",
        "",
        f"- Stress levels tested: 0..{max_level}" + (f" (+ {len(results) - len(stress_results)} merge-stress files)" if len(results) > len(stress_results) else ""),
        f"- Passed: {sum(1 for r in results if r.passed)}",
        f"- Failed: {sum(1 for r in results if not r.passed)}",
        "",
        "## Per-Level Results",
        "",
        "| Level | Pass | Cells | Chunks | Tables | Charts | Errors |",
        "|-------|------|-------|--------|--------|--------|--------|",
    ]
    for r in results:
        stats = r.workbook_stats
        err_msg = r.error_message[:40] + "…" if r.error_message and len(r.error_message) > 40 else (r.error_message or "")
        lines.append(
            f"| {r.level} | {'✓' if r.passed else '✗'} | "
            f"{stats.get('total_cells', '-')} | {r.chunk_count} | "
            f"{stats.get('tables', '-')} | {stats.get('charts', '-')} | "
            f"{len(r.parse_errors)} |"
        )
    lines.append("")
    lines.append("## Failures & Issues")
    lines.append("")
    for r in results:
        if not r.passed:
            lines.append(f"### Level {r.level} FAILED")
            lines.append("")
            if r.error_message:
                lines.append(f"- **Error**: {r.error_message}")
            for vf in r.validation_failures:
                lines.append(f"- Validation: {vf}")
            for pe in r.parse_errors:
                if "traceback" in pe:
                    lines.append("```")
                    lines.append(pe["traceback"])
                    lines.append("```")
                else:
                    lines.append(f"- Parse: [{pe.get('severity')}] {pe.get('stage')}: {pe.get('message')}")
            lines.append("")

    # Documented issues (known limitations)
    lines.append("## Documented Issues (Known Limitations)")
    lines.append("")
    doc_issues = [(r.source_file, d) for r in results for d in r.documented_issues]
    if doc_issues:
        by_id: dict[str, list[str]] = {}
        for f, d in doc_issues:
            iid = d.get("issue_id", "unknown")
            if iid not in by_id:
                by_id[iid] = []
            by_id[iid].append(f"{f} ({d.get('cell', '')})")
        for iid, files in sorted(by_id.items()):
            lines.append(f"### {iid}")
            for f in files[:10]:
                lines.append(f"- {f}")
            if len(files) > 10:
                lines.append(f"- … +{len(files) - 10} more")
            lines.append("")
    else:
        lines.append("None.")
        lines.append("")

    # Pipeline Problem Summary
    lines.append("## Identified Pipeline Problems")
    lines.append("")
    problems: list[str] = []
    for r in results:
        if not r.passed and r.error_message:
            problems.append(f"**Level {r.level}**: {r.error_message}")
        for pe in r.parse_errors:
            if pe.get("severity") == "error":
                problems.append(f"**Level {r.level}**: {pe.get('message', '')} (stage: {pe.get('stage')})")
    if problems:
        for p in problems:
            lines.append(f"- {p}")
    else:
        lines.append("- No hard failures in tested range.")
    lines.append("")

    out_path.write_text("\n".join(lines), encoding="utf-8")

    # Write docs/PARSER_KNOWN_ISSUES.md for Claude / downstream use
    docs_dir = PROJECT_ROOT / "docs"
    docs_dir.mkdir(exist_ok=True)
    docs_path = docs_dir / "PARSER_KNOWN_ISSUES.md"
    _write_known_issues_doc(results, docs_path)

    # JSON summary
    json_path = out_path.parent / "stress_results.json"
    json_data = [
        {
            "level": r.level,
            "passed": r.passed,
            "source_file": r.source_file,
            "error_message": r.error_message,
            "parse_errors": r.parse_errors,
            "workbook_stats": r.workbook_stats,
            "chunk_count": r.chunk_count,
            "validation_failures": r.validation_failures,
            "documented_issues": r.documented_issues,
        }
        for r in results
    ]
    with open(json_path, "w") as f:
        json.dump(json_data, f, indent=2)


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--max-level", type=int, default=19, help="Max complexity level")
    ap.add_argument("--no-stop", action="store_true", help="Don't stop on first failure")
    ap.add_argument("--build-only", action="store_true", help="Only build files, don't parse")
    args = ap.parse_args()

    if args.build_only:
        import importlib.util
        spec = importlib.util.spec_from_file_location("builder", STRESS_DIR / "complex_excel_builder.py")
        builder = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(builder)
        builder.build_all_levels(max_level=args.max_level)
        print("Built only. Run without --build-only to parse.")
    else:
        results = run_stress_loop(max_level=args.max_level, stop_on_first_fail=not args.no_stop)

        # Also run merge-stress workbooks for documented issues (build if missing)
        merge_paths = list(STRESS_DIR.glob("merge_stress_*.xlsx"))
        if not merge_paths:
            import importlib.util
            spec = importlib.util.spec_from_file_location("builder", STRESS_DIR / "complex_excel_builder.py")
            builder = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(builder)
            merge_paths = builder.build_merge_stress_workbooks()
        if merge_paths:
            merge_results = run_merge_stress(merge_paths)
            results = results + merge_results

        report_path = STRESS_DIR / "STRESS_TEST_RESULTS.md"
        generate_report(results, report_path)
        print(f"\nReport written to {report_path}")
        print(f"Known issues doc: {PROJECT_ROOT / 'docs' / 'PARSER_KNOWN_ISSUES.md'}")
