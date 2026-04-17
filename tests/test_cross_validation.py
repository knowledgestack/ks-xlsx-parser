"""
Cross-validation tests comparing parser output against python-calamine.

Calamine is a Rust-based Excel reader, completely independent from openpyxl.
These tests verify that our parser reads the same data that calamine does.
"""

from __future__ import annotations

import datetime

import pytest

from xlsx_parser.pipeline import parse_workbook

from tests.helpers.calamine_reader import CalamineResult
from tests.helpers.value_comparator import Mismatch, compare_cell_value, values_match


# ---------------------------------------------------------------------------
# Cross-validation on programmatic fixtures
# ---------------------------------------------------------------------------


@pytest.mark.crossval
class TestSheetNamesCrossVal:
    """Verify sheet names match between parser and calamine."""

    def test_sheet_names_match(self, programmatic_xlsx):
        parser_result = parse_workbook(path=programmatic_xlsx)
        calamine = CalamineResult.from_path(programmatic_xlsx)

        parser_names = [s.sheet_name for s in parser_result.workbook.sheets]
        assert parser_names == calamine.sheet_names, (
            f"Sheet names differ:\n  parser:   {parser_names}\n"
            f"  calamine: {calamine.sheet_names}"
        )

    def test_sheet_count_match(self, programmatic_xlsx):
        parser_result = parse_workbook(path=programmatic_xlsx)
        calamine = CalamineResult.from_path(programmatic_xlsx)
        assert len(parser_result.workbook.sheets) == len(calamine.sheet_names)


@pytest.mark.crossval
class TestCellValuesCrossVal:
    """Verify cell values match between parser and calamine."""

    def test_non_formula_values_match(self, programmatic_xlsx):
        parser_result = parse_workbook(path=programmatic_xlsx)
        calamine = CalamineResult.from_path(programmatic_xlsx)
        mismatches = _collect_mismatches(parser_result, calamine, formula_cells=False)
        assert len(mismatches) == 0, (
            f"{len(mismatches)} non-formula value mismatches:\n"
            + _format_mismatches(mismatches[:10])
        )

    def test_formula_computed_values_match(self, programmatic_xlsx):
        """For formula cells with cached values, parser's formula_value should
        match calamine's computed value. Programmatic fixtures often have no
        cached values, so we use a lenient threshold."""
        parser_result = parse_workbook(path=programmatic_xlsx)
        calamine = CalamineResult.from_path(programmatic_xlsx)
        mismatches = _collect_mismatches(parser_result, calamine, formula_cells=True)

        total_formulas = sum(
            1 for s in parser_result.workbook.sheets
            for c in s.cells.values()
            if c.formula
        )
        # Allow up to 100% mismatch for programmatic fixtures (no cached values)
        # This test is more meaningful for real-world files
        if total_formulas > 0 and len(mismatches) > 0:
            rate = len(mismatches) / total_formulas
            # Only fail if we have actual cached values but they don't match
            hard_mismatches = [
                m for m in mismatches
                if m.parser_value is not None and m.calamine_value is not None
            ]
            assert len(hard_mismatches) == 0, (
                f"{len(hard_mismatches)} formula value mismatches "
                f"(with cached values):\n"
                + _format_mismatches(hard_mismatches[:10])
            )


@pytest.mark.crossval
class TestDimensionsCrossVal:
    """Verify dimensions roughly match between parser and calamine."""

    def test_row_count_similar(self, programmatic_xlsx):
        parser_result = parse_workbook(path=programmatic_xlsx)
        calamine = CalamineResult.from_path(programmatic_xlsx)

        for sheet in parser_result.workbook.sheets:
            cal_sheet = calamine.sheets.get(sheet.sheet_name)
            if not cal_sheet or not sheet.used_range:
                continue
            parser_rows = sheet.used_range.row_count()
            # calamine total_height is the total row count of the sheet
            # For comparison, use the data area (start/end)
            if cal_sheet.start is not None and cal_sheet.end is not None:
                cal_rows = cal_sheet.end[0] - cal_sheet.start[0] + 1
                # Allow ±2 row difference (calamine may include trailing empty rows)
                assert abs(parser_rows - cal_rows) <= 2, (
                    f"Sheet '{sheet.sheet_name}' row count: "
                    f"parser={parser_rows}, calamine={cal_rows}"
                )

    def test_column_count_similar(self, programmatic_xlsx):
        parser_result = parse_workbook(path=programmatic_xlsx)
        calamine = CalamineResult.from_path(programmatic_xlsx)

        for sheet in parser_result.workbook.sheets:
            cal_sheet = calamine.sheets.get(sheet.sheet_name)
            if not cal_sheet or not sheet.used_range:
                continue
            parser_cols = sheet.used_range.col_count()
            if cal_sheet.start is not None and cal_sheet.end is not None:
                cal_cols = cal_sheet.end[1] - cal_sheet.start[1] + 1
                assert abs(parser_cols - cal_cols) <= 2, (
                    f"Sheet '{sheet.sheet_name}' col count: "
                    f"parser={parser_cols}, calamine={cal_cols}"
                )


@pytest.mark.crossval
class TestMergedRegionsCrossVal:
    """Verify merged regions match between parser and calamine."""

    def test_merged_region_count(self, programmatic_xlsx):
        parser_result = parse_workbook(path=programmatic_xlsx)
        calamine = CalamineResult.from_path(programmatic_xlsx)

        for sheet in parser_result.workbook.sheets:
            cal_sheet = calamine.sheets.get(sheet.sheet_name)
            if not cal_sheet or cal_sheet.merged_ranges is None:
                continue
            parser_count = len(sheet.merged_regions)
            cal_count = len(cal_sheet.merged_ranges)
            assert parser_count == cal_count, (
                f"Sheet '{sheet.sheet_name}' merge count: "
                f"parser={parser_count}, calamine={cal_count}"
            )

    def test_merged_region_ranges(self, programmatic_xlsx):
        parser_result = parse_workbook(path=programmatic_xlsx)
        calamine = CalamineResult.from_path(programmatic_xlsx)

        for sheet in parser_result.workbook.sheets:
            cal_sheet = calamine.sheets.get(sheet.sheet_name)
            if not cal_sheet or cal_sheet.merged_ranges is None:
                continue
            # Convert calamine ranges to comparable format
            # calamine: ((start_row, start_col), (end_row, end_col)) 0-indexed
            cal_ranges = set()
            for (sr, sc), (er, ec) in cal_sheet.merged_ranges:
                cal_ranges.add((sr + 1, sc + 1, er + 1, ec + 1))

            parser_ranges = set()
            for region in sheet.merged_regions:
                parser_ranges.add((
                    region.range.top_left.row,
                    region.range.top_left.col,
                    region.range.bottom_right.row,
                    region.range.bottom_right.col,
                ))

            assert parser_ranges == cal_ranges, (
                f"Sheet '{sheet.sheet_name}' merge ranges differ:\n"
                f"  parser:   {sorted(parser_ranges)}\n"
                f"  calamine: {sorted(cal_ranges)}"
            )


@pytest.mark.crossval
class TestMismatchRateCrossVal:
    """Overall mismatch rate must be below threshold."""

    def test_overall_mismatch_rate(self, programmatic_xlsx):
        parser_result = parse_workbook(path=programmatic_xlsx)
        calamine = CalamineResult.from_path(programmatic_xlsx)
        mismatches = _collect_mismatches(
            parser_result, calamine, formula_cells=False
        )
        total_cells = sum(
            s.cell_count() for s in parser_result.workbook.sheets
        )
        if total_cells > 0:
            rate = len(mismatches) / total_cells
            assert rate < 0.01, (
                f"Mismatch rate {rate:.1%} ({len(mismatches)}/{total_cells}) "
                f"exceeds 1% threshold"
            )


# ---------------------------------------------------------------------------
# Cross-validation on static files (examples + github datasets)
# ---------------------------------------------------------------------------


@pytest.mark.crossval
class TestSheetNamesStatic:
    def test_sheet_names_match(self, static_xlsx):
        parser_result = parse_workbook(path=static_xlsx)
        calamine = CalamineResult.from_path(static_xlsx)
        parser_names = [s.sheet_name for s in parser_result.workbook.sheets]
        assert parser_names == calamine.sheet_names


@pytest.mark.crossval
class TestCellValuesStatic:
    def test_non_formula_values_match(self, static_xlsx):
        parser_result = parse_workbook(path=static_xlsx)
        calamine = CalamineResult.from_path(static_xlsx)
        mismatches = _collect_mismatches(parser_result, calamine, formula_cells=False)
        total_cells = sum(s.cell_count() for s in parser_result.workbook.sheets)
        if total_cells > 0:
            rate = len(mismatches) / total_cells
            assert rate < 0.01, (
                f"{static_xlsx.name}: {len(mismatches)}/{total_cells} "
                f"({rate:.1%}) mismatches:\n"
                + _format_mismatches(mismatches[:10])
            )

    def test_formula_cached_values_match(self, static_xlsx):
        """For real-world files, formula cached values should match calamine.

        Threshold: <5% mismatch overall. A handful of files with highly nested
        dynamic-array or volatile formulas are known to exceed this because
        openpyxl doesn't always surface the latest cached value Excel wrote —
        we allow up to 15% for those, tracked in docs/PARSER_KNOWN_ISSUES.md.
        """
        known_loose_files = {
            "Walbridge Coatings 8.9.23.xlsx",  # openpyxl cached-value gap
        }
        threshold = 0.15 if static_xlsx.name in known_loose_files else 0.05

        parser_result = parse_workbook(path=static_xlsx)
        calamine = CalamineResult.from_path(static_xlsx)
        mismatches = _collect_mismatches(parser_result, calamine, formula_cells=True)
        hard_mismatches = [
            m for m in mismatches
            if m.parser_value is not None and m.calamine_value is not None
        ]
        total_formulas = sum(
            1 for s in parser_result.workbook.sheets
            for c in s.cells.values()
            if c.formula
        )
        if total_formulas > 0 and len(hard_mismatches) > 0:
            rate = len(hard_mismatches) / total_formulas
            assert rate < threshold, (
                f"{static_xlsx.name}: {len(hard_mismatches)}/{total_formulas} "
                f"formula mismatches ({rate:.1%}, threshold {threshold:.0%}):\n"
                + _format_mismatches(hard_mismatches[:10])
            )


@pytest.mark.crossval
class TestDimensionsStatic:
    def test_dimensions_similar(self, static_xlsx):
        parser_result = parse_workbook(path=static_xlsx)
        calamine = CalamineResult.from_path(static_xlsx)
        for sheet in parser_result.workbook.sheets:
            cal_sheet = calamine.sheets.get(sheet.sheet_name)
            if not cal_sheet or not sheet.used_range:
                continue
            if cal_sheet.start is not None and cal_sheet.end is not None:
                parser_rows = sheet.used_range.row_count()
                cal_rows = cal_sheet.end[0] - cal_sheet.start[0] + 1
                # Allow ±5 for real-world files (empty trailing rows)
                assert abs(parser_rows - cal_rows) <= 5, (
                    f"{static_xlsx.name} sheet '{sheet.sheet_name}' rows: "
                    f"parser={parser_rows}, calamine={cal_rows}"
                )


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _collect_mismatches(
    parser_result,
    calamine: CalamineResult,
    formula_cells: bool,
) -> list[Mismatch]:
    """Collect all mismatches between parser and calamine."""
    mismatches = []
    for sheet in parser_result.workbook.sheets:
        cal_sheet = calamine.sheets.get(sheet.sheet_name)
        if not cal_sheet:
            continue

        for cell in sheet.cells.values():
            # Filter by formula/non-formula
            if formula_cells and not cell.formula:
                continue
            if not formula_cells and cell.formula:
                continue

            # Skip merged slaves
            if cell.is_merged_slave:
                continue

            cal_val = cal_sheet.get_value(cell.coord.row, cell.coord.col)

            if not compare_cell_value(cell, cal_val):
                parser_val = (
                    cell.formula_value if cell.formula else cell.raw_value
                )
                mismatches.append(Mismatch(
                    sheet=sheet.sheet_name,
                    row=cell.coord.row,
                    col=cell.coord.col,
                    a1_ref=cell.a1_ref,
                    parser_value=parser_val,
                    calamine_value=cal_val,
                    category="formula" if cell.formula else "value",
                ))

    return mismatches


def _format_mismatches(mismatches: list[Mismatch]) -> str:
    """Format mismatch list for error messages."""
    lines = []
    for m in mismatches:
        lines.append(
            f"  {m.a1_ref}: parser={m.parser_value!r} ({type(m.parser_value).__name__}) "
            f"vs calamine={m.calamine_value!r} ({type(m.calamine_value).__name__})"
        )
    return "\n".join(lines)
