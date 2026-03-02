"""
Structural invariant tests for the xlsx_parser.

Tests properties that must always hold for any valid parse output,
regardless of the input file: merge structure, used range bounds,
cell identity, dependency graph, stats, tables, determinism.
"""

from __future__ import annotations

import json
import re

import pytest

from xlsx_parser.pipeline import parse_workbook
from xlsx_parser.models.common import EdgeType

from tests.helpers.invariant_checker import check_invariants


# ---------------------------------------------------------------------------
# Invariant tests on programmatic fixtures
# ---------------------------------------------------------------------------


@pytest.mark.invariant
class TestMergeInvariantsProgrammatic:
    """Merge invariants on programmatic fixtures."""

    def test_slave_has_master(self, programmatic_xlsx):
        result = parse_workbook(path=programmatic_xlsx)
        for sheet in result.workbook.sheets:
            for cell in sheet.cells.values():
                if cell.is_merged_slave:
                    assert cell.merge_master is not None, (
                        f"Cell {cell.a1_ref} is merged slave with no merge_master"
                    )

    def test_master_exists_and_flagged(self, programmatic_xlsx):
        result = parse_workbook(path=programmatic_xlsx)
        for sheet in result.workbook.sheets:
            for cell in sheet.cells.values():
                if cell.is_merged_slave and cell.merge_master:
                    master = sheet.get_cell(
                        cell.merge_master.row, cell.merge_master.col
                    )
                    assert master is not None, (
                        f"Slave {cell.a1_ref} master "
                        f"{cell.merge_master.to_a1()} missing"
                    )
                    assert master.is_merged_master, (
                        f"{cell.merge_master.to_a1()} referenced as master "
                        f"but is_merged_master=False"
                    )

    def test_master_is_top_left(self, programmatic_xlsx):
        result = parse_workbook(path=programmatic_xlsx)
        for sheet in result.workbook.sheets:
            for region in sheet.merged_regions:
                assert region.master == region.range.top_left, (
                    f"Region {region.range.to_a1()} master "
                    f"{region.master.to_a1()} != top_left "
                    f"{region.range.top_left.to_a1()}"
                )

    def test_not_both_master_and_slave(self, programmatic_xlsx):
        result = parse_workbook(path=programmatic_xlsx)
        for sheet in result.workbook.sheets:
            for cell in sheet.cells.values():
                assert not (cell.is_merged_master and cell.is_merged_slave), (
                    f"Cell {cell.a1_ref} is both master and slave"
                )


@pytest.mark.invariant
class TestUsedRangeInvariantsProgrammatic:
    """Used range invariants on programmatic fixtures."""

    def test_non_empty_sheet_has_used_range(self, programmatic_xlsx):
        result = parse_workbook(path=programmatic_xlsx)
        for sheet in result.workbook.sheets:
            if sheet.cells:
                assert sheet.used_range is not None, (
                    f"Sheet '{sheet.sheet_name}' has cells but no used_range"
                )

    def test_all_cells_within_used_range(self, programmatic_xlsx):
        result = parse_workbook(path=programmatic_xlsx)
        for sheet in result.workbook.sheets:
            if not sheet.used_range:
                continue
            for cell in sheet.cells.values():
                assert sheet.used_range.contains(cell.coord), (
                    f"{cell.a1_ref} outside used_range "
                    f"{sheet.used_range.to_a1()}"
                )

    def test_used_range_bounds_tight(self, programmatic_xlsx):
        result = parse_workbook(path=programmatic_xlsx)
        for sheet in result.workbook.sheets:
            if not sheet.cells or not sheet.used_range:
                continue
            rows = [c.coord.row for c in sheet.cells.values()]
            cols = [c.coord.col for c in sheet.cells.values()]
            assert sheet.used_range.top_left.row == min(rows)
            assert sheet.used_range.top_left.col == min(cols)
            assert sheet.used_range.bottom_right.row == max(rows)
            assert sheet.used_range.bottom_right.col == max(cols)


@pytest.mark.invariant
class TestCellIdentityInvariantsProgrammatic:
    """Cell identity invariants on programmatic fixtures."""

    def test_cell_ids_populated(self, programmatic_xlsx):
        result = parse_workbook(path=programmatic_xlsx)
        for sheet in result.workbook.sheets:
            for cell in sheet.cells.values():
                assert cell.cell_id, f"{cell.a1_ref} has empty cell_id"

    def test_cell_ids_unique(self, programmatic_xlsx):
        result = parse_workbook(path=programmatic_xlsx)
        all_ids = [
            c.cell_id
            for s in result.workbook.sheets
            for c in s.cells.values()
        ]
        assert len(all_ids) == len(set(all_ids)), "Duplicate cell_ids found"

    def test_cell_hash_format(self, programmatic_xlsx):
        hex_re = re.compile(r"^[0-9a-f]{16}$")
        result = parse_workbook(path=programmatic_xlsx)
        for sheet in result.workbook.sheets:
            for cell in sheet.cells.values():
                assert cell.cell_hash, f"{cell.a1_ref} has empty hash"
                assert hex_re.match(cell.cell_hash), (
                    f"{cell.a1_ref} hash '{cell.cell_hash}' not 16-char hex"
                )

    def test_cell_id_format(self, programmatic_xlsx):
        result = parse_workbook(path=programmatic_xlsx)
        for sheet in result.workbook.sheets:
            for cell in sheet.cells.values():
                expected = f"{cell.sheet_name}|{cell.coord.row}|{cell.coord.col}"
                assert cell.cell_id == expected, (
                    f"cell_id '{cell.cell_id}' != '{expected}'"
                )


@pytest.mark.invariant
class TestDependencyInvariantsProgrammatic:
    """Dependency graph invariants on programmatic fixtures."""

    def test_edge_sources_have_formulas(self, programmatic_xlsx):
        result = parse_workbook(path=programmatic_xlsx)
        sheet_map = {s.sheet_name: s for s in result.workbook.sheets}
        for edge in result.workbook.dependency_graph.edges:
            sheet = sheet_map.get(edge.source_sheet)
            assert sheet, f"Edge source sheet '{edge.source_sheet}' missing"
            cell = sheet.get_cell(edge.source_coord.row, edge.source_coord.col)
            assert cell, (
                f"Edge source {edge.source_sheet}!{edge.source_coord.to_a1()} "
                f"not in cells"
            )
            assert cell.formula is not None, (
                f"Edge source {cell.a1_ref} has no formula"
            )

    def test_cross_sheet_edges_cross_sheets(self, programmatic_xlsx):
        result = parse_workbook(path=programmatic_xlsx)
        for edge in result.workbook.dependency_graph.edges:
            if edge.edge_type == EdgeType.CROSS_SHEET:
                assert edge.target_sheet is not None
                assert edge.target_sheet != edge.source_sheet, (
                    f"Cross-sheet edge stays on {edge.source_sheet}"
                )

    def test_formula_count_matches_stats(self, programmatic_xlsx):
        result = parse_workbook(path=programmatic_xlsx)
        actual = sum(
            1 for s in result.workbook.sheets
            for c in s.cells.values()
            if c.formula is not None
        )
        assert result.workbook.total_formulas == actual


@pytest.mark.invariant
class TestStatsInvariantsProgrammatic:
    """Stats consistency invariants on programmatic fixtures."""

    def test_total_sheets(self, programmatic_xlsx):
        result = parse_workbook(path=programmatic_xlsx)
        assert result.workbook.total_sheets == len(result.workbook.sheets)

    def test_total_cells(self, programmatic_xlsx):
        result = parse_workbook(path=programmatic_xlsx)
        actual = sum(s.cell_count() for s in result.workbook.sheets)
        assert result.workbook.total_cells == actual


@pytest.mark.invariant
class TestTableInvariantsProgrammatic:
    """Table invariants on programmatic fixtures."""

    def test_table_sheet_exists(self, programmatic_xlsx):
        result = parse_workbook(path=programmatic_xlsx)
        sheet_names = {s.sheet_name for s in result.workbook.sheets}
        for table in result.workbook.tables:
            assert table.sheet_name in sheet_names, (
                f"Table '{table.table_name}' sheet "
                f"'{table.sheet_name}' missing"
            )

    def test_table_column_count(self, programmatic_xlsx):
        result = parse_workbook(path=programmatic_xlsx)
        for table in result.workbook.tables:
            assert len(table.columns) == table.ref_range.col_count(), (
                f"Table '{table.table_name}': {len(table.columns)} cols "
                f"vs range {table.ref_range.col_count()}"
            )

    def test_table_names_unique(self, programmatic_xlsx):
        result = parse_workbook(path=programmatic_xlsx)
        names = [t.table_name for t in result.workbook.tables]
        assert len(names) == len(set(names)), f"Duplicate table names: {names}"


@pytest.mark.invariant
class TestDeterminismProgrammatic:
    """Determinism invariants on programmatic fixtures."""

    def test_deterministic_hashes(self, programmatic_xlsx):
        r1 = parse_workbook(path=programmatic_xlsx)
        r2 = parse_workbook(path=programmatic_xlsx)
        assert r1.workbook.workbook_hash == r2.workbook.workbook_hash
        assert r1.workbook.workbook_id == r2.workbook.workbook_id

        ids1 = sorted(c.cell_id for s in r1.workbook.sheets for c in s.cells.values())
        ids2 = sorted(c.cell_id for s in r2.workbook.sheets for c in s.cells.values())
        assert ids1 == ids2

        hashes1 = sorted(c.cell_hash for s in r1.workbook.sheets for c in s.cells.values())
        hashes2 = sorted(c.cell_hash for s in r2.workbook.sheets for c in s.cells.values())
        assert hashes1 == hashes2

    def test_json_round_trip(self, programmatic_xlsx):
        result = parse_workbook(path=programmatic_xlsx)
        data = result.to_json()
        json_str = json.dumps(data)
        parsed = json.loads(json_str)
        assert "workbook" in parsed
        assert "chunks" in parsed
        assert parsed["workbook"]["workbook_hash"] == result.workbook.workbook_hash


@pytest.mark.invariant
class TestWorkbookHashInvariants:
    """Workbook-level hash invariants."""

    def test_workbook_hash_format(self, programmatic_xlsx):
        result = parse_workbook(path=programmatic_xlsx)
        hex_re = re.compile(r"^[0-9a-f]{16}$")
        assert hex_re.match(result.workbook.workbook_hash), (
            f"workbook_hash '{result.workbook.workbook_hash}' not 16-char hex"
        )

    def test_workbook_id_populated(self, programmatic_xlsx):
        result = parse_workbook(path=programmatic_xlsx)
        assert result.workbook.workbook_id, "workbook_id is empty"

    def test_sheet_ids_populated(self, programmatic_xlsx):
        result = parse_workbook(path=programmatic_xlsx)
        for sheet in result.workbook.sheets:
            assert sheet.sheet_id, (
                f"Sheet '{sheet.sheet_name}' has empty sheet_id"
            )


# ---------------------------------------------------------------------------
# Same invariants on static files (examples + github datasets)
# ---------------------------------------------------------------------------


@pytest.mark.invariant
class TestAllInvariantsStatic:
    """Run full invariant checker against each static xlsx file."""

    def test_all_invariants_pass(self, static_xlsx):
        result = parse_workbook(path=static_xlsx)
        violations = check_invariants(result.workbook)
        assert len(violations) == 0, (
            f"{len(violations)} violations in {static_xlsx.name}:\n"
            + "\n".join(violations[:10])
        )

    def test_deterministic_hashes(self, static_xlsx):
        r1 = parse_workbook(path=static_xlsx)
        r2 = parse_workbook(path=static_xlsx)
        assert r1.workbook.workbook_hash == r2.workbook.workbook_hash

    def test_json_serializable(self, static_xlsx):
        result = parse_workbook(path=static_xlsx)
        data = result.to_json()
        json.dumps(data)  # must not raise
