"""
Tests for formula reference parsing and dependency graph construction.

Covers A1 refs, ranges, cross-sheet refs, external refs, structured refs,
and circular reference detection.
"""

import pytest

from xlsx_parser.formula.formula_parser import FormulaParser
from xlsx_parser.models import CellCoord, DependencyGraph, DependencyEdgeDTO, EdgeType


class TestFormulaParser:
    """Test formula reference extraction."""

    def setup_method(self):
        self.parser = FormulaParser()

    def test_simple_cell_ref(self):
        refs = self.parser.parse("A1+B1", "Sheet1")
        assert len(refs) == 2
        assert refs[0].coord.row == 1
        assert refs[0].coord.col == 1
        assert refs[1].coord.row == 1
        assert refs[1].coord.col == 2

    def test_range_ref(self):
        refs = self.parser.parse("SUM(A1:A10)", "Sheet1")
        assert len(refs) == 1
        assert refs[0].range is not None
        assert refs[0].range.top_left.to_a1() == "A1"
        assert refs[0].range.bottom_right.to_a1() == "A10"

    def test_cross_sheet_ref(self):
        refs = self.parser.parse("Sheet2!A1+Sheet2!B1", "Sheet1")
        assert len(refs) == 2
        assert refs[0].sheet_name == "Sheet2"
        assert refs[1].sheet_name == "Sheet2"

    def test_quoted_sheet_ref(self):
        refs = self.parser.parse("'My Sheet'!A1", "Sheet1")
        assert len(refs) == 1
        assert refs[0].sheet_name == "My Sheet"

    def test_absolute_refs(self):
        refs = self.parser.parse("$A$1+$B$2:$C$3", "Sheet1")
        assert len(refs) == 2
        assert refs[0].coord.to_a1() == "A1"
        assert refs[1].range.to_a1() == "B2:C3"

    def test_external_ref(self):
        refs = self.parser.parse("[Budget.xlsx]Sheet1!A1", "Sheet1")
        external = [r for r in refs if r.is_external]
        assert len(external) == 1
        assert external[0].external_workbook == "Budget.xlsx"
        assert external[0].sheet_name == "Sheet1"

    def test_structured_table_ref(self):
        refs = self.parser.parse("SalesData[Revenue]", "Sheet1")
        structured = [r for r in refs if r.is_structured]
        assert len(structured) == 1
        assert structured[0].table_name == "SalesData"

    def test_complex_formula(self):
        formula = "IF(Inputs!B1>0,SUM(A1:A10)*Inputs!B3,0)"
        refs = self.parser.parse(formula, "Calcs")
        # Should find: Inputs!B1, A1:A10, Inputs!B3
        assert len(refs) >= 3

    def test_no_refs(self):
        refs = self.parser.parse("1+2+3", "Sheet1")
        assert len(refs) == 0

    def test_function_not_treated_as_table(self):
        refs = self.parser.parse("SUM(A1:A10)", "Sheet1")
        structured = [r for r in refs if r.is_structured]
        assert len(structured) == 0


class TestDependencyGraph:
    """Test dependency graph construction and traversal."""

    def test_upstream_traversal(self):
        graph = DependencyGraph()
        # B1 = A1 + A2
        graph.add_edge(DependencyEdgeDTO(
            source_sheet="S1", source_coord=CellCoord(row=1, col=2),
            target_sheet="S1", target_coord=CellCoord(row=1, col=1),
            target_ref_string="A1", edge_type=EdgeType.CELL_TO_CELL,
        ))
        graph.add_edge(DependencyEdgeDTO(
            source_sheet="S1", source_coord=CellCoord(row=1, col=2),
            target_sheet="S1", target_coord=CellCoord(row=2, col=1),
            target_ref_string="A2", edge_type=EdgeType.CELL_TO_CELL,
        ))
        graph.build_indexes()

        upstream = graph.get_upstream("S1", CellCoord(row=1, col=2))
        assert len(upstream) == 2

    def test_downstream_traversal(self):
        graph = DependencyGraph()
        # B1 depends on A1
        graph.add_edge(DependencyEdgeDTO(
            source_sheet="S1", source_coord=CellCoord(row=1, col=2),
            target_sheet="S1", target_coord=CellCoord(row=1, col=1),
            target_ref_string="A1", edge_type=EdgeType.CELL_TO_CELL,
        ))
        graph.build_indexes()

        downstream = graph.get_downstream("S1", CellCoord(row=1, col=1))
        assert len(downstream) == 1
        assert downstream[0].source_coord.to_a1() == "B1"

    def test_circular_ref_detection(self):
        graph = DependencyGraph()
        # A1 = B1, B1 = A1
        graph.add_edge(DependencyEdgeDTO(
            source_sheet="S1", source_coord=CellCoord(row=1, col=1),
            target_sheet="S1", target_coord=CellCoord(row=1, col=2),
            target_ref_string="B1", edge_type=EdgeType.CELL_TO_CELL,
        ))
        graph.add_edge(DependencyEdgeDTO(
            source_sheet="S1", source_coord=CellCoord(row=1, col=2),
            target_sheet="S1", target_coord=CellCoord(row=1, col=1),
            target_ref_string="A1", edge_type=EdgeType.CELL_TO_CELL,
        ))
        graph.build_indexes()

        circular = graph.detect_circular_refs()
        assert len(circular) > 0
        assert graph.has_circular_refs

    def test_cross_sheet_edge(self):
        graph = DependencyGraph()
        graph.add_edge(DependencyEdgeDTO(
            source_sheet="Sheet2", source_coord=CellCoord(row=1, col=1),
            target_sheet="Sheet1", target_coord=CellCoord(row=1, col=1),
            target_ref_string="Sheet1!A1", edge_type=EdgeType.CROSS_SHEET,
        ))
        graph.build_indexes()

        upstream = graph.get_upstream("Sheet2", CellCoord(row=1, col=1))
        assert len(upstream) == 1
        assert upstream[0].edge_type == EdgeType.CROSS_SHEET
