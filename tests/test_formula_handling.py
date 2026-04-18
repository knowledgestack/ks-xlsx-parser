"""
Tests for formula extraction, parsing, dependency graph, and rendering.

Verifies that the parser correctly:
- Extracts formula strings from cells (stripping the leading '=')
- Sets data_type to 'f' for formula cells
- Provides a display_value fallback showing the formula when no computed value exists
- Parses formula references (A1-style, cross-sheet, range, external, structured)
- Builds the dependency graph with correct edge types and cycle detection
- Counts formulas in workbook stats
- Renders formula cells in chunk output
"""

import pytest

from formula.dependency_builder import DependencyBuilder
from formula.formula_parser import FormulaParser, ParsedReference
from models.common import BlockType, CellCoord, EdgeType
from parsers import WorkbookParser
from pipeline import parse_workbook


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _parse(path):
    """Parse a workbook and return the WorkbookDTO."""
    return WorkbookParser(path=path).parse()


def _get_formula_cells(sheet):
    """Return all cells that have a formula in the given sheet."""
    return [c for c in sheet.cells.values() if c.formula is not None]


# ---------------------------------------------------------------------------
# FormulaParser unit tests
# ---------------------------------------------------------------------------


class TestFormulaParserCellRefs:
    """Test FormulaParser extraction of A1-style cell references."""

    def setup_method(self):
        self.parser = FormulaParser()

    def test_simple_cell_ref(self):
        refs = self.parser.parse("A1+B1", "Sheet1")
        ref_strs = {r.ref_string for r in refs}
        assert "A1" in ref_strs
        assert "B1" in ref_strs

    def test_absolute_cell_ref(self):
        refs = self.parser.parse("$A$1+$B$1", "Sheet1")
        assert len(refs) == 2
        for r in refs:
            assert r.coord is not None

    def test_range_ref(self):
        refs = self.parser.parse("SUM(A1:B10)", "Sheet1")
        range_refs = [r for r in refs if r.range is not None]
        assert len(range_refs) == 1
        assert range_refs[0].range.top_left == CellCoord(row=1, col=1)
        assert range_refs[0].range.bottom_right == CellCoord(row=10, col=2)

    def test_multiple_range_refs(self):
        refs = self.parser.parse("SUM(A1:A10)+SUM(B1:B10)", "Sheet1")
        range_refs = [r for r in refs if r.range is not None]
        assert len(range_refs) == 2

    def test_mixed_refs(self):
        refs = self.parser.parse("A1+SUM(B1:B10)+C5", "Sheet1")
        cell_refs = [r for r in refs if r.coord is not None]
        range_refs = [r for r in refs if r.range is not None]
        assert len(cell_refs) == 2  # A1, C5
        assert len(range_refs) == 1  # B1:B10


class TestFormulaParserCrossSheet:
    """Test FormulaParser extraction of cross-sheet references."""

    def setup_method(self):
        self.parser = FormulaParser()

    def test_simple_cross_sheet(self):
        refs = self.parser.parse("Sheet2!A1", "Sheet1")
        assert len(refs) == 1
        assert refs[0].sheet_name == "Sheet2"
        assert refs[0].coord == CellCoord(row=1, col=1)

    def test_quoted_sheet_name(self):
        refs = self.parser.parse("'My Sheet'!B5", "Sheet1")
        assert len(refs) == 1
        assert refs[0].sheet_name == "My Sheet"

    def test_cross_sheet_range(self):
        refs = self.parser.parse("SUM(Revenue!A1:A10)", "Summary")
        assert len(refs) == 1
        assert refs[0].sheet_name == "Revenue"
        assert refs[0].range is not None

    def test_multiple_sheets(self):
        refs = self.parser.parse("Revenue!A1+Costs!A1", "Summary")
        sheets = {r.sheet_name for r in refs}
        assert "Revenue" in sheets
        assert "Costs" in sheets


class TestFormulaParserExternalRefs:
    """Test FormulaParser extraction of external workbook references."""

    def setup_method(self):
        self.parser = FormulaParser()

    def test_external_ref(self):
        refs = self.parser.parse("[Budget.xlsx]Sheet1!A1", "Sheet1")
        ext_refs = [r for r in refs if r.is_external]
        assert len(ext_refs) == 1
        assert ext_refs[0].external_workbook == "Budget.xlsx"

    def test_external_range_ref(self):
        refs = self.parser.parse("[Data.xlsx]Sales!A1:C10", "Sheet1")
        ext_refs = [r for r in refs if r.is_external]
        assert len(ext_refs) == 1
        assert ext_refs[0].range is not None


class TestFormulaParserStructuredRefs:
    """Test FormulaParser extraction of structured table references."""

    def setup_method(self):
        self.parser = FormulaParser()

    def test_table_column_ref(self):
        refs = self.parser.parse("SalesTable[Amount]", "Sheet1")
        struct_refs = [r for r in refs if r.is_structured]
        assert len(struct_refs) == 1
        assert struct_refs[0].table_name == "SalesTable"

    def test_function_not_mistaken_for_table(self):
        refs = self.parser.parse("SUM(A1:A10)", "Sheet1")
        struct_refs = [r for r in refs if r.is_structured]
        assert len(struct_refs) == 0


class TestFormulaParserEdgeCases:
    """Test edge cases in formula parsing."""

    def setup_method(self):
        self.parser = FormulaParser()

    def test_empty_formula(self):
        refs = self.parser.parse("", "Sheet1")
        assert refs == []

    def test_constant_only(self):
        refs = self.parser.parse("42", "Sheet1")
        # No cell references expected
        cell_refs = [r for r in refs if r.coord is not None or r.range is not None]
        assert len(cell_refs) == 0

    def test_nested_functions(self):
        refs = self.parser.parse('IF(A1>0,SUM(B1:B10),AVERAGE(C1:C10))', "Sheet1")
        assert len(refs) >= 3  # A1, B1:B10, C1:C10

    def test_dedup_repeated_ref(self):
        refs = self.parser.parse("A1+A1+A1", "Sheet1")
        # Should be deduped
        assert len(refs) == 1


# ---------------------------------------------------------------------------
# Cell-level formula extraction tests
# ---------------------------------------------------------------------------


class TestSimpleFormulaExtraction:
    """Test that formula strings are correctly extracted from cells."""

    def test_arithmetic_formulas_extracted(self, simple_formulas):
        wb = _parse(simple_formulas)
        sheet = wb.sheets[0]
        formula_cells = _get_formula_cells(sheet)
        assert len(formula_cells) > 0

    def test_formula_string_strips_equals(self, simple_formulas):
        wb = _parse(simple_formulas)
        sheet = wb.sheets[0]
        # C2 = "=A2+B2"
        c2 = sheet.get_cell(2, 3)
        assert c2 is not None
        assert c2.formula == "A2+B2"
        assert not c2.formula.startswith("=")

    def test_data_type_is_f(self, simple_formulas):
        wb = _parse(simple_formulas)
        sheet = wb.sheets[0]
        c2 = sheet.get_cell(2, 3)
        assert c2.data_type == "f"

    def test_sum_formula(self, simple_formulas):
        wb = _parse(simple_formulas)
        sheet = wb.sheets[0]
        # Row 7 col 1 = "=SUM(A2:A6)"
        a7 = sheet.get_cell(7, 1)
        assert a7 is not None
        assert a7.formula == "SUM(A2:A6)"

    def test_average_formula(self, simple_formulas):
        wb = _parse(simple_formulas)
        sheet = wb.sheets[0]
        d7 = sheet.get_cell(7, 4)
        assert d7 is not None
        assert d7.formula == "AVERAGE(D2:D6)"

    def test_max_formula(self, simple_formulas):
        wb = _parse(simple_formulas)
        sheet = wb.sheets[0]
        e7 = sheet.get_cell(7, 5)
        assert e7 is not None
        assert e7.formula == "MAX(E2:E6)"

    def test_min_formula(self, simple_formulas):
        wb = _parse(simple_formulas)
        sheet = wb.sheets[0]
        f7 = sheet.get_cell(7, 6)
        assert f7 is not None
        assert f7.formula == "MIN(F2:F6)"

    def test_count_formula(self, simple_formulas):
        wb = _parse(simple_formulas)
        sheet = wb.sheets[0]
        a8 = sheet.get_cell(8, 1)
        assert a8 is not None
        assert a8.formula == "COUNT(A2:A6)"

    def test_counta_formula(self, simple_formulas):
        wb = _parse(simple_formulas)
        sheet = wb.sheets[0]
        b8 = sheet.get_cell(8, 2)
        assert b8 is not None
        assert b8.formula == "COUNTA(B2:B6)"

    def test_all_arithmetic_ops(self, simple_formulas):
        wb = _parse(simple_formulas)
        sheet = wb.sheets[0]
        c2 = sheet.get_cell(2, 3)  # A2+B2
        d2 = sheet.get_cell(2, 4)  # A2-B2
        e2 = sheet.get_cell(2, 5)  # A2*B2
        f2 = sheet.get_cell(2, 6)  # A2/B2
        assert "+" in c2.formula
        assert "-" in d2.formula
        assert "*" in e2.formula
        assert "/" in f2.formula

    def test_formula_count_in_workbook(self, simple_formulas):
        wb = _parse(simple_formulas)
        assert wb.total_formulas > 0
        # 5 rows x 4 formula cols + 6 summary + 2 count = 28
        assert wb.total_formulas == 28

    def test_display_value_for_formula_cells(self, simple_formulas):
        wb = _parse(simple_formulas)
        sheet = wb.sheets[0]
        # Formula cells in programmatic files have no computed value;
        # display_value should fall back to showing the formula
        c2 = sheet.get_cell(2, 3)
        assert c2.display_value is not None
        assert c2.display_value == "=A2+B2"

    def test_non_formula_cells_unaffected(self, simple_formulas):
        wb = _parse(simple_formulas)
        sheet = wb.sheets[0]
        a2 = sheet.get_cell(2, 1)  # value = 10
        assert a2.formula is None
        assert a2.data_type != "f"
        assert a2.raw_value == 10


class TestNestedFormulas:
    """Test nested IF, conditional, and aggregate formulas."""

    def test_nested_if_extracted(self, nested_formulas):
        wb = _parse(nested_formulas)
        sheet = wb.sheets[0]
        b2 = sheet.get_cell(2, 2)
        assert b2 is not None
        assert b2.formula is not None
        assert "IF" in b2.formula

    def test_nested_if_structure(self, nested_formulas):
        wb = _parse(nested_formulas)
        sheet = wb.sheets[0]
        b2 = sheet.get_cell(2, 2)
        # Should contain nested IF calls
        assert b2.formula.count("IF") >= 4

    def test_and_function_in_formula(self, nested_formulas):
        wb = _parse(nested_formulas)
        sheet = wb.sheets[0]
        c2 = sheet.get_cell(2, 3)
        assert c2 is not None
        assert "AND" in c2.formula

    def test_iferror_formula(self, nested_formulas):
        wb = _parse(nested_formulas)
        sheet = wb.sheets[0]
        d2 = sheet.get_cell(2, 4)
        assert d2 is not None
        assert "IFERROR" in d2.formula

    def test_countif_formula(self, nested_formulas):
        wb = _parse(nested_formulas)
        sheet = wb.sheets[0]
        b11 = sheet.get_cell(11, 2)
        assert b11 is not None
        assert "COUNTIF" in b11.formula

    def test_sumif_formula(self, nested_formulas):
        wb = _parse(nested_formulas)
        sheet = wb.sheets[0]
        b12 = sheet.get_cell(12, 2)
        assert b12 is not None
        assert "SUMIF" in b12.formula

    def test_all_formula_cells_have_type_f(self, nested_formulas):
        wb = _parse(nested_formulas)
        sheet = wb.sheets[0]
        for cell in _get_formula_cells(sheet):
            assert cell.data_type == "f", f"Cell {cell.coord.to_a1()} should be type 'f'"

    def test_display_value_shows_formula(self, nested_formulas):
        wb = _parse(nested_formulas)
        sheet = wb.sheets[0]
        b2 = sheet.get_cell(2, 2)
        assert b2.display_value is not None
        assert b2.display_value.startswith("=")
        assert "IF" in b2.display_value


class TestCrossSheetFormulas:
    """Test formulas that reference cells on other sheets."""

    def test_cross_sheet_ref_extracted(self, cross_sheet_formulas):
        wb = _parse(cross_sheet_formulas)
        summary = [s for s in wb.sheets if s.sheet_name == "Summary"][0]
        b2 = summary.get_cell(2, 2)  # =Revenue!A2
        assert b2 is not None
        assert b2.formula is not None
        assert "Revenue!" in b2.formula

    def test_cross_sheet_formula_string(self, cross_sheet_formulas):
        wb = _parse(cross_sheet_formulas)
        summary = [s for s in wb.sheets if s.sheet_name == "Summary"][0]
        b2 = summary.get_cell(2, 2)
        assert b2.formula == "Revenue!A2"

    def test_costs_sheet_ref(self, cross_sheet_formulas):
        wb = _parse(cross_sheet_formulas)
        summary = [s for s in wb.sheets if s.sheet_name == "Summary"][0]
        c2 = summary.get_cell(2, 3)  # =Costs!A2
        assert c2 is not None
        assert "Costs!" in c2.formula

    def test_local_ref_on_summary(self, cross_sheet_formulas):
        wb = _parse(cross_sheet_formulas)
        summary = [s for s in wb.sheets if s.sheet_name == "Summary"][0]
        d2 = summary.get_cell(2, 4)  # =B2-C2
        assert d2.formula == "B2-C2"
        assert "!" not in d2.formula

    def test_three_sheets_present(self, cross_sheet_formulas):
        wb = _parse(cross_sheet_formulas)
        assert len(wb.sheets) == 3

    def test_source_sheets_have_no_formulas(self, cross_sheet_formulas):
        wb = _parse(cross_sheet_formulas)
        revenue = [s for s in wb.sheets if s.sheet_name == "Revenue"][0]
        costs = [s for s in wb.sheets if s.sheet_name == "Costs"][0]
        assert len(_get_formula_cells(revenue)) == 0
        assert len(_get_formula_cells(costs)) == 0

    def test_summary_formula_count(self, cross_sheet_formulas):
        wb = _parse(cross_sheet_formulas)
        summary = [s for s in wb.sheets if s.sheet_name == "Summary"][0]
        formulas = _get_formula_cells(summary)
        # 4 quarters x 4 formula cols + 4 totals = 20
        assert len(formulas) == 20


class TestTextFormulas:
    """Test text manipulation formulas."""

    def test_concatenate_formula(self, text_formulas):
        wb = _parse(text_formulas)
        sheet = wb.sheets[0]
        c2 = sheet.get_cell(2, 3)
        assert c2 is not None
        assert "CONCATENATE" in c2.formula

    def test_upper_formula(self, text_formulas):
        wb = _parse(text_formulas)
        sheet = wb.sheets[0]
        d2 = sheet.get_cell(2, 4)
        assert d2 is not None
        assert "UPPER" in d2.formula

    def test_left_function(self, text_formulas):
        wb = _parse(text_formulas)
        sheet = wb.sheets[0]
        e2 = sheet.get_cell(2, 5)
        assert e2 is not None
        assert "LEFT" in e2.formula

    def test_len_formula(self, text_formulas):
        wb = _parse(text_formulas)
        sheet = wb.sheets[0]
        b5 = sheet.get_cell(5, 2)
        assert b5 is not None
        assert "LEN" in b5.formula

    def test_trim_formula(self, text_formulas):
        wb = _parse(text_formulas)
        sheet = wb.sheets[0]
        c5 = sheet.get_cell(5, 3)
        assert c5 is not None
        assert "TRIM" in c5.formula


class TestLookupFormulas:
    """Test VLOOKUP and INDEX/MATCH formulas."""

    def test_vlookup_extracted(self, lookup_formulas):
        wb = _parse(lookup_formulas)
        sheet = wb.sheets[0]
        f2 = sheet.get_cell(2, 6)
        assert f2 is not None
        assert "VLOOKUP" in f2.formula

    def test_vlookup_arguments_preserved(self, lookup_formulas):
        wb = _parse(lookup_formulas)
        sheet = wb.sheets[0]
        f2 = sheet.get_cell(2, 6)
        assert "E2" in f2.formula
        assert "A2:C5" in f2.formula
        assert "FALSE" in f2.formula

    def test_index_match_extracted(self, lookup_formulas):
        wb = _parse(lookup_formulas)
        sheet = wb.sheets[0]
        g2 = sheet.get_cell(2, 7)
        assert g2 is not None
        assert "INDEX" in g2.formula
        assert "MATCH" in g2.formula

    def test_multiple_vlookups(self, lookup_formulas):
        wb = _parse(lookup_formulas)
        sheet = wb.sheets[0]
        f2 = sheet.get_cell(2, 6)
        f3 = sheet.get_cell(3, 6)
        assert f2.formula is not None
        assert f3.formula is not None
        # Different lookup values
        assert "E2" in f2.formula
        assert "E3" in f3.formula


# ---------------------------------------------------------------------------
# Dependency graph tests
# ---------------------------------------------------------------------------


class TestDependencyGraph:
    """Test the dependency graph builder."""

    def test_simple_deps_built(self, simple_formulas):
        wb = _parse(simple_formulas)
        graph = wb.dependency_graph
        assert len(graph.edges) > 0

    def test_arithmetic_creates_cell_to_cell_edges(self, simple_formulas):
        wb = _parse(simple_formulas)
        graph = wb.dependency_graph
        cell_edges = [e for e in graph.edges if e.edge_type == EdgeType.CELL_TO_CELL]
        assert len(cell_edges) > 0

    def test_sum_creates_range_edge(self, simple_formulas):
        wb = _parse(simple_formulas)
        graph = wb.dependency_graph
        range_edges = [e for e in graph.edges if e.edge_type == EdgeType.CELL_TO_RANGE]
        assert len(range_edges) > 0

    def test_cross_sheet_edges(self, cross_sheet_formulas):
        wb = _parse(cross_sheet_formulas)
        graph = wb.dependency_graph
        cross_edges = [e for e in graph.edges if e.edge_type == EdgeType.CROSS_SHEET]
        assert len(cross_edges) > 0

    def test_cross_sheet_targets_correct_sheets(self, cross_sheet_formulas):
        wb = _parse(cross_sheet_formulas)
        graph = wb.dependency_graph
        cross_edges = [e for e in graph.edges if e.edge_type == EdgeType.CROSS_SHEET]
        target_sheets = {e.target_sheet for e in cross_edges}
        assert "Revenue" in target_sheets
        assert "Costs" in target_sheets

    def test_upstream_traversal(self, simple_formulas):
        wb = _parse(simple_formulas)
        graph = wb.dependency_graph
        # C2 = A2 + B2 → should have upstream to A2, B2
        upstream = graph.get_upstream("Formulas", CellCoord(row=2, col=3))
        assert len(upstream) >= 2

    def test_downstream_traversal(self, simple_formulas):
        wb = _parse(simple_formulas)
        graph = wb.dependency_graph
        # A2 is referenced by C2, D2, E2, F2 and SUM(A2:A6) row 7
        downstream = graph.get_downstream("Formulas", CellCoord(row=2, col=1))
        assert len(downstream) >= 1

    def test_edge_source_coords(self, simple_formulas):
        wb = _parse(simple_formulas)
        graph = wb.dependency_graph
        # Every edge should have a valid source
        for edge in graph.edges:
            assert edge.source_sheet is not None
            assert edge.source_coord is not None
            assert edge.source_coord.row > 0
            assert edge.source_coord.col > 0


class TestCircularReferences:
    """Test circular reference detection."""

    def test_direct_circular_detected(self, circular_ref_formulas):
        wb = _parse(circular_ref_formulas)
        graph = wb.dependency_graph
        circular = graph.detect_circular_refs()
        assert len(circular) > 0

    def test_circular_cells_identified(self, circular_ref_formulas):
        wb = _parse(circular_ref_formulas)
        graph = wb.dependency_graph
        circular = graph.detect_circular_refs()
        # A1 <-> B1 should be circular
        circular_strs = {str(c) for c in circular}
        has_a1 = any("A1" in c for c in circular_strs)
        has_b1 = any("B1" in c for c in circular_strs)
        assert has_a1 or has_b1

    def test_indirect_circular_detected(self, circular_ref_formulas):
        wb = _parse(circular_ref_formulas)
        graph = wb.dependency_graph
        circular = graph.detect_circular_refs()
        # A3 -> C3 -> B3 -> A3 should be circular
        circular_strs = {str(c) for c in circular}
        has_row3 = any("3" in c for c in circular_strs)
        assert has_row3

    def test_non_circular_chain_not_flagged(self, circular_ref_formulas):
        wb = _parse(circular_ref_formulas)
        graph = wb.dependency_graph
        circular = graph.detect_circular_refs()
        # A5, B5, C5 form a chain but NOT circular
        circular_strs = {str(c) for c in circular}
        has_a5 = any("A5" in c for c in circular_strs)
        assert not has_a5

    def test_has_circular_refs_property(self, circular_ref_formulas):
        wb = _parse(circular_ref_formulas)
        graph = wb.dependency_graph
        graph.detect_circular_refs()
        assert graph.has_circular_refs is True

    def test_formula_strings_still_extracted_for_circular(self, circular_ref_formulas):
        wb = _parse(circular_ref_formulas)
        sheet = wb.sheets[0]
        a1 = sheet.get_cell(1, 1)
        b1 = sheet.get_cell(1, 2)
        assert a1.formula == "B1+1"
        assert b1.formula == "A1+1"


# ---------------------------------------------------------------------------
# Mixed formula types & cross-sheet dependency tests
# ---------------------------------------------------------------------------


class TestMixedFormulaTypes:
    """Test workbook with mixed formula types across sheets."""

    def test_data_sheet_has_no_formulas(self, mixed_formula_types):
        wb = _parse(mixed_formula_types)
        data = [s for s in wb.sheets if s.sheet_name == "Data"][0]
        assert len(_get_formula_cells(data)) == 0

    def test_analysis_sheet_has_formulas(self, mixed_formula_types):
        wb = _parse(mixed_formula_types)
        analysis = [s for s in wb.sheets if s.sheet_name == "Analysis"][0]
        formulas = _get_formula_cells(analysis)
        assert len(formulas) == 7  # B2 through B8

    def test_cross_sheet_sum(self, mixed_formula_types):
        wb = _parse(mixed_formula_types)
        analysis = [s for s in wb.sheets if s.sheet_name == "Analysis"][0]
        b2 = analysis.get_cell(2, 2)
        assert b2.formula == "SUM(Data!B2:B7)"

    def test_cross_sheet_average(self, mixed_formula_types):
        wb = _parse(mixed_formula_types)
        analysis = [s for s in wb.sheets if s.sheet_name == "Analysis"][0]
        b3 = analysis.get_cell(3, 2)
        assert b3.formula == "AVERAGE(Data!B2:B7)"

    def test_countif_with_cross_sheet(self, mixed_formula_types):
        wb = _parse(mixed_formula_types)
        analysis = [s for s in wb.sheets if s.sheet_name == "Analysis"][0]
        b6 = analysis.get_cell(6, 2)
        assert b6.formula is not None
        assert "COUNTIF" in b6.formula
        assert "Data!" in b6.formula

    def test_local_formula_refs(self, mixed_formula_types):
        wb = _parse(mixed_formula_types)
        analysis = [s for s in wb.sheets if s.sheet_name == "Analysis"][0]
        b8 = analysis.get_cell(8, 2)
        assert b8.formula == "B4-B5"

    def test_dependency_edges_cross_sheets(self, mixed_formula_types):
        wb = _parse(mixed_formula_types)
        graph = wb.dependency_graph
        cross_edges = [e for e in graph.edges if e.edge_type == EdgeType.CROSS_SHEET]
        assert len(cross_edges) > 0
        target_sheets = {e.target_sheet for e in cross_edges}
        assert "Data" in target_sheets


# ---------------------------------------------------------------------------
# Pipeline integration tests
# ---------------------------------------------------------------------------


class TestFormulaPipeline:
    """End-to-end tests for formula handling through the pipeline."""

    def test_total_formulas_counted(self, simple_formulas):
        result = parse_workbook(path=simple_formulas)
        assert result.workbook.total_formulas > 0

    def test_formula_cells_in_json(self, simple_formulas):
        result = parse_workbook(path=simple_formulas)
        json_data = result.to_json()
        # Chunks should contain formula cell content
        all_text = " ".join(c["render_text"] for c in json_data["chunks"])
        # Formula display values should appear in rendered text
        assert len(all_text) > 0

    def test_cross_sheet_json_has_formulas(self, cross_sheet_formulas):
        result = parse_workbook(path=cross_sheet_formulas)
        assert result.workbook.total_formulas > 0

    def test_nested_formula_json(self, nested_formulas):
        result = parse_workbook(path=nested_formulas)
        assert result.workbook.total_formulas > 0
        json_data = result.to_json()
        assert json_data["workbook"]["total_formulas"] > 0

    def test_formula_chunk_has_render_text(self, simple_formulas):
        result = parse_workbook(path=simple_formulas)
        for chunk in result.chunks:
            assert chunk.render_text is not None
            assert len(chunk.render_text) > 0

    def test_formula_chunk_has_render_html(self, simple_formulas):
        result = parse_workbook(path=simple_formulas)
        for chunk in result.chunks:
            assert chunk.render_html is not None
            assert "<table" in chunk.render_html.lower()

    def test_deterministic_formula_hashing(self, simple_formulas):
        r1 = parse_workbook(path=simple_formulas)
        r2 = parse_workbook(path=simple_formulas)
        assert r1.workbook.workbook_hash == r2.workbook.workbook_hash
        for c1, c2 in zip(r1.chunks, r2.chunks):
            assert c1.chunk_id == c2.chunk_id
            assert c1.content_hash == c2.content_hash

    def test_circular_ref_workbook_parses(self, circular_ref_formulas):
        result = parse_workbook(path=circular_ref_formulas)
        assert result.workbook.total_formulas > 0
        assert result.total_chunks >= 1

    def test_lookup_formulas_in_pipeline(self, lookup_formulas):
        result = parse_workbook(path=lookup_formulas)
        assert result.workbook.total_formulas > 0
        json_data = result.to_json()
        assert json_data["total_chunks"] >= 1


class TestFormulaInBlocks:
    """Test that formula cells are correctly reflected in block metadata."""

    def test_block_formula_count(self, simple_formulas):
        result = parse_workbook(path=simple_formulas)
        # The block should report formula_count > 0
        from chunking.segmenter import LayoutSegmenter
        sheet = result.workbook.sheets[0]
        tables = [t for t in result.workbook.tables if t.sheet_name == sheet.sheet_name]
        segmenter = LayoutSegmenter(sheet, tables=tables)
        blocks = segmenter.segment()
        total_formulas_in_blocks = sum(b.formula_count for b in blocks)
        assert total_formulas_in_blocks > 0

    def test_cross_sheet_block_has_formulas(self, cross_sheet_formulas):
        result = parse_workbook(path=cross_sheet_formulas)
        from chunking.segmenter import LayoutSegmenter
        summary = [s for s in result.workbook.sheets if s.sheet_name == "Summary"][0]
        tables = [t for t in result.workbook.tables if t.sheet_name == summary.sheet_name]
        segmenter = LayoutSegmenter(summary, tables=tables)
        blocks = segmenter.segment()
        formula_blocks = [b for b in blocks if b.formula_count > 0]
        assert len(formula_blocks) >= 1
