"""
Tests for workbook, sheet, and cell parsing.

Uses programmatically generated fixture workbooks to verify
correct extraction of values, formulas, styles, merges,
tables, comments, data validations, and sheet properties.
"""

import pytest

from xlsx_parser.parsers import WorkbookParser


class TestSimpleWorkbook:
    """Test basic cell value and formula extraction."""

    def test_parse_cell_values(self, simple_workbook):
        parser = WorkbookParser(path=simple_workbook)
        result = parser.parse()

        assert result.total_sheets == 1
        assert result.total_cells > 0

        sheet = result.sheets[0]
        assert sheet.sheet_name == "Sheet1"

        # Check header cell
        a1 = sheet.get_cell(1, 1)
        assert a1 is not None
        assert a1.raw_value == "Name"
        assert a1.style is not None
        assert a1.style.font.bold is True

    def test_parse_formula(self, simple_workbook):
        parser = WorkbookParser(path=simple_workbook)
        result = parser.parse()

        sheet = result.sheets[0]
        b4 = sheet.get_cell(4, 2)  # B4 = =B2-B3
        assert b4 is not None
        assert b4.formula is not None
        assert "B2" in b4.formula and "B3" in b4.formula

    def test_parse_number_format(self, simple_workbook):
        parser = WorkbookParser(path=simple_workbook)
        result = parser.parse()

        sheet = result.sheets[0]
        b2 = sheet.get_cell(2, 2)  # Revenue = 1000
        assert b2 is not None
        assert b2.raw_value == 1000

    def test_workbook_hash_deterministic(self, simple_workbook):
        r1 = WorkbookParser(path=simple_workbook).parse()
        r2 = WorkbookParser(path=simple_workbook).parse()
        assert r1.workbook_hash == r2.workbook_hash

    def test_cell_ids_populated(self, simple_workbook):
        parser = WorkbookParser(path=simple_workbook)
        result = parser.parse()
        sheet = result.sheets[0]
        for cell in sheet.cells.values():
            assert cell.cell_id != ""
            assert cell.cell_hash != ""


class TestMergedCells:
    """Test merged cell handling."""

    def test_merge_regions_detected(self, merged_cells_workbook):
        result = WorkbookParser(path=merged_cells_workbook).parse()
        sheet = result.sheets[0]
        assert len(sheet.merged_regions) >= 2

    def test_merge_master_annotated(self, merged_cells_workbook):
        result = WorkbookParser(path=merged_cells_workbook).parse()
        sheet = result.sheets[0]
        a1 = sheet.get_cell(1, 1)  # Master of A1:D1 merge
        assert a1 is not None
        assert a1.is_merged_master is True
        assert a1.merge_col_extent == 4

    def test_merge_slave_annotated(self, merged_cells_workbook):
        result = WorkbookParser(path=merged_cells_workbook).parse()
        sheet = result.sheets[0]
        b1 = sheet.get_cell(1, 2)  # Slave in A1:D1 merge
        assert b1 is not None
        assert b1.is_merged_slave is True


class TestFormulas:
    """Test formula extraction and cross-sheet references."""

    def test_cross_sheet_formulas(self, formula_workbook):
        result = WorkbookParser(path=formula_workbook).parse()
        assert result.total_sheets == 3

        calcs = result.sheets[1]
        b1 = calcs.get_cell(1, 2)
        assert b1 is not None
        assert b1.formula is not None
        assert "Inputs" in b1.formula

    def test_dependency_graph_built(self, formula_workbook):
        result = WorkbookParser(path=formula_workbook).parse()
        assert len(result.dependency_graph.edges) > 0

    def test_named_ranges_extracted(self, formula_workbook):
        result = WorkbookParser(path=formula_workbook).parse()
        names = {nr.name for nr in result.named_ranges}
        assert "Price" in names
        assert "Quantity" in names


class TestTables:
    """Test Excel ListObject table extraction."""

    def test_table_detected(self, table_workbook):
        result = WorkbookParser(path=table_workbook).parse()
        assert len(result.tables) == 1

    def test_table_properties(self, table_workbook):
        result = WorkbookParser(path=table_workbook).parse()
        table = result.tables[0]
        assert table.table_name == "SalesData"
        assert len(table.columns) == 7
        assert table.columns[0].name == "Product"

    def test_table_range(self, table_workbook):
        result = WorkbookParser(path=table_workbook).parse()
        table = result.tables[0]
        assert table.ref_range.to_a1() == "A1:G5"


class TestConditionalFormatting:
    """Test conditional formatting rule extraction."""

    def test_rules_extracted(self, conditional_format_workbook):
        result = WorkbookParser(path=conditional_format_workbook).parse()
        sheet = result.sheets[0]
        assert len(sheet.conditional_format_rules) > 0
        rule = sheet.conditional_format_rules[0]
        assert rule.rule_type == "cellIs"
        assert rule.operator == "greaterThan"


class TestDataValidation:
    """Test data validation extraction."""

    def test_validation_extracted(self, data_validation_workbook):
        result = WorkbookParser(path=data_validation_workbook).parse()
        sheet = result.sheets[0]
        assert len(sheet.data_validations) > 0
        dv = sheet.data_validations[0]
        assert dv.validation_type == "list"


class TestMultiSheet:
    """Test multi-sheet workbooks including hidden sheets."""

    def test_all_sheets_parsed(self, multi_sheet_workbook):
        result = WorkbookParser(path=multi_sheet_workbook).parse()
        assert result.total_sheets == 3

    def test_hidden_sheet_detected(self, multi_sheet_workbook):
        result = WorkbookParser(path=multi_sheet_workbook).parse()
        hidden = [s for s in result.sheets if s.properties.is_hidden]
        assert len(hidden) == 1
        assert hidden[0].sheet_name == "Hidden"


class TestHiddenRowsCols:
    """Test hidden row and column detection."""

    def test_hidden_row_detected(self, hidden_rows_cols_workbook):
        result = WorkbookParser(path=hidden_rows_cols_workbook).parse()
        sheet = result.sheets[0]
        assert 3 in sheet.hidden_rows

    def test_hidden_col_detected(self, hidden_rows_cols_workbook):
        result = WorkbookParser(path=hidden_rows_cols_workbook).parse()
        sheet = result.sheets[0]
        assert 2 in sheet.hidden_cols  # Column B = col 2


class TestComments:
    """Test cell comment extraction."""

    def test_comments_extracted(self, comment_workbook):
        result = WorkbookParser(path=comment_workbook).parse()
        sheet = result.sheets[0]
        a1 = sheet.get_cell(1, 1)
        assert a1 is not None
        assert a1.comment_text == "Total annual revenue"
        assert a1.comment_author == "Analyst"


class TestSparseSheet:
    """Test handling of large sparse sheets."""

    def test_sparse_cells_extracted(self, large_sparse_workbook):
        result = WorkbookParser(path=large_sparse_workbook).parse()
        sheet = result.sheets[0]
        assert sheet.cell_count() == 4  # A1, B1, Z100, CV1000
        assert sheet.get_cell(1000, 100) is not None

    def test_used_range_spans_sparse(self, large_sparse_workbook):
        result = WorkbookParser(path=large_sparse_workbook).parse()
        sheet = result.sheets[0]
        assert sheet.used_range is not None
        assert sheet.used_range.bottom_right.row == 1000


class TestFreezePane:
    """Test freeze pane detection."""

    def test_freeze_pane_extracted(self, freeze_panes_workbook):
        result = WorkbookParser(path=freeze_panes_workbook).parse()
        sheet = result.sheets[0]
        assert sheet.properties.freeze_pane == "A2"


class TestStyledWorkbook:
    """Test rich formatting extraction."""

    def test_font_color_extracted(self, styled_workbook):
        result = WorkbookParser(path=styled_workbook).parse()
        sheet = result.sheets[0]
        a1 = sheet.get_cell(1, 1)
        assert a1 is not None
        assert a1.style is not None
        assert a1.style.font.bold is True

    def test_fill_extracted(self, styled_workbook):
        result = WorkbookParser(path=styled_workbook).parse()
        sheet = result.sheets[0]
        a1 = sheet.get_cell(1, 1)
        assert a1.style.fill is not None

    def test_border_extracted(self, styled_workbook):
        result = WorkbookParser(path=styled_workbook).parse()
        sheet = result.sheets[0]
        a1 = sheet.get_cell(1, 1)
        assert a1.style.border is not None


class TestWideSheet:
    """Test wide sheets with many columns."""

    def test_100_columns_parsed(self, wide_workbook):
        result = WorkbookParser(path=wide_workbook).parse()
        sheet = result.sheets[0]
        # Should have 100 header cells + 100*4 data cells = 500
        assert sheet.cell_count() == 500
        cell_100 = sheet.get_cell(1, 100)
        assert cell_100 is not None
        assert cell_100.raw_value == "Col100"


class TestHyperlinks:
    """Test hyperlink extraction."""

    def test_hyperlink_extracted(self, hyperlink_workbook):
        result = WorkbookParser(path=hyperlink_workbook).parse()
        sheet = result.sheets[0]
        a1 = sheet.get_cell(1, 1)
        assert a1 is not None
        assert a1.hyperlink == "https://www.google.com"
