"""
Tests for HTML and text rendering.

Verifies correct rendering of merged cells, formatting,
headers, and coordinate annotations.
"""

import pytest

from chunking.segmenter import LayoutSegmenter
from parsers import WorkbookParser
from rendering.html_renderer import HtmlRenderer
from rendering.text_renderer import TextRenderer


class TestHtmlRendering:
    """Test HTML table rendering."""

    def test_basic_html_output(self, simple_workbook):
        result = WorkbookParser(path=simple_workbook).parse()
        sheet = result.sheets[0]
        segmenter = LayoutSegmenter(sheet)
        blocks = segmenter.segment()

        renderer = HtmlRenderer(sheet)
        html = renderer.render_block(blocks[0])

        assert "<table" in html
        assert "</table>" in html
        assert "data-sheet=" in html

    def test_merged_cell_rowspan_colspan(self, merged_cells_workbook):
        result = WorkbookParser(path=merged_cells_workbook).parse()
        sheet = result.sheets[0]
        segmenter = LayoutSegmenter(sheet)
        blocks = segmenter.segment()

        renderer = HtmlRenderer(sheet)
        html = renderer.render_block(blocks[0])

        assert 'colspan="4"' in html or "colspan" in html

    def test_bold_rendered_as_style(self, styled_workbook):
        result = WorkbookParser(path=styled_workbook).parse()
        sheet = result.sheets[0]
        segmenter = LayoutSegmenter(sheet)
        blocks = segmenter.segment()

        renderer = HtmlRenderer(sheet)
        html = renderer.render_block(blocks[0])

        assert "font-weight:bold" in html

    def test_data_ref_attributes(self, simple_workbook):
        result = WorkbookParser(path=simple_workbook).parse()
        sheet = result.sheets[0]
        segmenter = LayoutSegmenter(sheet)
        blocks = segmenter.segment()

        renderer = HtmlRenderer(sheet)
        html = renderer.render_block(blocks[0])

        assert 'data-ref="A1"' in html


class TestTextRendering:
    """Test plain text / markdown rendering."""

    def test_basic_text_output(self, simple_workbook):
        result = WorkbookParser(path=simple_workbook).parse()
        sheet = result.sheets[0]
        segmenter = LayoutSegmenter(sheet)
        blocks = segmenter.segment()

        renderer = TextRenderer(sheet)
        text = renderer.render_block(blocks[0])

        assert "Sheet1" in text
        assert "|" in text  # Table-like format

    def test_formula_annotation(self, simple_workbook):
        result = WorkbookParser(path=simple_workbook).parse()
        sheet = result.sheets[0]
        segmenter = LayoutSegmenter(sheet)
        blocks = segmenter.segment()

        renderer = TextRenderer(sheet)
        text = renderer.render_block(blocks[0])

        # Formula cells display the formula string (prefixed with =)
        assert "=" in text
        # The formula display value should appear in the rendered text
        assert "B2" in text or "b2" in text.lower()

    def test_text_includes_range(self, simple_workbook):
        result = WorkbookParser(path=simple_workbook).parse()
        sheet = result.sheets[0]
        segmenter = LayoutSegmenter(sheet)
        blocks = segmenter.segment()

        renderer = TextRenderer(sheet)
        text = renderer.render_block(blocks[0])

        # Should include the A1-style range
        assert "!" in text  # Sheet1!range format

    def test_numeric_cells_use_scientific_notation_not_truncation(self):
        """Long numeric values use scientific notation instead of truncating with ..."""
        from models.sheet import SheetDTO
        from models.cell import CellDTO
        from models.common import CellCoord, CellRange
        from models.block import BlockDTO
        from models.common import BlockType

        # Create a sheet with a numeric cell whose display_value would exceed column width.
        # Column width is min(max_cell_len, 30), so we need a number that formats to >30 chars.
        coord = CellCoord(row=1, col=1)
        cell = CellDTO(
            coord=coord,
            sheet_name="Test",
            raw_value=0.002668,
            display_value="0.002668000000000000000000000000",  # 32 chars - would truncate
        )
        sheet = SheetDTO(
            sheet_name="Test",
            sheet_index=0,
            cells={"1,1": cell},
            hidden_rows=set(),
            hidden_cols=set(),
        )

        rng = CellRange(
            top_left=CellCoord(row=1, col=1),
            bottom_right=CellCoord(row=1, col=1),
        )
        block = BlockDTO(
            sheet_name="Test",
            block_index=0,
            cell_range=rng,
            block_type=BlockType.TABLE,
        )

        renderer = TextRenderer(sheet)
        text = renderer.render_block(block)

        # Number should appear in scientific notation (full precision) rather than truncated with …
        assert "2.668000e-03" in text
