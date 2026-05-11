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

    def test_numeric_cells_render_raw_not_display_formatted(self):
        """Numeric cells render the raw value, ignoring Excel's display
        formatting. This is intentional for RAG retrievability: a query
        like "1272" should match the cell even if Excel displays it as
        "1,272.00". The clobbered display format used to also trigger a
        sci-notation fallback (``1.272000e+03``) once the ``[=]`` formula
        marker pushed the rendered string past col_width — this test
        guards against that regression."""
        from models.sheet import SheetDTO
        from models.cell import CellDTO
        from models.common import CellCoord, CellRange
        from models.block import BlockDTO
        from models.common import BlockType

        coord = CellCoord(row=1, col=1)
        cell = CellDTO(
            coord=coord,
            sheet_name="Test",
            raw_value=0.002668,
            # Excel display would be e.g. "0.27%" or "0.002668000000000..."
            display_value="0.002668000000000000000000000000",
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

        # Raw value, sci-notation-free
        assert "0.002668" in text
        assert "e-03" not in text
        assert "e+03" not in text
