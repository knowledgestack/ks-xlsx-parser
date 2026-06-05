"""
Tests for HTML and text rendering.

Verifies correct rendering of merged cells, formatting,
headers, and coordinate annotations.
"""


from ks_xlsx_parser.chunking.segmenter import LayoutSegmenter
from ks_xlsx_parser.parsers import WorkbookParser
from ks_xlsx_parser.rendering.html_renderer import HtmlRenderer
from ks_xlsx_parser.rendering.text_renderer import TextRenderer


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

    def test_hidden_cells_included_and_flagged(self, hidden_rows_cols_workbook):
        """Hidden rows/columns are emitted in the HTML (flagged
        `data-hidden`) rather than dropped — matching the text renderer."""
        from ks_xlsx_parser.pipeline import parse_workbook

        html = "\n".join(
            c.render_html for c in parse_workbook(str(hidden_rows_cols_workbook)).chunks
        )
        assert 'data-hidden="true"' in html
        assert "R3C1" in html  # content of hidden row 3
        assert "R1C2" in html  # content of hidden column B

    def test_hidden_sheet_flagged_on_table(self, multi_sheet_workbook):
        """Tables from a hidden worksheet carry `data-sheet-hidden`."""
        from ks_xlsx_parser.pipeline import parse_workbook

        chunks = parse_workbook(str(multi_sheet_workbook)).chunks
        hidden_html = [c.render_html for c in chunks if c.sheet_name == "Hidden"]
        assert hidden_html
        assert all('data-sheet-hidden="true"' in h for h in hidden_html)


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
        from ks_xlsx_parser.models.block import BlockDTO
        from ks_xlsx_parser.models.cell import CellDTO
        from ks_xlsx_parser.models.common import BlockType, CellCoord, CellRange
        from ks_xlsx_parser.models.sheet import SheetDTO

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

    def test_table_header_uses_real_names_not_column_letters(self, table_workbook):
        """Option A: the grid header holds the table's real column names while
        Excel column letters move to a `cols:` map on the bracket line, with a
        leading `row` gutter — so downstream 'find the real header' logic sees
        'Product', not 'A'."""
        from ks_xlsx_parser.pipeline import parse_workbook

        chunks = parse_workbook(str(table_workbook)).chunks
        text = next(c.render_text for c in chunks if "Product" in c.render_text)

        grid_lines = [ln for ln in text.splitlines() if ln.startswith("|")]
        header_cells = [c.strip() for c in grid_lines[0].split("|")[1:-1]]
        # Gutter first, then the *real* header names — not bare column letters.
        assert header_cells[0] == "row"
        assert header_cells[1] == "Product"
        # Column letters are published as a name→letter map instead.
        assert "cols: A=Product" in text

    def test_hidden_rows_and_cols_are_extracted_and_flagged(
        self, hidden_rows_cols_workbook
    ):
        """Hidden rows/columns are rendered (not dropped) and flagged
        `[hidden]` — in the `cols:` map for columns, the gutter for rows."""
        from ks_xlsx_parser.pipeline import parse_workbook

        text = "\n".join(
            c.render_text for c in parse_workbook(str(hidden_rows_cols_workbook)).chunks
        )
        assert "R1C2" in text  # a cell in hidden column B
        assert "R3C1" in text  # a cell in hidden row 3
        assert "[hidden]" in text

    def test_hidden_rows_and_cols_recorded_in_chunk_metadata(
        self, hidden_rows_cols_workbook
    ):
        """Hidden rows/columns are stored as structured chunk metadata
        (scoped to the chunk's range), not just inline text markers."""
        from ks_xlsx_parser.pipeline import parse_workbook

        chunks = parse_workbook(str(hidden_rows_cols_workbook)).chunks
        hidden_rows = {r for c in chunks for r in c.metadata.get("hidden_rows", [])}
        hidden_cols = {col for c in chunks for col in c.metadata.get("hidden_cols", [])}
        assert 3 in hidden_rows  # row 3 is hidden
        assert "B" in hidden_cols  # column B is hidden

    def test_hidden_sheet_marked_in_metadata_and_render(self, multi_sheet_workbook):
        """A hidden worksheet is still parsed and chunked, and every chunk
        from it carries `sheet_hidden` metadata and a `[hidden sheet]` render
        marker; visible-sheet chunks carry neither."""
        from ks_xlsx_parser.pipeline import parse_workbook

        chunks = parse_workbook(str(multi_sheet_workbook)).chunks
        hidden = [c for c in chunks if c.sheet_name == "Hidden"]
        visible = [c for c in chunks if c.sheet_name != "Hidden"]

        assert hidden, "hidden sheet should still be parsed and chunked"
        assert all(c.metadata.get("sheet_hidden") is True for c in hidden)
        assert all("[hidden sheet]" in c.render_text for c in hidden)
        assert all("sheet_hidden" not in c.metadata for c in visible)
        assert all("[hidden sheet]" not in c.render_text for c in visible)
