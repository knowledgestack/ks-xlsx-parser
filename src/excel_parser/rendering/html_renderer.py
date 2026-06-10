"""
HTML renderer for sheet blocks.

Produces HTML table representations of blocks with:
- rowspan/colspan for merged cells
- Inline styles for formatting (bold, colors, alignment)
- Number format preservation
- Header row semantics (<thead>/<th>)
- Coordinate annotations as data attributes
"""

from __future__ import annotations

import html
import logging

from excel_parser.analysis.header_detector import find_header_span
from excel_parser.models.block import BlockDTO
from excel_parser.models.cell import CellStyle
from excel_parser.models.common import BlockType, col_number_to_letter
from excel_parser.models.sheet import SheetDTO

logger = logging.getLogger(__name__)


class HtmlRenderer:
    """
    Renders a block of cells as an HTML table.

    Handles merged cells (rowspan/colspan), style preservation,
    and header detection. Outputs self-contained HTML fragments
    suitable for embedding in RAG retrieval results.
    """

    def __init__(self, sheet: SheetDTO):
        self._sheet = sheet

    def render_block(self, block: BlockDTO) -> str:
        """
        Render a block as an HTML table.

        The header row is located by :func:`find_header_span` (not assumed to be
        row 1), so non-bold headers and headers below a title row are detected.
        Title rows above the header become a ``<caption>``.
        """
        rng = block.cell_range
        header_row = self._header_row(block)
        rows = list(range(rng.top_left.row, rng.bottom_right.row + 1))
        if header_row is not None:
            title_rows = [r for r in rows if r < header_row]
            data_rows = [r for r in rows if r > header_row]
        else:
            title_rows = []
            data_rows = rows
        return self._render_table(
            block, header_row, title_rows, data_rows, rng.to_a1()
        )

    def render_window(
        self,
        block: BlockDTO,
        data_start: int,
        data_end: int,
        banner: str | None = None,
        include_title: bool = False,
    ) -> str:
        """Render one row-window of a block, repeating the header (see TextRenderer)."""
        header_row = self._header_row(block)
        title_rows = (
            list(range(block.cell_range.top_left.row, header_row))
            if include_title and header_row is not None
            else []
        )
        data_rows = list(range(data_start, data_end + 1))
        rng = block.cell_range
        slice_label = (
            f"{col_number_to_letter(rng.top_left.col)}{data_start}:"
            f"{col_number_to_letter(rng.bottom_right.col)}{data_end}"
        )
        return self._render_table(
            block, header_row, title_rows, data_rows, slice_label, banner=banner
        )

    def _header_row(self, block: BlockDTO) -> int | None:
        if block.block_type not in (BlockType.TABLE, BlockType.ASSUMPTIONS_TABLE):
            return None
        span = find_header_span(self._sheet, block.cell_range)
        return span.top if span else None

    def _render_table(
        self,
        block: BlockDTO,
        header_row: int | None,
        title_rows: list[int],
        data_rows: list[int],
        range_label: str,
        banner: str | None = None,
    ) -> str:
        rng = block.cell_range
        cols = range(rng.top_left.col, rng.bottom_right.col + 1)

        # Track which cells are covered by a merge and should be skipped
        skip_cells: set[tuple[int, int]] = set()
        for merge in self._sheet.merged_regions:
            mr = merge.range
            if not self._overlaps(rng, mr):
                continue
            master = merge.master
            for r in range(mr.top_left.row, mr.bottom_right.row + 1):
                for c in range(mr.top_left.col, mr.bottom_right.col + 1):
                    if (r, c) != (master.row, master.col):
                        skip_cells.add((r, c))

        sheet_hidden_attr = (
            ' data-sheet-hidden="true"' if self._sheet.properties.is_hidden else ""
        )
        parts: list[str] = []
        if banner:
            parts.append(f"<p class=\"partial-table-note\">{html.escape(banner)}</p>")
        parts.append(
            f'<table data-sheet="{html.escape(block.sheet_name)}" '
            f'data-range="{range_label}" '
            f'data-block-type="{block.block_type.value}"{sheet_hidden_attr}>'
        )

        # Title/caption rows above the header are emitted as a <caption>.
        title_text = []
        for row in title_rows:
            vals = []
            for col in cols:
                cell = self._sheet.get_cell(row, col)
                if cell and cell.display_value is not None:
                    vals.append(str(cell.display_value))
                elif cell and cell.raw_value is not None:
                    vals.append(str(cell.raw_value))
            if vals:
                title_text.append(" ".join(vals))
        if title_text:
            parts.append(f"<caption>{html.escape(' / '.join(title_text))}</caption>")

        def _render_row(row: int, tag: str) -> None:
            row_hidden = row in self._sheet.hidden_rows
            parts.append('<tr data-hidden="true">' if row_hidden else "<tr>")
            for col in cols:
                if (row, col) in skip_cells:
                    continue
                cell = self._sheet.get_cell(row, col)
                cell_ref = f"{col_number_to_letter(col)}{row}"
                rowspan = colspan = 1
                if cell and cell.is_merged_master:
                    rowspan = cell.merge_extent or 1
                    colspan = cell.merge_col_extent or 1
                attrs = [f'data-ref="{cell_ref}"']
                if row_hidden or col in self._sheet.hidden_cols:
                    attrs.append('data-hidden="true"')
                if rowspan > 1:
                    attrs.append(f'rowspan="{rowspan}"')
                if colspan > 1:
                    attrs.append(f'colspan="{colspan}"')
                style = self._cell_style_to_css(cell.style if cell else None)
                if style:
                    attrs.append(f'style="{style}"')
                content = ""
                if cell and cell.display_value is not None:
                    content = html.escape(str(cell.display_value))
                elif cell and cell.raw_value is not None:
                    content = html.escape(str(cell.raw_value))
                parts.append(f"<{tag} {' '.join(attrs)}>{content}</{tag}>")
            parts.append("</tr>")

        if header_row is not None:
            parts.append("<thead>")
            _render_row(header_row, "th")
            parts.append("</thead>")
        if data_rows:
            if header_row is not None:
                parts.append("<tbody>")
            for row in data_rows:
                _render_row(row, "td")
            if header_row is not None:
                parts.append("</tbody>")

        parts.append("</table>")
        return "\n".join(parts)

    def _cell_style_to_css(self, style: CellStyle | None) -> str:
        """Convert a CellStyle to inline CSS."""
        if not style:
            return ""

        css_parts: list[str] = []

        if style.font:
            if style.font.bold:
                css_parts.append("font-weight:bold")
            if style.font.italic:
                css_parts.append("font-style:italic")
            if style.font.color and not style.font.color.startswith(("theme:", "indexed:")):
                css_parts.append(f"color:#{style.font.color}")
            if style.font.size:
                css_parts.append(f"font-size:{style.font.size}pt")

        if style.fill and style.fill.fg_color:
            color = style.fill.fg_color
            if not color.startswith(("theme:", "indexed:")):
                css_parts.append(f"background-color:#{color}")

        if style.alignment:
            if style.alignment.horizontal:
                css_parts.append(f"text-align:{style.alignment.horizontal}")
            if style.alignment.vertical:
                css_parts.append(f"vertical-align:{style.alignment.vertical}")

        return ";".join(css_parts)

    @staticmethod
    def _overlaps(a, b) -> bool:
        """Check if two CellRange objects overlap."""
        return not (
            a.bottom_right.row < b.top_left.row
            or a.top_left.row > b.bottom_right.row
            or a.bottom_right.col < b.top_left.col
            or a.top_left.col > b.bottom_right.col
        )
