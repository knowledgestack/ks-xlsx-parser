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

from models.block import BlockDTO
from models.cell import CellDTO, CellStyle
from models.common import BlockType, CellCoord, col_number_to_letter
from models.sheet import SheetDTO

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

        Args:
            block: The block to render.

        Returns:
            HTML string with a <table> element.
        """
        rng = block.cell_range
        rows = range(rng.top_left.row, rng.bottom_right.row + 1)
        cols = range(rng.top_left.col, rng.bottom_right.col + 1)

        # Track which cells are covered by a merge and should be skipped
        skip_cells: set[tuple[int, int]] = set()
        for merge in self._sheet.merged_regions:
            mr = merge.range
            # Only consider merges that overlap this block
            if not self._overlaps(rng, mr):
                continue
            master = merge.master
            for r in range(mr.top_left.row, mr.bottom_right.row + 1):
                for c in range(mr.top_left.col, mr.bottom_right.col + 1):
                    if (r, c) != (master.row, master.col):
                        skip_cells.add((r, c))

        # Determine if first row should be treated as header
        is_first_row_header = block.block_type in (
            BlockType.TABLE,
            BlockType.ASSUMPTIONS_TABLE,
        )

        parts: list[str] = []
        parts.append(
            f'<table data-sheet="{html.escape(block.sheet_name)}" '
            f'data-range="{rng.to_a1()}" '
            f'data-block-type="{block.block_type.value}">'
        )

        for row_idx, row in enumerate(rows):
            if row in self._sheet.hidden_rows:
                continue

            is_header_row = row_idx == 0 and is_first_row_header
            tag = "th" if is_header_row else "td"
            wrapper = "thead" if is_header_row else "tbody"

            if row_idx == 0 and is_header_row:
                parts.append("<thead>")
            elif row_idx == 1 and is_first_row_header:
                parts.append("<tbody>")

            parts.append("<tr>")
            for col in cols:
                if col in self._sheet.hidden_cols:
                    continue
                if (row, col) in skip_cells:
                    continue

                cell = self._sheet.get_cell(row, col)
                cell_ref = f"{col_number_to_letter(col)}{row}"

                # Check if this is a merge master
                rowspan = 1
                colspan = 1
                if cell and cell.is_merged_master:
                    rowspan = cell.merge_extent or 1
                    colspan = cell.merge_col_extent or 1

                # Build cell attributes
                attrs = [f'data-ref="{cell_ref}"']
                if rowspan > 1:
                    attrs.append(f'rowspan="{rowspan}"')
                if colspan > 1:
                    attrs.append(f'colspan="{colspan}"')

                # Build inline style
                style = self._cell_style_to_css(cell.style if cell else None)
                if style:
                    attrs.append(f'style="{style}"')

                # Cell content
                content = ""
                if cell and cell.display_value is not None:
                    content = html.escape(str(cell.display_value))
                elif cell and cell.raw_value is not None:
                    content = html.escape(str(cell.raw_value))

                attrs_str = " ".join(attrs)
                parts.append(f"<{tag} {attrs_str}>{content}</{tag}>")

            parts.append("</tr>")

            if row_idx == 0 and is_first_row_header:
                parts.append("</thead>")

        # Close tbody if we opened it
        if is_first_row_header and len(list(rows)) > 1:
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
