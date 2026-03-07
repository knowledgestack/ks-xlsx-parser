"""
Plain text / markdown renderer for sheet blocks.

Produces human-readable text representations of blocks for RAG
retrieval. Includes coordinate headers, aligned columns, and
semantic markers for formulas and key cells.
"""

from __future__ import annotations

import logging

from ..models.block import BlockDTO
from ..models.common import BlockType, col_number_to_letter
from ..models.chart import ChartDTO
from ..models.sheet import SheetDTO

logger = logging.getLogger(__name__)


class TextRenderer:
    """
    Renders blocks as plain text with coordinate context.

    Produces compact, human-readable text suitable for RAG embedding.
    Includes column headers, row labels, and semantic annotations.
    """

    def __init__(self, sheet: SheetDTO):
        self._sheet = sheet

    def render_block(self, block: BlockDTO) -> str:
        """
        Render a block as plain text with coordinate context.

        Format:
            [Sheet1!A1:D10] (table: "SalesData")
            | A        | B       | C      | D       |
            |----------|---------|--------|---------|
            | Product  | Q1      | Q2     | Q3      |
            | Widget A | 100     | 150    | 200     |
            ...
        """
        rng = block.cell_range
        rows = range(rng.top_left.row, rng.bottom_right.row + 1)
        cols = range(rng.top_left.col, rng.bottom_right.col + 1)

        lines: list[str] = []

        # Header with location and type
        type_label = block.block_type.value.replace("_", " ")
        header = f"[{block.sheet_name}!{rng.to_a1()}] ({type_label})"
        if block.table_name:
            header += f' table: "{block.table_name}"'
        lines.append(header)

        # Compute column widths
        col_widths: dict[int, int] = {}
        for col in cols:
            col_letter = col_number_to_letter(col)
            max_width = len(col_letter)
            for row in rows:
                cell = self._sheet.get_cell(row, col)
                if cell:
                    val = cell.display_value or (str(cell.raw_value) if cell.raw_value is not None else "")
                    max_width = max(max_width, len(val))
            col_widths[col] = min(max_width, 30)  # Cap at 30 for alignment; text may overflow

        # Column header row
        col_headers = []
        for col in cols:
            if col in self._sheet.hidden_cols:
                continue
            letter = col_number_to_letter(col)
            col_headers.append(letter.ljust(col_widths[col]))
        lines.append("| " + " | ".join(col_headers) + " |")
        lines.append(
            "|-" + "-|-".join("-" * col_widths[c] for c in cols if c not in self._sheet.hidden_cols) + "-|"
        )

        # Data rows
        is_first_data = True
        for row in rows:
            if row in self._sheet.hidden_rows:
                continue

            values = []
            for col in cols:
                if col in self._sheet.hidden_cols:
                    continue
                cell = self._sheet.get_cell(row, col)
                val = ""
                if cell:
                    if cell.display_value is not None:
                        val = str(cell.display_value)
                    elif cell.raw_value is not None:
                        val = str(cell.raw_value)

                    # Annotate formulas with a marker (unless display already shows the formula)
                    if cell.formula and not val.startswith("="):
                        val = f"{val} [=]"

                # For long numeric values: use scientific notation (preserves precision).
                # Text strings are never truncated.
                if len(val) > col_widths[col]:
                    raw = cell.raw_value
                    if isinstance(raw, (int, float)):
                        val = f"{float(raw):.6e}"
                values.append(val.ljust(col_widths[col]))

            line = "| " + " | ".join(values) + " |"
            lines.append(line)

            # Add separator after first row if it looks like a header
            if is_first_data and block.block_type in (
                BlockType.TABLE,
                BlockType.ASSUMPTIONS_TABLE,
            ):
                lines.append(
                    "|-"
                    + "-|-".join(
                        "-" * col_widths[c]
                        for c in cols
                        if c not in self._sheet.hidden_cols
                    )
                    + "-|"
                )
            is_first_data = False

        return "\n".join(lines)

    @staticmethod
    def render_chart_summary(chart: ChartDTO) -> str:
        """Render a chart as a text summary for RAG."""
        return chart.summary_text or chart.generate_summary()
