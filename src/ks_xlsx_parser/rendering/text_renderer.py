"""
Plain text / markdown renderer for sheet blocks.

Produces human-readable text representations of blocks for RAG
retrieval. Includes coordinate headers, aligned columns, and
semantic markers for formulas and key cells.
"""

from __future__ import annotations

import datetime as _dt
import logging

from ks_xlsx_parser.models.block import BlockDTO
from ks_xlsx_parser.models.chart import ChartDTO
from ks_xlsx_parser.models.common import BlockType, col_number_to_letter
from ks_xlsx_parser.models.sheet import SheetDTO

logger = logging.getLogger(__name__)


def _flatten_cell_text(val: str) -> str:
    """Collapse embedded line breaks so a cell stays on one row of the
    Markdown grid. Excel headers often contain `\\n` to wrap text visually
    (e.g. ``"租金\\n天数"``); rendered into ``| ... |`` rows verbatim they
    rip the grid apart."""
    if "\n" not in val and "\r" not in val:
        return val
    return val.replace("\r\n", " ").replace("\n", " ").replace("\r", " ")


def _format_number_for_retrieval(raw: int | float) -> str:
    """Render a numeric raw value in a retrieval-friendly form.

    Excel's ``display_value`` honours the cell's number-format string,
    which produces ``"1,272.00"`` for 1272 or ``"6%"`` for 0.06. Those
    are great for humans but defeat substring-match retrieval — a user
    asking "what was the value in 2020?" types ``1272``, not ``1,272.00``.

    Rules:
      - Integer-valued floats → ``str(int(v))``  (1272.0 → "1272")
      - Integers → ``str(v)``                    (1272   → "1272")
      - Floats → ``g`` format up to 10 significant digits, trailing
        zeros trimmed. Avoids both sci-notation for ordinary magnitudes
        and trailing ``.0`` noise.
    """
    if isinstance(raw, bool):  # bool is a subclass of int
        return "TRUE" if raw else "FALSE"
    if isinstance(raw, int):
        return str(raw)
    # float
    if raw == int(raw) and abs(raw) < 1e16:
        return str(int(raw))
    return f"{raw:.10g}"


def _cell_render_value(cell) -> str:
    """Pick the string form of `cell` that's best for RAG retrieval.

    For *numeric* cells we ignore the display-formatted string and emit
    the raw value verbatim — Excel's commas, percent signs, trailing
    zeros, and currency symbols all defeat substring search.

    For dates we emit ISO ``YYYY-MM-DD`` (no time component) which is
    both human-readable and matches the date format that openpyxl /
    pandas surface when reading the answer file.

    Strings and everything else fall back to ``display_value``.
    """
    if cell is None:
        return ""
    raw = cell.raw_value

    if isinstance(raw, (_dt.date, _dt.datetime)):
        if isinstance(raw, _dt.datetime):
            if raw.hour == 0 and raw.minute == 0 and raw.second == 0:
                return raw.date().isoformat()
            return raw.isoformat(sep=" ")
        return raw.isoformat()

    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        return _format_number_for_retrieval(raw)

    if cell.display_value is not None:
        return str(cell.display_value)
    if raw is not None:
        return str(raw)
    return ""


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
        Render a block as a plain-text / Markdown table with coordinate
        context.

        The grid is a *standard* Markdown table whose header row holds the
        block's real column names (for ``TABLE`` / ``ASSUMPTIONS_TABLE``
        blocks, mirroring :class:`HtmlRenderer`'s ``<thead>`` behaviour).
        Excel column letters are published on the bracket line as a
        ``cols:`` map rather than occupying the header row, and a leading
        ``row`` gutter carries the Excel row number of every line. Together
        the ``cols:`` map and the row gutter let an agent reconstruct a full
        A1 reference (column ``Amount`` + row ``3`` → ``B3``) without the
        column letters masquerading as the table's headers and defeating
        downstream header detection.

        Hidden rows and columns are *included* (not dropped) and flagged
        ``[hidden]`` — in the gutter for rows, in the ``cols:`` map for
        columns.

        Format::

            [Sheet1!A1:D3] (table) cols: A=Product, B=Q1, C=Q2, D=Q3
            | row | Product  | Q1  | Q2  | Q3  |
            |-----|----------|-----|-----|-----|
            | 2   | Widget A | 100 | 150 | 200 |
        """
        rng = block.cell_range
        rows = list(range(rng.top_left.row, rng.bottom_right.row + 1))
        cols = list(range(rng.top_left.col, rng.bottom_right.col + 1))

        # Mirror HtmlRenderer: for these block types the first row carries the
        # real column names, so it becomes the Markdown header row.
        first_row_is_header = block.block_type in (
            BlockType.TABLE,
            BlockType.ASSUMPTIONS_TABLE,
        )

        lines: list[str] = []

        # --- Bracket line + column-letter map ------------------------------
        type_label = block.block_type.value.replace("_", " ")
        header = f"[{block.sheet_name}!{rng.to_a1()}] ({type_label})"
        if self._sheet.properties.is_hidden:
            header += " [hidden sheet]"
        if block.table_name:
            header += f' table: "{block.table_name}"'

        # Column letters live here (not as a grid row) so the grid header is
        # free to hold real names while an agent can still map name → letter.
        col_descs: list[str] = []
        for col in cols:
            desc = col_number_to_letter(col)
            if first_row_is_header:
                name_cell = self._sheet.get_cell(rng.top_left.row, col)
                name = (
                    _flatten_cell_text(_cell_render_value(name_cell))
                    if name_cell
                    else ""
                )
                if name:
                    desc = f"{desc}={name}"
            if col in self._sheet.hidden_cols:
                desc += " [hidden]"
            col_descs.append(desc)
        header += " cols: " + ", ".join(col_descs)
        lines.append(header)

        # Compute column widths using the SAME rendering rules the data
        # rows will use, including the trailing `[=]` formula marker.
        # Otherwise `[=]` inflates a cell past col_width post-hoc.
        col_widths: dict[int, int] = {}
        for col in cols:
            max_width = len(col_number_to_letter(col))
            for row in rows:
                cell = self._sheet.get_cell(row, col)
                if cell is None:
                    continue
                val = _cell_render_value(cell)
                if cell.formula and not val.startswith("="):
                    val = f"{val} [=]"
                val = _flatten_cell_text(val)
                max_width = max(max_width, len(val))
            col_widths[col] = min(max_width, 30)  # Cap at 30 for alignment; text may overflow

        # Row-number gutter: gives every value a row coordinate so an agent
        # can form a full A1 reference. Hidden rows are flagged here.
        gutter_header = "row"
        gutter: dict[int, str] = {}
        for row in rows:
            label = str(row)
            if row in self._sheet.hidden_rows:
                label += " [hidden]"
            gutter[row] = label
        gutter_width = max([len(gutter_header), *(len(g) for g in gutter.values())])

        def _row(gutter_cell: str, values: list[str]) -> str:
            return "| " + " | ".join([gutter_cell.ljust(gutter_width), *values]) + " |"

        def _sep() -> str:
            return (
                "|-"
                + "-|-".join(["-" * gutter_width, *("-" * col_widths[c] for c in cols)])
                + "-|"
            )

        def _cells(row: int) -> list[str]:
            values = []
            for col in cols:
                cell = self._sheet.get_cell(row, col)
                val = _cell_render_value(cell) if cell else ""
                if cell and cell.formula and not val.startswith("="):
                    val = f"{val} [=]"
                # Markdown rows are single-line; collapse embedded newlines
                # (common in headers like "租金\n天数") so they don't break the grid.
                val = _flatten_cell_text(val)
                # Full retrieval value (no truncation); alignment may overflow.
                values.append(val.ljust(col_widths[col]))
            return values

        # Header row: real first row for tables, else Excel column letters.
        if first_row_is_header:
            lines.append(_row(gutter_header, _cells(rng.top_left.row)))
            lines.append(_sep())
            data_rows = rows[1:]
        else:
            letters = [col_number_to_letter(c).ljust(col_widths[c]) for c in cols]
            lines.append(_row(gutter_header, letters))
            lines.append(_sep())
            data_rows = rows

        # Data rows (hidden rows/cols included; hidden rows flagged in gutter).
        for row in data_rows:
            lines.append(_row(gutter[row], _cells(row)))

        return "\n".join(lines)

    @staticmethod
    def render_chart_summary(chart: ChartDTO) -> str:
        """Render a chart as a text summary for RAG."""
        return chart.summary_text or chart.generate_summary()
