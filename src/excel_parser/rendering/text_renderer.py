"""
Plain text / markdown renderer for sheet blocks.

Produces human-readable text representations of blocks for RAG
retrieval. Includes coordinate headers, aligned columns, and
semantic markers for formulas and key cells.
"""

from __future__ import annotations

import datetime as _dt
import logging

from excel_parser.analysis.header_detector import find_header_span
from excel_parser.models.block import BlockDTO
from excel_parser.models.chart import ChartDTO
from excel_parser.models.common import BlockType, CellCoord, col_number_to_letter
from excel_parser.models.sheet import SheetDTO

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
        block's real column names — located by :func:`find_header_span`, so a
        non-bold header or a header sitting below a title/caption row is found
        rather than blindly assuming row 1. Rows above the header render as
        ``title:`` caption lines. Excel column letters are published on the
        bracket line as a ``cols:`` map rather than occupying the header row,
        and a leading ``row`` gutter carries the Excel row number of every line.
        Together the ``cols:`` map and the row gutter let an agent reconstruct a
        full A1 reference (column ``Amount`` + row ``3`` → ``B3``).

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
        header_row = self._header_row(block)
        if header_row is not None:
            title_rows = list(range(rng.top_left.row, header_row))
            data_rows = list(range(header_row + 1, rng.bottom_right.row + 1))
        else:
            title_rows = []
            data_rows = list(range(rng.top_left.row, rng.bottom_right.row + 1))
        return self._render_grid(
            block, header_row, title_rows, data_rows,
            f"{block.sheet_name}!{rng.to_a1()}",
        )

    def render_window(
        self,
        block: BlockDTO,
        data_start: int,
        data_end: int,
        banner: str | None = None,
        include_title: bool = False,
    ) -> str:
        """
        Render one row-window of a (large) block, repeating the block's header
        so the window is self-contained for retrieval.

        ``data_start``/``data_end`` are inclusive Excel rows of the data slice;
        the header row (if any) is rendered above them regardless of where the
        slice begins. ``banner`` is an optional partial-table notice prepended
        to the output (see :class:`ChunkBuilder`). ``include_title`` renders the
        table's title/caption rows (above the header) — set on the first part so
        the title is not lost when a titled table is windowed.
        """
        header_row = self._header_row(block)
        title_rows = (
            list(range(block.cell_range.top_left.row, header_row))
            if include_title and header_row is not None
            else []
        )
        data_rows = list(range(data_start, data_end + 1))
        slice_label = (
            f"{block.sheet_name}!"
            f"{CellCoord(row=data_start, col=block.cell_range.top_left.col).to_a1()}:"
            f"{CellCoord(row=data_end, col=block.cell_range.bottom_right.col).to_a1()}"
        )
        return self._render_grid(
            block, header_row, title_rows, data_rows, slice_label, banner=banner
        )

    def _header_row(self, block: BlockDTO) -> int | None:
        """Excel row of the block's column header, or None if it has none."""
        if block.block_type not in (BlockType.TABLE, BlockType.ASSUMPTIONS_TABLE):
            return None
        span = find_header_span(self._sheet, block.cell_range)
        return span.top if span else None

    def _render_grid(
        self,
        block: BlockDTO,
        header_row: int | None,
        title_rows: list[int],
        data_rows: list[int],
        range_label: str,
        banner: str | None = None,
    ) -> str:
        cols = list(range(block.cell_range.top_left.col, block.cell_range.bottom_right.col + 1))
        width_rows = ([header_row] if header_row is not None else []) + title_rows + data_rows

        lines: list[str] = []
        if banner:
            lines.append(banner)

        # --- Bracket line + column-letter map ------------------------------
        type_label = block.block_type.value.replace("_", " ")
        header = f"[{range_label}] ({type_label})"
        if self._sheet.properties.is_hidden:
            header += " [hidden sheet]"
        if block.table_name:
            header += f' table: "{block.table_name}"'

        # Column letters live here (not as a grid row) so the grid header is
        # free to hold real names while an agent can still map name → letter.
        col_descs: list[str] = []
        for col in cols:
            desc = col_number_to_letter(col)
            if header_row is not None:
                name_cell = self._sheet.get_cell(header_row, col)
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

        # Title/caption rows above the header (rendered outside the grid).
        for row in title_rows:
            vals = [
                _flatten_cell_text(_cell_render_value(self._sheet.get_cell(row, c)))
                for c in cols
            ]
            text = " ".join(v for v in vals if v).strip()
            if text:
                lines.append(f"title (row {row}): {text}")

        # Compute column widths using the SAME rendering rules the data
        # rows will use, including the trailing `[=]` formula marker.
        col_widths: dict[int, int] = {}
        for col in cols:
            max_width = len(col_number_to_letter(col))
            for row in width_rows:
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
        for row in width_rows:
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
                values.append(val.ljust(col_widths[col]))
            return values

        # Header row: the real header for tables, else Excel column letters.
        if header_row is not None:
            lines.append(_row(gutter_header, _cells(header_row)))
        else:
            letters = [col_number_to_letter(c).ljust(col_widths[c]) for c in cols]
            lines.append(_row(gutter_header, letters))
        lines.append(_sep())

        # Data rows (hidden rows/cols included; hidden rows flagged in gutter).
        for row in data_rows:
            lines.append(_row(gutter[row], _cells(row)))

        return "\n".join(lines)

    @staticmethod
    def render_chart_summary(chart: ChartDTO) -> str:
        """Render a chart as a text summary for RAG."""
        return chart.summary_text or chart.generate_summary()
