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
from ks_xlsx_parser.models.common import BlockType, CellCoord, col_number_to_letter
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

    For *numeric* cells we emit the raw value verbatim — Excel's
    commas, percent signs, trailing zeros, and currency symbols all
    defeat substring search. **When the cell carries a number-format
    that meaningfully changes the displayed string (e.g. 0.06 → "6%",
    1272 → "1,272.00", 46022 → "2025-12-31"), we additionally append
    the formatted form in `[brackets]`** so substring match hits
    either the raw or the displayed shape — the question may quote
    either, and `answer.xlsx` may use the display form even though
    `input.xlsx` keeps the raw.

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
        raw_str = _format_number_for_retrieval(raw)
        # If a meaningful number-format produced a different display
        # string, emit both forms. Skip when the displayed form is
        # identical to the raw (no information added) or trivially
        # convertible (just trailing zeros), to keep render_text terse.
        disp = (cell.display_value or "").strip()
        if disp and disp != raw_str and not _is_trivial_format_diff(raw_str, disp):
            return f"{raw_str} [{disp}]"
        return raw_str

    if cell.display_value is not None:
        return str(cell.display_value)
    if raw is not None:
        return str(raw)
    return ""


def _is_trivial_format_diff(raw_str: str, display_str: str) -> bool:
    """True if `display_str` adds no retrieval value over `raw_str`.

    Trivial: ``"1272"`` → ``"1272.0"`` / ``"1272.00"`` (trailing zeros
    only, no other formatting change). The displayed form contributes
    no new tokens substring-search could hit.

    NOT trivial: ``"1272"`` → ``"1,272.00"`` (thousands separator), or
    ``"0.06"`` → ``"6%"``, or ``"1272"`` → ``"$1,272"``. Each of these
    surfaces a distinct token a user might quote.
    """
    if raw_str == display_str:
        return True
    # Trim trailing zeros after a decimal point on the displayed form.
    # If what remains equals the raw, the only difference was insignificant.
    if "." in display_str:
        head, tail = display_str.split(".", 1)
        if tail.rstrip("0") == "" and head == raw_str:
            return True
    return False


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
            r1    | Product  | Q1      | Q2     | Q3      |
            r2    | Widget A | 100     | 150    | 200     |
            ...

        Per-row `r<N>` prefix carries the sheet row number so a
        downstream LLM consumer can compute cell coordinates
        deterministically (block header gives the A1 range; per-row
        anchors close the gap to (row, col)). The prefix width is
        sized to the largest row number in the block.
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

        # Row-anchor width — `r<N>` plus padding. Sized once per block
        # so all rows align under a constant-width column.
        row_anchor_width = max(len(f"r{r}") for r in rows)
        row_anchor_pad = " " * row_anchor_width  # blank slot for header / separator

        # Build slave→master lookup for merged regions on this sheet.
        # Slave cells (everything in a merged range except the master)
        # render the master's value with a `←` propagation marker so
        # the chunk's text contains the visible value at every position
        # it appears in Excel, not just the top-left of the region.
        merged_master: dict[tuple[int, int], CellCoord] = {}
        for region in self._sheet.merged_regions:
            mr = region.range
            master = region.master
            for r in range(mr.top_left.row, mr.bottom_right.row + 1):
                for c in range(mr.top_left.col, mr.bottom_right.col + 1):
                    if r == master.row and c == master.col:
                        continue
                    merged_master[(r, c)] = master

        def _value_for(row: int, col: int) -> tuple[str, bool]:
            """Return (rendered string, is_propagated_from_master)."""
            cell = self._sheet.get_cell(row, col)
            if cell is not None and not cell.is_merged_slave:
                val = _cell_render_value(cell)
                if cell.formula and not val.startswith("="):
                    val = f"{val} [=]"
                return _flatten_cell_text(val), False
            # Slave: propagate the master's value.
            master = merged_master.get((row, col))
            if master is None:
                return "", False
            master_cell = self._sheet.get_cell(master.row, master.col)
            if master_cell is None:
                return "", False
            mval = _cell_render_value(master_cell)
            if master_cell.formula and not mval.startswith("="):
                mval = f"{mval} [=]"
            return _flatten_cell_text(f"← {mval}"), True

        # Compute column widths using the SAME rendering rules the data
        # rows will use, including the trailing `[=]` formula marker
        # AND the merged-cell `←` propagation marker. Otherwise these
        # inflate a cell past col_width post-hoc and spuriously trigger
        # the long-value fallback below.
        col_widths: dict[int, int] = {}
        for col in cols:
            col_letter = col_number_to_letter(col)
            max_width = len(col_letter)
            for row in rows:
                val, _ = _value_for(row, col)
                max_width = max(max_width, len(val))
            col_widths[col] = min(max_width, 30)  # Cap at 30 for alignment; text may overflow

        # Column header row — leading blank slot matches the row-anchor width.
        col_headers = []
        for col in cols:
            if col in self._sheet.hidden_cols:
                continue
            letter = col_number_to_letter(col)
            col_headers.append(letter.ljust(col_widths[col]))
        lines.append(row_anchor_pad + " | " + " | ".join(col_headers) + " |")
        lines.append(
            row_anchor_pad
            + " |-"
            + "-|-".join("-" * col_widths[c] for c in cols if c not in self._sheet.hidden_cols)
            + "-|"
        )

        # Data rows
        is_first_data = True
        for row in rows:
            if row in self._sheet.hidden_rows:
                continue

            anchor = f"r{row}".ljust(row_anchor_width)

            values = []
            for col in cols:
                if col in self._sheet.hidden_cols:
                    continue
                val, _ = _value_for(row, col)
                # Long-value fallback: only triggers if the rendered string
                # genuinely exceeds the (now consistently-computed) column
                # width — i.e. the column was capped at 30. We still emit
                # the full retrieval value (no truncation) and let the
                # alignment overflow; truncating destroys retrievability.
                values.append(val.ljust(col_widths[col]))

            line = anchor + " | " + " | ".join(values) + " |"
            lines.append(line)

            # Add separator after first row if it looks like a header
            if is_first_data and block.block_type in (
                BlockType.TABLE,
                BlockType.ASSUMPTIONS_TABLE,
            ):
                lines.append(
                    row_anchor_pad
                    + " |-"
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
