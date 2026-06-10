"""
Shared header-row detection.

A single source of truth for "which row(s) of this block are the column header",
used by the segmenter (block classification + table stitching) and both renderers
(so a whole-table chunk and a windowed part render the header identically).

The detector does NOT assume the header is the block's first row. It recognises a
header by two independent signals, so it works for headers that are not bold and
for headers that sit below a title/caption row:

  - *styled*    — a majority of the row's non-empty cells are bold or filled.
  - *contrast*  — at least two columns where the row's cell is a text label while
                  the cells below it are predominantly numeric/date (data).

Under-detection is deliberately preferred over over-detection: a row is only a
header if it clearly labels data beneath it, which keeps the segmenter from
gluing unrelated regions together.
"""
from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass

from excel_parser.models.common import CellRange
from excel_parser.models.sheet import SheetDTO

# How many rows from the top of a block to scan for a header (lets a header sit
# below a title/caption without being missed, while bounding the work).
DEFAULT_MAX_SCAN = 4
# Minimum number of label-over-data columns for the "contrast" signal to fire.
_MIN_CONTRAST_COLS = 2
# Maximum number of rows a column header band may span. Bounds the downward
# extension (find_header_span) so a runaway scan can never swallow data rows.
MAX_HEADER_ROWS = 6


@dataclass(frozen=True)
class HeaderSpan:
    """Inclusive Excel row range of a block's column header."""

    top: int
    bottom: int


def _is_data_value(cell) -> bool:
    """True if the cell holds a number or date (the stuff that sits under labels)."""
    if cell is None:
        return False
    raw = cell.raw_value
    if isinstance(raw, bool):  # bool is a subclass of int; treat as data anyway
        return True
    return isinstance(raw, (int, float, _dt.date, _dt.datetime))


def _is_styled(cell) -> bool:
    if cell is None or cell.style is None:
        return False
    font = cell.style.font
    if font and font.bold:
        return True
    fill = cell.style.fill
    return bool(fill and fill.fg_color)


def _row_nonempty(sheet: SheetDTO, row: int, c0: int, c1: int) -> list:
    cells = []
    for col in range(c0, c1 + 1):
        cell = sheet.get_cell(row, col)
        if cell and not cell.is_empty:
            cells.append(cell)
    return cells


def _is_data_row(sheet: SheetDTO, row: int, c0: int, c1: int) -> bool:
    """True if the row's non-empty cells are a majority numbers/dates (i.e. data)."""
    cells = _row_nonempty(sheet, row, c0, c1)
    if not cells:
        return False
    return sum(1 for c in cells if _is_data_value(c)) * 2 >= len(cells)


def _is_styled_row(sheet: SheetDTO, row: int, c0: int, c1: int) -> bool:
    """True if a majority of the row's non-empty cells are bold or filled."""
    cells = _row_nonempty(sheet, row, c0, c1)
    if not cells:
        return False
    return sum(1 for c in cells if _is_styled(c)) * 2 >= len(cells)


def _is_header_row(sheet: SheetDTO, row: int, c0: int, c1: int, bot: int) -> bool:
    """Whether ``row`` looks like a multi-column header labelling the data below."""
    nonempty = _row_nonempty(sheet, row, c0, c1)
    if len(nonempty) < 2:
        return False  # a single labelled cell is a title/section label, not a header

    styled = sum(1 for c in nonempty if _is_styled(c))
    if styled * 2 >= len(nonempty):
        return True

    # Contrast: count columns where this row holds a text label and the cells
    # below it are a majority of data values (numbers/dates).
    label_over_data = 0
    for col in range(c0, c1 + 1):
        head = sheet.get_cell(row, col)
        if head is None or head.is_empty or not isinstance(head.raw_value, str):
            continue
        below = [
            sheet.get_cell(r, col)
            for r in range(row + 1, bot + 1)
        ]
        below = [c for c in below if c and not c.is_empty]
        if not below:
            continue
        if sum(1 for c in below if _is_data_value(c)) * 2 >= len(below):
            label_over_data += 1
            if label_over_data >= _MIN_CONTRAST_COLS:
                return True
    return False


def find_header_span(
    sheet: SheetDTO,
    cell_range: CellRange,
    max_scan: int = DEFAULT_MAX_SCAN,
) -> HeaderSpan | None:
    """
    Find the header row(s) of a block, or ``None`` if it has no recognisable
    header (e.g. a free-text block, a single-row region, or a key/value list).

    Scans the first ``max_scan`` rows from the top:
      - single-cell rows are treated as titles/captions and skipped over;
      - the first multi-column row is the decision point — if it labels the data
        below (styled or contrast) it is the header; if it is itself data, the
        block has no header and we stop.
    """
    top = cell_range.top_left.row
    bot = cell_range.bottom_right.row
    c0 = cell_range.top_left.col
    c1 = cell_range.bottom_right.col
    if bot <= top:
        return None  # need at least one data row beneath a header

    last_candidate = min(top + max_scan - 1, bot - 1)
    for r in range(top, last_candidate + 1):
        nonempty = _row_nonempty(sheet, r, c0, c1)
        if len(nonempty) < 2:
            continue  # title/caption/section row — keep looking below it
        if _is_header_row(sheet, r, c0, c1, bot):
            return HeaderSpan(top=r, bottom=_header_band_bottom(sheet, r, c0, c1, bot))
        # First multi-column row is itself data → there is no header above it.
        return None
    return None


def _header_band_bottom(sheet: SheetDTO, r: int, c0: int, c1: int, bot: int) -> int:
    """
    Given the first header row ``r``, return the last row of the header band.

    A multi-row header is a run of contiguous *multi-column* label rows sitting
    directly above the data. A continuation row only joins the band if it is
    styled (bold/filled) like a header — genuine stacked headers are uniformly
    styled, whereas the first data row below a one-row header is not, which keeps
    single-row headers exactly one row. The band also ends at the first data row
    (numeric/date majority), a blank row, or a single-cell row (a section divider
    or sub-title). Bounded by :data:`MAX_HEADER_ROWS`.
    """
    # Only extend when the header row itself is styled; an unstyled header was
    # found by the contrast signal alone and has no styling trail to follow.
    if not _is_styled_row(sheet, r, c0, c1):
        return r
    bottom = r
    limit = min(bot - 1, r + MAX_HEADER_ROWS - 1)
    for rr in range(r + 1, limit + 1):
        if len(_row_nonempty(sheet, rr, c0, c1)) < 2:
            break  # blank or single-cell row (divider/title) terminates the band
        if _is_data_row(sheet, rr, c0, c1):
            break  # data has begun — header band is everything above this row
        if not _is_styled_row(sheet, rr, c0, c1):
            break  # unstyled row → no longer part of the styled header band
        bottom = rr
    return bottom


def has_header(sheet: SheetDTO, cell_range: CellRange) -> bool:
    """Convenience: whether the block has a recognisable column header."""
    return find_header_span(sheet, cell_range) is not None
