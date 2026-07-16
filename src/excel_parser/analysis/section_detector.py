"""
Deterministic section detection inside a table.

Real-world tables are very often divided into sections (Region › Sub-region ›
rows; REVENUE / EXPENSES; month groups). This module finds those sections from
the *file's own structure* — no layout inference, no AI — using two signals, in
priority order:

  Tier 1 — Excel row outline / grouping levels (Data → Group). Deterministic and
           natively nesting-aware: a row at outline level L is a child of the
           nearest preceding row at a lower level. Used whenever outline levels
           vary within the table's data region.
  Tier 2 — Merged-row dividers: a row whose cells are merged horizontally from
           the table's left edge across at least half its width is a flat
           (single-level) section header. Used when there is no outline grouping.

We deliberately do NOT infer sections from "a lone label with empty data
columns" — a legitimate data row can have all value columns blank, so that
signal produces false positives. (That heuristic is available as an explicit,
opt-in Tier 3 elsewhere, never the default.) Returning no sections is correct
when the file carries no structural section signal.
"""
from __future__ import annotations

from dataclasses import dataclass

from excel_parser.analysis.header_detector import HeaderSpan
from excel_parser.models.common import CellRange
from excel_parser.models.sheet import SheetDTO


@dataclass(frozen=True)
class Section:
    """A section within a table.

    ``level`` is the nesting depth (0 = top-level). ``top_row`` is the section
    header row; ``bottom_row`` is the last row that belongs to the section
    (inclusive). A nested section's row range is contained within its parent's.
    """

    label: str
    level: int
    top_row: int
    bottom_row: int


def _row_label(sheet: SheetDTO, row: int, c0: int, c1: int) -> str:
    """The section label = the first non-empty cell value in the row."""
    for col in range(c0, c1 + 1):
        cell = sheet.get_cell(row, col)
        if cell and not cell.is_empty and cell.raw_value is not None:
            return str(cell.raw_value).strip()
    return ""


def _merged_divider_rows(sheet: SheetDTO, cell_range: CellRange) -> list[int]:
    """Rows spanned by a horizontal merge from the table's left edge across at
    least half its width — flat section dividers."""
    c0 = cell_range.top_left.col
    c1 = cell_range.bottom_right.col
    width = c1 - c0 + 1
    min_span = max(2, (width + 1) // 2)
    rows: list[int] = []
    for mr in sheet.merged_regions:
        r = mr.range
        if r.top_left.row != r.bottom_right.row:
            continue  # multi-row merge — not a row divider
        if r.top_left.col != c0:
            continue  # must start at the table's left edge
        span = min(r.bottom_right.col, c1) - r.top_left.col + 1
        if span >= min_span:
            rows.append(r.top_left.row)
    return sorted(rows)


def find_sections(
    sheet: SheetDTO,
    cell_range: CellRange,
    header_span: HeaderSpan | None,
) -> list[Section]:
    """
    Find the sections of a table block, in document order. Returns an empty list
    when the file carries no structural section signal (no outline grouping and
    no merged-row dividers).
    """
    top = cell_range.top_left.row
    bot = cell_range.bottom_right.row
    c0 = cell_range.top_left.col
    c1 = cell_range.bottom_right.col
    data_start = (header_span.bottom + 1) if header_span else top
    if bot < data_start:
        return []

    levels = sheet.row_outline_levels
    data_levels = {r: levels.get(r, 0) for r in range(data_start, bot + 1)}
    has_outline = len(set(data_levels.values())) > 1

    sections: list[Section] = []

    if has_outline:
        # Tier 1: outline tree. A row introduces a section if a row below it
        # (before the level returns to <= its own) sits at a deeper level.
        for r in range(data_start, bot + 1):
            level = data_levels[r]
            introduces = False
            end = bot
            for rr in range(r + 1, bot + 1):
                if data_levels[rr] <= level:
                    end = rr - 1
                    break
                introduces = True  # at least one deeper row before the level drops
            if not introduces:
                continue
            label = _row_label(sheet, r, c0, c1)
            if label:
                sections.append(
                    Section(label=label, level=level, top_row=r, bottom_row=end)
                )
    else:
        # Tier 2: merged-row dividers → flat sections.
        dividers = [d for d in _merged_divider_rows(sheet, cell_range) if d >= data_start]
        for i, d in enumerate(dividers):
            label = _row_label(sheet, d, c0, c1)
            if not label:
                continue
            end = (dividers[i + 1] - 1) if i + 1 < len(dividers) else bot
            sections.append(Section(label=label, level=0, top_row=d, bottom_row=end))

    return sections


def section_path_at(sections: list[Section], row: int) -> list[Section]:
    """The ancestor chain of sections containing ``row``, shallowest first."""
    path = [s for s in sections if s.top_row <= row <= s.bottom_row]
    path.sort(key=lambda s: (s.level, s.top_row))
    return path
