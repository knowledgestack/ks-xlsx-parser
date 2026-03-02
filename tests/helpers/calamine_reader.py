"""
Wrapper around python-calamine for cross-validation reads.

Provides a standardized CalamineResult that can be compared against
our parser's WorkbookDTO output.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from python_calamine import CalamineWorkbook


@dataclass
class CalamineSheetData:
    """Parsed data from a single sheet via calamine."""

    name: str
    rows: list[list[Any]]  # raw rows from to_python()
    total_height: int
    total_width: int
    start: tuple[int, int] | None  # (row, col) 0-indexed top-left of data
    end: tuple[int, int] | None  # (row, col) 0-indexed bottom-right of data
    merged_ranges: list[tuple[tuple[int, int], tuple[int, int]]] | None = None

    def get_value(self, row_1based: int, col_1based: int) -> Any:
        """Get cell value using 1-based coordinates (matching Excel/openpyxl)."""
        row_idx = row_1based - 1
        col_idx = col_1based - 1
        if 0 <= row_idx < len(self.rows):
            row = self.rows[row_idx]
            if 0 <= col_idx < len(row):
                return row[col_idx]
        return None


@dataclass
class CalamineResult:
    """Complete calamine parse result for cross-validation."""

    sheet_names: list[str] = field(default_factory=list)
    sheets: dict[str, CalamineSheetData] = field(default_factory=dict)

    @classmethod
    def from_path(cls, path: str | Path) -> CalamineResult:
        """Parse an xlsx file with calamine and return structured result."""
        wb = CalamineWorkbook.from_path(str(path))
        result = cls(sheet_names=list(wb.sheet_names))

        for name in wb.sheet_names:
            sheet = wb.get_sheet_by_name(name)
            # skip_empty_area=False so indices map directly to 1-based coords
            rows = sheet.to_python(skip_empty_area=False)

            merged = None
            try:
                merged = sheet.merged_cell_ranges
            except Exception:
                pass  # not all formats support this

            result.sheets[name] = CalamineSheetData(
                name=name,
                rows=rows,
                total_height=sheet.total_height,
                total_width=sheet.total_width,
                start=sheet.start,
                end=sheet.end,
                merged_ranges=merged,
            )

        wb.close()
        return result
