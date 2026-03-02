"""
Sheet-level DTO capturing all data within a single worksheet.

A SheetDTO aggregates cells, tables, charts, shapes, and layout metadata.
It serves as the primary container for per-sheet data and is the unit
of parallel processing in the pipeline.
"""

from __future__ import annotations

from typing import Any

from pydantic import Field

from .cell import CellDTO
from .common import (
    BoundingBox,
    CellCoord,
    CellRange,
    ParseError,
    StableModel,
    compute_hash,
)


class ConditionalFormatRule(StableModel):
    """A single conditional formatting rule on a sheet."""

    model_config = {"frozen": True, "extra": "forbid"}

    ranges: list[str]  # A1-style range strings
    rule_type: str  # "cellIs", "colorScale", "dataBar", "iconSet", "expression", etc.
    operator: str | None = None  # "greaterThan", "lessThan", "between", etc.
    formula: str | None = None
    priority: int | None = None
    stop_if_true: bool = False
    format_description: str | None = None  # Human-readable summary


class DataValidationRule(StableModel):
    """A data validation rule applied to a range of cells."""

    model_config = {"frozen": True, "extra": "forbid"}

    ranges: list[str]
    validation_type: str  # "list", "whole", "decimal", "date", "textLength", "custom"
    operator: str | None = None
    formula1: str | None = None
    formula2: str | None = None
    allow_blank: bool = True
    show_error_message: bool = False
    error_title: str | None = None
    error_message: str | None = None
    prompt_title: str | None = None
    prompt_message: str | None = None


class MergedRegion(StableModel):
    """A merged cell region with its master cell and extent."""

    model_config = {"frozen": True, "extra": "forbid"}

    range: CellRange
    master: CellCoord


class SheetProperties(StableModel):
    """Non-data properties of a worksheet."""

    model_config = {"frozen": True, "extra": "forbid"}

    is_hidden: bool = False
    tab_color: str | None = None
    default_row_height: float | None = None
    default_col_width: float | None = None
    freeze_pane: str | None = None  # A1-style ref of the freeze pane split
    print_area: str | None = None
    print_title_rows: str | None = None
    print_title_cols: str | None = None
    auto_filter_range: str | None = None
    zoom_scale: int | None = None
    sheet_protection: bool = False
    right_to_left: bool = False


class SheetDTO(StableModel):
    """
    Complete representation of a single worksheet.

    Stores cells in a sparse dict keyed by (row, col) for memory efficiency
    on large/sparse sheets. Also holds tables, charts, shapes, and all
    sheet-level metadata.
    """

    model_config = {"frozen": False, "extra": "forbid"}

    # Identity
    sheet_name: str
    sheet_index: int  # 0-based position in workbook
    sheet_id: str = Field(default="", description="Deterministic ID")

    # Cells: sparse storage keyed by "row,col"
    cells: dict[str, CellDTO] = Field(default_factory=dict)

    # Dimensions
    used_range: CellRange | None = None  # Actual used range (computed, not from XML)
    row_heights: dict[int, float] = Field(default_factory=dict)  # row_num → height in points
    col_widths: dict[int, float] = Field(default_factory=dict)  # col_num → width in characters

    # Hidden rows and columns
    hidden_rows: set[int] = Field(default_factory=set)
    hidden_cols: set[int] = Field(default_factory=set)

    # Merges
    merged_regions: list[MergedRegion] = Field(default_factory=list)

    # Sheet-level metadata
    properties: SheetProperties = Field(default_factory=SheetProperties)
    conditional_format_rules: list[ConditionalFormatRule] = Field(default_factory=list)
    data_validations: list[DataValidationRule] = Field(default_factory=list)

    # Errors collected during parsing
    errors: list[ParseError] = Field(default_factory=list)

    def get_cell(self, row: int, col: int) -> CellDTO | None:
        """Retrieve a cell by row and column number."""
        return self.cells.get(f"{row},{col}")

    def set_cell(self, cell: CellDTO) -> None:
        """Store a cell in the sparse cell dict."""
        self.cells[f"{cell.coord.row},{cell.coord.col}"] = cell

    def cell_count(self) -> int:
        """Number of non-empty cells stored."""
        return len(self.cells)

    def compute_used_range(self) -> CellRange | None:
        """Compute the actual used range from stored cells."""
        if not self.cells:
            return None
        min_row = min_col = float("inf")
        max_row = max_col = 0
        for cell in self.cells.values():
            min_row = min(min_row, cell.coord.row)
            min_col = min(min_col, cell.coord.col)
            max_row = max(max_row, cell.coord.row)
            max_col = max(max_col, cell.coord.col)
        return CellRange(
            top_left=CellCoord(row=int(min_row), col=int(min_col)),
            bottom_right=CellCoord(row=int(max_row), col=int(max_col)),
        )

    def compute_bounding_box(self, cell_range: CellRange) -> BoundingBox:
        """
        Compute pixel bounding box for a cell range using row heights
        and column widths. Defaults: row height=15pt, col width=8.43 chars≈64pt.
        """
        default_row_h = self.properties.default_row_height or 15.0
        default_col_w = (self.properties.default_col_width or 8.43) * 7.5  # chars to points approx

        x = sum(
            self.col_widths.get(c, default_col_w)
            for c in range(1, cell_range.top_left.col)
        )
        y = sum(
            self.row_heights.get(r, default_row_h)
            for r in range(1, cell_range.top_left.row)
        )
        width = sum(
            self.col_widths.get(c, default_col_w)
            for c in range(cell_range.top_left.col, cell_range.bottom_right.col + 1)
        )
        height = sum(
            self.row_heights.get(r, default_row_h)
            for r in range(cell_range.top_left.row, cell_range.bottom_right.row + 1)
        )
        return BoundingBox(x=x, y=y, width=width, height=height)

    def finalize(self, workbook_hash: str) -> None:
        """Compute IDs and hashes for the sheet and all its cells."""
        self.sheet_id = compute_hash(workbook_hash, self.sheet_name, str(self.sheet_index))
        self.used_range = self.compute_used_range()
        for cell in self.cells.values():
            cell.finalize()
