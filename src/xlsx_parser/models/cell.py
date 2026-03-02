"""
Cell-level DTOs capturing value, formula, formatting, and metadata.

A CellDTO is the atomic unit of the workbook representation. It stores
the raw value, the display-formatted value, the formula (if any), and
a snapshot of the cell's style. Each cell has a stable content hash
for change detection and citation.
"""

from __future__ import annotations

from typing import Any

from pydantic import Field

from .common import CellAnnotation, CellCoord, StableModel, compute_hash


class FontStyle(StableModel):
    """Font attributes for a cell."""

    name: str | None = None
    size: float | None = None
    bold: bool = False
    italic: bool = False
    underline: str | None = None  # "single", "double", etc.
    strikethrough: bool = False
    color: str | None = None  # hex color string e.g. "FF0000"


class FillStyle(StableModel):
    """Fill/background attributes for a cell."""

    pattern_type: str | None = None  # "solid", "gray125", etc.
    fg_color: str | None = None
    bg_color: str | None = None


class BorderSide(StableModel):
    """A single border edge."""

    style: str | None = None  # "thin", "medium", "thick", "dashed", etc.
    color: str | None = None


class BorderStyle(StableModel):
    """All four border edges of a cell."""

    left: BorderSide | None = None
    right: BorderSide | None = None
    top: BorderSide | None = None
    bottom: BorderSide | None = None


class AlignmentStyle(StableModel):
    """Text alignment attributes."""

    horizontal: str | None = None  # "left", "center", "right", "justify"
    vertical: str | None = None  # "top", "center", "bottom"
    wrap_text: bool = False
    text_rotation: int | None = None
    indent: int | None = None


class CellStyle(StableModel):
    """Complete style snapshot for a cell."""

    font: FontStyle | None = None
    fill: FillStyle | None = None
    border: BorderStyle | None = None
    alignment: AlignmentStyle | None = None
    number_format: str | None = None  # e.g., "#,##0.00", "yyyy-mm-dd"
    protection_locked: bool | None = None
    protection_hidden: bool | None = None


class CellDTO(StableModel):
    """
    Complete representation of a single cell.

    Captures both the data content (raw value, display value, formula)
    and the presentation (style). Merged cells are represented by having
    `is_merged_slave=True` with a reference to the master cell coordinate.
    """

    model_config = {"frozen": False, "extra": "forbid"}

    # Location
    coord: CellCoord
    sheet_name: str

    # Values
    raw_value: Any = None  # Python-native value (str, int, float, datetime, bool, None)
    display_value: str | None = None  # Formatted string as it would appear in Excel
    data_type: str | None = None  # "s" (string), "n" (number), "d" (date), "b" (bool), "f" (formula), "e" (error)

    # Formula
    formula: str | None = None  # Raw formula string without leading '='
    formula_value: Any = None  # Computed value from data_only pass

    # Style
    style: CellStyle | None = None

    # Merge info
    is_merged_master: bool = False
    is_merged_slave: bool = False
    merge_master: CellCoord | None = None  # For slaves: coord of the master cell
    merge_extent: int | None = None  # For masters: number of rows spanned
    merge_col_extent: int | None = None  # For masters: number of cols spanned

    # Comment
    comment_text: str | None = None
    comment_author: str | None = None

    # Conditional formatting (evaluated rules that apply to this cell)
    conditional_formats: list[str] | None = None

    # Data validation
    data_validation: str | None = None  # Serialized validation rule

    # Hyperlink
    hyperlink: str | None = None

    # Cell annotation (Stage 1: Excellent Algorithm)
    annotation: CellAnnotation | None = None
    annotation_confidence: float | None = None  # 0.0-1.0

    # Stable ID and hash
    cell_id: str = Field(default="", description="Deterministic ID: sheet|row|col")
    cell_hash: str = Field(default="", description="Content hash of value+formula")

    def compute_id(self) -> str:
        """Generate deterministic cell ID."""
        return f"{self.sheet_name}|{self.coord.row}|{self.coord.col}"

    def compute_cell_hash(self) -> str:
        """Generate content hash from identity + content fields."""
        return compute_hash(
            self.sheet_name,
            str(self.coord.row),
            str(self.coord.col),
            str(self.raw_value) if self.raw_value is not None else "",
            self.formula or "",
        )

    def finalize(self) -> None:
        """Compute and set the cell_id and cell_hash fields."""
        self.cell_id = self.compute_id()
        self.cell_hash = self.compute_cell_hash()

    @property
    def is_empty(self) -> bool:
        """A cell is empty if it has no value and no formula."""
        return self.raw_value is None and self.formula is None

    @property
    def a1_ref(self) -> str:
        """Full A1-style reference including sheet name."""
        return f"{self.sheet_name}!{self.coord.to_a1()}"
