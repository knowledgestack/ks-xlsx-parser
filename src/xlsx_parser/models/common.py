"""
Common types and utilities shared across all DTOs.

Provides the base model, hashing utilities, and shared enums/types
used throughout the xlsx_parser data model.
"""

from __future__ import annotations

import enum
from typing import Any

import xxhash
from pydantic import BaseModel, Field


class StableModel(BaseModel):
    """
    Base model for all DTOs. Provides deterministic JSON serialization
    and a stable hash computation method.

    All models inherit from this to ensure consistent behavior:
    - Deterministic field ordering in JSON output
    - xxhash64-based content hashing
    - Immutable-by-default configuration
    """

    model_config = {"frozen": True, "extra": "forbid"}

    def stable_json(self) -> str:
        """Return JSON with deterministic key ordering for hashing."""
        return self.model_dump_json(exclude_none=True)

    def compute_stable_hash(self) -> str:
        """Compute xxhash64 hex digest of the stable JSON representation."""
        return xxhash.xxh64(self.stable_json().encode("utf-8")).hexdigest()


class CellCoord(StableModel):
    """A single cell coordinate (1-indexed row and column)."""

    row: int = Field(ge=1, description="1-indexed row number")
    col: int = Field(ge=1, description="1-indexed column number")

    def to_a1(self) -> str:
        """Convert to A1-style reference (e.g., row=1, col=3 → 'C1')."""
        result = []
        c = self.col
        while c > 0:
            c, remainder = divmod(c - 1, 26)
            result.append(chr(65 + remainder))
        return "".join(reversed(result)) + str(self.row)

    def __str__(self) -> str:
        return self.to_a1()


class CellRange(StableModel):
    """A rectangular range of cells defined by top-left and bottom-right corners."""

    top_left: CellCoord
    bottom_right: CellCoord

    def to_a1(self) -> str:
        """Convert to A1-style range (e.g., 'A1:C10')."""
        return f"{self.top_left.to_a1()}:{self.bottom_right.to_a1()}"

    def contains(self, coord: CellCoord) -> bool:
        """Check if a cell coordinate falls within this range."""
        return (
            self.top_left.row <= coord.row <= self.bottom_right.row
            and self.top_left.col <= coord.col <= self.bottom_right.col
        )

    def row_count(self) -> int:
        return self.bottom_right.row - self.top_left.row + 1

    def col_count(self) -> int:
        return self.bottom_right.col - self.top_left.col + 1

    def __str__(self) -> str:
        return self.to_a1()


class BoundingBox(StableModel):
    """
    Pixel-level bounding box computed from row heights and column widths.
    Origin is top-left corner of the sheet (0, 0). Units are points (1/72 inch).
    """

    x: float = Field(description="Left edge in points")
    y: float = Field(description="Top edge in points")
    width: float = Field(description="Width in points")
    height: float = Field(description="Height in points")


class CellAnnotation(str, enum.Enum):
    """Cell role annotation assigned during Stage 1."""

    DATA = "data"
    LABEL = "label"


class BlockType(str, enum.Enum):
    """Classification of a detected layout block."""

    TABLE = "table"
    CALCULATION_BLOCK = "calculation_block"
    ASSUMPTIONS_TABLE = "assumptions_table"
    RESULTS_BLOCK = "results_block"
    TEXT_BLOCK = "text_block"
    HEADER = "header"
    CHART_ANCHOR = "chart_anchor"
    IMAGE_ANCHOR = "image_anchor"
    MIXED = "mixed"
    EMPTY = "empty"
    LABEL_BLOCK = "label_block"
    DATA_BLOCK = "data_block"
    LIGHT_BLOCK = "light_block"


class EdgeType(str, enum.Enum):
    """Type of formula dependency edge."""

    CELL_TO_CELL = "cell_to_cell"
    CELL_TO_RANGE = "cell_to_range"
    CROSS_SHEET = "cross_sheet"
    EXTERNAL = "external"
    STRUCTURED_REF = "structured_ref"
    NAMED_RANGE = "named_range"


class ChartType(str, enum.Enum):
    """Supported chart types extracted from OOXML."""

    BAR = "bar"
    COLUMN = "column"
    LINE = "line"
    PIE = "pie"
    AREA = "area"
    SCATTER = "scatter"
    DOUGHNUT = "doughnut"
    RADAR = "radar"
    BUBBLE = "bubble"
    SURFACE = "surface"
    STOCK = "stock"
    COMBO = "combo"
    UNKNOWN = "unknown"


class Severity(str, enum.Enum):
    """Severity level for parse errors/warnings."""

    INFO = "info"
    WARNING = "warning"
    ERROR = "error"


class ParseError(StableModel):
    """A non-fatal error encountered during parsing."""

    model_config = {"frozen": True, "extra": "forbid"}

    severity: Severity
    stage: str = Field(description="Pipeline stage where the error occurred")
    message: str
    sheet_name: str | None = None
    cell_ref: str | None = None
    detail: dict[str, Any] | None = None


def compute_hash(*parts: str) -> str:
    """Compute xxhash64 hex digest from concatenated string parts."""
    hasher = xxhash.xxh64()
    for part in parts:
        hasher.update(part.encode("utf-8"))
    return hasher.hexdigest()


def col_letter_to_number(col_str: str) -> int:
    """Convert column letter(s) to 1-indexed number. E.g., 'A'→1, 'Z'→26, 'AA'→27."""
    result = 0
    for char in col_str.upper():
        result = result * 26 + (ord(char) - ord("A") + 1)
    return result


def col_number_to_letter(col_num: int) -> str:
    """Convert 1-indexed column number to letter(s). E.g., 1→'A', 27→'AA'."""
    result = []
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        result.append(chr(65 + remainder))
    return "".join(reversed(result))
