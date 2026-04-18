"""
Common types and utilities shared across all DTOs.

Provides the base model, hashing utilities, and shared enums/types
used throughout the xlsx_parser data model.
"""

from __future__ import annotations

import enum
from dataclasses import dataclass
from functools import lru_cache
from typing import Any

import xxhash
from pydantic import BaseModel, Field


@lru_cache(maxsize=1 << 16)
def _col_to_letters(col: int) -> str:
    """Convert 1-indexed column number to A1 letters (e.g. 1→'A', 27→'AA')."""
    letters: list[str] = []
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        letters.append(chr(65 + remainder))
    return "".join(reversed(letters))


@lru_cache(maxsize=1 << 18)
def addr_to_a1(row: int, col: int) -> str:
    """Convert (row, col) to A1-style reference (e.g. row=1,col=3 → 'C1').

    Cached: a workbook of 100k distinct cells hits the cache on every
    subsequent lookup inside the dependency graph, chunking, and indexing
    paths.
    """
    return _col_to_letters(col) + str(row)


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


@dataclass(frozen=True, slots=True, eq=True)
class CellCoord:
    """A single cell coordinate (1-indexed row and column).

    **Not a Pydantic model** — frozen slotted dataclass. Profiling showed
    339k Pydantic inits on Walbridge contributed ~0.65 s of parse time;
    dataclass construction is ~2.2× faster with the same immutability
    and equality semantics. Validation of ``row >= 1`` / ``col >= 1`` is
    dropped: all producers in this codebase build coords from parsed
    OOXML or openpyxl cells that are already 1-indexed.
    """

    row: int
    col: int

    def to_a1(self) -> str:
        """Convert to A1-style reference (e.g., row=1, col=3 → 'C1')."""
        return addr_to_a1(self.row, self.col)

    def __str__(self) -> str:
        return addr_to_a1(self.row, self.col)

    # Pydantic-compat: `model_dump` is occasionally called on deeply-nested
    # DTOs that contain a CellCoord. Keeping a shim avoids touching those
    # callers. Emits a plain dict identical to pydantic's default output.
    def model_dump(self, **_kwargs: Any) -> dict[str, int]:
        return {"row": self.row, "col": self.col}


@dataclass(frozen=True, slots=True, eq=True)
class CellRange:
    """A rectangular range of cells defined by top-left and bottom-right corners.

    Like :class:`CellCoord`, a slotted frozen dataclass rather than a
    Pydantic model — construction is hot-path (formula dep graph, chunking).
    """

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

    def model_dump(self, **_kwargs: Any) -> dict[str, Any]:
        return {
            "top_left": self.top_left.model_dump(),
            "bottom_right": self.bottom_right.model_dump(),
        }


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


class CalculationMode(str, enum.Enum):
    """Workbook calculation mode."""

    AUTO = "auto"
    MANUAL = "manual"
    SEMI_AUTOMATIC = "semiAutomatic"


class DateSystem(str, enum.Enum):
    """Workbook date system base."""

    DATE_1900 = "1900"
    DATE_1904 = "1904"


class PivotLayoutType(str, enum.Enum):
    """PivotTable layout type."""

    COMPACT = "compact"
    TABULAR = "tabular"
    OUTLINE = "outline"


class SheetPurpose(str, enum.Enum):
    """Detected purpose of a worksheet."""

    INPUT = "input"
    CALCULATION = "calculation"
    DASHBOARD = "dashboard"
    RAW_DATA = "raw_data"
    LOOKUP = "lookup"
    REPORT = "report"
    TEMPLATE = "template"
    CONFIG = "config"
    UNKNOWN = "unknown"


class RichTextRun(StableModel):
    """A single run of text within a rich-text cell."""

    model_config = {"frozen": True, "extra": "forbid"}

    text: str
    bold: bool = False
    italic: bool = False
    underline: str | None = None
    strikethrough: bool = False
    color: str | None = None
    font_name: str | None = None
    font_size: float | None = None


class FilterCriteria(StableModel):
    """A single autofilter criterion on a column."""

    model_config = {"frozen": True, "extra": "forbid"}

    col_index: int  # 0-based column offset within the filter range
    filter_type: str = "values"  # "values", "custom", "top10", "dynamic", "color"
    values: list[str] = Field(default_factory=list)
    operator: str | None = None
    custom_value: str | None = None


class SortKey(StableModel):
    """A sort key applied to the sheet."""

    model_config = {"frozen": True, "extra": "forbid"}

    col_index: int
    descending: bool = False


class PivotField(StableModel):
    """A field in a PivotTable (row, column, filter, or value)."""

    model_config = {"frozen": True, "extra": "forbid"}

    name: str
    field_index: int | None = None
    subtotals: list[str] = Field(default_factory=list)
    items: list[str] = Field(default_factory=list)


class PivotValueField(StableModel):
    """A value/measure field in a PivotTable."""

    model_config = {"frozen": True, "extra": "forbid"}

    name: str
    source_field: str | None = None
    aggregation: str = "sum"  # "sum", "count", "average", "max", "min", etc.
    number_format: str | None = None


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
