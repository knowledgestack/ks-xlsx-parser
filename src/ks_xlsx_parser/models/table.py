"""
Table DTO for Excel ListObject tables.

Represents a named table defined in Excel with headers, data body,
optional totals row, structured references, and banding styles.
Tables are first-class objects for RAG retrieval.
"""

from __future__ import annotations

from pydantic import Field

from .common import CellRange, StableModel, compute_hash


class TableColumn(StableModel):
    """A single column definition within a table."""

    name: str
    column_index: int  # 0-based index within the table
    totals_function: str | None = None  # "sum", "count", "average", etc.
    totals_value: str | None = None
    data_type_hint: str | None = None  # Inferred from data: "numeric", "text", "date"


class TableDTO(StableModel):
    """
    An Excel ListObject table.

    Tables are defined regions with named columns, optional headers row,
    optional totals row, and structured reference support. They are key
    RAG retrieval units.
    """

    model_config = {"frozen": False, "extra": "forbid"}

    # Identity
    table_name: str
    display_name: str
    sheet_name: str
    table_id: str = Field(default="", description="Deterministic ID")

    # Range
    ref_range: CellRange  # Full range including headers and totals
    header_range: CellRange | None = None  # Just the header row
    data_range: CellRange | None = None  # Just the data body
    totals_range: CellRange | None = None  # Just the totals row (if present)

    # Columns
    columns: list[TableColumn] = Field(default_factory=list)

    # Style
    style_name: str | None = None  # e.g., "TableStyleMedium2"
    show_first_column: bool = False
    show_last_column: bool = False
    show_row_stripes: bool = True
    show_column_stripes: bool = False

    # Auto-filter
    has_auto_filter: bool = False
    has_totals_row: bool = False

    # Hash
    content_hash: str = Field(default="")

    def finalize(self, workbook_hash: str) -> None:
        """Compute stable ID and content hash."""
        self.table_id = compute_hash(
            workbook_hash, self.sheet_name, self.table_name
        )
        col_sig = "|".join(c.name for c in self.columns)
        self.content_hash = compute_hash(
            self.sheet_name,
            self.table_name,
            self.ref_range.to_a1(),
            col_sig,
        )
