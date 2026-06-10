"""
Table structure models for Stage 3 (Solid Table Identification).

A TableStructure represents a logical table composed of multiple regions
(header, body, footer, row labels) assembled from adjacent blocks.
"""

from __future__ import annotations

import enum

from pydantic import Field

from .common import CellRange, StableModel, compute_hash


class TableRegionRole(str, enum.Enum):
    """Role of a region within a table structure."""

    HEADER = "header"
    BODY = "body"
    FOOTER = "footer"
    ROW_LABEL = "row_label"


class TableRegion(StableModel):
    """A contiguous region within a table structure."""

    model_config = {"frozen": False, "extra": "forbid"}

    role: TableRegionRole
    cell_range: CellRange
    source_block_id: str | None = None  # Block that contributed this region


class TableStructure(StableModel):
    """
    A logical table assembled from label and data blocks.

    Created during Stage 3 by associating label blocks (headers, row labels,
    footers) with adjacent data blocks. Represents the explicit
    header-body-footer structure of a table.
    """

    model_config = {"frozen": False, "extra": "forbid"}

    structure_id: str = Field(default="")
    sheet_name: str = ""
    regions: list[TableRegion] = Field(default_factory=list)
    source_block_ids: list[str] = Field(default_factory=list)
    overall_range: CellRange | None = None

    def finalize(self, workbook_hash: str) -> None:
        """Compute stable ID."""
        range_str = self.overall_range.to_a1() if self.overall_range else ""
        self.structure_id = compute_hash(
            workbook_hash, self.sheet_name, range_str,
            ",".join(self.source_block_ids),
        )

    @property
    def header_region(self) -> TableRegion | None:
        for r in self.regions:
            if r.role == TableRegionRole.HEADER:
                return r
        return None

    @property
    def body_region(self) -> TableRegion | None:
        for r in self.regions:
            if r.role == TableRegionRole.BODY:
                return r
        return None

    @property
    def footer_region(self) -> TableRegion | None:
        for r in self.regions:
            if r.role == TableRegionRole.FOOTER:
                return r
        return None

    @property
    def row_label_region(self) -> TableRegion | None:
        for r in self.regions:
            if r.role == TableRegionRole.ROW_LABEL:
                return r
        return None
