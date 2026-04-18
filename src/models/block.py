"""
Block and Chunk DTOs for layout segmentation and RAG retrieval.

A Block is a contiguous region of a sheet identified by the segmentation
algorithm. A Chunk is a RAG-ready unit derived from one or more blocks,
enriched with rendered content, dependency context, and metadata.
"""

from __future__ import annotations

from typing import Any

from pydantic import Field

from .common import (
    BlockType,
    BoundingBox,
    CellCoord,
    CellRange,
    StableModel,
    compute_hash,
)


class BlockDTO(StableModel):
    """
    A logical block identified by the layout segmentation algorithm.

    Represents a contiguous region of a sheet with a semantic type
    (table, calculation block, header, etc.). Blocks are the
    intermediate unit between raw cells and RAG chunks.
    """

    model_config = {"frozen": False, "extra": "forbid"}

    # Identity
    block_index: int  # Index within the sheet's block list
    sheet_name: str
    block_id: str = Field(default="", description="Deterministic ID")

    # Type
    block_type: BlockType = BlockType.MIXED

    # Coordinates
    cell_range: CellRange
    bounding_box: BoundingBox | None = None

    # Content summary
    cell_count: int = 0
    formula_count: int = 0
    has_merges: bool = False
    has_formatting: bool = False

    # Key cells (e.g., bold/colored output cells, headers)
    key_cells: list[CellCoord] = Field(default_factory=list)

    # Overlapping named ranges
    named_ranges: list[str] = Field(default_factory=list)

    # Table reference (if this block corresponds to an Excel table)
    table_name: str | None = None

    # Parent/child relationships (Stage 5: table grouping)
    parent_block_id: str | None = None
    child_block_ids: list[str] = Field(default_factory=list)

    # Density and annotation counts (Stage 2/4)
    density: float | None = None
    label_cell_count: int = 0
    data_cell_count: int = 0

    # Table structure reference (Stage 3)
    table_structure_id: str | None = None

    # Hash
    content_hash: str = Field(default="")

    def finalize(self, workbook_hash: str) -> None:
        """Compute stable ID and content hash."""
        self.block_id = compute_hash(
            workbook_hash,
            self.sheet_name,
            str(self.block_index),
            self.cell_range.to_a1(),
        )
        self.content_hash = compute_hash(
            self.sheet_name,
            self.block_type.value,
            self.cell_range.to_a1(),
            str(self.cell_count),
            str(self.formula_count),
        )


class DependencySummary(StableModel):
    """Compact summary of a chunk's dependency context."""

    upstream_refs: list[str] = Field(default_factory=list)  # A1-style refs
    downstream_refs: list[str] = Field(default_factory=list)
    cross_sheet_refs: list[str] = Field(default_factory=list)
    has_circular: bool = False


class ChunkDTO(StableModel):
    """
    A RAG-ready chunk derived from a block or group of blocks.

    This is the primary unit stored in the vector store. It includes
    rendered content (HTML and text), dependency context, token count,
    and full provenance metadata for citations.
    """

    model_config = {"frozen": False, "extra": "forbid"}

    # Identity
    chunk_id: str = Field(default="", description="Deterministic chunk ID")
    chunk_index: int = 0  # Global index within the workbook

    # Source provenance
    source_uri: str = ""  # e.g., "workbook.xlsx#Sheet1!A1:D20"
    workbook_hash: str = ""
    sheet_name: str = ""
    block_type: BlockType = BlockType.MIXED

    # Coordinates
    top_left_cell: str = ""  # A1-style
    bottom_right_cell: str = ""  # A1-style
    cell_range: CellRange | None = None

    # Key cells (notable outputs, headers, etc.)
    key_cells: list[str] = Field(default_factory=list)  # A1-style refs

    # Named ranges overlapping this chunk
    named_ranges: list[str] = Field(default_factory=list)

    # Dependencies
    dependency_summary: DependencySummary = Field(default_factory=DependencySummary)

    # Rendered content
    render_html: str = ""
    render_text: str = ""

    # Token count estimate (for embedding budget management)
    token_count: int = 0

    # Hash
    content_hash: str = ""

    # Navigation: prev/next chunk IDs for sequential traversal
    prev_chunk_id: str | None = None
    next_chunk_id: str | None = None

    # Metadata bag for extensibility
    metadata: dict[str, Any] = Field(default_factory=dict)

    def finalize(self, workbook_hash: str, workbook_path: str) -> None:
        """Compute chunk ID, source URI, and content hash."""
        self.workbook_hash = workbook_hash
        self.source_uri = f"{workbook_path}#{self.sheet_name}!{self.top_left_cell}:{self.bottom_right_cell}"
        self.content_hash = compute_hash(
            workbook_hash,
            self.sheet_name,
            self.block_type.value,
            self.top_left_cell,
            self.bottom_right_cell,
            self.render_text,
        )
        self.chunk_id = compute_hash(
            workbook_hash,
            self.sheet_name,
            self.block_type.value,
            self.top_left_cell,
            self.bottom_right_cell,
        )
