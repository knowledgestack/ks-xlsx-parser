"""
Tree structure models for Stage 7 (Recursive Light Table Identification).

Represents the hierarchical structure of a spreadsheet, from leaf blocks
up through tables and table groups to the sheet root.
"""

from __future__ import annotations

import enum

from pydantic import Field

from .common import CellRange, StableModel, compute_hash


class TreeNodeType(str, enum.Enum):
    """Type of node in the spreadsheet structure tree."""

    SHEET = "sheet"
    TABLE_GROUP = "table_group"
    TABLE = "table"
    BLOCK = "block"


class TreeNode(StableModel):
    """
    A node in the spreadsheet structure tree.

    Built bottom-up during Stage 7. Leaf nodes wrap individual blocks,
    table nodes wrap TableStructures, group nodes wrap related tables,
    and the root is the sheet node.
    """

    model_config = {"frozen": False, "extra": "forbid"}

    node_id: str = Field(default="")
    node_type: TreeNodeType
    sheet_name: str = ""
    cell_range: CellRange | None = None

    # Tree relationships
    parent_id: str | None = None
    children_ids: list[str] = Field(default_factory=list)
    depth: int = 0

    # References to source objects
    block_id: str | None = None
    table_structure_id: str | None = None

    # Metadata
    label: str = ""  # Human-readable label

    def finalize(self, workbook_hash: str) -> None:
        """Compute stable node ID."""
        range_str = self.cell_range.to_a1() if self.cell_range else ""
        self.node_id = compute_hash(
            workbook_hash, self.sheet_name,
            self.node_type.value, range_str,
            str(self.depth),
        )
