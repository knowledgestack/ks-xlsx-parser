"""
Stage 7: Recursive Light Table Identification.

Builds a hierarchical tree from blocks, table structures, and table groups.
The tree represents the complete structural hierarchy of a sheet, from
leaf blocks up to the sheet root.
"""

from __future__ import annotations

import logging

from ..models.block import BlockDTO
from ..models.common import CellCoord, CellRange, compute_hash
from ..models.sheet import SheetDTO
from ..models.table_structure import TableStructure
from ..models.tree import TreeNode, TreeNodeType

logger = logging.getLogger(__name__)


class TreeBuilder:
    """
    Builds a structure tree for a sheet.

    Bottom-up construction:
    1. Leaf nodes: individual blocks
    2. Table nodes: wrap TableStructures
    3. Group nodes: wrap parent blocks (table groups)
    4. Sheet root: contains all top-level nodes
    """

    def __init__(self, sheet: SheetDTO, workbook_hash: str = ""):
        self._sheet = sheet
        self._workbook_hash = workbook_hash

    def build_tree(
        self,
        blocks: list[BlockDTO],
        structures: list[TableStructure],
    ) -> list[TreeNode]:
        """
        Build the complete tree for one sheet.

        Returns:
            All TreeNode objects. The root has depth=0 and node_type=SHEET.
        """
        all_nodes: list[TreeNode] = []
        block_to_node: dict[str, TreeNode] = {}

        # Step 1: Create block-level leaf nodes
        for block in blocks:
            if block.parent_block_id:
                # This block is a child of a group; skip for now
                continue

            node = TreeNode(
                node_type=TreeNodeType.BLOCK,
                sheet_name=self._sheet.sheet_name,
                cell_range=block.cell_range,
                block_id=block.block_id,
                label=f"Block:{block.block_type.value}",
            )
            node.finalize(self._workbook_hash)
            all_nodes.append(node)
            block_to_node[block.block_id] = node

        # Step 2: Create table nodes from structures
        structure_nodes: list[TreeNode] = []
        used_block_ids: set[str] = set()

        for structure in structures:
            # Only create table nodes for structures with explicit header/footer
            has_header = structure.header_region is not None
            has_footer = structure.footer_region is not None
            has_row_labels = structure.row_label_region is not None

            if not (has_header or has_footer or has_row_labels):
                continue

            table_node = TreeNode(
                node_type=TreeNodeType.TABLE,
                sheet_name=self._sheet.sheet_name,
                cell_range=structure.overall_range,
                table_structure_id=structure.structure_id,
                label=f"Table:{structure.overall_range.to_a1() if structure.overall_range else '?'}",
            )
            table_node.finalize(self._workbook_hash)

            # Add block nodes as children
            for bid in structure.source_block_ids:
                if bid in block_to_node:
                    child_node = block_to_node[bid]
                    child_node.parent_id = table_node.node_id
                    table_node.children_ids.append(child_node.node_id)
                    used_block_ids.add(bid)

            structure_nodes.append(table_node)
            all_nodes.append(table_node)

        # Step 3: Create group nodes from parent blocks
        parent_blocks = [b for b in blocks if b.child_block_ids]
        for parent_block in parent_blocks:
            group_node = TreeNode(
                node_type=TreeNodeType.TABLE_GROUP,
                sheet_name=self._sheet.sheet_name,
                cell_range=parent_block.cell_range,
                block_id=parent_block.block_id,
                label=f"Group:{parent_block.cell_range.to_a1()}",
            )
            group_node.finalize(self._workbook_hash)

            # Find child table nodes or block nodes
            for child_bid in parent_block.child_block_ids:
                # Check if any table node wraps this block
                for tnode in structure_nodes:
                    if child_bid in [
                        bn.block_id for bn in all_nodes
                        if bn.block_id and bn.parent_id == tnode.node_id
                    ]:
                        if tnode.parent_id is None:
                            tnode.parent_id = group_node.node_id
                            group_node.children_ids.append(tnode.node_id)
                        break
                else:
                    # Direct block child
                    if child_bid in block_to_node:
                        child_node = block_to_node[child_bid]
                        if child_node.parent_id is None:
                            child_node.parent_id = group_node.node_id
                            group_node.children_ids.append(child_node.node_id)

            all_nodes.append(group_node)

        # Step 4: Create sheet root
        root = TreeNode(
            node_type=TreeNodeType.SHEET,
            sheet_name=self._sheet.sheet_name,
            cell_range=self._sheet.used_range,
            label=f"Sheet:{self._sheet.sheet_name}",
        )
        root.finalize(self._workbook_hash)

        # Attach orphan nodes to root
        for node in all_nodes:
            if node.parent_id is None and node.node_id != root.node_id:
                node.parent_id = root.node_id
                root.children_ids.append(node.node_id)

        all_nodes.append(root)

        # Compute depths via BFS from root
        self._compute_depths(root, all_nodes)

        logger.info(
            "Sheet '%s': built tree with %d nodes (root depth 0)",
            self._sheet.sheet_name,
            len(all_nodes),
        )
        return all_nodes

    def _compute_depths(self, root: TreeNode, all_nodes: list[TreeNode]) -> None:
        """Compute depth for all nodes via BFS from root."""
        node_map = {n.node_id: n for n in all_nodes}
        root.depth = 0
        queue = [root]

        while queue:
            current = queue.pop(0)
            for child_id in current.children_ids:
                child = node_map.get(child_id)
                if child:
                    child.depth = current.depth + 1
                    queue.append(child)
