"""
Stage 8: Template Extraction.

Converts the structure tree into templates with degrees of freedom (DOFs).
Each template node classifies cells as CONSTANT (labels), DOF (data),
or FORMULA, and stores structural constraints.
"""

from __future__ import annotations

import logging

from models.common import CellAnnotation, CellCoord
from models.sheet import SheetDTO
from models.template import (
    DOFType,
    TemplateCellSpec,
    TemplateConstraint,
    TemplateNode,
)
from models.tree import TreeNode, TreeNodeType

logger = logging.getLogger(__name__)


class TemplateExtractor:
    """
    Extracts templates from the structure tree.

    Walks the tree and creates TemplateNode objects for each TABLE or
    TABLE_GROUP node. Cells are classified based on annotation:
    - LABEL annotation → CONSTANT (fixed value)
    - DATA annotation → DOF (degree of freedom, varies across documents)
    - Has formula → FORMULA
    """

    def __init__(self, sheet: SheetDTO, workbook_hash: str = ""):
        self._sheet = sheet
        self._workbook_hash = workbook_hash

    def extract(self, tree_nodes: list[TreeNode]) -> list[TemplateNode]:
        """
        Extract templates from tree nodes.

        Returns:
            List of TemplateNode objects with cell specifications
            and structural constraints.
        """
        node_map = {n.node_id: n for n in tree_nodes}
        templates: list[TemplateNode] = []

        for node in tree_nodes:
            if node.node_type in (TreeNodeType.TABLE, TreeNodeType.TABLE_GROUP, TreeNodeType.BLOCK):
                template = self._extract_template_node(node, node_map)
                if template:
                    templates.append(template)

        logger.info(
            "Sheet '%s': extracted %d template nodes",
            self._sheet.sheet_name,
            len(templates),
        )
        return templates

    def _extract_template_node(
        self, node: TreeNode, node_map: dict[str, TreeNode]
    ) -> TemplateNode | None:
        """Extract a TemplateNode from a TreeNode."""
        if node.cell_range is None:
            return None

        rng = node.cell_range
        cell_specs: list[TemplateCellSpec] = []

        for row in range(rng.top_left.row, rng.bottom_right.row + 1):
            for col in range(rng.top_left.col, rng.bottom_right.col + 1):
                cell = self._sheet.get_cell(row, col)
                if cell is None:
                    continue

                coord = CellCoord(row=row, col=col)

                # Classify cell
                if cell.formula:
                    dof_type = DOFType.FORMULA
                    spec = TemplateCellSpec(
                        coord=coord,
                        dof_type=dof_type,
                        formula_text=cell.formula,
                    )
                elif cell.annotation == CellAnnotation.LABEL:
                    dof_type = DOFType.CONSTANT
                    spec = TemplateCellSpec(
                        coord=coord,
                        dof_type=dof_type,
                        constant_value=cell.raw_value,
                    )
                else:
                    dof_type = DOFType.DOF
                    data_type = "string"
                    if isinstance(cell.raw_value, (int, float)) and not isinstance(cell.raw_value, bool):
                        data_type = "number"
                    elif isinstance(cell.raw_value, bool):
                        data_type = "boolean"
                    spec = TemplateCellSpec(
                        coord=coord,
                        dof_type=dof_type,
                        data_type_hint=data_type,
                    )

                cell_specs.append(spec)

        # Build constraints
        constraints: list[TemplateConstraint] = []
        constraints.append(TemplateConstraint(
            constraint_type="row_count",
            expected_value=rng.row_count(),
            description=f"Expected {rng.row_count()} rows",
        ))
        constraints.append(TemplateConstraint(
            constraint_type="col_count",
            expected_value=rng.col_count(),
            description=f"Expected {rng.col_count()} columns",
        ))

        # For table groups, add sub-table count constraint
        if node.node_type == TreeNodeType.TABLE_GROUP:
            child_tables = [
                node_map[cid] for cid in node.children_ids
                if cid in node_map and node_map[cid].node_type == TreeNodeType.TABLE
            ]
            if child_tables:
                constraints.append(TemplateConstraint(
                    constraint_type="sub_table_count",
                    expected_value=len(child_tables),
                    description=f"Expected {len(child_tables)} sub-tables",
                ))

        # Children
        children_ids = [
            cid for cid in node.children_ids
            if cid in node_map
        ]

        template = TemplateNode(
            sheet_name=self._sheet.sheet_name,
            cell_range=rng,
            tree_node_id=node.node_id,
            cell_specs=cell_specs,
            constraints=constraints,
            children_ids=children_ids,
        )
        template.finalize(self._workbook_hash)

        return template
