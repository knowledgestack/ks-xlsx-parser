"""
Template models for Stages 8-10 (Template Extraction, Comparison, Export).

Templates represent the structure of a spreadsheet with cells classified
as constants (labels), degrees of freedom (data), or formulas. Templates
can be compared across documents and exported as importable Python classes.
"""

from __future__ import annotations

import enum
from typing import Any

from pydantic import Field

from .common import CellCoord, CellRange, StableModel, compute_hash


class DOFType(str, enum.Enum):
    """Classification of a cell in a template."""

    CONSTANT = "constant"   # Label cell - value is fixed across documents
    DOF = "dof"             # Data cell - value varies (degree of freedom)
    FORMULA = "formula"     # Formula cell - computed from other cells


class TemplateCellSpec(StableModel):
    """Specification of a single cell within a template."""

    model_config = {"frozen": False, "extra": "forbid"}

    coord: CellCoord
    dof_type: DOFType
    constant_value: Any = None  # For CONSTANT cells: the expected value
    formula_text: str | None = None  # For FORMULA cells: the formula
    data_type_hint: str | None = None  # Expected data type for DOF cells
    is_flexible: bool = False  # True if this cell was promoted to DOF during comparison


class TemplateConstraint(StableModel):
    """A structural constraint on a template region."""

    model_config = {"frozen": True, "extra": "forbid"}

    constraint_type: str  # "row_count", "col_count", "sub_table_count", "value_range"
    expected_value: Any = None
    min_value: Any = None
    max_value: Any = None
    description: str = ""


class TemplateNode(StableModel):
    """
    A node in the template tree.

    Created during Stage 8 from the structure tree. Each node contains
    cell specifications and structural constraints.
    """

    model_config = {"frozen": False, "extra": "forbid"}

    template_node_id: str = Field(default="")
    sheet_name: str = ""
    cell_range: CellRange | None = None
    tree_node_id: str | None = None  # Reference to source TreeNode

    # Cell specifications
    cell_specs: list[TemplateCellSpec] = Field(default_factory=list)

    # Structural constraints
    constraints: list[TemplateConstraint] = Field(default_factory=list)

    # Children for hierarchical templates
    children_ids: list[str] = Field(default_factory=list)

    # Summary stats
    total_constants: int = 0
    total_dofs: int = 0
    total_formulas: int = 0

    def finalize(self, workbook_hash: str) -> None:
        """Compute stable ID and summary stats."""
        range_str = self.cell_range.to_a1() if self.cell_range else ""
        self.template_node_id = compute_hash(
            workbook_hash, self.sheet_name, range_str,
        )
        self.total_constants = sum(
            1 for s in self.cell_specs if s.dof_type == DOFType.CONSTANT
        )
        self.total_dofs = sum(
            1 for s in self.cell_specs if s.dof_type == DOFType.DOF
        )
        self.total_formulas = sum(
            1 for s in self.cell_specs if s.dof_type == DOFType.FORMULA
        )


class DOFConflict(StableModel):
    """A conflict found during multi-document template comparison."""

    model_config = {"frozen": True, "extra": "forbid"}

    coord: CellCoord
    sheet_name: str = ""
    values_found: list[str] = Field(default_factory=list)
    resolution: str = ""  # "promoted_to_dof", "majority_constant", etc.


class GeneralizedTemplate(StableModel):
    """
    A generalized template produced by comparing multiple documents.

    Contains the most general template that accommodates all training
    documents, with DOFs added where conflicts were found.
    """

    model_config = {"frozen": False, "extra": "forbid"}

    template_id: str = Field(default="")
    source_files: list[str] = Field(default_factory=list)
    template_nodes: list[TemplateNode] = Field(default_factory=list)
    conflicts: list[DOFConflict] = Field(default_factory=list)

    total_constants: int = 0
    total_dofs: int = 0
    total_formulas: int = 0
    needs_reanalysis: bool = False  # True if too many DOFs

    def finalize(self) -> None:
        """Compute summary stats."""
        self.total_constants = sum(n.total_constants for n in self.template_nodes)
        self.total_dofs = sum(n.total_dofs for n in self.template_nodes)
        self.total_formulas = sum(n.total_formulas for n in self.template_nodes)
        self.template_id = compute_hash(
            ",".join(sorted(self.source_files)),
            str(self.total_constants),
            str(self.total_dofs),
        )
