"""
Stage 9: Multi-Document DOF Comparison.

Compares templates from multiple documents to find the most general
template. Uses cell-level alignment to detect conflicts between
documents and promotes conflicting constants to DOFs.
"""

from __future__ import annotations

import logging
from collections import defaultdict

from ..models.common import CellCoord, compute_hash
from ..models.template import (
    DOFConflict,
    DOFType,
    GeneralizedTemplate,
    TemplateCellSpec,
    TemplateConstraint,
    TemplateNode,
)

logger = logging.getLogger(__name__)

# DOF threshold: if total DOFs exceed this, the template needs re-analysis
_DEFAULT_DOF_THRESHOLD = 50


class TemplateComparator:
    """
    Compares templates from multiple documents to produce a generalized template.

    For each cell position across all input templates:
    - If all documents agree on CONSTANT with same value → CONSTANT
    - If all documents agree on FORMULA with same formula → FORMULA
    - If any documents disagree → promote to DOF

    Tracks all conflicts for audit trail.
    """

    def __init__(self, dof_threshold: int = _DEFAULT_DOF_THRESHOLD):
        self._dof_threshold = dof_threshold

    def compare(
        self,
        template_sets: list[tuple[str, list[TemplateNode]]],
    ) -> GeneralizedTemplate:
        """
        Compare templates from multiple documents.

        Args:
            template_sets: List of (source_file, template_nodes) tuples.

        Returns:
            GeneralizedTemplate with merged cell specs and conflict records.
        """
        if not template_sets:
            return GeneralizedTemplate()

        source_files = [s[0] for s in template_sets]

        if len(template_sets) == 1:
            # Single document: generalized template is the same
            return GeneralizedTemplate(
                source_files=source_files,
                template_nodes=template_sets[0][1],
            )

        # Align templates by sheet name and cell range
        # Group template nodes by (sheet_name, range_a1)
        node_groups: dict[str, list[tuple[str, TemplateNode]]] = defaultdict(list)
        for source_file, nodes in template_sets:
            for node in nodes:
                key = f"{node.sheet_name}|{node.cell_range.to_a1() if node.cell_range else '?'}"
                node_groups[key].append((source_file, node))

        merged_nodes: list[TemplateNode] = []
        all_conflicts: list[DOFConflict] = []

        for group_key, group in node_groups.items():
            merged_node, conflicts = self._merge_template_nodes(group)
            merged_nodes.append(merged_node)
            all_conflicts.extend(conflicts)

        result = GeneralizedTemplate(
            source_files=source_files,
            template_nodes=merged_nodes,
            conflicts=all_conflicts,
        )
        result.finalize()

        # Check DOF threshold
        if result.total_dofs > self._dof_threshold:
            result.needs_reanalysis = True
            logger.warning(
                "Generalized template has %d DOFs (threshold: %d), re-analysis recommended",
                result.total_dofs,
                self._dof_threshold,
            )

        logger.info(
            "Compared %d documents: %d constants, %d DOFs, %d formulas, %d conflicts",
            len(template_sets),
            result.total_constants,
            result.total_dofs,
            result.total_formulas,
            len(all_conflicts),
        )
        return result

    def _merge_template_nodes(
        self,
        group: list[tuple[str, TemplateNode]],
    ) -> tuple[TemplateNode, list[DOFConflict]]:
        """Merge multiple template nodes from different documents."""
        # Use first node as base
        base_source, base = group[0]
        conflicts: list[DOFConflict] = []

        # Index specs by coord
        spec_by_coord: dict[str, list[tuple[str, TemplateCellSpec]]] = defaultdict(list)
        for source, node in group:
            for spec in node.cell_specs:
                coord_key = f"{spec.coord.row},{spec.coord.col}"
                spec_by_coord[coord_key].append((source, spec))

        # Merge each cell position
        merged_specs: list[TemplateCellSpec] = []
        for coord_key, specs in spec_by_coord.items():
            merged_spec, conflict = self._merge_cell_specs(specs, base.sheet_name)
            merged_specs.append(merged_spec)
            if conflict:
                conflicts.append(conflict)

        # Also detect cells that appear in some docs but not others
        all_coords = set()
        doc_coords: dict[str, set[str]] = defaultdict(set)
        for source, node in group:
            for spec in node.cell_specs:
                ck = f"{spec.coord.row},{spec.coord.col}"
                all_coords.add(ck)
                doc_coords[source].add(ck)

        for ck in all_coords:
            if any(ck not in dc for dc in doc_coords.values()):
                if ck not in {f"{s.coord.row},{s.coord.col}" for s in merged_specs}:
                    # Cell exists in some docs but not all → DOF
                    row, col = map(int, ck.split(","))
                    merged_specs.append(TemplateCellSpec(
                        coord=CellCoord(row=row, col=col),
                        dof_type=DOFType.DOF,
                        is_flexible=True,
                    ))
                    conflicts.append(DOFConflict(
                        coord=CellCoord(row=row, col=col),
                        sheet_name=base.sheet_name,
                        values_found=["present", "absent"],
                        resolution="promoted_to_dof",
                    ))

        # Merge constraints (take maximums/minimums)
        merged_constraints = self._merge_constraints(
            [node.constraints for _, node in group]
        )

        result = TemplateNode(
            sheet_name=base.sheet_name,
            cell_range=base.cell_range,
            cell_specs=merged_specs,
            constraints=merged_constraints,
        )
        result.finalize("")

        return result, conflicts

    def _merge_cell_specs(
        self,
        specs: list[tuple[str, TemplateCellSpec]],
        sheet_name: str,
    ) -> tuple[TemplateCellSpec, DOFConflict | None]:
        """Merge cell specs from multiple documents."""
        _, first = specs[0]
        coord = first.coord

        if len(specs) == 1:
            return first, None

        # Check for type consensus
        types = {s.dof_type for _, s in specs}

        if len(types) == 1 and DOFType.CONSTANT in types:
            # All agree on CONSTANT - check if values match
            values = {
                str(s.constant_value) if s.constant_value is not None else ""
                for _, s in specs
            }
            if len(values) == 1:
                return first, None

            # Values differ → promote to DOF
            conflict = DOFConflict(
                coord=coord,
                sheet_name=sheet_name,
                values_found=sorted(values),
                resolution="promoted_to_dof",
            )
            return TemplateCellSpec(
                coord=coord,
                dof_type=DOFType.DOF,
                is_flexible=True,
            ), conflict

        if len(types) == 1 and DOFType.FORMULA in types:
            # All agree on FORMULA - check if formulas match
            formulas = {s.formula_text or "" for _, s in specs}
            if len(formulas) == 1:
                return first, None

            # Formulas differ → keep as formula with first formula
            return first, None

        if DOFType.DOF in types:
            # Any DOF → all DOF
            return TemplateCellSpec(
                coord=coord,
                dof_type=DOFType.DOF,
                is_flexible=True,
            ), None

        # Mixed types → promote to DOF
        conflict = DOFConflict(
            coord=coord,
            sheet_name=sheet_name,
            values_found=[f"{s.dof_type.value}" for _, s in specs],
            resolution="promoted_to_dof",
        )
        return TemplateCellSpec(
            coord=coord,
            dof_type=DOFType.DOF,
            is_flexible=True,
        ), conflict

    def _merge_constraints(
        self,
        constraint_sets: list[list[TemplateConstraint]],
    ) -> list[TemplateConstraint]:
        """Merge constraints from multiple template nodes."""
        by_type: dict[str, list[TemplateConstraint]] = defaultdict(list)
        for constraints in constraint_sets:
            for c in constraints:
                by_type[c.constraint_type].append(c)

        merged = []
        for ctype, constraints in by_type.items():
            expected_values = [c.expected_value for c in constraints if c.expected_value is not None]
            if expected_values:
                min_val = min(expected_values)
                max_val = max(expected_values)
                if min_val == max_val:
                    merged.append(TemplateConstraint(
                        constraint_type=ctype,
                        expected_value=min_val,
                        description=constraints[0].description,
                    ))
                else:
                    merged.append(TemplateConstraint(
                        constraint_type=ctype,
                        min_value=min_val,
                        max_value=max_val,
                        description=f"{ctype} ranges from {min_val} to {max_val}",
                    ))
        return merged
