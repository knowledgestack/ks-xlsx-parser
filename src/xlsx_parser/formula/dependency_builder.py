"""
Dependency graph builder.

Scans all cells with formulas across all sheets, parses their references,
and constructs a DependencyGraph with typed edges. Handles cross-sheet,
external, structured, and named range references.
"""

from __future__ import annotations

import logging
from typing import TYPE_CHECKING

from ..models.common import CellCoord, EdgeType
from ..models.dependency import DependencyEdgeDTO, DependencyGraph
from .formula_parser import FormulaParser

if TYPE_CHECKING:
    from ..models.sheet import SheetDTO
    from ..models.workbook import NamedRangeDTO

logger = logging.getLogger(__name__)


class DependencyBuilder:
    """
    Builds a formula dependency graph from parsed sheets.

    Iterates over all cells with formulas, parses their references
    using FormulaParser, and constructs DependencyEdgeDTO objects.
    After building, runs cycle detection to annotate circular refs.
    """

    def __init__(
        self,
        sheets: list[SheetDTO],
        named_ranges: list[NamedRangeDTO] | None = None,
    ):
        self._sheets = sheets
        self._named_ranges = {nr.name: nr for nr in (named_ranges or [])}
        self._parser = FormulaParser()

    def build(self) -> DependencyGraph:
        """
        Build the complete dependency graph.

        Returns:
            A DependencyGraph with all edges and indexes built,
            and circular references detected.
        """
        graph = DependencyGraph()

        for sheet in self._sheets:
            for cell in sheet.cells.values():
                if cell.formula:
                    self._process_formula(graph, cell.sheet_name, cell.coord, cell.formula)

        graph.build_indexes()

        # Detect circular references
        circular = graph.detect_circular_refs()
        if circular:
            logger.warning(
                "Detected %d cells in circular references: %s",
                len(circular),
                list(circular)[:10],
            )

        logger.info(
            "Dependency graph built: %d edges, %d circular refs",
            len(graph.edges),
            len(circular),
        )

        return graph

    def _process_formula(
        self,
        graph: DependencyGraph,
        source_sheet: str,
        source_coord: CellCoord,
        formula: str,
    ) -> None:
        """Parse a formula and add all its reference edges to the graph."""
        try:
            refs = self._parser.parse(formula, source_sheet)
        except Exception as e:
            logger.debug("Failed to parse formula '%s': %s", formula, e)
            return

        for ref in refs:
            edge_type = self._classify_edge(ref, source_sheet)

            edge = DependencyEdgeDTO(
                source_sheet=source_sheet,
                source_coord=source_coord,
                target_sheet=ref.sheet_name,
                target_coord=ref.coord,
                target_range=ref.range,
                target_ref_string=ref.ref_string,
                edge_type=edge_type,
                external_workbook=ref.external_workbook,
                named_range_name=ref.named_range,
            )
            graph.add_edge(edge)

    def _classify_edge(self, ref, source_sheet: str) -> EdgeType:
        """Determine the edge type based on the reference properties."""
        if ref.is_external:
            return EdgeType.EXTERNAL
        if ref.is_structured:
            return EdgeType.STRUCTURED_REF
        if ref.named_range:
            return EdgeType.NAMED_RANGE
        if ref.sheet_name and ref.sheet_name != source_sheet:
            return EdgeType.CROSS_SHEET
        if ref.range:
            return EdgeType.CELL_TO_RANGE
        return EdgeType.CELL_TO_CELL
