"""
Stage 6: Pattern Table Splitting.

Splits tables with non-prime body dimensions by detecting repeating
label patterns. When labels repeat at regular intervals, the table
is split into sub-tables and variable labels are reclassified as data.
"""

from __future__ import annotations

import logging
import math
from collections import defaultdict

from ..models.block import BlockDTO
from ..models.common import BlockType, CellAnnotation, CellCoord, CellRange
from ..models.sheet import SheetDTO
from ..models.table_structure import TableRegion, TableRegionRole, TableStructure

logger = logging.getLogger(__name__)

# Minimum stability score to consider a pattern as repeating
_STABILITY_THRESHOLD = 0.8

# Minimum body rows to attempt pattern splitting
_MIN_BODY_ROWS = 4


class PatternSplitter:
    """
    Splits tables at repeating label patterns.

    For tables with non-prime body row counts, checks if label columns
    repeat at regular factor intervals. If so, splits the table into
    sub-tables. Labels that differ between periods are reclassified
    from LABEL to DATA (they become degrees of freedom).
    """

    def __init__(self, sheet: SheetDTO):
        self._sheet = sheet

    def split(
        self,
        blocks: list[BlockDTO],
        structures: list[TableStructure],
    ) -> tuple[list[BlockDTO], list[TableStructure]]:
        """
        Attempt pattern-based splitting on all table structures.

        Returns:
            Updated (blocks, structures) with split tables added.
        """
        new_blocks: list[BlockDTO] = []
        new_structures: list[TableStructure] = []
        split_count = 0

        for structure in structures:
            body = structure.body_region
            if body is None:
                continue

            body_rows = body.cell_range.row_count()
            if body_rows < _MIN_BODY_ROWS:
                continue

            # Get non-trivial factors
            factors = self._get_factors(body_rows)
            if not factors:
                continue

            # Try each factor to find repeating label patterns
            best_factor = None
            best_stability = 0.0

            for factor in factors:
                stability = self._compute_label_stability(
                    body.cell_range, factor
                )
                if stability > best_stability:
                    best_stability = stability
                    best_factor = factor

            if best_factor is not None and best_stability >= _STABILITY_THRESHOLD:
                # Split the table
                sub_structures, sub_blocks = self._split_at_factor(
                    structure, body.cell_range, best_factor
                )
                new_structures.extend(sub_structures)
                new_blocks.extend(sub_blocks)
                split_count += 1

                # Reclassify variable labels as data
                self._reclassify_variable_labels(body.cell_range, best_factor)

        blocks.extend(new_blocks)
        structures.extend(new_structures)

        if split_count:
            logger.info(
                "Sheet '%s': split %d tables by repeating label patterns",
                self._sheet.sheet_name,
                split_count,
            )

        return blocks, structures

    def _get_factors(self, n: int) -> list[int]:
        """Get non-trivial factors of n (excluding 1 and n)."""
        factors = []
        for i in range(2, int(math.sqrt(n)) + 1):
            if n % i == 0:
                factors.append(i)
                if i != n // i:
                    factors.append(n // i)
        return sorted(factors)

    def _compute_label_stability(
        self, body_range: CellRange, period: int
    ) -> float:
        """
        Compute how stable label column values are across periods.

        Checks the first column of the body. For each period boundary,
        compares the label values to the first period. Score is the
        fraction of positions where values match across all periods.

        Returns:
            Stability score from 0.0 (all different) to 1.0 (all same).
        """
        start_row = body_range.top_left.row
        first_col = body_range.top_left.col
        total_rows = body_range.row_count()
        num_periods = total_rows // period

        if num_periods < 2:
            return 0.0

        # Get label values for first period
        first_period = []
        for offset in range(period):
            cell = self._sheet.get_cell(start_row + offset, first_col)
            val = str(cell.raw_value).strip().lower() if cell and cell.raw_value is not None else ""
            first_period.append(val)

        # Compare with subsequent periods
        matches = 0
        total = 0
        for p in range(1, num_periods):
            for offset in range(period):
                row = start_row + p * period + offset
                cell = self._sheet.get_cell(row, first_col)
                val = str(cell.raw_value).strip().lower() if cell and cell.raw_value is not None else ""
                total += 1
                if val == first_period[offset]:
                    matches += 1

        return matches / total if total > 0 else 0.0

    def _split_at_factor(
        self,
        structure: TableStructure,
        body_range: CellRange,
        factor: int,
    ) -> tuple[list[TableStructure], list[BlockDTO]]:
        """Split a table body at the given factor interval."""
        start_row = body_range.top_left.row
        num_sub = body_range.row_count() // factor
        sub_structures = []
        sub_blocks = []

        for i in range(num_sub):
            sub_start = start_row + i * factor
            sub_end = sub_start + factor - 1
            sub_range = CellRange(
                top_left=CellCoord(row=sub_start, col=body_range.top_left.col),
                bottom_right=CellCoord(row=sub_end, col=body_range.bottom_right.col),
            )

            # Create sub-block
            cell_count = 0
            formula_count = 0
            for r in range(sub_start, sub_end + 1):
                for c in range(body_range.top_left.col, body_range.bottom_right.col + 1):
                    cell = self._sheet.get_cell(r, c)
                    if cell:
                        cell_count += 1
                        if cell.formula:
                            formula_count += 1

            sub_block = BlockDTO(
                block_index=0,  # Will be re-indexed
                sheet_name=self._sheet.sheet_name,
                block_type=BlockType.TABLE,
                cell_range=sub_range,
                cell_count=cell_count,
                formula_count=formula_count,
            )
            sub_blocks.append(sub_block)

            # Create sub-structure with header from parent
            regions = [TableRegion(
                role=TableRegionRole.BODY,
                cell_range=sub_range,
                source_block_id=sub_block.block_id,
            )]

            # Inherit header from parent if available
            if structure.header_region:
                regions.append(TableRegion(
                    role=TableRegionRole.HEADER,
                    cell_range=structure.header_region.cell_range,
                    source_block_id=structure.header_region.source_block_id,
                ))

            sub_structure = TableStructure(
                sheet_name=self._sheet.sheet_name,
                regions=regions,
                source_block_ids=[sub_block.block_id],
                overall_range=sub_range,
            )
            sub_structures.append(sub_structure)

        return sub_structures, sub_blocks

    def _reclassify_variable_labels(
        self, body_range: CellRange, period: int
    ) -> None:
        """Reclassify label cells that vary between periods as DATA."""
        start_row = body_range.top_left.row
        first_col = body_range.top_left.col
        num_periods = body_range.row_count() // period

        if num_periods < 2:
            return

        # Find which positions have variable values
        for offset in range(period):
            values = set()
            for p in range(num_periods):
                row = start_row + p * period + offset
                cell = self._sheet.get_cell(row, first_col)
                if cell and cell.raw_value is not None:
                    values.add(str(cell.raw_value).strip().lower())

            # If values differ, reclassify all cells at this position as DATA
            if len(values) > 1:
                for p in range(num_periods):
                    row = start_row + p * period + offset
                    cell = self._sheet.get_cell(row, first_col)
                    if cell and cell.annotation == CellAnnotation.LABEL:
                        cell.annotation = CellAnnotation.DATA
                        cell.annotation_confidence = 0.5  # Moderate confidence
