"""
Stage 4: Light Block Identification.

Detects sparse (light) blocks and associates them with nearby tables
using spatial proximity rules. Handles offset labels, footnotes,
and annotation rows.
"""

from __future__ import annotations

import logging

from models.block import BlockDTO
from models.common import BlockType, CellCoord, CellRange
from models.table_structure import TableRegion, TableRegionRole, TableStructure

logger = logging.getLogger(__name__)

# Density threshold: blocks below this are considered "light"
_DENSITY_THRESHOLD = 0.8

# Maximum distance (in rows/cols) for proximity association
_MAX_ASSOCIATION_DISTANCE = 3


class LightBlockDetector:
    """
    Detects sparse blocks and associates them with nearby tables.

    A "light block" is one with density < 0.8 (less than 80% of cells
    in its bounding rectangle are filled). Light blocks are associated
    with the nearest table structure within a proximity threshold.
    """

    def detect_and_associate(
        self,
        blocks: list[BlockDTO],
        structures: list[TableStructure],
    ) -> tuple[list[BlockDTO], list[TableStructure]]:
        """
        Detect light blocks and associate them with nearby tables.

        Args:
            blocks: All blocks (after annotation-based splitting).
            structures: Table structures from Stage 3.

        Returns:
            Updated (blocks, structures) with light blocks identified
            and associated with nearby tables.
        """
        # Compute density for blocks that don't have it yet
        for block in blocks:
            if block.density is None:
                area = block.cell_range.row_count() * block.cell_range.col_count()
                block.density = round(
                    block.cell_count / area, 4
                ) if area > 0 else 0.0

        # Identify light blocks
        light_blocks = []
        solid_blocks = []
        for block in blocks:
            if block.table_name:
                # Excel ListObject tables are always solid
                solid_blocks.append(block)
            elif block.density is not None and block.density < _DENSITY_THRESHOLD:
                block.block_type = BlockType.LIGHT_BLOCK
                light_blocks.append(block)
            else:
                solid_blocks.append(block)

        if not light_blocks:
            logger.info("No light blocks detected")
            return blocks, structures

        # Associate each light block with the nearest table
        for light_block in light_blocks:
            best_structure = None
            best_distance = float("inf")

            for structure in structures:
                if structure.overall_range is None:
                    continue
                dist = self._manhattan_distance(
                    light_block.cell_range, structure.overall_range
                )
                if dist < best_distance and dist <= _MAX_ASSOCIATION_DISTANCE:
                    best_distance = dist
                    best_structure = structure

            if best_structure is not None:
                # Classify the light block's role
                role = self._classify_light_block_role(
                    light_block, best_structure
                )
                # Add as a region to the table structure
                best_structure.regions.append(TableRegion(
                    role=role,
                    cell_range=light_block.cell_range,
                    source_block_id=light_block.block_id,
                ))
                best_structure.source_block_ids.append(light_block.block_id)
                # Update overall range
                all_ranges = [r.cell_range for r in best_structure.regions]
                best_structure.overall_range = self._merge_ranges(all_ranges)

                light_block.table_structure_id = best_structure.structure_id

        logger.info(
            "Detected %d light blocks, associated %d with tables",
            len(light_blocks),
            sum(1 for b in light_blocks if b.table_structure_id),
        )
        return blocks, structures

    def _manhattan_distance(self, a: CellRange, b: CellRange) -> int:
        """
        Compute Manhattan distance between nearest edges of two ranges.
        Returns 0 if ranges overlap or are adjacent.
        """
        row_gap = max(
            0,
            max(a.top_left.row - b.bottom_right.row, b.top_left.row - a.bottom_right.row) - 1,
        )
        col_gap = max(
            0,
            max(a.top_left.col - b.bottom_right.col, b.top_left.col - a.bottom_right.col) - 1,
        )
        return row_gap + col_gap

    def _classify_light_block_role(
        self, block: BlockDTO, structure: TableStructure
    ) -> TableRegionRole:
        """
        Classify a light block's role relative to a table structure.

        - Below the table body → footer (footnote)
        - Above the table body → header (offset header)
        - Left of the table body → row_label
        - Otherwise → footer (default for annotations)
        """
        body = structure.body_region
        if body is None:
            return TableRegionRole.FOOTER

        body_range = body.cell_range

        # Check if block is below the body
        if block.cell_range.top_left.row > body_range.bottom_right.row:
            return TableRegionRole.FOOTER

        # Check if block is above the body
        if block.cell_range.bottom_right.row < body_range.top_left.row:
            return TableRegionRole.HEADER

        # Check if block is left of the body
        if block.cell_range.bottom_right.col < body_range.top_left.col:
            return TableRegionRole.ROW_LABEL

        return TableRegionRole.FOOTER

    @staticmethod
    def _merge_ranges(ranges: list[CellRange]) -> CellRange:
        """Compute the bounding range of multiple ranges."""
        min_row = min(r.top_left.row for r in ranges)
        min_col = min(r.top_left.col for r in ranges)
        max_row = max(r.bottom_right.row for r in ranges)
        max_col = max(r.bottom_right.col for r in ranges)
        return CellRange(
            top_left=CellCoord(row=min_row, col=min_col),
            bottom_right=CellCoord(row=max_row, col=max_col),
        )
