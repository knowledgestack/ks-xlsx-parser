"""
Stage 2: Solid Block Identification by Annotation Type.

Splits spatial blocks into contiguous regions of the same annotation type
(DATA or LABEL), producing annotation-homogeneous blocks.
"""

from __future__ import annotations

import logging
from collections import defaultdict

from models.block import BlockDTO
from models.common import BlockType, CellAnnotation, CellCoord, CellRange
from models.sheet import SheetDTO

logger = logging.getLogger(__name__)


class BlockSplitter:
    """
    Splits blocks into annotation-homogeneous sub-blocks.

    After Stage 1 annotates cells as DATA or LABEL, this stage
    re-groups blocks so that each block contains only one annotation type.
    Uses flood-fill to find contiguous regions of same annotation.
    """

    def __init__(self, sheet: SheetDTO):
        self._sheet = sheet

    def split_blocks(self, blocks: list[BlockDTO]) -> list[BlockDTO]:
        """
        Split blocks into annotation-homogeneous sub-blocks.

        Blocks that are already homogeneous (all DATA or all LABEL) are kept.
        Mixed blocks are split using flood-fill on annotation type.
        Excel table blocks (with table_name) are preserved as-is.

        Returns:
            New list of blocks with annotation-based splitting applied.
        """
        result: list[BlockDTO] = []

        for block in blocks:
            # Preserve Excel ListObject table blocks
            if block.table_name:
                self._update_annotation_counts(block)
                result.append(block)
                continue

            # Get cells in this block
            cells_in_block = self._get_cells_in_range(block.cell_range)
            if not cells_in_block:
                result.append(block)
                continue

            # Check if block is already homogeneous
            annotations = {
                c.annotation for c in cells_in_block if c.annotation is not None
            }
            if len(annotations) <= 1:
                self._update_annotation_counts(block)
                result.append(block)
                continue

            # Split into contiguous same-annotation regions
            sub_blocks = self._flood_fill_split(cells_in_block, block.sheet_name)
            result.extend(sub_blocks)

        # Re-index
        for idx, block in enumerate(result):
            block.block_index = idx

        logger.info(
            "Sheet '%s': split %d blocks into %d annotation-homogeneous blocks",
            self._sheet.sheet_name,
            len(blocks),
            len(result),
        )
        return result

    def _flood_fill_split(
        self, cells: list, sheet_name: str
    ) -> list[BlockDTO]:
        """
        Split cells into contiguous same-annotation regions using flood fill.
        """
        # Build grid lookup
        cell_map: dict[tuple[int, int], object] = {}
        for cell in cells:
            cell_map[(cell.coord.row, cell.coord.col)] = cell

        visited: set[tuple[int, int]] = set()
        regions: list[list] = []

        for cell in cells:
            pos = (cell.coord.row, cell.coord.col)
            if pos in visited:
                continue

            annotation = cell.annotation or CellAnnotation.DATA
            region = []
            stack = [pos]

            while stack:
                current = stack.pop()
                if current in visited:
                    continue
                visited.add(current)

                current_cell = cell_map.get(current)
                if current_cell is None:
                    continue

                current_ann = current_cell.annotation or CellAnnotation.DATA
                if current_ann != annotation:
                    continue

                region.append(current_cell)

                # Check 4-connected neighbors
                r, c = current
                for nr, nc in [(r - 1, c), (r + 1, c), (r, c - 1), (r, c + 1)]:
                    if (nr, nc) not in visited and (nr, nc) in cell_map:
                        neighbor = cell_map[(nr, nc)]
                        neighbor_ann = neighbor.annotation or CellAnnotation.DATA
                        if neighbor_ann == annotation:
                            stack.append((nr, nc))

            if region:
                regions.append(region)

        # Convert regions to blocks
        blocks = []
        for region in regions:
            min_row = min(c.coord.row for c in region)
            max_row = max(c.coord.row for c in region)
            min_col = min(c.coord.col for c in region)
            max_col = max(c.coord.col for c in region)

            cell_range = CellRange(
                top_left=CellCoord(row=min_row, col=min_col),
                bottom_right=CellCoord(row=max_row, col=max_col),
            )

            annotation = region[0].annotation or CellAnnotation.DATA
            block_type = (
                BlockType.LABEL_BLOCK
                if annotation == CellAnnotation.LABEL
                else BlockType.DATA_BLOCK
            )

            label_count = sum(
                1 for c in region if c.annotation == CellAnnotation.LABEL
            )
            data_count = len(region) - label_count
            formula_count = sum(1 for c in region if c.formula)
            has_merges = any(c.is_merged_master or c.is_merged_slave for c in region)
            has_formatting = any(c.style is not None for c in region)

            key_cells = []
            for c in region:
                if c.style and c.style.font and c.style.font.bold:
                    key_cells.append(c.coord)

            # Compute density
            area = cell_range.row_count() * cell_range.col_count()
            density = len(region) / area if area > 0 else 0.0

            block = BlockDTO(
                block_index=len(blocks),
                sheet_name=sheet_name,
                block_type=block_type,
                cell_range=cell_range,
                bounding_box=self._sheet.compute_bounding_box(cell_range),
                cell_count=len(region),
                formula_count=formula_count,
                has_merges=has_merges,
                has_formatting=has_formatting,
                key_cells=key_cells[:20],
                density=round(density, 4),
                label_cell_count=label_count,
                data_cell_count=data_count,
            )
            blocks.append(block)

        return blocks

    def _get_cells_in_range(self, rng: CellRange) -> list:
        """Get all cells within a range."""
        result = []
        for cell in self._sheet.cells.values():
            if rng.contains(cell.coord):
                result.append(cell)
        return result

    def _update_annotation_counts(self, block: BlockDTO) -> None:
        """Update label/data counts on an existing block."""
        cells = self._get_cells_in_range(block.cell_range)
        block.label_cell_count = sum(
            1 for c in cells if c.annotation == CellAnnotation.LABEL
        )
        block.data_cell_count = sum(
            1 for c in cells if c.annotation == CellAnnotation.DATA
        )
        area = block.cell_range.row_count() * block.cell_range.col_count()
        block.density = round(block.cell_count / area, 4) if area > 0 else 0.0
