"""
Stage 3: Solid Table Identification Pass 1.

Groups label blocks near data blocks into table structures with explicit
header-body-footer relationships using spatial proximity rules.
"""

from __future__ import annotations

import logging

from models.block import BlockDTO
from models.common import BlockType, CellCoord, CellRange, compute_hash
from models.sheet import SheetDTO
from models.table_structure import TableRegion, TableRegionRole, TableStructure

logger = logging.getLogger(__name__)

# Maximum gap (in rows/cols) for label-data association
_MAX_HEADER_GAP = 0  # Header must be directly adjacent (0 blank rows between)
_MAX_FOOTER_GAP = 1  # Footer can be 1 row below
_MAX_ROW_LABEL_GAP = 0  # Row labels must be directly adjacent
_MIN_OVERLAP_RATIO = 0.5  # At least 50% column/row overlap required


class TableAssembler:
    """
    Assembles table structures from label and data blocks.

    Uses proximity rules to associate label blocks with nearby data blocks:
    - Label block directly above DATA block with overlapping columns → header
    - Label block directly left of DATA block with overlapping rows → row labels
    - Label block directly below DATA block → footer
    """

    def __init__(self, sheet: SheetDTO):
        self._sheet = sheet

    def assemble(self, blocks: list[BlockDTO]) -> list[TableStructure]:
        """
        Assemble table structures from blocks.

        Returns:
            List of TableStructure objects with header/body/footer regions.
        """
        # Separate label and data blocks
        label_blocks = [
            b for b in blocks
            if b.block_type in (BlockType.LABEL_BLOCK, BlockType.HEADER)
        ]
        data_blocks = [
            b for b in blocks
            if b.block_type in (
                BlockType.DATA_BLOCK, BlockType.TABLE,
                BlockType.CALCULATION_BLOCK, BlockType.ASSUMPTIONS_TABLE,
                BlockType.RESULTS_BLOCK, BlockType.MIXED,
            )
        ]

        if not data_blocks:
            return []

        used_labels: set[int] = set()  # Indices of label blocks already assigned
        structures: list[TableStructure] = []

        for data_block in data_blocks:
            regions: list[TableRegion] = []
            source_ids: list[str] = [data_block.block_id]

            # Body region from the data block
            regions.append(TableRegion(
                role=TableRegionRole.BODY,
                cell_range=data_block.cell_range,
                source_block_id=data_block.block_id,
            ))

            # Find header: label block directly above with column overlap
            for i, label in enumerate(label_blocks):
                if i in used_labels:
                    continue
                if self._is_header(label, data_block):
                    regions.append(TableRegion(
                        role=TableRegionRole.HEADER,
                        cell_range=label.cell_range,
                        source_block_id=label.block_id,
                    ))
                    source_ids.append(label.block_id)
                    used_labels.add(i)
                    break

            # Find row labels: label block directly left with row overlap
            for i, label in enumerate(label_blocks):
                if i in used_labels:
                    continue
                if self._is_row_label(label, data_block):
                    regions.append(TableRegion(
                        role=TableRegionRole.ROW_LABEL,
                        cell_range=label.cell_range,
                        source_block_id=label.block_id,
                    ))
                    source_ids.append(label.block_id)
                    used_labels.add(i)
                    break

            # Find footer: label block directly below with column overlap
            for i, label in enumerate(label_blocks):
                if i in used_labels:
                    continue
                if self._is_footer(label, data_block):
                    regions.append(TableRegion(
                        role=TableRegionRole.FOOTER,
                        cell_range=label.cell_range,
                        source_block_id=label.block_id,
                    ))
                    source_ids.append(label.block_id)
                    used_labels.add(i)
                    break

            # Compute overall range
            all_ranges = [r.cell_range for r in regions]
            overall = self._merge_ranges(all_ranges)

            structure = TableStructure(
                sheet_name=self._sheet.sheet_name,
                regions=regions,
                source_block_ids=source_ids,
                overall_range=overall,
            )
            structures.append(structure)

            # Update block references
            data_block.table_structure_id = structure.structure_id

        logger.info(
            "Sheet '%s': assembled %d table structures from %d data blocks",
            self._sheet.sheet_name,
            len(structures),
            len(data_blocks),
        )
        return structures

    def _is_header(self, label: BlockDTO, data: BlockDTO) -> bool:
        """Check if label block is a header for the data block."""
        # Label must end within _MAX_HEADER_GAP rows above data
        row_gap = data.cell_range.top_left.row - label.cell_range.bottom_right.row - 1
        if row_gap < 0 or row_gap > _MAX_HEADER_GAP:
            return False

        # Must have sufficient column overlap
        return self._col_overlap_ratio(label, data) >= _MIN_OVERLAP_RATIO

    def _is_row_label(self, label: BlockDTO, data: BlockDTO) -> bool:
        """Check if label block is a row label for the data block."""
        # Label must end within _MAX_ROW_LABEL_GAP cols left of data
        col_gap = data.cell_range.top_left.col - label.cell_range.bottom_right.col - 1
        if col_gap < 0 or col_gap > _MAX_ROW_LABEL_GAP:
            return False

        # Must have sufficient row overlap
        return self._row_overlap_ratio(label, data) >= _MIN_OVERLAP_RATIO

    def _is_footer(self, label: BlockDTO, data: BlockDTO) -> bool:
        """Check if label block is a footer for the data block."""
        # Label must start within _MAX_FOOTER_GAP rows below data
        row_gap = label.cell_range.top_left.row - data.cell_range.bottom_right.row - 1
        if row_gap < 0 or row_gap > _MAX_FOOTER_GAP:
            return False

        # Must have sufficient column overlap
        return self._col_overlap_ratio(label, data) >= _MIN_OVERLAP_RATIO

    @staticmethod
    def _col_overlap_ratio(a: BlockDTO, b: BlockDTO) -> float:
        """Compute column overlap ratio between two blocks."""
        overlap_start = max(a.cell_range.top_left.col, b.cell_range.top_left.col)
        overlap_end = min(a.cell_range.bottom_right.col, b.cell_range.bottom_right.col)
        overlap = max(0, overlap_end - overlap_start + 1)
        a_width = a.cell_range.col_count()
        return overlap / a_width if a_width > 0 else 0.0

    @staticmethod
    def _row_overlap_ratio(a: BlockDTO, b: BlockDTO) -> float:
        """Compute row overlap ratio between two blocks."""
        overlap_start = max(a.cell_range.top_left.row, b.cell_range.top_left.row)
        overlap_end = min(a.cell_range.bottom_right.row, b.cell_range.bottom_right.row)
        overlap = max(0, overlap_end - overlap_start + 1)
        a_height = a.cell_range.row_count()
        return overlap / a_height if a_height > 0 else 0.0

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
