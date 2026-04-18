"""
Stage 5: Solid Table Pass 2 - Table Grouping.

Groups adjacent tables with similar structure into parent table structures,
adding one level of hierarchical abstraction. Uses structural similarity
scoring based on column count, header overlap, alignment, and formula patterns.
"""

from __future__ import annotations

import logging
from collections import defaultdict

from models.block import BlockDTO
from models.common import BlockType, CellCoord, CellRange, compute_hash
from models.sheet import SheetDTO
from models.table_structure import TableStructure

logger = logging.getLogger(__name__)

# Similarity threshold for grouping tables
_SIMILARITY_THRESHOLD = 0.7

# Component weights for similarity score
_W_COL_COUNT = 0.30
_W_HEADER_OVERLAP = 0.30
_W_ALIGNMENT = 0.20
_W_FORMULA_PATTERN = 0.20


class TableGrouper:
    """
    Groups structurally similar adjacent tables into parent table structures.

    Computes structural similarity between all table pairs and greedily
    merges pairs above the similarity threshold using Union-Find.
    """

    def __init__(self, sheet: SheetDTO):
        self._sheet = sheet

    def group_tables(
        self,
        blocks: list[BlockDTO],
        structures: list[TableStructure],
    ) -> tuple[list[BlockDTO], list[TableStructure]]:
        """
        Group similar adjacent tables into parent structures.

        Returns:
            Updated (blocks, structures) with parent blocks and structures added.
        """
        if len(structures) < 2:
            return blocks, structures

        # Compute pairwise similarity
        pairs: list[tuple[int, int, float]] = []
        for i in range(len(structures)):
            for j in range(i + 1, len(structures)):
                sim = self._structural_similarity(structures[i], structures[j])
                if sim >= _SIMILARITY_THRESHOLD:
                    pairs.append((i, j, sim))

        if not pairs:
            logger.info("No table groups found (no similar pairs)")
            return blocks, structures

        # Sort by similarity descending
        pairs.sort(key=lambda p: p[2], reverse=True)

        # Union-Find for grouping
        parent = list(range(len(structures)))

        def find(x: int) -> int:
            while parent[x] != x:
                parent[x] = parent[parent[x]]
                x = parent[x]
            return x

        def union(x: int, y: int) -> None:
            px, py = find(x), find(y)
            if px != py:
                parent[px] = py

        for i, j, sim in pairs:
            union(i, j)

        # Build groups
        groups: dict[int, list[int]] = defaultdict(list)
        for i in range(len(structures)):
            groups[find(i)].append(i)

        # Create parent blocks and structures for multi-member groups
        new_parent_blocks: list[BlockDTO] = []
        for root, members in groups.items():
            if len(members) < 2:
                continue

            member_structures = [structures[m] for m in members]
            member_block_ids = []
            for s in member_structures:
                member_block_ids.extend(s.source_block_ids)

            # Find member blocks
            member_blocks = [
                b for b in blocks if b.block_id in member_block_ids
            ]

            # Compute overall range
            all_ranges = [
                s.overall_range for s in member_structures
                if s.overall_range is not None
            ]
            if not all_ranges:
                continue

            overall_range = self._merge_ranges(all_ranges)

            # Create parent block
            parent_block = BlockDTO(
                block_index=len(blocks) + len(new_parent_blocks),
                sheet_name=self._sheet.sheet_name,
                block_type=BlockType.TABLE,
                cell_range=overall_range,
                bounding_box=self._sheet.compute_bounding_box(overall_range),
                cell_count=sum(b.cell_count for b in member_blocks),
                formula_count=sum(b.formula_count for b in member_blocks),
                child_block_ids=[b.block_id for b in member_blocks],
            )

            # Set parent references on child blocks
            for child in member_blocks:
                child.parent_block_id = parent_block.block_id

            new_parent_blocks.append(parent_block)

            # Create parent table structure
            parent_structure = TableStructure(
                sheet_name=self._sheet.sheet_name,
                regions=[r for s in member_structures for r in s.regions],
                source_block_ids=member_block_ids,
                overall_range=overall_range,
            )
            structures.append(parent_structure)

        blocks.extend(new_parent_blocks)

        logger.info(
            "Grouped %d tables into %d groups",
            len(structures),
            len([g for g in groups.values() if len(g) > 1]),
        )
        return blocks, structures

    def _structural_similarity(
        self, a: TableStructure, b: TableStructure
    ) -> float:
        """
        Compute structural similarity between two table structures.

        Components:
        - Column count match (30%): 1.0 if same, decreases with difference
        - Header text Jaccard overlap (30%): overlap of header cell values
        - Alignment (20%): same starting row or column
        - Formula pattern (20%): similar formula-to-cell ratio
        """
        if a.overall_range is None or b.overall_range is None:
            return 0.0

        # Column count similarity
        a_cols = a.overall_range.col_count()
        b_cols = b.overall_range.col_count()
        col_sim = 1.0 - abs(a_cols - b_cols) / max(a_cols, b_cols, 1)

        # Header overlap (Jaccard similarity)
        a_headers = self._get_header_values(a)
        b_headers = self._get_header_values(b)
        if a_headers or b_headers:
            intersection = len(a_headers & b_headers)
            union = len(a_headers | b_headers)
            header_sim = intersection / union if union > 0 else 0.0
        else:
            header_sim = 0.5  # No headers, neutral

        # Alignment: check if tables share start row or column
        alignment_sim = 0.0
        if a.overall_range.top_left.row == b.overall_range.top_left.row:
            alignment_sim += 0.5
        if a.overall_range.top_left.col == b.overall_range.top_left.col:
            alignment_sim += 0.5

        # Formula pattern: compare formula ratio in body regions
        a_ratio = self._formula_ratio(a)
        b_ratio = self._formula_ratio(b)
        formula_sim = 1.0 - abs(a_ratio - b_ratio)

        return (
            _W_COL_COUNT * col_sim
            + _W_HEADER_OVERLAP * header_sim
            + _W_ALIGNMENT * alignment_sim
            + _W_FORMULA_PATTERN * formula_sim
        )

    def _get_header_values(self, structure: TableStructure) -> set[str]:
        """Get the set of cell values in the header region."""
        header = structure.header_region
        if header is None:
            return set()

        values = set()
        rng = header.cell_range
        for row in range(rng.top_left.row, rng.bottom_right.row + 1):
            for col in range(rng.top_left.col, rng.bottom_right.col + 1):
                cell = self._sheet.get_cell(row, col)
                if cell and cell.raw_value is not None:
                    values.add(str(cell.raw_value).strip().lower())
        return values

    def _formula_ratio(self, structure: TableStructure) -> float:
        """Compute the ratio of formula cells in the body region."""
        body = structure.body_region
        if body is None:
            return 0.0

        rng = body.cell_range
        total = 0
        formulas = 0
        for row in range(rng.top_left.row, rng.bottom_right.row + 1):
            for col in range(rng.top_left.col, rng.bottom_right.col + 1):
                cell = self._sheet.get_cell(row, col)
                if cell:
                    total += 1
                    if cell.formula:
                        formulas += 1
        return formulas / total if total > 0 else 0.0

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
