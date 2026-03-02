"""
Layout segmentation algorithm.

Identifies logical "blocks" within a worksheet by analyzing cell content,
formatting patterns, blank row/column gaps, merged regions, borders,
and Excel table definitions. Produces BlockDTO objects with bounding
coordinates and semantic type classifications.

Algorithm overview:
1. Overlay known Excel tables as pre-defined blocks.
2. Find connected components of non-empty cells (using blank row/col gaps as separators).
3. For each connected component, classify its type by heuristics:
   - Has formulas referencing other cells → calculation_block
   - All text, no formulas, bold headers → text_block or header
   - Contains named ranges like "assumptions" → assumptions_table
   - Has bold/colored output cells with formulas → results_block
4. Split large blocks at internal blank rows/cols if they appear to contain
   multiple logical sections.
"""

from __future__ import annotations

import logging
from collections import defaultdict

from ..models.block import BlockDTO
from ..models.common import BlockType, CellCoord, CellRange
from ..models.sheet import SheetDTO
from ..models.table import TableDTO

logger = logging.getLogger(__name__)

# Minimum gap (in rows/cols) to consider a boundary between blocks
DEFAULT_GAP_ROWS = 1
DEFAULT_GAP_COLS = 1


class LayoutSegmenter:
    """
    Segments a worksheet into logical blocks.

    Uses a combination of Excel table definitions, blank row/column
    gap detection, style continuity analysis, and content heuristics
    to identify coherent regions of a sheet.
    """

    def __init__(
        self,
        sheet: SheetDTO,
        tables: list[TableDTO] | None = None,
        named_range_names: list[str] | None = None,
        gap_rows: int = DEFAULT_GAP_ROWS,
        gap_cols: int = DEFAULT_GAP_COLS,
    ):
        """
        Args:
            sheet: The parsed SheetDTO to segment.
            tables: Excel table definitions on this sheet.
            named_range_names: Named ranges overlapping this sheet.
            gap_rows: Number of consecutive blank rows to split on.
            gap_cols: Number of consecutive blank columns to split on.
        """
        self._sheet = sheet
        self._tables = [t for t in (tables or []) if t.sheet_name == sheet.sheet_name]
        self._named_ranges = named_range_names or []
        self._gap_rows = gap_rows
        self._gap_cols = gap_cols

    def segment(self) -> list[BlockDTO]:
        """
        Segment the sheet into blocks.

        Returns:
            Ordered list of BlockDTO objects covering the sheet's used range.
        """
        if not self._sheet.cells:
            return []

        used = self._sheet.used_range or self._sheet.compute_used_range()
        if not used:
            return []

        # Stage 0 enhancement: compute adaptive gap thresholds
        adaptive_row_gap, adaptive_col_gap = self._compute_adaptive_gaps(used)

        blocks: list[BlockDTO] = []

        # Step 1: Create blocks from Excel table definitions
        table_ranges: list[CellRange] = []
        for table in self._tables:
            block = self._table_to_block(table, len(blocks))
            blocks.append(block)
            table_ranges.append(table.ref_range)

        # Step 2: Find connected components from remaining cells
        non_table_cells = self._cells_outside_ranges(table_ranges)
        if non_table_cells:
            components = self._find_connected_components(
                non_table_cells, adaptive_row_gap, adaptive_col_gap
            )
            # Step 2b: Split components at style boundaries
            refined_components = []
            for component in components:
                sub_components = self._detect_style_boundaries(component)
                refined_components.extend(sub_components)

            for component_cells in refined_components:
                block = self._classify_component(component_cells, len(blocks))
                blocks.append(block)

        # Sort blocks by position (top-left corner)
        blocks.sort(key=lambda b: (b.cell_range.top_left.row, b.cell_range.top_left.col))

        # Re-index after sorting
        for idx, block in enumerate(blocks):
            block.block_index = idx

        logger.info(
            "Sheet '%s' segmented into %d blocks",
            self._sheet.sheet_name,
            len(blocks),
        )
        return blocks

    def _compute_adaptive_gaps(self, used: CellRange) -> tuple[int, int]:
        """
        Compute adaptive gap thresholds based on sheet density.

        Very dense sheets (>0.9) use slightly larger row gaps to avoid
        over-splitting tightly packed data. Column gaps are never increased
        because even single-column separators typically indicate real boundaries.
        """
        area = used.row_count() * used.col_count()
        if area == 0:
            return self._gap_rows, self._gap_cols

        density = len(self._sheet.cells) / area

        # Only increase row gap for extremely dense sheets
        if density > 0.9:
            row_gap = max(self._gap_rows, 2)
        else:
            row_gap = self._gap_rows

        # Never increase column gap - column separators are reliable boundaries
        col_gap = self._gap_cols

        return row_gap, col_gap

    def _detect_style_boundaries(self, cells: list) -> list[list]:
        """
        Split a component at persistent fill-color discontinuities.

        Only splits on fill/background color changes (not bold, which is
        expected for header rows). Requires the change to persist for 3+
        rows and both sides of the boundary must have 3+ rows to avoid
        splitting headers from their data.
        """
        if len(cells) <= 1:
            return [cells]

        # Group cells by row
        rows: dict[int, list] = defaultdict(list)
        for cell in cells:
            rows[cell.coord.row].append(cell)

        sorted_row_nums = sorted(rows.keys())
        if len(sorted_row_nums) <= 5:
            # Too few rows to meaningfully split by style
            return [cells]

        # Compute fill-only style signature per row (ignore bold)
        def _row_fill_sig(row_cells: list) -> str:
            parts = []
            for c in sorted(row_cells, key=lambda x: x.coord.col):
                fg = ""
                if c.style and c.style.fill and c.style.fill.fg_color:
                    fg = c.style.fill.fg_color
                parts.append(fg)
            return ";".join(parts)

        signatures = {r: _row_fill_sig(rows[r]) for r in sorted_row_nums}

        # Find split points: persistent fill color changes (3+ rows on each side)
        split_rows: list[int] = []
        for i in range(3, len(sorted_row_nums) - 2):
            curr_row = sorted_row_nums[i]
            prev_row = sorted_row_nums[i - 1]
            if signatures[curr_row] != signatures[prev_row]:
                # Verify persistence: check 2 more rows after the change
                next1 = sorted_row_nums[i + 1] if i + 1 < len(sorted_row_nums) else None
                next2 = sorted_row_nums[i + 2] if i + 2 < len(sorted_row_nums) else None
                if (
                    next1 is not None
                    and next2 is not None
                    and signatures.get(next1) == signatures[curr_row]
                    and signatures.get(next2) == signatures[curr_row]
                ):
                    split_rows.append(curr_row)

        if not split_rows:
            return [cells]

        # Split cells into groups at split rows
        components = []
        current_cells = []
        split_set = set(split_rows)
        for row_num in sorted_row_nums:
            if row_num in split_set and current_cells:
                components.append(current_cells)
                current_cells = []
            current_cells.extend(rows[row_num])
        if current_cells:
            components.append(current_cells)

        return [c for c in components if c]

    def segment_with_details(self) -> tuple[list[BlockDTO], list[list]]:
        """
        Segment the sheet and also return raw connected components.

        Returns:
            A tuple of (classified blocks, raw connected components).
            The raw components are the pre-classification cell lists
            for stages that need to inspect them directly.
        """
        blocks = self.segment()
        # Re-compute connected components for inspection
        table_ranges = [t.ref_range for t in self._tables]
        non_table_cells = self._cells_outside_ranges(table_ranges)
        components = self._find_connected_components(non_table_cells) if non_table_cells else []
        return blocks, components

    def _table_to_block(self, table: TableDTO, index: int) -> BlockDTO:
        """Convert an Excel table definition into a BlockDTO."""
        cells_in_range = self._count_cells_in_range(table.ref_range)
        formula_count = self._count_formulas_in_range(table.ref_range)
        has_merges = self._has_merges_in_range(table.ref_range)

        return BlockDTO(
            block_index=index,
            sheet_name=self._sheet.sheet_name,
            block_type=BlockType.TABLE,
            cell_range=table.ref_range,
            bounding_box=self._sheet.compute_bounding_box(table.ref_range),
            cell_count=cells_in_range,
            formula_count=formula_count,
            has_merges=has_merges,
            has_formatting=True,
            table_name=table.table_name,
            named_ranges=self._overlapping_named_ranges(table.ref_range),
        )

    def _cells_outside_ranges(self, ranges: list[CellRange]) -> dict[str, object]:
        """Return cells not covered by any of the given ranges."""
        result = {}
        for key, cell in self._sheet.cells.items():
            coord = cell.coord
            inside = any(r.contains(coord) for r in ranges)
            if not inside:
                result[key] = cell
        return result

    def _find_connected_components(
        self,
        cells: dict,
        gap_rows: int | None = None,
        gap_cols: int | None = None,
    ) -> list[list]:
        """
        Find connected components of non-empty cells using blank row/col gaps.

        Two cells are in the same component if they are within `gap_rows`
        rows and `gap_cols` columns of each other (i.e., there is no gap
        of blank rows/cols between them that exceeds the threshold).
        """
        if not cells:
            return []

        effective_gap_rows = gap_rows if gap_rows is not None else self._gap_rows
        effective_gap_cols = gap_cols if gap_cols is not None else self._gap_cols

        # Build a set of occupied rows and columns
        occupied_rows: set[int] = set()
        occupied_cols: set[int] = set()
        for cell in cells.values():
            occupied_rows.add(cell.coord.row)
            occupied_cols.add(cell.coord.col)

        # Find row gaps: stretches of empty rows that split the sheet
        sorted_rows = sorted(occupied_rows)
        row_groups: list[set[int]] = []
        current_group: set[int] = {sorted_rows[0]}
        for i in range(1, len(sorted_rows)):
            gap = sorted_rows[i] - sorted_rows[i - 1] - 1
            if gap >= effective_gap_rows:
                row_groups.append(current_group)
                current_group = set()
            current_group.add(sorted_rows[i])
        row_groups.append(current_group)

        # Within each row group, find column gaps
        components: list[list] = []
        for row_group in row_groups:
            # Get cells in this row group
            group_cells = [
                c for c in cells.values() if c.coord.row in row_group
            ]
            if not group_cells:
                continue

            # Find column groups within this row group
            group_cols = sorted({c.coord.col for c in group_cells})
            col_groups: list[set[int]] = []
            current_cols: set[int] = {group_cols[0]}
            for i in range(1, len(group_cols)):
                gap = group_cols[i] - group_cols[i - 1] - 1
                if gap >= effective_gap_cols:
                    col_groups.append(current_cols)
                    current_cols = set()
                current_cols.add(group_cols[i])
            col_groups.append(current_cols)

            for col_group in col_groups:
                component = [
                    c for c in group_cells if c.coord.col in col_group
                ]
                if component:
                    components.append(component)

        return components

    def _classify_component(self, cells: list, index: int) -> BlockDTO:
        """
        Classify a connected component of cells into a block type.

        Heuristics:
        - If >50% cells have formulas → calculation_block
        - If first row is bold/merged and rest are values → table (without ListObject)
        - If cell values contain keywords like "assumption" → assumptions_table
        - If has bold output cells with formatting emphasis → results_block
        - Otherwise → mixed or text_block
        """
        # Compute bounding range
        min_row = min(c.coord.row for c in cells)
        max_row = max(c.coord.row for c in cells)
        min_col = min(c.coord.col for c in cells)
        max_col = max(c.coord.col for c in cells)

        cell_range = CellRange(
            top_left=CellCoord(row=min_row, col=min_col),
            bottom_right=CellCoord(row=max_row, col=max_col),
        )

        formula_count = sum(1 for c in cells if c.formula)
        total = len(cells)
        has_merges = any(c.is_merged_master or c.is_merged_slave for c in cells)

        # Check for bold first row (header heuristic)
        first_row_cells = [c for c in cells if c.coord.row == min_row]
        has_bold_header = any(
            c.style and c.style.font and c.style.font.bold
            for c in first_row_cells
        )

        # Check for emphasized output cells (bold, colored)
        key_cells = []
        for c in cells:
            if c.style and c.style.font:
                if c.style.font.bold or (c.style.fill and c.style.fill.fg_color):
                    key_cells.append(c.coord)

        # Check for assumption-related keywords
        has_assumption_keyword = any(
            isinstance(c.raw_value, str) and any(
                kw in c.raw_value.lower()
                for kw in ("assumption", "input", "parameter", "scenario")
            )
            for c in cells
        )

        # Classify
        block_type = BlockType.MIXED
        row_span = max_row - min_row + 1

        if row_span == 1 and has_bold_header and not formula_count:
            block_type = BlockType.HEADER
        elif has_assumption_keyword and formula_count < total * 0.3:
            block_type = BlockType.ASSUMPTIONS_TABLE
        elif formula_count > total * 0.5:
            if key_cells and has_bold_header:
                block_type = BlockType.RESULTS_BLOCK
            else:
                block_type = BlockType.CALCULATION_BLOCK
        elif has_bold_header and total > 3:
            block_type = BlockType.TABLE
        elif all(
            isinstance(c.raw_value, str) or c.raw_value is None
            for c in cells
        ):
            block_type = BlockType.TEXT_BLOCK

        has_formatting = any(c.style is not None for c in cells)

        return BlockDTO(
            block_index=index,
            sheet_name=self._sheet.sheet_name,
            block_type=block_type,
            cell_range=cell_range,
            bounding_box=self._sheet.compute_bounding_box(cell_range),
            cell_count=total,
            formula_count=formula_count,
            has_merges=has_merges,
            has_formatting=has_formatting,
            key_cells=key_cells[:20],  # Limit to prevent huge lists
            named_ranges=self._overlapping_named_ranges(cell_range),
        )

    def _count_cells_in_range(self, rng: CellRange) -> int:
        """Count non-empty cells within a range."""
        count = 0
        for cell in self._sheet.cells.values():
            if rng.contains(cell.coord):
                count += 1
        return count

    def _count_formulas_in_range(self, rng: CellRange) -> int:
        """Count cells with formulas within a range."""
        count = 0
        for cell in self._sheet.cells.values():
            if rng.contains(cell.coord) and cell.formula:
                count += 1
        return count

    def _has_merges_in_range(self, rng: CellRange) -> bool:
        """Check if any merged regions overlap with the given range."""
        for merge in self._sheet.merged_regions:
            if self._ranges_overlap(rng, merge.range):
                return True
        return False

    def _overlapping_named_ranges(self, rng: CellRange) -> list[str]:
        """Find named ranges that overlap with the given range."""
        # In v1, we just return names that were passed in.
        # Full range intersection requires parsing named range refs.
        return list(self._named_ranges)

    @staticmethod
    def _ranges_overlap(a: CellRange, b: CellRange) -> bool:
        """Check if two cell ranges overlap."""
        return not (
            a.bottom_right.row < b.top_left.row
            or a.top_left.row > b.bottom_right.row
            or a.bottom_right.col < b.top_left.col
            or a.top_left.col > b.bottom_right.col
        )
