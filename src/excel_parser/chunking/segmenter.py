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

from excel_parser.analysis.header_detector import find_header_span
from excel_parser.models.block import BlockDTO
from excel_parser.models.common import BlockType, CellCoord, CellRange
from excel_parser.models.sheet import SheetDTO
from excel_parser.models.table import TableDTO

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
            for component_cells in components:
                block = self._classify_component(component_cells, len(blocks))
                blocks.append(block)

        # Sort blocks by position (top-left corner)
        blocks.sort(key=lambda b: (b.cell_range.top_left.row, b.cell_range.top_left.col))

        # Merge standalone title/header rows into the block that follows them
        blocks = self._merge_title_blocks(blocks)

        # Re-stitch table fragments severed by single-column "section/note"
        # rows and blank spacer rows, so a header is never split from its body.
        blocks = self._stitch_vertical_tables(blocks)

        # Re-sort: stitching can extend an anchor's top-left column leftward, so
        # re-establish positional order before assigning indices / nav pointers.
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

    # Maximum number of rows a standalone title/header block can have to be
    # eligible for merging into the following block.
    _MAX_TITLE_ROWS = 3

    def _merge_title_blocks(self, blocks: list[BlockDTO]) -> list[BlockDTO]:
        """
        Merge standalone title/header blocks into their following data block.

        A block is treated as a "title" when ALL of the following hold:
          - has ≤ _MAX_TITLE_ROWS rows,
          - has no formulas and no explicit Excel table name,
          - is classified as HEADER, TEXT_BLOCK, TABLE, or MIXED,
          - sits directly above the next block (≤1 blank row for HEADER,
            0 blank rows for other types),
          - the next block is strictly taller (more rows) than the title,
          - their column ranges overlap.

        When merged the following block's cell_range is extended upward to
        include the title rows.  Cell counts and key_cells are combined.
        """
        if len(blocks) <= 1:
            return blocks

        absorbed: set[int] = set()  # indices of title blocks consumed
        merged: list[BlockDTO] = list(blocks)  # mutable copy

        for i in range(len(merged) - 1):
            blk = merged[i]

            if i in absorbed:
                continue

            row_span = blk.cell_range.bottom_right.row - blk.cell_range.top_left.row + 1
            if row_span > self._MAX_TITLE_ROWS:
                continue

            # Must look like a title, not a real data table
            if blk.formula_count > 0:
                continue
            if blk.table_name:  # explicit Excel table — don't merge away
                continue

            is_header = blk.block_type == BlockType.HEADER
            is_other_title = blk.block_type in (
                BlockType.TEXT_BLOCK,
                BlockType.TABLE,
                BlockType.MIXED,
            )
            if not (is_header or is_other_title):
                continue

            # HEADER blocks tolerate 1 blank row; others must be contiguous
            max_gap = 1 if is_header else 0

            # Find the next non-absorbed block
            j = i + 1
            while j < len(merged) and j in absorbed:
                j += 1
            if j >= len(merged):
                continue

            nxt = merged[j]
            gap = nxt.cell_range.top_left.row - blk.cell_range.bottom_right.row - 1
            if gap > max_gap:
                continue

            # The next block must be taller — we merge small into large,
            # not two equally-sized neighbours
            nxt_rows = nxt.cell_range.bottom_right.row - nxt.cell_range.top_left.row + 1
            if nxt_rows <= row_span:
                continue

            # Column ranges must overlap
            if (blk.cell_range.bottom_right.col < nxt.cell_range.top_left.col
                    or blk.cell_range.top_left.col > nxt.cell_range.bottom_right.col):
                continue

            # Merge: extend nxt upward
            new_top_row = min(blk.cell_range.top_left.row, nxt.cell_range.top_left.row)
            new_top_col = min(blk.cell_range.top_left.col, nxt.cell_range.top_left.col)
            new_bot_row = max(blk.cell_range.bottom_right.row, nxt.cell_range.bottom_right.row)
            new_bot_col = max(blk.cell_range.bottom_right.col, nxt.cell_range.bottom_right.col)

            nxt.cell_range = CellRange(
                top_left=CellCoord(row=new_top_row, col=new_top_col),
                bottom_right=CellCoord(row=new_bot_row, col=new_bot_col),
            )
            nxt.cell_count += blk.cell_count
            nxt.has_merges = nxt.has_merges or blk.has_merges
            nxt.has_formatting = nxt.has_formatting or blk.has_formatting
            nxt.key_cells = (blk.key_cells + nxt.key_cells)[:20]
            nxt.bounding_box = self._sheet.compute_bounding_box(nxt.cell_range)

            absorbed.add(i)
            logger.debug(
                "Merged title block %d (%s) into block %d (%s)",
                i, blk.cell_range.to_a1(), j, nxt.cell_range.to_a1(),
            )

        return [b for idx, b in enumerate(merged) if idx not in absorbed]

    # Maximum blank-row gap (per hop) tolerated when stitching table fragments.
    _MAX_STITCH_GAP_ROWS = 1
    # Minimum column-overlap ratio for a fragment to count as a continuation body.
    _MIN_CONTINUATION_OVERLAP = 0.5
    # Block types that can anchor a vertical-stitch chain (data-bearing regions).
    _ANCHOR_TYPES = (
        BlockType.TABLE,
        BlockType.ASSUMPTIONS_TABLE,
        BlockType.RESULTS_BLOCK,
        BlockType.CALCULATION_BLOCK,
        BlockType.MIXED,
    )

    def _stitch_vertical_tables(self, blocks: list[BlockDTO]) -> list[BlockDTO]:
        """
        Stitch a table that was fragmented by single-column "section/note" rows
        and blank spacer rows back into one block.

        Blank rows make :meth:`_find_connected_components` split a single logical
        table into a header+top fragment, an isolated single-column row, and an
        orphaned body fragment — so the header ends up in a different chunk from
        the rest of the table.  :meth:`_merge_title_blocks` cannot repair this
        (it only folds a small block *down* into a taller one).

        Starting from each data-bearing "anchor" block, this pass walks downward
        and classifies each following block into one of three roles:

          - CONTINUATION: a multi-column fragment that re-aligns with the anchor's
            columns and has no header of its own → absorbed; the chain re-anchors.
          - BRIDGE: a narrow interstitial row (e.g. a single-column section label)
            → held pending; absorbed only if a continuation follows within budget.
          - STOP: a block with its own bold header (a genuine new table), a block
            too far below, or a chart/Excel-table block → ends the chain. Any
            pending bridges are left as separate blocks.

        A block in a non-overlapping column range (e.g. a far-offset footnote) is
        skipped over without ending the chain and without being absorbed.

        Blocks are assumed sorted by ``(top_left.row, top_left.col)``.  The anchor
        always survives and keeps its ``block_type``/``table_name``, so the
        renderer still treats the anchor's first row as the table header.
        """
        if len(blocks) <= 1:
            return blocks

        absorbed: set[int] = set()
        merged = list(blocks)

        for i in range(len(merged)):
            if i in absorbed:
                continue
            anchor = merged[i]
            if anchor.block_type not in self._ANCHOR_TYPES:
                continue
            # Only stitch when the anchor has a genuine column header to protect —
            # that is the header we must keep attached to its body. Without this
            # gate we'd merge independent stacked sections that each carry only a
            # single-column title (e.g. ASSUMPTIONS + RESULTS, or two tables that
            # each open with a merged title cell).
            if find_header_span(self._sheet, anchor.cell_range) is None:
                continue

            # Running bottom of the chain so far (advances over bridges too, so
            # per-hop gaps are measured correctly across blank spacer rows).
            last_bottom = anchor.cell_range.bottom_right.row
            pending: list[int] = []  # bridge indices held, not yet committed

            j = i + 1
            while j < len(merged):
                if j in absorbed:
                    j += 1
                    continue
                cand = merged[j]

                gap = cand.cell_range.top_left.row - last_bottom - 1
                if gap > self._MAX_STITCH_GAP_ROWS:
                    break  # too far below: end the chain (pending bridges stay)
                if cand.table_name or cand.block_type in (
                    BlockType.CHART_ANCHOR,
                    BlockType.IMAGE_ANCHOR,
                ):
                    break

                overlap = self._col_overlap_ratio(
                    anchor.cell_range, cand.cell_range
                )
                if overlap == 0.0:
                    # Far-offset block (e.g. a note in another column): leave it
                    # alone and keep scanning without ending the chain.
                    j += 1
                    continue

                # A genuine new table has its own column header. A bold
                # single-cell row (e.g. "REVENUE", "SUBTOTAL") is a section/
                # subtotal label that belongs inside the table — find_header_span
                # requires >=2 labelling cells, so a lone label never reads as a
                # header, and a continuation's data row never does either.
                if find_header_span(self._sheet, cand.cell_range) is not None:
                    break  # genuine new table with its own header

                is_continuation = (
                    cand.cell_range.col_count() >= 2
                    and overlap >= self._MIN_CONTINUATION_OVERLAP
                )
                if is_continuation:
                    for p in pending:
                        self._absorb(anchor, merged[p])
                        absorbed.add(p)
                    pending = []
                    self._absorb(anchor, cand)
                    absorbed.add(j)
                    last_bottom = anchor.cell_range.bottom_right.row
                else:
                    pending.append(j)
                    last_bottom = max(last_bottom, cand.cell_range.bottom_right.row)
                j += 1

            if absorbed:
                logger.debug(
                    "Stitched %d fragment(s) into anchor %s",
                    len(absorbed),
                    anchor.cell_range.to_a1(),
                )

        return [b for idx, b in enumerate(merged) if idx not in absorbed]

    def _absorb(self, anchor: BlockDTO, other: BlockDTO) -> None:
        """
        Fold ``other`` into ``anchor`` by extending the anchor's range to cover
        both.  The anchor survives and keeps its type/table_name.

        Unlike :meth:`_merge_title_blocks`, this sums ``formula_count`` —
        continuation bodies may contain formulas, and ``formula_count`` feeds
        ``BlockDTO.finalize``'s content hash and the dependency summary.
        """
        a = anchor.cell_range
        b = other.cell_range
        anchor.cell_range = CellRange(
            top_left=CellCoord(
                row=min(a.top_left.row, b.top_left.row),
                col=min(a.top_left.col, b.top_left.col),
            ),
            bottom_right=CellCoord(
                row=max(a.bottom_right.row, b.bottom_right.row),
                col=max(a.bottom_right.col, b.bottom_right.col),
            ),
        )
        anchor.cell_count += other.cell_count
        anchor.formula_count += other.formula_count
        anchor.has_merges = anchor.has_merges or other.has_merges
        anchor.has_formatting = anchor.has_formatting or other.has_formatting
        anchor.key_cells = (anchor.key_cells + other.key_cells)[:20]
        anchor.bounding_box = self._sheet.compute_bounding_box(anchor.cell_range)

    @staticmethod
    def _col_overlap_ratio(anchor: CellRange, cand: CellRange) -> float:
        """Fraction of the anchor's columns that overlap the candidate's."""
        lo = max(anchor.top_left.col, cand.top_left.col)
        hi = min(anchor.bottom_right.col, cand.bottom_right.col)
        if hi < lo:
            return 0.0
        width = anchor.col_count()
        return (hi - lo + 1) / width if width > 0 else 0.0

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

        # Check for bold first row (used for the standalone-HEADER and
        # RESULTS_BLOCK heuristics below).
        first_row_cells = [c for c in cells if c.coord.row == min_row]
        has_bold_header = any(
            c.style and c.style.font and c.style.font.bold
            for c in first_row_cells
        )

        # Detect a real column header anywhere in the first rows (not just a bold
        # row 1) — recognises non-bold, fill-styled, and title-above headers.
        has_table_header = find_header_span(self._sheet, cell_range) is not None

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
        elif has_table_header and total > 3:
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
