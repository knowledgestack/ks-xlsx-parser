"""
Chunk builder for RAG retrieval.

Converts segmented blocks into ChunkDTO objects by:
1. Rendering each block to HTML and plain text
2. Computing token counts
3. Summarizing dependency context
4. Setting prev/next navigation pointers
5. Computing deterministic chunk IDs and hashes
"""

from __future__ import annotations

import logging

from excel_parser.analysis.header_detector import find_header_span
from excel_parser.analysis.section_detector import (
    Section,
    find_sections,
    section_path_at,
)
from excel_parser.models.block import BlockDTO, ChunkDTO, DependencySummary
from excel_parser.models.common import (
    BlockType,
    CellCoord,
    CellRange,
    EdgeType,
    col_number_to_letter,
    compute_hash,
)
from excel_parser.models.sheet import SheetDTO
from excel_parser.models.workbook import WorkbookDTO
from excel_parser.rendering.html_renderer import HtmlRenderer
from excel_parser.rendering.text_renderer import TextRenderer

logger = logging.getLogger(__name__)

# Approximate tokens per character for English text (conservative estimate)
CHARS_PER_TOKEN = 4

# Default per-chunk token budget. Tables under this stay a single whole chunk
# ("most tables should come into the LLM"); larger tables are split into row
# windows that each repeat the header and are flagged as partial.
DEFAULT_MAX_CHUNK_TOKENS = 2000


class ChunkBuilder:
    """
    Builds RAG-ready chunks from segmented blocks.

    For each block:
    - Renders HTML and plain text representations
    - Estimates token count
    - Summarizes upstream/downstream dependencies
    - Assigns navigation pointers (prev/next)
    - Computes deterministic chunk IDs

    Usage:
        builder = ChunkBuilder(workbook_dto)
        chunks = builder.build_all()
    """

    def __init__(
        self,
        workbook: WorkbookDTO,
        max_chunk_tokens: int | None = DEFAULT_MAX_CHUNK_TOKENS,
    ):
        self._workbook = workbook
        self._dep_graph = workbook.dependency_graph
        # Per-chunk token budget; None disables windowing (one chunk per block).
        self._max_chunk_tokens = max_chunk_tokens
        # Circular-ref detection is O(V+E) and does not depend on which block
        # we're looking at — cache it once per workbook to avoid re-running it
        # per chunk (otherwise ~O(chunks × V+E) on dense models).
        self._circular_refs_cache: set[str] | None = None

    def _circular_refs(self) -> set[str]:
        if self._circular_refs_cache is None:
            self._circular_refs_cache = self._dep_graph.detect_circular_refs()
        return self._circular_refs_cache

    def build_all(self) -> list[ChunkDTO]:
        """
        Build chunks for all blocks across all sheets.

        Returns:
            Ordered list of ChunkDTO objects with navigation pointers.
        """
        all_chunks: list[ChunkDTO] = []

        for sheet in self._workbook.sheets:
            # Segment the sheet
            from .segmenter import LayoutSegmenter

            sheet_tables = [
                t for t in self._workbook.tables
                if t.sheet_name == sheet.sheet_name
            ]
            sheet_named = [
                nr.name for nr in self._workbook.named_ranges
                if nr.scope_sheet == sheet.sheet_name or nr.scope_sheet is None
            ]

            segmenter = LayoutSegmenter(
                sheet=sheet,
                tables=sheet_tables,
                named_range_names=sheet_named,
            )
            blocks = segmenter.segment()

            # Finalize blocks
            for block in blocks:
                block.finalize(self._workbook.workbook_hash)

            # Render and build chunks
            html_renderer = HtmlRenderer(sheet)
            text_renderer = TextRenderer(sheet)

            for block in blocks:
                all_chunks.extend(
                    self._block_to_chunks(
                        block, sheet, html_renderer, text_renderer
                    )
                )

        # Add chart summary chunks
        for chart in self._workbook.charts:
            chunk = self._chart_to_chunk(chart)
            all_chunks.append(chunk)

        # Assign global indexes and navigation pointers
        for idx, chunk in enumerate(all_chunks):
            chunk.chunk_index = idx
            chunk.finalize(
                self._workbook.workbook_hash,
                self._workbook.file_path or self._workbook.filename,
            )

        # Set prev/next pointers
        for i in range(len(all_chunks)):
            if i > 0:
                all_chunks[i].prev_chunk_id = all_chunks[i - 1].chunk_id
            if i < len(all_chunks) - 1:
                all_chunks[i].next_chunk_id = all_chunks[i + 1].chunk_id

        logger.info("Built %d chunks from workbook", len(all_chunks))
        return all_chunks

    def _block_to_chunks(
        self,
        block: BlockDTO,
        sheet: SheetDTO,
        html_renderer: HtmlRenderer,
        text_renderer: TextRenderer,
    ) -> list[ChunkDTO]:
        """
        Convert a block into one or more chunks.

        Most blocks become a single whole-table chunk. A block whose rendered
        text exceeds ``max_chunk_tokens`` is split into row windows that each
        repeat the header and are flagged as a partial view of the table (so an
        agent knows the chunk is one slice, not the whole table). Each window's
        ``cell_range`` is the *data slice only* — the repeated header is render
        sugar — keeping chunk coordinates non-overlapping and provenance clean.
        """
        # Deterministic section structure for this table (outline levels /
        # merged dividers). Empty for non-table blocks or unsectioned tables.
        sections, sections_meta = self._detect_sections(block, sheet)

        whole_text = self._render_text(text_renderer, block)
        windows = None
        if self._max_chunk_tokens:
            token_count = max(len(whole_text) // CHARS_PER_TOKEN, 1)
            if token_count > self._max_chunk_tokens:
                windows = self._plan_windows(
                    block, sheet, text_renderer, token_count
                )

        if not windows:
            whole_html = self._render_html(html_renderer, block)
            return [self._make_chunk(block, sheet, whole_text, whole_html,
                                     block.cell_range, None,
                                     sections_meta=sections_meta)]

        rng = block.cell_range
        min_col, max_col = rng.top_left.col, rng.bottom_right.col
        table_range = f"{block.sheet_name}!{rng.to_a1()}"
        # Deterministic table id shared by every part (hash-derived, not a counter).
        table_id = compute_hash(
            self._workbook.workbook_hash, block.sheet_name, rng.to_a1()
        )
        data_lo, data_hi = windows[0][0], windows[-1][1]
        n = len(windows)
        chunks: list[ChunkDTO] = []
        for i, (ws, we) in enumerate(windows, 1):
            banner = (
                f"[PART {i}/{n} of table {table_range} — rows {ws}–{we} "
                f"of {data_lo}–{data_hi}; header repeated below]"
            )
            # Active section ancestors whose header row sits ABOVE this part (so
            # a part that starts mid-section still names its section). Sections
            # whose header is inside the part render inline, so they're excluded.
            section_path_meta = None
            active = [s for s in section_path_at(sections, ws) if s.top_row < ws]
            if active:
                banner += "\nsection: " + " › ".join(s.label for s in active)
                section_path_meta = [
                    {"label": s.label, "level": s.level} for s in active
                ]
            include_title = i == 1  # keep the table's title on the first part only
            wt = self._render_text(
                text_renderer, block, window=(ws, we, banner, include_title)
            )
            wh = self._render_html(
                html_renderer, block, window=(ws, we, banner, include_title)
            )
            wrange = CellRange(
                top_left=CellCoord(row=ws, col=min_col),
                bottom_right=CellCoord(row=we, col=max_col),
            )
            table_part = {
                "table_id": table_id,
                "part": i,
                "of": n,
                "row_start": ws,
                "row_end": we,
                "table_range": table_range,
                "is_partial": True,
            }
            chunks.append(self._make_chunk(
                block, sheet, wt, wh, wrange, table_part,
                sections_meta=sections_meta, section_path_meta=section_path_meta,
            ))
        return chunks

    # Block types that can carry sections (data-bearing tables).
    _SECTIONABLE = (
        BlockType.TABLE,
        BlockType.ASSUMPTIONS_TABLE,
        BlockType.RESULTS_BLOCK,
        BlockType.CALCULATION_BLOCK,
        BlockType.MIXED,
        BlockType.DATA_BLOCK,
    )

    def _detect_sections(
        self, block: BlockDTO, sheet: SheetDTO
    ) -> tuple[list[Section], list[dict] | None]:
        """Find a table's sections and a JSON-serialisable hierarchy for metadata."""
        if block.block_type not in self._SECTIONABLE:
            return [], None
        header_span = None
        if block.block_type in (BlockType.TABLE, BlockType.ASSUMPTIONS_TABLE):
            header_span = find_header_span(sheet, block.cell_range)
        sections = find_sections(sheet, block.cell_range, header_span)
        if not sections:
            return [], None
        meta = [
            {
                "label": s.label,
                "level": s.level,
                "row_start": s.top_row,
                "row_end": s.bottom_row,
            }
            for s in sections[:200]  # cap to bound metadata size on huge tables
        ]
        return sections, meta

    @staticmethod
    def _render_text(renderer: TextRenderer, block: BlockDTO, window=None) -> str:
        try:
            if window is not None:
                start, end, banner, include_title = window
                return renderer.render_window(block, start, end, banner, include_title)
            return renderer.render_block(block)
        except Exception as e:  # noqa: BLE001 - rendering must never abort chunking
            logger.warning("Text rendering failed for block %s: %s", block.block_id, e)
            return f"[render error: {e}]"

    @staticmethod
    def _render_html(renderer: HtmlRenderer, block: BlockDTO, window=None) -> str:
        try:
            if window is not None:
                start, end, banner, include_title = window
                return renderer.render_window(block, start, end, banner, include_title)
            return renderer.render_block(block)
        except Exception as e:  # noqa: BLE001
            logger.warning("HTML rendering failed for block %s: %s", block.block_id, e)
            return f"<!-- render error: {e} -->"

    # Fraction of the budget targeted when sizing windows, leaving headroom for
    # the approximation in the per-row estimate (column widths grow as a window
    # spans wider values). The budget is therefore an approximate target, not a
    # hard cap — windows may run modestly over.
    _WINDOW_BUDGET_MARGIN = 0.9

    def _plan_windows(
        self,
        block: BlockDTO,
        sheet: SheetDTO,
        text_renderer: TextRenderer,
        whole_token_count: int,
    ) -> list[tuple[int, int]] | None:
        """
        Split a block's data rows into windows that each (approximately) fit the
        token budget. Returns inclusive (start_row, end_row) data ranges, or None
        if the block cannot/should not be windowed (e.g. a single data row).

        Sizing accounts for the fixed per-window overhead — the repeated header,
        bracket, ``cols:`` map and banner are re-paid by every part — so a part
        targets ``max_chunk_tokens`` rather than the whole-table row average
        (which amortises the header once and would make every window overflow).
        The per-row cost is taken from the *whole* render's global column widths,
        an upper bound for any window (a window's local widths are no wider), and
        a margin absorbs the residual approximation.
        """
        budget = self._max_chunk_tokens
        if not budget:
            return None
        rng = block.cell_range
        bot = rng.bottom_right.row
        header_row = None
        if block.block_type in (BlockType.TABLE, BlockType.ASSUMPTIONS_TABLE):
            span = find_header_span(sheet, rng)
            header_row = span.top if span else None
        data_lo = (header_row + 1) if header_row is not None else rng.top_left.row
        num_data = bot - data_lo + 1
        if num_data <= 1:
            return None  # nothing to window

        # Per-row cost using GLOBAL column widths (the whole render). A window's
        # local widths are <= global, so this over-estimates → windows stay under.
        total_rows = bot - rng.top_left.row + 1
        per_row = max(1, whole_token_count // total_rows)

        # Fixed overhead a window re-pays (header + bracket + cols + banner),
        # measured with a representative banner.
        sample_banner = (
            f"[PART 0/0 of table {block.sheet_name}!{rng.to_a1()} — "
            f"rows 0000-0000 of 0000-0000; header repeated below]"
        )
        overhead = max(
            len(text_renderer.render_window(block, data_lo, data_lo, sample_banner))
            // CHARS_PER_TOKEN,
            1,
        )

        avail = int(budget * self._WINDOW_BUDGET_MARGIN) - overhead
        rows_per = max(1, avail // per_row) if avail > 0 else 1
        if rows_per >= num_data:
            return None  # fits in a single window — keep whole

        windows: list[tuple[int, int]] = []
        r = data_lo
        while r <= bot:
            windows.append((r, min(r + rows_per - 1, bot)))
            r += rows_per
        return windows

    def _make_chunk(
        self,
        block: BlockDTO,
        sheet: SheetDTO,
        render_text: str,
        render_html: str,
        rng: CellRange,
        table_part: dict | None,
        sections_meta: list[dict] | None = None,
        section_path_meta: list[dict] | None = None,
    ) -> ChunkDTO:
        """Assemble a ChunkDTO for a given (data) range and pre-rendered content."""
        token_count = max(len(render_text) // CHARS_PER_TOKEN, 1)
        dep_summary = self._build_dependency_summary(sheet, rng)

        key_cells = [
            f"{sheet.sheet_name}!{coord.to_a1()}"
            for coord in block.key_cells
            if rng.contains(coord)
        ]

        # Hidden-content provenance, scoped to this chunk's range.
        metadata: dict[str, object] = {}
        if sheet.properties.is_hidden:
            metadata["sheet_hidden"] = True
        hidden_rows = sorted(
            r for r in sheet.hidden_rows
            if rng.top_left.row <= r <= rng.bottom_right.row
        )
        if hidden_rows:
            metadata["hidden_rows"] = hidden_rows
        hidden_cols = sorted(
            c for c in sheet.hidden_cols
            if rng.top_left.col <= c <= rng.bottom_right.col
        )
        if hidden_cols:
            metadata["hidden_cols"] = [col_number_to_letter(c) for c in hidden_cols]
        if table_part is not None:
            metadata["table_part"] = table_part
        # Section structure (deterministic — outline levels / merged dividers).
        # `sections` is the table's section hierarchy; `section_path` is the
        # active ancestor chain for a windowed part whose section header sits
        # above the part (so a mid-section part still names its section).
        if sections_meta:
            metadata["sections"] = sections_meta
        if section_path_meta:
            metadata["section_path"] = section_path_meta

        return ChunkDTO(
            sheet_name=block.sheet_name,
            block_type=block.block_type,
            top_left_cell=rng.top_left.to_a1(),
            bottom_right_cell=rng.bottom_right.to_a1(),
            cell_range=rng,
            key_cells=key_cells,
            named_ranges=block.named_ranges,
            dependency_summary=dep_summary,
            render_html=render_html,
            render_text=render_text,
            token_count=token_count,
            metadata=metadata,
        )

    def _chart_to_chunk(self, chart) -> ChunkDTO:
        """Convert a chart into a RAG chunk."""
        summary = chart.summary_text or chart.generate_summary()
        token_count = max(len(summary) // CHARS_PER_TOKEN, 1)

        # Determine chart position range
        top_left = "A1"
        bottom_right = "A1"
        if chart.anchor:
            from excel_parser.models.common import col_number_to_letter
            top_left = f"{col_number_to_letter(chart.anchor.from_col + 1)}{chart.anchor.from_row + 1}"
            if chart.anchor.to_col is not None and chart.anchor.to_row is not None:
                bottom_right = f"{col_number_to_letter(chart.anchor.to_col + 1)}{chart.anchor.to_row + 1}"

        html_content = f'<div class="chart-summary" data-chart-type="{chart.chart_type.value}">'
        html_content += f"<h4>{summary.split(chr(10))[0]}</h4>"
        html_content += f"<pre>{summary}</pre></div>"

        return ChunkDTO(
            sheet_name=chart.sheet_name,
            block_type="chart_anchor",
            top_left_cell=top_left,
            bottom_right_cell=bottom_right,
            render_html=html_content,
            render_text=summary,
            token_count=token_count,
            metadata={"chart_id": chart.chart_id, "chart_type": chart.chart_type.value},
        )

    def _build_dependency_summary(
        self, sheet: SheetDTO, rng: CellRange
    ) -> DependencySummary:
        """Build a compact dependency summary for a cell range."""
        upstream: set[str] = set()
        downstream: set[str] = set()
        cross_sheet: set[str] = set()
        has_circular = False

        for row in range(rng.top_left.row, rng.bottom_right.row + 1):
            for col in range(rng.top_left.col, rng.bottom_right.col + 1):
                cell = sheet.get_cell(row, col)
                if not cell or not cell.formula:
                    continue

                coord = CellCoord(row=row, col=col)

                # Upstream deps (what this cell references)
                for edge in self._dep_graph.get_upstream(
                    sheet.sheet_name, coord, max_depth=2
                ):
                    ref = edge.target_ref_string
                    upstream.add(ref)
                    if edge.edge_type == EdgeType.CROSS_SHEET:
                        cross_sheet.add(ref)

                # Downstream deps (what references this cell)
                for edge in self._dep_graph.get_downstream(
                    sheet.sheet_name, coord, max_depth=1
                ):
                    downstream.add(
                        f"{edge.source_sheet}!{edge.source_coord.to_a1()}"
                    )

        # Check for circular refs (cached once per workbook)
        circular = self._circular_refs()
        if circular:
            for row in range(rng.top_left.row, rng.bottom_right.row + 1):
                if has_circular:
                    break
                for col in range(rng.top_left.col, rng.bottom_right.col + 1):
                    key = f"{sheet.sheet_name}!{CellCoord(row=row, col=col).to_a1()}"
                    if key in circular:
                        has_circular = True
                        break

        return DependencySummary(
            upstream_refs=sorted(upstream)[:50],
            downstream_refs=sorted(downstream)[:50],
            cross_sheet_refs=sorted(cross_sheet)[:20],
            has_circular=has_circular,
        )
