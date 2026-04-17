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

from ..models.block import BlockDTO, ChunkDTO, DependencySummary
from ..models.common import CellCoord, EdgeType
from ..models.dependency import DependencyGraph
from ..models.sheet import SheetDTO
from ..models.table import TableDTO
from ..models.workbook import NamedRangeDTO, WorkbookDTO
from ..rendering.html_renderer import HtmlRenderer
from ..rendering.text_renderer import TextRenderer

logger = logging.getLogger(__name__)

# Approximate tokens per character for English text (conservative estimate)
CHARS_PER_TOKEN = 4


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

    def __init__(self, workbook: WorkbookDTO):
        self._workbook = workbook
        self._dep_graph = workbook.dependency_graph
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
                chunk = self._block_to_chunk(
                    block, sheet, html_renderer, text_renderer
                )
                all_chunks.append(chunk)

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

    def _block_to_chunk(
        self,
        block: BlockDTO,
        sheet: SheetDTO,
        html_renderer: HtmlRenderer,
        text_renderer: TextRenderer,
    ) -> ChunkDTO:
        """Convert a block into a chunk with rendered content and metadata."""
        # Render
        try:
            render_html = html_renderer.render_block(block)
        except Exception as e:
            logger.warning("HTML rendering failed for block %s: %s", block.block_id, e)
            render_html = f"<!-- render error: {e} -->"

        try:
            render_text = text_renderer.render_block(block)
        except Exception as e:
            logger.warning("Text rendering failed for block %s: %s", block.block_id, e)
            render_text = f"[render error: {e}]"

        # Token count estimate
        token_count = max(len(render_text) // CHARS_PER_TOKEN, 1)

        # Dependency summary
        dep_summary = self._build_dependency_summary(block, sheet)

        # Key cells as A1 refs
        key_cells = [
            f"{sheet.sheet_name}!{coord.to_a1()}"
            for coord in block.key_cells
        ]

        return ChunkDTO(
            sheet_name=block.sheet_name,
            block_type=block.block_type,
            top_left_cell=block.cell_range.top_left.to_a1(),
            bottom_right_cell=block.cell_range.bottom_right.to_a1(),
            cell_range=block.cell_range,
            key_cells=key_cells,
            named_ranges=block.named_ranges,
            dependency_summary=dep_summary,
            render_html=render_html,
            render_text=render_text,
            token_count=token_count,
        )

    def _chart_to_chunk(self, chart) -> ChunkDTO:
        """Convert a chart into a RAG chunk."""
        summary = chart.summary_text or chart.generate_summary()
        token_count = max(len(summary) // CHARS_PER_TOKEN, 1)

        # Determine chart position range
        top_left = "A1"
        bottom_right = "A1"
        if chart.anchor:
            from ..models.common import col_number_to_letter
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
        self, block: BlockDTO, sheet: SheetDTO
    ) -> DependencySummary:
        """Build a compact dependency summary for a block."""
        upstream: set[str] = set()
        downstream: set[str] = set()
        cross_sheet: set[str] = set()
        has_circular = False

        rng = block.cell_range
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
