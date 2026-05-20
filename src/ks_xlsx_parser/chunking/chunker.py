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
import os

from ks_xlsx_parser.models.block import BlockDTO, ChunkDTO, DependencySummary
from ks_xlsx_parser.models.common import CellCoord, CellRange, EdgeType
from ks_xlsx_parser.models.sheet import SheetDTO
from ks_xlsx_parser.models.workbook import WorkbookDTO
from ks_xlsx_parser.rendering.html_renderer import HtmlRenderer
from ks_xlsx_parser.rendering.text_renderer import TextRenderer

logger = logging.getLogger(__name__)

# Approximate tokens per character for English text (conservative estimate)
CHARS_PER_TOKEN = 4

# Cluster-04 chunk-size cap. A block whose rendered text exceeds this
# many chars gets split into row-group children, each with a tight A1
# range covering only its rows.
#
# Default is intentionally HIGH (effectively off for any reasonable
# workbook). Empirical reason: on the 50-sample stratified for the
# cluster-04 large-sheet pattern, every budget tight enough to actually
# fragment moderate tables (2k–4k chars) regressed at least one
# instance because the embedding cannot discriminate between
# same-shape row-group children — splitting one chunk into N
# look-alike chunks makes the embedding pick the wrong child first.
# An 8k budget was neutral (0 wins, 0 regressions). Above that no
# real-world tables we measured split. So keep tables whole by
# default; the splitter is plumbing for downstream consumers who
# need hard per-chunk token economy (e.g. an embedding model with a
# strict 512-token ceiling). They opt in via
# ``KS_CHUNK_BUDGET_CHARS=2000`` and validate against THEIR corpus.
def _budget_chars() -> int:
    """Read budget per call so tests can monkeypatch the env var."""
    return int(os.environ.get("KS_CHUNK_BUDGET_CHARS", "100000"))

# Below this row count, splitting can't produce ≥2 children meaningfully
# (need at least 2 data rows to split into 2 groups).
MIN_ROWS_TO_SPLIT = 3


def _split_block_by_rows(
    block: BlockDTO,
    sheet: SheetDTO,
    text_renderer: TextRenderer,
    budget_chars: int | None = None,
) -> list[BlockDTO]:
    """Split an oversize block into contiguous row-group sub-blocks.

    Returns ``[block]`` unchanged when the block fits the budget. Otherwise
    returns N sub-blocks with non-overlapping, contiguous row ranges
    summing to the original block's row coverage. Each sub-block carries
    a copy of the parent's block-level metadata (block_type, table_name,
    named_ranges, density flags) and a subset of ``key_cells`` whose rows
    fall inside the slice.

    Header preservation: this v1 emits row-group children WITHOUT
    repeating the parent's header row(s). The cluster-04 doc calls
    out "every group needs the header rows OR a header-summary block";
    we accept the simpler invariant here because the alternative
    (overlapping cell_range across siblings, or out-of-band header
    text whose cells live outside the chunk's claimed range) breaks
    cluster-02's range-tightening invariant. Header-context recovery
    is a follow-up.
    """
    if budget_chars is None:
        budget_chars = _budget_chars()
    rng = block.cell_range
    n_rows = rng.bottom_right.row - rng.top_left.row + 1
    if n_rows < MIN_ROWS_TO_SPLIT:
        return [block]

    # Render once to measure. text_renderer is fast (~ms per 1000-row block);
    # we already had to render in _block_to_chunk anyway — this just moves
    # the work earlier.
    full_text = text_renderer.render_block(block)
    if len(full_text) <= budget_chars:
        return [block]

    # Estimate rows-per-child from observed chars/row (averaged over the
    # full render, including header overhead). Ceil to ensure children
    # don't individually exceed the budget on the high end.
    avg_chars_per_row = max(1.0, len(full_text) / n_rows)
    rows_per_group = max(2, int(budget_chars / avg_chars_per_row))

    sub_blocks: list[BlockDTO] = []
    cur = rng.top_left.row
    while cur <= rng.bottom_right.row:
        end = min(cur + rows_per_group - 1, rng.bottom_right.row)
        sub_range = CellRange(
            top_left=CellCoord(row=cur, col=rng.top_left.col),
            bottom_right=CellCoord(row=end, col=rng.bottom_right.col),
        )
        # Carry only the key_cells whose row falls inside this slice.
        # Important for downstream metadata fidelity (key_cells encode
        # "notable output cells" — bold, colored, etc.).
        sub_key_cells = [k for k in block.key_cells if cur <= k.row <= end]
        sub = BlockDTO(
            block_index=block.block_index,
            sheet_name=block.sheet_name,
            block_type=block.block_type,
            cell_range=sub_range,
            bounding_box=block.bounding_box,
            cell_count=block.cell_count,
            formula_count=block.formula_count,
            has_merges=block.has_merges,
            has_formatting=block.has_formatting,
            key_cells=sub_key_cells,
            named_ranges=block.named_ranges,
            table_name=block.table_name,
            parent_block_id=block.block_id or None,
            density=block.density,
            label_cell_count=block.label_cell_count,
            data_cell_count=block.data_cell_count,
            table_structure_id=block.table_structure_id,
        )
        sub_blocks.append(sub)
        cur = end + 1

    return sub_blocks


def _tight_content_bbox(
    block: BlockDTO, sheet: SheetDTO,
) -> tuple[int, int, int, int] | None:
    """Bounding box of cells with non-empty content within ``block.cell_range``.

    Returns ``(r0, c0, r1, c1)`` or ``None`` if the entire range is
    empty. "Non-empty" = ``raw_value`` is not None OR ``display_value``
    has a non-whitespace string. Merged-slave cells count as empty
    (their master is what carries the value).

    Used to clip overclaiming chunk ranges before emission so that
    geometric overlap scoring (cluster-02 cohort) is honest. A sheet
    that has data in A1:E50 but whose segmenter handed us A1:XFD500
    used to over-claim coverage of any GT cell in the empty area;
    after the clip, the chunk only claims what it can actually
    answer questions about.

    Iterates ``sheet.cells`` rather than the dense range — many sheets
    have styled-empty cells stored across XFD columns, and the dense
    loop costs ~65k iterations per row for those.
    """
    rng = block.cell_range
    r0_, c0_ = rng.top_left.row, rng.top_left.col
    r1_, c1_ = rng.bottom_right.row, rng.bottom_right.col

    min_row = min_col = None
    max_row = max_col = None
    # sheet.cells is dict[str, CellDTO] keyed "row,col"; iterate values
    # and read coord off the DTO. Avoids iterating the dense range when
    # styled-empty cells span XFD width (16384 cols).
    for cell in sheet.cells.values():
        r, c = cell.coord.row, cell.coord.col
        if not (r0_ <= r <= r1_ and c0_ <= c <= c1_):
            continue
        if cell.is_merged_slave:
            continue
        if cell.raw_value is None and not (
            cell.display_value and cell.display_value.strip()
        ):
            continue
        if min_row is None or r < min_row:
            min_row = r
        if max_row is None or r > max_row:
            max_row = r
        if min_col is None or c < min_col:
            min_col = c
        if max_col is None or c > max_col:
            max_col = c

    if min_row is None:
        return None
    return (min_row, min_col, max_row, max_col)


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

            # Cluster-04 size cap: replace oversize blocks with row-group
            # children before chunk emission. Finalize the children with
            # the workbook hash so their block_ids are stable IDs.
            expanded_blocks: list[BlockDTO] = []
            for block in blocks:
                children = _split_block_by_rows(block, sheet, text_renderer)
                if len(children) > 1:
                    # Re-finalize each child with the workbook hash so its
                    # block_id is deterministic. (The parent kept its own
                    # block_id from the earlier finalize; children inherit
                    # parent_block_id and get a fresh ID off the child's
                    # narrower cell_range.)
                    for child in children:
                        child.finalize(self._workbook.workbook_hash)
                expanded_blocks.extend(children)
            blocks = expanded_blocks

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

        # Cluster-02 invariant: the chunk's claimed A1 range must be a
        # tight bbox over cells with content. Overclaiming inflates
        # geometric recall for the parser but lies to downstream
        # consumers (a UI citation that highlights an empty area).
        # We clip but never widen — the renderer already only outputs
        # cells inside `block.cell_range`, so a narrowed range still
        # contains every cell that contributed to `render_text`.
        tight = _tight_content_bbox(block, sheet)
        if tight is not None:
            r0, c0, r1, c1 = tight
            chunk_range = CellRange(
                top_left=CellCoord(row=r0, col=c0),
                bottom_right=CellCoord(row=r1, col=c1),
            )
        else:
            # Block has zero non-empty cells (styled-empty cells across
            # XFD columns is a real shape on this corpus). Claim only the
            # top-left cell so the chunk is honest about being empty.
            tl = block.cell_range.top_left
            chunk_range = CellRange(top_left=tl, bottom_right=tl)

        return ChunkDTO(
            sheet_name=block.sheet_name,
            block_type=block.block_type,
            top_left_cell=chunk_range.top_left.to_a1(),
            bottom_right_cell=chunk_range.bottom_right.to_a1(),
            cell_range=chunk_range,
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
            from ks_xlsx_parser.models.common import col_number_to_letter
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
