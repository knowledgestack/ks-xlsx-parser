"""
End-to-end parsing pipeline.

This is the primary public API for the xlsx_parser. It orchestrates
the full 11-stage Excellent Algorithm:
  Stage 0: Sheet Chunking (adaptive gaps + style boundaries)
  Stage 1: Cell Annotation (feature-based scorer)
  Stage 2: Solid Block ID (annotation-based splitting)
  Stage 3: Table Assembly (label-data association)
  Stage 4: Light Block Detection (sparse block association)
  Stage 5: Table Grouping (structural similarity)
  Stage 6: Pattern Splitting (repeating labels)
  Stage 7: Tree Building (recursive hierarchy)
  Stage 8: Template Extraction (DOF identification)
  Render + Store

Usage:
    from xlsx_parser.pipeline import parse_workbook
    result = parse_workbook("path/to/workbook.xlsx")
    print(result.workbook.total_cells)
    for chunk in result.chunks:
        print(chunk.source_uri, chunk.render_text[:100])
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from .analysis.light_block_detector import LightBlockDetector
from .analysis.pattern_splitter import PatternSplitter
from .analysis.table_assembler import TableAssembler
from .analysis.table_grouper import TableGrouper
from .analysis.template_extractor import TemplateExtractor
from .analysis.tree_builder import TreeBuilder
from .annotation.block_splitter import BlockSplitter
from .annotation.cell_annotator import CellAnnotator
from .chunking.chunker import ChunkBuilder
from .comparison.template_comparator import TemplateComparator
from .export.model_exporter import ModelExporter
from .models.block import BlockDTO, ChunkDTO
from .models.table_structure import TableStructure
from .models.template import GeneralizedTemplate, TemplateNode
from .models.tree import TreeNode
from .models.workbook import WorkbookDTO
from .parsers.workbook_parser import WorkbookParser
from .storage.serializer import WorkbookSerializer

logger = logging.getLogger(__name__)


@dataclass
class ParseResult:
    """Complete output of the parsing pipeline."""

    workbook: WorkbookDTO
    chunks: list[ChunkDTO] = field(default_factory=list)
    serializer: WorkbookSerializer | None = None

    @property
    def total_chunks(self) -> int:
        return len(self.chunks)

    @property
    def total_tokens(self) -> int:
        return sum(c.token_count for c in self.chunks)

    def to_json(self) -> dict[str, Any]:
        """Serialize the full result to a JSON-compatible dict."""
        return {
            "workbook": {
                "workbook_id": self.workbook.workbook_id,
                "filename": self.workbook.filename,
                "workbook_hash": self.workbook.workbook_hash,
                "total_sheets": self.workbook.total_sheets,
                "total_cells": self.workbook.total_cells,
                "total_formulas": self.workbook.total_formulas,
                "parse_duration_ms": self.workbook.parse_duration_ms,
                "errors": [e.model_dump(exclude_none=True) for e in self.workbook.errors],
            },
            "chunks": [
                {
                    "chunk_id": c.chunk_id,
                    "source_uri": c.source_uri,
                    "sheet_name": c.sheet_name,
                    "block_type": c.block_type if isinstance(c.block_type, str) else c.block_type.value,
                    "top_left": c.top_left_cell,
                    "bottom_right": c.bottom_right_cell,
                    "token_count": c.token_count,
                    "content_hash": c.content_hash,
                    "render_text": c.render_text,
                }
                for c in self.chunks
            ],
            "total_chunks": self.total_chunks,
            "total_tokens": self.total_tokens,
        }


def parse_workbook(
    path: str | Path | None = None,
    content: bytes | None = None,
    filename: str | None = None,
    max_cells_per_sheet: int = 2_000_000,
) -> ParseResult:
    """
    Parse an Excel workbook through the full Excellent Algorithm pipeline.

    Stages 0-8 are executed for single-document parsing.
    For stages 9-10, use compare_workbooks() and export_importer().

    Args:
        path: Path to the .xlsx file.
        content: Raw bytes (alternative to path).
        filename: Display filename when using content.
        max_cells_per_sheet: Safety limit per sheet.

    Returns:
        ParseResult with workbook DTO, chunks, and serializer.
    """
    # Load + Parse
    parser = WorkbookParser(
        path=path,
        content=content,
        filename=filename,
        max_cells_per_sheet=max_cells_per_sheet,
    )
    workbook = parser.parse()

    # Run Excellent Algorithm stages 0-8 per sheet
    all_table_structures: list[TableStructure] = []
    all_tree_nodes: list[TreeNode] = []
    all_template_nodes: list[TemplateNode] = []

    for sheet in workbook.sheets:
        # Stage 0: Sheet Chunking (handled inside ChunkBuilder/LayoutSegmenter)
        # The LayoutSegmenter now uses adaptive gaps + style boundaries

        # Stage 1: Cell Annotation
        annotator = CellAnnotator(sheet)
        annotator.annotate()

        # Stage 0+2: Segment then split by annotation
        from .chunking.segmenter import LayoutSegmenter
        sheet_tables = [t for t in workbook.tables if t.sheet_name == sheet.sheet_name]
        sheet_named = [
            nr.name for nr in workbook.named_ranges
            if nr.scope_sheet == sheet.sheet_name or nr.scope_sheet is None
        ]
        segmenter = LayoutSegmenter(
            sheet=sheet, tables=sheet_tables, named_range_names=sheet_named,
        )
        blocks = segmenter.segment()

        # Finalize block IDs
        for block in blocks:
            block.finalize(workbook.workbook_hash)

        # Stage 2: Solid Block Identification (split by annotation)
        splitter = BlockSplitter(sheet)
        blocks = splitter.split_blocks(blocks)

        # Re-finalize after splitting
        for block in blocks:
            block.finalize(workbook.workbook_hash)

        # Stage 3: Table Assembly
        assembler = TableAssembler(sheet)
        structures = assembler.assemble(blocks)

        # Finalize structures
        for s in structures:
            s.finalize(workbook.workbook_hash)

        # Stage 4: Light Block Detection
        detector = LightBlockDetector()
        blocks, structures = detector.detect_and_associate(blocks, structures)

        # Stage 5: Table Grouping
        grouper = TableGrouper(sheet)
        blocks, structures = grouper.group_tables(blocks, structures)

        # Re-finalize blocks after grouping
        for block in blocks:
            block.finalize(workbook.workbook_hash)

        # Stage 6: Pattern Splitting
        pattern_splitter = PatternSplitter(sheet)
        blocks, structures = pattern_splitter.split(blocks, structures)

        # Stage 7: Tree Building
        tree_builder = TreeBuilder(sheet, workbook.workbook_hash)
        tree_nodes = tree_builder.build_tree(blocks, structures)

        # Stage 8: Template Extraction
        extractor = TemplateExtractor(sheet, workbook.workbook_hash)
        template_nodes = extractor.extract(tree_nodes)

        all_table_structures.extend(structures)
        all_tree_nodes.extend(tree_nodes)
        all_template_nodes.extend(template_nodes)

    # Store results on workbook
    workbook.table_structures = all_table_structures
    workbook.tree_nodes = all_tree_nodes
    workbook.template_nodes = all_template_nodes

    # Render chunks (uses original segmentation internally)
    chunk_builder = ChunkBuilder(workbook)
    chunks = chunk_builder.build_all()

    # Prepare serializer
    serializer = WorkbookSerializer(workbook, chunks)

    return ParseResult(
        workbook=workbook,
        chunks=chunks,
        serializer=serializer,
    )


def compare_workbooks(
    paths: list[str | Path],
    dof_threshold: int = 50,
) -> GeneralizedTemplate:
    """
    Stage 9: Compare templates from multiple workbooks.

    Parses each workbook, extracts templates, and compares them
    to produce a generalized template with DOFs where conflicts exist.

    Args:
        paths: Paths to .xlsx files.
        dof_threshold: Maximum DOFs before re-analysis is recommended.

    Returns:
        GeneralizedTemplate with merged cell specs and conflict records.
    """
    template_sets: list[tuple[str, list[TemplateNode]]] = []

    for path in paths:
        result = parse_workbook(path=path)
        filename = result.workbook.filename
        templates = result.workbook.template_nodes
        template_sets.append((filename, templates))

    comparator = TemplateComparator(dof_threshold=dof_threshold)
    return comparator.compare(template_sets)


def export_importer(
    template: GeneralizedTemplate,
    output_path: str | Path,
    class_name: str = "GeneratedImporter",
) -> Path:
    """
    Stage 10: Export a generalized template as an importable Python class.

    Args:
        template: GeneralizedTemplate from compare_workbooks().
        output_path: Path for the generated Python file.
        class_name: Name for the generated class.

    Returns:
        Path to the generated file.
    """
    exporter = ModelExporter()
    return exporter.export_to_file(template, output_path, class_name)
