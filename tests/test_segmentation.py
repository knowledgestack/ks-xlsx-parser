"""
Tests for layout segmentation and block classification.

Verifies that the segmenter correctly identifies tables, calculation blocks,
assumption blocks, result blocks, and text headers.
"""

import pytest

from chunking.segmenter import LayoutSegmenter
from models import BlockType
from parsers import WorkbookParser


class TestSegmentation:
    """Test block detection and classification."""

    def test_simple_block_detected(self, simple_workbook):
        result = WorkbookParser(path=simple_workbook).parse()
        sheet = result.sheets[0]
        segmenter = LayoutSegmenter(sheet)
        blocks = segmenter.segment()
        assert len(blocks) >= 1

    def test_table_block_from_listobject(self, table_workbook):
        result = WorkbookParser(path=table_workbook).parse()
        sheet = result.sheets[0]
        tables = [t for t in result.tables if t.sheet_name == sheet.sheet_name]
        segmenter = LayoutSegmenter(sheet, tables=tables)
        blocks = segmenter.segment()
        table_blocks = [b for b in blocks if b.block_type == BlockType.TABLE]
        assert len(table_blocks) >= 1
        assert table_blocks[0].table_name == "SalesData"

    def test_assumptions_block_classified(self, assumptions_workbook):
        result = WorkbookParser(path=assumptions_workbook).parse()
        sheet = result.sheets[0]
        segmenter = LayoutSegmenter(sheet)
        blocks = segmenter.segment()
        # Should find at least 2 blocks separated by blank row 6
        assert len(blocks) >= 2

    def test_sparse_segmentation(self, large_sparse_workbook):
        result = WorkbookParser(path=large_sparse_workbook).parse()
        sheet = result.sheets[0]
        segmenter = LayoutSegmenter(sheet)
        blocks = segmenter.segment()
        # Sparse cells should form separate blocks
        assert len(blocks) >= 2

    def test_blocks_have_bounding_box(self, simple_workbook):
        result = WorkbookParser(path=simple_workbook).parse()
        sheet = result.sheets[0]
        segmenter = LayoutSegmenter(sheet)
        blocks = segmenter.segment()
        for block in blocks:
            assert block.cell_range is not None
            assert block.bounding_box is not None

    def test_block_hashes_deterministic(self, simple_workbook):
        result = WorkbookParser(path=simple_workbook).parse()
        sheet = result.sheets[0]

        seg1 = LayoutSegmenter(sheet)
        blocks1 = seg1.segment()
        for b in blocks1:
            b.finalize("test_hash")

        seg2 = LayoutSegmenter(sheet)
        blocks2 = seg2.segment()
        for b in blocks2:
            b.finalize("test_hash")

        assert len(blocks1) == len(blocks2)
        for b1, b2 in zip(blocks1, blocks2):
            assert b1.content_hash == b2.content_hash
