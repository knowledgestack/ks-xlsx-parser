"""
Tests for the end-to-end parsing pipeline.

Verifies that the full pipeline produces correct chunks with
rendered content, dependency summaries, and deterministic hashes.
"""

import json

import pytest

from pipeline import parse_workbook


class TestEndToEndPipeline:
    """Test the full parse pipeline."""

    def test_simple_pipeline(self, simple_workbook):
        result = parse_workbook(path=simple_workbook)

        assert result.workbook.total_sheets == 1
        assert result.workbook.total_cells > 0
        assert result.total_chunks > 0
        assert result.total_tokens > 0

    def test_chunks_have_rendered_content(self, simple_workbook):
        result = parse_workbook(path=simple_workbook)

        for chunk in result.chunks:
            assert chunk.render_html != ""
            assert chunk.render_text != ""
            assert chunk.token_count > 0

    def test_chunks_have_source_uri(self, simple_workbook):
        result = parse_workbook(path=simple_workbook)

        for chunk in result.chunks:
            assert chunk.source_uri != ""
            assert chunk.sheet_name != ""
            assert chunk.chunk_id != ""
            assert chunk.content_hash != ""

    def test_chunk_navigation(self, simple_workbook):
        result = parse_workbook(path=simple_workbook)

        if len(result.chunks) > 1:
            assert result.chunks[0].next_chunk_id is not None
            assert result.chunks[-1].prev_chunk_id is not None

    def test_formula_workbook_pipeline(self, formula_workbook):
        result = parse_workbook(path=formula_workbook)

        assert result.workbook.total_sheets == 3
        assert result.workbook.total_formulas > 0
        assert len(result.workbook.dependency_graph.edges) > 0
        assert len(result.workbook.named_ranges) >= 2

    def test_table_workbook_pipeline(self, table_workbook):
        result = parse_workbook(path=table_workbook)

        assert len(result.workbook.tables) == 1
        assert result.total_chunks > 0

    def test_deterministic_output(self, simple_workbook):
        r1 = parse_workbook(path=simple_workbook)
        r2 = parse_workbook(path=simple_workbook)

        assert r1.workbook.workbook_hash == r2.workbook.workbook_hash
        assert r1.total_chunks == r2.total_chunks
        for c1, c2 in zip(r1.chunks, r2.chunks):
            assert c1.chunk_id == c2.chunk_id
            assert c1.content_hash == c2.content_hash

    def test_to_json(self, simple_workbook):
        result = parse_workbook(path=simple_workbook)
        j = result.to_json()

        assert "workbook" in j
        assert "chunks" in j
        assert j["total_chunks"] > 0
        # Should be JSON-serializable
        json.dumps(j)

    def test_parse_from_bytes(self, simple_workbook):
        content = simple_workbook.read_bytes()
        result = parse_workbook(content=content, filename="test.xlsx")

        assert result.workbook.total_cells > 0
        assert result.workbook.filename == "test.xlsx"

    def test_serializer_records(self, simple_workbook):
        result = parse_workbook(path=simple_workbook)
        serializer = result.serializer

        wb_record = serializer.to_workbook_record()
        assert wb_record["id"] != ""
        assert wb_record["file_hash"] != ""

        sheet_records = serializer.to_sheet_records()
        assert len(sheet_records) == 1

        chunk_records = serializer.to_chunk_records()
        assert len(chunk_records) > 0

        vector_entries = serializer.to_vector_store_entries()
        assert len(vector_entries) > 0
        for entry in vector_entries:
            assert "id" in entry
            assert "text" in entry
            assert "metadata" in entry


class TestAssumptionsPipeline:
    """Test the pipeline on an assumptions/results workbook."""

    def test_multiple_blocks_detected(self, assumptions_workbook):
        result = parse_workbook(path=assumptions_workbook)
        assert result.total_chunks >= 2

    def test_dependency_context_in_chunks(self, assumptions_workbook):
        result = parse_workbook(path=assumptions_workbook)
        # At least one chunk should have upstream dependencies
        has_deps = any(
            len(c.dependency_summary.upstream_refs) > 0
            for c in result.chunks
        )
        assert has_deps
