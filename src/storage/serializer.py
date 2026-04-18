"""
Serialization and storage mapping.

Converts WorkbookDTO + ChunkDTO objects into storage-ready representations:
- JSON dicts for Postgres JSONB columns
- Flat records for relational tables
- Vector store entries for embedding

Also provides schema definitions for the target Postgres tables.
"""

from __future__ import annotations

import json
import logging
from typing import Any

from models.block import ChunkDTO
from models.workbook import WorkbookDTO

logger = logging.getLogger(__name__)


# Postgres schema DDL (for reference / migration scripts)
POSTGRES_SCHEMA = """
CREATE TABLE IF NOT EXISTS workbooks (
    id TEXT PRIMARY KEY,
    file_hash TEXT NOT NULL UNIQUE,
    filename TEXT NOT NULL,
    properties JSONB,
    total_sheets INT,
    total_cells INT,
    total_formulas INT,
    parse_duration_ms FLOAT,
    errors JSONB,
    created_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS sheets (
    id TEXT PRIMARY KEY,
    workbook_id TEXT NOT NULL REFERENCES workbooks(id),
    sheet_name TEXT NOT NULL,
    sheet_index INT NOT NULL,
    used_range TEXT,
    cell_count INT,
    properties JSONB,
    hidden_rows JSONB,
    hidden_cols JSONB,
    merged_regions JSONB,
    conditional_formats JSONB,
    data_validations JSONB,
    errors JSONB
);

CREATE TABLE IF NOT EXISTS blocks (
    id TEXT PRIMARY KEY,
    sheet_id TEXT NOT NULL REFERENCES sheets(id),
    workbook_id TEXT NOT NULL REFERENCES workbooks(id),
    block_type TEXT NOT NULL,
    top_left TEXT NOT NULL,
    bottom_right TEXT NOT NULL,
    cell_count INT,
    formula_count INT,
    content_hash TEXT NOT NULL,
    render_html TEXT,
    render_text TEXT,
    token_count INT,
    metadata JSONB,
    prev_block_id TEXT,
    next_block_id TEXT
);

CREATE TABLE IF NOT EXISTS cells (
    id TEXT PRIMARY KEY,
    sheet_id TEXT NOT NULL REFERENCES sheets(id),
    row_num INT NOT NULL,
    col_num INT NOT NULL,
    raw_value TEXT,
    display_value TEXT,
    formula TEXT,
    formula_value TEXT,
    data_type TEXT,
    style JSONB,
    comment_text TEXT,
    hyperlink TEXT,
    is_merged_master BOOLEAN DEFAULT FALSE,
    is_merged_slave BOOLEAN DEFAULT FALSE
);

CREATE TABLE IF NOT EXISTS dependencies (
    id TEXT PRIMARY KEY,
    source_sheet TEXT NOT NULL,
    source_cell TEXT NOT NULL,
    target_sheet TEXT,
    target_ref TEXT NOT NULL,
    edge_type TEXT NOT NULL,
    external_workbook TEXT,
    named_range_name TEXT
);

CREATE TABLE IF NOT EXISTS charts (
    id TEXT PRIMARY KEY,
    sheet_id TEXT NOT NULL REFERENCES sheets(id),
    chart_type TEXT NOT NULL,
    title TEXT,
    series JSONB,
    axes JSONB,
    anchor JSONB,
    summary_text TEXT,
    content_hash TEXT
);

CREATE TABLE IF NOT EXISTS tables_def (
    id TEXT PRIMARY KEY,
    sheet_id TEXT NOT NULL REFERENCES sheets(id),
    table_name TEXT NOT NULL,
    display_name TEXT,
    ref_range TEXT NOT NULL,
    columns JSONB,
    has_totals_row BOOLEAN DEFAULT FALSE,
    style_name TEXT,
    content_hash TEXT
);

CREATE TABLE IF NOT EXISTS named_ranges (
    id TEXT PRIMARY KEY,
    workbook_id TEXT NOT NULL REFERENCES workbooks(id),
    name TEXT NOT NULL,
    ref_string TEXT NOT NULL,
    scope_sheet TEXT,
    is_hidden BOOLEAN DEFAULT FALSE
);

CREATE INDEX IF NOT EXISTS idx_sheets_workbook ON sheets(workbook_id);
CREATE INDEX IF NOT EXISTS idx_blocks_sheet ON blocks(sheet_id);
CREATE INDEX IF NOT EXISTS idx_cells_sheet ON cells(sheet_id);
CREATE INDEX IF NOT EXISTS idx_deps_source ON dependencies(source_sheet, source_cell);
CREATE INDEX IF NOT EXISTS idx_deps_target ON dependencies(target_sheet, target_ref);
CREATE INDEX IF NOT EXISTS idx_blocks_hash ON blocks(content_hash);
"""


class WorkbookSerializer:
    """
    Serializes parsed workbook data into storage-ready formats.

    Provides methods to convert WorkbookDTO and ChunkDTO objects
    into flat dicts suitable for Postgres insertion and vector
    store upsert.
    """

    def __init__(self, workbook: WorkbookDTO, chunks: list[ChunkDTO]):
        self._workbook = workbook
        self._chunks = chunks

    def to_workbook_record(self) -> dict[str, Any]:
        """Serialize workbook-level data for the `workbooks` table."""
        wb = self._workbook
        return {
            "id": wb.workbook_id,
            "file_hash": wb.workbook_hash,
            "filename": wb.filename,
            "properties": json.loads(wb.properties.model_dump_json(exclude_none=True)),
            "total_sheets": wb.total_sheets,
            "total_cells": wb.total_cells,
            "total_formulas": wb.total_formulas,
            "parse_duration_ms": wb.parse_duration_ms,
            "errors": [json.loads(e.model_dump_json(exclude_none=True)) for e in wb.errors],
        }

    def to_sheet_records(self) -> list[dict[str, Any]]:
        """Serialize sheet data for the `sheets` table."""
        records = []
        for sheet in self._workbook.sheets:
            records.append({
                "id": sheet.sheet_id,
                "workbook_id": self._workbook.workbook_id,
                "sheet_name": sheet.sheet_name,
                "sheet_index": sheet.sheet_index,
                "used_range": sheet.used_range.to_a1() if sheet.used_range else None,
                "cell_count": sheet.cell_count(),
                "properties": json.loads(sheet.properties.model_dump_json(exclude_none=True)),
                "hidden_rows": sorted(sheet.hidden_rows),
                "hidden_cols": sorted(sheet.hidden_cols),
                "merged_regions": [
                    {"range": m.range.to_a1(), "master": m.master.to_a1()}
                    for m in sheet.merged_regions
                ],
                "conditional_formats": [
                    json.loads(cf.model_dump_json(exclude_none=True))
                    for cf in sheet.conditional_format_rules
                ],
                "data_validations": [
                    json.loads(dv.model_dump_json(exclude_none=True))
                    for dv in sheet.data_validations
                ],
                "errors": [json.loads(e.model_dump_json(exclude_none=True)) for e in sheet.errors],
            })
        return records

    def to_chunk_records(self) -> list[dict[str, Any]]:
        """
        Serialize chunks for the `blocks` table.

        Each record includes the rendered content and metadata
        needed for RAG retrieval.
        """
        records = []
        for chunk in self._chunks:
            records.append({
                "id": chunk.chunk_id,
                "sheet_name": chunk.sheet_name,
                "workbook_hash": chunk.workbook_hash,
                "block_type": chunk.block_type if isinstance(chunk.block_type, str) else chunk.block_type.value,
                "top_left": chunk.top_left_cell,
                "bottom_right": chunk.bottom_right_cell,
                "source_uri": chunk.source_uri,
                "content_hash": chunk.content_hash,
                "render_html": chunk.render_html,
                "render_text": chunk.render_text,
                "token_count": chunk.token_count,
                "key_cells": chunk.key_cells,
                "named_ranges": chunk.named_ranges,
                "dependency_summary": json.loads(
                    chunk.dependency_summary.model_dump_json(exclude_none=True)
                ),
                "prev_chunk_id": chunk.prev_chunk_id,
                "next_chunk_id": chunk.next_chunk_id,
                "metadata": chunk.metadata,
            })
        return records

    def to_vector_store_entries(self) -> list[dict[str, Any]]:
        """
        Prepare entries for a vector store (e.g., Pinecone, Weaviate, pgvector).

        Each entry includes the text to embed and metadata for filtering.
        """
        entries = []
        for chunk in self._chunks:
            entries.append({
                "id": chunk.chunk_id,
                "text": chunk.render_text,
                "metadata": {
                    "workbook_hash": chunk.workbook_hash,
                    "sheet_name": chunk.sheet_name,
                    "block_type": chunk.block_type if isinstance(chunk.block_type, str) else chunk.block_type.value,
                    "source_uri": chunk.source_uri,
                    "top_left": chunk.top_left_cell,
                    "bottom_right": chunk.bottom_right_cell,
                    "token_count": chunk.token_count,
                    "content_hash": chunk.content_hash,
                    "has_formulas": chunk.dependency_summary.upstream_refs != [],
                },
            })
        return entries

    @staticmethod
    def get_schema_ddl() -> str:
        """Return the Postgres schema DDL for all tables."""
        return POSTGRES_SCHEMA
