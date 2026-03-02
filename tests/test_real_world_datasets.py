"""
Tests against real-world Excel datasets from GitHub.

Source: https://github.com/rohanmistry231/Practice-Datasets-for-Excel

Validates that the parser produces correct, complete JSON output for
a variety of public datasets covering different shapes, sizes, and
content types (numeric, text, dates, mixed).
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from xlsx_parser.chunking.segmenter import LayoutSegmenter
from xlsx_parser.models import BlockType
from xlsx_parser.parsers import WorkbookParser
from xlsx_parser.pipeline import parse_workbook
from xlsx_parser.storage.serializer import WorkbookSerializer


FIXTURES_DIR = Path(__file__).parent / "fixtures" / "github_datasets"

# Each entry: (filename, expected_sheets, expected_min_rows, expected_header_sample)
DATASET_CATALOG = [
    ("iris.xlsx", 1, 150, ["sepal_length", "sepal_width", "petal_length"]),
    ("titanic.xlsx", 1, 891, ["PassengerId", "Survived", "Pclass"]),
    ("boston.xlsx", 1, 506, ["CRIM", "ZN", "INDUS"]),
    ("world_happiness_2019.xlsx", 1, 156, ["Overall rank", "Country or region", "Score"]),
    ("bestsellers.xlsx", 1, 550, ["Name", "Author", "User Rating"]),
    ("superstore.xlsx", 3, 1952, ["Row ID", "Order Priority", "Discount"]),
    ("worldcups.xlsx", 1, 20, ["Year", "Country", "Winner"]),
    ("breast_cancer.xlsx", 1, 569, ["id", "diagnosis", "radius_mean"]),
    ("apple_stock.xlsx", 1, 10016, ["Date", "Open", "High"]),
    ("winequality_red.xlsx", 1, 1599, None),  # semicolon-separated header, skip header check
]


def _fixture_path(name: str) -> Path:
    return FIXTURES_DIR / name


# ---------------------------------------------------------------------------
# Parametrized: every dataset parses without error
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    "filename,expected_sheets,expected_min_rows,expected_headers",
    DATASET_CATALOG,
    ids=[d[0].replace(".xlsx", "") for d in DATASET_CATALOG],
)
class TestDatasetParsing:
    """Core parsing validation across all datasets."""

    def test_parses_without_error(self, filename, expected_sheets, expected_min_rows, expected_headers):
        """Parser completes without raising an exception."""
        result = parse_workbook(path=_fixture_path(filename))
        assert result.workbook is not None

    def test_correct_sheet_count(self, filename, expected_sheets, expected_min_rows, expected_headers):
        """Workbook has the expected number of sheets."""
        result = parse_workbook(path=_fixture_path(filename))
        assert len(result.workbook.sheets) == expected_sheets

    def test_minimum_data_rows(self, filename, expected_sheets, expected_min_rows, expected_headers):
        """First sheet has at least the expected number of data rows."""
        result = parse_workbook(path=_fixture_path(filename))
        sheet = result.workbook.sheets[0]
        if sheet.used_range:
            data_rows = sheet.used_range.row_count() - 1  # minus header row
            assert data_rows >= expected_min_rows

    def test_headers_detected(self, filename, expected_sheets, expected_min_rows, expected_headers):
        """First row contains the expected column headers."""
        if expected_headers is None:
            pytest.skip("Header check skipped for this dataset")
        result = parse_workbook(path=_fixture_path(filename))
        sheet = result.workbook.sheets[0]
        first_row = sheet.used_range.top_left.row
        actual_headers = []
        for col in range(sheet.used_range.top_left.col, sheet.used_range.bottom_right.col + 1):
            cell = sheet.get_cell(first_row, col)
            if cell and cell.raw_value is not None:
                actual_headers.append(str(cell.raw_value))
        for expected in expected_headers:
            assert expected in actual_headers, (
                f"Expected header '{expected}' not found in {actual_headers[:10]}"
            )

    def test_produces_chunks(self, filename, expected_sheets, expected_min_rows, expected_headers):
        """Pipeline produces at least one chunk per sheet."""
        result = parse_workbook(path=_fixture_path(filename))
        assert result.total_chunks >= expected_sheets


# ---------------------------------------------------------------------------
# JSON serialization
# ---------------------------------------------------------------------------


class TestJsonSerialization:
    """Verify JSON output is valid, complete, and contains expected fields."""

    @pytest.mark.parametrize("filename", [d[0] for d in DATASET_CATALOG],
                             ids=[d[0].replace(".xlsx", "") for d in DATASET_CATALOG])
    def test_to_json_valid(self, filename):
        """to_json() returns a dict that round-trips through json.dumps/loads."""
        result = parse_workbook(path=_fixture_path(filename))
        data = result.to_json()
        json_str = json.dumps(data)
        roundtripped = json.loads(json_str)
        assert roundtripped["total_chunks"] == result.total_chunks

    @pytest.mark.parametrize("filename", [d[0] for d in DATASET_CATALOG],
                             ids=[d[0].replace(".xlsx", "") for d in DATASET_CATALOG])
    def test_to_json_has_required_keys(self, filename):
        """JSON output contains all required top-level keys."""
        result = parse_workbook(path=_fixture_path(filename))
        data = result.to_json()
        assert "workbook" in data
        assert "chunks" in data
        assert "total_chunks" in data
        assert "total_tokens" in data

    @pytest.mark.parametrize("filename", [d[0] for d in DATASET_CATALOG],
                             ids=[d[0].replace(".xlsx", "") for d in DATASET_CATALOG])
    def test_workbook_metadata_in_json(self, filename):
        """Workbook section has all required metadata fields."""
        result = parse_workbook(path=_fixture_path(filename))
        wb_json = result.to_json()["workbook"]
        assert wb_json["workbook_id"]
        assert wb_json["filename"]
        assert wb_json["workbook_hash"]
        assert isinstance(wb_json["total_sheets"], int)
        assert isinstance(wb_json["total_cells"], int)
        assert isinstance(wb_json["errors"], list)

    @pytest.mark.parametrize("filename", [d[0] for d in DATASET_CATALOG],
                             ids=[d[0].replace(".xlsx", "") for d in DATASET_CATALOG])
    def test_chunk_json_has_required_keys(self, filename):
        """Each chunk in JSON has all required fields."""
        result = parse_workbook(path=_fixture_path(filename))
        for chunk in result.to_json()["chunks"]:
            assert "chunk_id" in chunk
            assert "source_uri" in chunk
            assert "sheet_name" in chunk
            assert "block_type" in chunk
            assert "top_left" in chunk
            assert "bottom_right" in chunk
            assert "render_text" in chunk
            assert chunk["render_text"]  # not empty

    @pytest.mark.parametrize("filename", [d[0] for d in DATASET_CATALOG],
                             ids=[d[0].replace(".xlsx", "") for d in DATASET_CATALOG])
    def test_chunk_render_text_contains_data(self, filename):
        """Rendered text in chunks contains actual cell data, not just structure."""
        result = parse_workbook(path=_fixture_path(filename))
        sheet = result.workbook.sheets[0]
        # Get a data value from the sheet (short values to avoid semicolon-delimited lines)
        if sheet.used_range:
            first_data_row = sheet.used_range.top_left.row + 1
            for col in range(sheet.used_range.top_left.col, sheet.used_range.bottom_right.col + 1):
                cell = sheet.get_cell(first_data_row, col)
                if cell and cell.display_value and 2 < len(str(cell.display_value)) <= 30:
                    # At least one chunk should contain this value
                    found = any(
                        str(cell.display_value) in c.render_text
                        for c in result.chunks
                    )
                    assert found, f"Value '{cell.display_value}' not found in any chunk render_text"
                    return
        pytest.skip("No suitable data value found to check")


# ---------------------------------------------------------------------------
# Serializer records (Postgres-ready)
# ---------------------------------------------------------------------------


class TestSerializerRecords:
    """Verify WorkbookSerializer produces valid storage records."""

    @pytest.mark.parametrize("filename", [d[0] for d in DATASET_CATALOG],
                             ids=[d[0].replace(".xlsx", "") for d in DATASET_CATALOG])
    def test_workbook_record(self, filename):
        """Workbook record has all required fields for Postgres."""
        result = parse_workbook(path=_fixture_path(filename))
        serializer = WorkbookSerializer(result.workbook, result.chunks)
        rec = serializer.to_workbook_record()
        assert rec["id"]
        assert rec["file_hash"]
        assert rec["filename"]
        assert isinstance(rec["total_sheets"], int)
        assert isinstance(rec["total_cells"], int)
        # Ensure JSON-serializable
        json.dumps(rec)

    @pytest.mark.parametrize(
        "filename,expected_sheets",
        [(d[0], d[1]) for d in DATASET_CATALOG],
        ids=[d[0].replace(".xlsx", "") for d in DATASET_CATALOG],
    )
    def test_sheet_records_count(self, filename, expected_sheets):
        """Correct number of sheet records produced."""
        result = parse_workbook(path=_fixture_path(filename))
        serializer = WorkbookSerializer(result.workbook, result.chunks)
        sheets = serializer.to_sheet_records()
        assert len(sheets) == expected_sheets
        for s in sheets:
            assert s["sheet_name"]
            assert s["workbook_id"]
            json.dumps(s)

    @pytest.mark.parametrize("filename", [d[0] for d in DATASET_CATALOG],
                             ids=[d[0].replace(".xlsx", "") for d in DATASET_CATALOG])
    def test_chunk_records(self, filename):
        """Chunk records are valid and JSON-serializable."""
        result = parse_workbook(path=_fixture_path(filename))
        serializer = WorkbookSerializer(result.workbook, result.chunks)
        chunks = serializer.to_chunk_records()
        assert len(chunks) >= 1
        for c in chunks:
            assert c["id"]
            assert c["sheet_name"]
            assert c["block_type"]
            assert c["render_text"]
            json.dumps(c)

    @pytest.mark.parametrize("filename", [d[0] for d in DATASET_CATALOG],
                             ids=[d[0].replace(".xlsx", "") for d in DATASET_CATALOG])
    def test_vector_store_entries(self, filename):
        """Vector store entries have text and metadata for embedding."""
        result = parse_workbook(path=_fixture_path(filename))
        serializer = WorkbookSerializer(result.workbook, result.chunks)
        entries = serializer.to_vector_store_entries()
        assert len(entries) >= 1
        for e in entries:
            assert e["id"]
            assert e["text"]
            assert e["metadata"]["workbook_hash"]
            assert e["metadata"]["sheet_name"]
            assert e["metadata"]["source_uri"]
            json.dumps(e)


# ---------------------------------------------------------------------------
# Layout detection on real data
# ---------------------------------------------------------------------------


class TestRealWorldLayout:
    """Verify layout segmentation works correctly on real datasets."""

    @pytest.mark.parametrize("filename", [d[0] for d in DATASET_CATALOG],
                             ids=[d[0].replace(".xlsx", "") for d in DATASET_CATALOG])
    def test_blocks_have_valid_ranges(self, filename):
        """All detected blocks have non-degenerate cell ranges."""
        result = WorkbookParser(path=_fixture_path(filename)).parse()
        for sheet in result.sheets:
            tables = [t for t in result.tables if t.sheet_name == sheet.sheet_name]
            segmenter = LayoutSegmenter(sheet, tables=tables)
            blocks = segmenter.segment()
            for block in blocks:
                assert block.cell_range is not None
                assert block.cell_range.row_count() >= 1
                assert block.cell_range.col_count() >= 1
                assert block.cell_count > 0

    @pytest.mark.parametrize("filename", [d[0] for d in DATASET_CATALOG],
                             ids=[d[0].replace(".xlsx", "") for d in DATASET_CATALOG])
    def test_blocks_have_valid_types(self, filename):
        """All block types are valid BlockType enum values."""
        result = WorkbookParser(path=_fixture_path(filename)).parse()
        for sheet in result.sheets:
            tables = [t for t in result.tables if t.sheet_name == sheet.sheet_name]
            segmenter = LayoutSegmenter(sheet, tables=tables)
            blocks = segmenter.segment()
            valid_types = set(BlockType)
            for block in blocks:
                assert block.block_type in valid_types

    def test_superstore_multi_sheet_layout(self):
        """SuperStore has 3 sheets, each producing at least one block."""
        result = WorkbookParser(path=_fixture_path("superstore.xlsx")).parse()
        assert len(result.sheets) == 3
        for sheet in result.sheets:
            tables = [t for t in result.tables if t.sheet_name == sheet.sheet_name]
            segmenter = LayoutSegmenter(sheet, tables=tables)
            blocks = segmenter.segment()
            assert len(blocks) >= 1, f"Sheet '{sheet.sheet_name}' has no blocks"

    def test_world_happiness_has_table(self):
        """World Happiness dataset has an Excel ListObject table."""
        result = WorkbookParser(path=_fixture_path("world_happiness_2019.xlsx")).parse()
        assert len(result.tables) >= 1
        table = result.tables[0]
        assert table.table_name
        assert table.ref_range is not None


# ---------------------------------------------------------------------------
# Determinism on real data
# ---------------------------------------------------------------------------


class TestRealWorldDeterminism:
    """Parsing the same file twice produces identical output."""

    @pytest.mark.parametrize("filename", ["iris.xlsx", "worldcups.xlsx", "bestsellers.xlsx"],
                             ids=["iris", "worldcups", "bestsellers"])
    def test_deterministic_json(self, filename):
        """Two parses of the same file produce identical JSON (excluding timing)."""
        r1 = parse_workbook(path=_fixture_path(filename))
        r2 = parse_workbook(path=_fixture_path(filename))
        j1 = r1.to_json()
        j2 = r2.to_json()
        # parse_duration_ms varies between runs; exclude from comparison
        j1["workbook"]["parse_duration_ms"] = 0
        j2["workbook"]["parse_duration_ms"] = 0
        assert json.dumps(j1, sort_keys=True) == json.dumps(j2, sort_keys=True)

    @pytest.mark.parametrize("filename", ["iris.xlsx", "worldcups.xlsx", "bestsellers.xlsx"],
                             ids=["iris", "worldcups", "bestsellers"])
    def test_deterministic_hashes(self, filename):
        """Chunk IDs and content hashes are stable across runs."""
        r1 = parse_workbook(path=_fixture_path(filename))
        r2 = parse_workbook(path=_fixture_path(filename))
        assert r1.total_chunks == r2.total_chunks
        for c1, c2 in zip(r1.chunks, r2.chunks):
            assert c1.chunk_id == c2.chunk_id
            assert c1.content_hash == c2.content_hash


# ---------------------------------------------------------------------------
# Specific dataset content validation
# ---------------------------------------------------------------------------


class TestDatasetContent:
    """Spot-check specific known values in well-known datasets."""

    def test_iris_species_values(self):
        """Iris dataset contains known species names."""
        result = parse_workbook(path=_fixture_path("iris.xlsx"))
        sheet = result.workbook.sheets[0]
        species_col = None
        # Find the species column
        for col in range(1, 20):
            cell = sheet.get_cell(1, col)
            if cell and cell.raw_value == "species":
                species_col = col
                break
        assert species_col is not None, "species column not found"
        # Check known species
        species_values = set()
        for row in range(2, 152):
            cell = sheet.get_cell(row, species_col)
            if cell and cell.raw_value:
                species_values.add(cell.raw_value)
        assert "setosa" in species_values
        assert "versicolor" in species_values
        assert "virginica" in species_values

    def test_worldcups_has_known_winners(self):
        """WorldCups dataset contains known World Cup winners."""
        result = parse_workbook(path=_fixture_path("worldcups.xlsx"))
        sheet = result.workbook.sheets[0]
        winner_col = None
        for col in range(1, 20):
            cell = sheet.get_cell(1, col)
            if cell and cell.raw_value == "Winner":
                winner_col = col
                break
        assert winner_col is not None, "Winner column not found"
        winners = set()
        for row in range(2, 25):
            cell = sheet.get_cell(row, winner_col)
            if cell and cell.raw_value:
                winners.add(cell.raw_value)
        assert "Brazil" in winners
        assert "Germany" in winners

    def test_titanic_numeric_columns(self):
        """Titanic dataset has numeric columns (Survived, Pclass, Age)."""
        result = parse_workbook(path=_fixture_path("titanic.xlsx"))
        sheet = result.workbook.sheets[0]
        # Check Survived column has 0/1 values
        survived_col = None
        for col in range(1, 30):
            cell = sheet.get_cell(1, col)
            if cell and cell.raw_value == "Survived":
                survived_col = col
                break
        assert survived_col is not None
        cell_val = sheet.get_cell(2, survived_col)
        assert cell_val is not None
        assert cell_val.raw_value in (0, 1, 0.0, 1.0)

    def test_apple_stock_date_column(self):
        """Apple stock dataset has a Date column with date values."""
        result = parse_workbook(path=_fixture_path("apple_stock.xlsx"))
        sheet = result.workbook.sheets[0]
        date_col = None
        for col in range(1, 10):
            cell = sheet.get_cell(1, col)
            if cell and cell.raw_value == "Date":
                date_col = col
                break
        assert date_col is not None
        # Check that at least one date cell has a date-like display value
        date_cell = sheet.get_cell(2, date_col)
        assert date_cell is not None
        assert date_cell.display_value is not None

    def test_superstore_multiple_sheets_content(self):
        """SuperStore has Orders, Returns, and Users sheets with distinct content."""
        result = parse_workbook(path=_fixture_path("superstore.xlsx"))
        sheet_names = {s.sheet_name for s in result.workbook.sheets}
        assert "Orders" in sheet_names
        assert "Returns" in sheet_names
        assert "Users" in sheet_names

        # Orders sheet should be large
        orders = next(s for s in result.workbook.sheets if s.sheet_name == "Orders")
        assert orders.cell_count() > 40000

        # Users sheet should be small
        users = next(s for s in result.workbook.sheets if s.sheet_name == "Users")
        assert users.cell_count() <= 20
