"""
Tests for DTOs and model utilities.

Covers CellCoord, CellRange, hashing, and serialization.
"""

from models import (
    CellCoord,
    CellRange,
    compute_hash,
    col_letter_to_number,
    col_number_to_letter,
)


class TestCellCoord:
    """Test cell coordinate conversions."""

    def test_a1_simple(self):
        assert CellCoord(row=1, col=1).to_a1() == "A1"

    def test_a1_double_letter(self):
        assert CellCoord(row=1, col=27).to_a1() == "AA1"

    def test_a1_large_col(self):
        assert CellCoord(row=100, col=26).to_a1() == "Z100"

    def test_a1_triple_letter(self):
        # Column 703 = AAA
        assert CellCoord(row=1, col=703).to_a1() == "AAA1"


class TestCellRange:
    """Test cell range operations."""

    def test_a1_range(self):
        r = CellRange(
            top_left=CellCoord(row=1, col=1),
            bottom_right=CellCoord(row=10, col=3),
        )
        assert r.to_a1() == "A1:C10"

    def test_contains(self):
        r = CellRange(
            top_left=CellCoord(row=2, col=2),
            bottom_right=CellCoord(row=5, col=5),
        )
        assert r.contains(CellCoord(row=3, col=3))
        assert not r.contains(CellCoord(row=1, col=1))
        assert not r.contains(CellCoord(row=6, col=3))

    def test_row_col_count(self):
        r = CellRange(
            top_left=CellCoord(row=1, col=1),
            bottom_right=CellCoord(row=5, col=3),
        )
        assert r.row_count() == 5
        assert r.col_count() == 3


class TestColumnConversion:
    """Test column letter ↔ number conversion."""

    def test_letter_to_number(self):
        assert col_letter_to_number("A") == 1
        assert col_letter_to_number("Z") == 26
        assert col_letter_to_number("AA") == 27
        assert col_letter_to_number("AZ") == 52
        assert col_letter_to_number("BA") == 53

    def test_number_to_letter(self):
        assert col_number_to_letter(1) == "A"
        assert col_number_to_letter(26) == "Z"
        assert col_number_to_letter(27) == "AA"
        assert col_number_to_letter(52) == "AZ"

    def test_roundtrip(self):
        for n in [1, 10, 26, 27, 100, 256, 702, 703]:
            assert col_letter_to_number(col_number_to_letter(n)) == n


class TestHashing:
    """Test deterministic hashing."""

    def test_same_input_same_hash(self):
        h1 = compute_hash("sheet1", "A1", "100")
        h2 = compute_hash("sheet1", "A1", "100")
        assert h1 == h2

    def test_different_input_different_hash(self):
        h1 = compute_hash("sheet1", "A1", "100")
        h2 = compute_hash("sheet1", "A1", "200")
        assert h1 != h2

    def test_hash_is_hex_string(self):
        h = compute_hash("test")
        assert isinstance(h, str)
        assert len(h) == 16  # xxhash64 hex digest
        int(h, 16)  # Should not raise
