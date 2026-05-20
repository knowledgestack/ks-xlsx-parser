"""Regression tests for `scripts.eval_retrieval.parse_position_spec`.

Covers cluster 00 in `docs/planning/recall-90/` — SpreadsheetBench
ground-truth strings that the harness historically failed to parse,
causing the parser to be blamed for benchmark-spec typos.

Reference table of malformed strings actually present in the dataset:
    50442    "RESULTS 1'!G17"
    49490    "Sheet1'!H3:H58"
    49036    "Dashboard'!B8"
    55427    "Compiled and located schools da'!B2:B1461"
    48975    "Output'!B11:B17"
    37456    'G12：J15'                                 (fullwidth colon)
    184-6    "'RAWDATA!'A1:P6,'OUTPUT!'A1:P6'"          (quote past `!`)
    164-22   "'YEAR1!'A1:G1478,'YEAR2!'A1:G1480'"       (same shape)
"""
from __future__ import annotations

import pytest

from scripts.eval_retrieval import parse_position_spec


@pytest.mark.parametrize(
    "spec, expected_sheet, expected_box",
    [
        ("RESULTS 1'!G17",                              "RESULTS 1",                         (17, 7, 17, 7)),
        ("Sheet1'!H3:H58",                              "Sheet1",                            (3, 8, 58, 8)),
        ("Dashboard'!B8",                               "Dashboard",                         (8, 2, 8, 2)),
        ("Compiled and located schools da'!B2:B1461",   "Compiled and located schools da",   (2, 2, 1461, 2)),
        ("Output'!B11:B17",                             "Output",                            (11, 2, 17, 2)),
    ],
)
def test_unmatched_closing_quote_recovers_sheet(spec, expected_sheet, expected_box):
    """`Foo'!A1` (stray closing quote, no opening) must yield ('Foo', A1)."""
    regions = parse_position_spec(spec, default_sheet="DEFAULT")
    assert regions, f"no regions for {spec!r}"
    assert regions[0][0] == expected_sheet
    assert regions[0][1] == expected_box


def test_fullwidth_colon_in_range():
    """SpreadsheetBench's Chinese-Excel entries use the fullwidth colon `：`."""
    regions = parse_position_spec("G12：J15", default_sheet="S")
    assert regions, "fullwidth colon range did not parse"
    sheet, box = regions[0]
    assert sheet == "S"
    assert box == (12, 7, 15, 10)


@pytest.mark.parametrize(
    "spec, expected",
    [
        (
            "'RAWDATA!'A1:P6,'OUTPUT!'A1:P6'",
            [("RAWDATA", (1, 1, 6, 16)), ("OUTPUT", (1, 1, 6, 16))],
        ),
        (
            "'YEAR1!'A1:G1478,'YEAR2!'A1:G1480'",
            [("YEAR1", (1, 1, 1478, 7)), ("YEAR2", (1, 1, 1480, 7))],
        ),
    ],
)
def test_quote_past_separator_multi_region(spec, expected):
    """`'Sheet!'A1:B2,'Other!'C3:D4'` — quote AFTER the `!`, comma-joined."""
    regions = parse_position_spec(spec, default_sheet=None)
    assert regions == expected, f"got {regions!r}"


# Controls — well-formed strings must keep working after the fix.

@pytest.mark.parametrize(
    "spec, expected",
    [
        ("A1:D10", [(None, (1, 1, 10, 4))]),
        ("'Sheet1'!A1:D10", [("Sheet1", (1, 1, 10, 4))]),
        ("Sheet1!A1:B2,Sheet2!C3:D4", [("Sheet1", (1, 1, 2, 2)), ("Sheet2", (3, 3, 4, 4))]),
        ("'summary'!D10", [("summary", (10, 4, 10, 4))]),
    ],
)
def test_well_formed_specs_still_parse(spec, expected):
    regions = parse_position_spec(spec, default_sheet=None)
    assert regions == expected


def test_empty_spec_returns_empty():
    assert parse_position_spec("", default_sheet="S") == []
    assert parse_position_spec("   ", default_sheet="S") == []


@pytest.mark.parametrize(
    "spec, expected",
    [
        # Whitespace around the colon — 6 dataset entries have this shape.
        ("A1: A89",     [(None, (1, 1, 89, 1))]),
        ("C1: E13",     [(None, (1, 3, 13, 5))]),
        ("A1 : D384",   [(None, (1, 1, 384, 4))]),
        # Fullwidth colon on bare range (matches 565-19 / 37456)
        ("D1：E7",      [(None, (1, 4, 7, 5))]),
    ],
)
def test_bare_range_whitespace_and_fullwidth(spec, expected):
    """SpreadsheetBench data_position strings sometimes use ' : ' or '：'.

    Six instances in the corpus had unparseable data_position before the
    parse_range whitespace fix: 330-23, 334-11, 347-49, 353-29, 565-19, 37456.
    """
    assert parse_position_spec(spec, default_sheet=None) == expected
