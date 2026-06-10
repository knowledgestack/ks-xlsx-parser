"""
Tests for column-header span detection, including multi-row header bands.

The detector must (a) keep single-row headers exactly one row, (b) extend over
contiguous *multi-column* label rows that sit above the data, and (c) stop at
the first data row or a single-cell section divider — never swallowing data or
dividers into the header band.
"""
from __future__ import annotations

import io

from openpyxl import Workbook
from openpyxl.styles import Font

from excel_parser.analysis.header_detector import find_header_span
from excel_parser.parsers.workbook_parser import WorkbookParser


def _sheet(build):
    wb = Workbook()
    ws = wb.active
    ws.title = "S"
    build(ws)
    buf = io.BytesIO()
    wb.save(buf)
    return WorkbookParser(content=buf.getvalue(), filename="x.xlsx").parse().sheets[0]


def _span(sheet):
    return find_header_span(sheet, sheet.compute_used_range())


def test_single_row_header_stays_one_row():
    def build(ws):
        for ci, h in enumerate(["Item", "Q1", "Q2"], 1):
            ws.cell(row=1, column=ci, value=h).font = Font(bold=True)
        ws.cell(row=2, column=1, value="A")
        ws.cell(row=2, column=2, value=1)
        ws.cell(row=3, column=1, value="B")
        ws.cell(row=3, column=2, value=2)

    span = _span(_sheet(build))
    assert span is not None and (span.top, span.bottom) == (1, 1)


def test_multi_row_header_band():
    # Two stacked label rows (group label + sub-label) above numeric data.
    def build(ws):
        for ci, h in enumerate(["", "Revenue", "Revenue", "Cost"], 1):
            ws.cell(row=1, column=ci, value=h).font = Font(bold=True)
        for ci, h in enumerate(["Region", "2019", "2020", "2020"], 1):
            ws.cell(row=2, column=ci, value=h).font = Font(bold=True)
        ws.cell(row=3, column=1, value="North")
        for ci in range(2, 5):
            ws.cell(row=3, column=ci, value=ci * 10)
        ws.cell(row=4, column=1, value="South")
        for ci in range(2, 5):
            ws.cell(row=4, column=ci, value=ci * 20)

    span = _span(_sheet(build))
    assert span is not None and (span.top, span.bottom) == (1, 2)


def test_band_stops_at_data_row():
    def build(ws):
        for ci, h in enumerate(["Item", "V1", "V2"], 1):
            ws.cell(row=1, column=ci, value=h).font = Font(bold=True)
        # row 2 is already data — band must not extend past row 1
        ws.cell(row=2, column=1, value="A")
        ws.cell(row=2, column=2, value=1)
        ws.cell(row=2, column=3, value=2)

    span = _span(_sheet(build))
    assert span is not None and (span.top, span.bottom) == (1, 1)


def test_band_stops_at_single_cell_divider():
    # A single-cell "NORTH" divider directly under the header must NOT be
    # absorbed into the header band (it's a section divider, not a header row).
    def build(ws):
        for ci, h in enumerate(["Item", "V1", "V2"], 1):
            ws.cell(row=1, column=ci, value=h).font = Font(bold=True)
        ws.cell(row=2, column=1, value="NORTH")          # single-cell divider
        ws.cell(row=3, column=1, value="Widget")
        ws.cell(row=3, column=2, value=1)

    span = _span(_sheet(build))
    assert span is not None and (span.top, span.bottom) == (1, 1)


def test_single_row_header_above_text_data_stays_one_row():
    # Regression guard: a bold one-row header over a *text* first data row must
    # stay one row. The data row is unstyled, so the styled-band extension stops.
    def build(ws):
        for ci, h in enumerate(["Name", "City", "Role"], 1):
            ws.cell(row=1, column=ci, value=h).font = Font(bold=True)
        ws.cell(row=2, column=1, value="Alice")   # all-text data row, not bold
        ws.cell(row=2, column=2, value="Berlin")
        ws.cell(row=2, column=3, value="Eng")

    span = _span(_sheet(build))
    assert span is not None and (span.top, span.bottom) == (1, 1)


def test_styled_multi_row_header_band():
    # Two bold label rows over data → both rows are the header band.
    def build(ws):
        for ci, h in enumerate(["", "2019", "2020"], 1):
            ws.cell(row=1, column=ci, value=h).font = Font(bold=True)
        for ci, h in enumerate(["Region", "Rev", "Rev"], 1):
            ws.cell(row=2, column=ci, value=h).font = Font(bold=True)
        ws.cell(row=3, column=1, value="North")
        ws.cell(row=3, column=2, value=10)
        ws.cell(row=3, column=3, value=20)

    span = _span(_sheet(build))
    assert span is not None and (span.top, span.bottom) == (1, 2)


def test_no_header_for_pure_data_block():
    def build(ws):
        for r in range(1, 4):
            for c in range(1, 4):
                ws.cell(row=r, column=c, value=r * c)

    assert _span(_sheet(build)) is None
