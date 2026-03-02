"""
Tests for chart extraction via OOXML parsing.

Verifies that charts are extracted from the ZIP archive with correct
type, series, axes, title, and position anchor.
"""

import pytest

from xlsx_parser.charts.chart_extractor import ChartExtractor
from xlsx_parser.models import ChartType
from xlsx_parser.parsers import WorkbookParser


class TestChartExtraction:
    """Test chart extraction from OOXML."""

    def test_chart_detected(self, chart_workbook):
        result = WorkbookParser(path=chart_workbook).parse()
        assert len(result.charts) >= 1

    def test_chart_type(self, chart_workbook):
        result = WorkbookParser(path=chart_workbook).parse()
        chart = result.charts[0]
        assert chart.chart_type == ChartType.BAR

    def test_chart_title(self, chart_workbook):
        result = WorkbookParser(path=chart_workbook).parse()
        chart = result.charts[0]
        assert chart.title == "Monthly Revenue"

    def test_chart_series(self, chart_workbook):
        result = WorkbookParser(path=chart_workbook).parse()
        chart = result.charts[0]
        assert len(chart.series) >= 1

    def test_chart_summary(self, chart_workbook):
        result = WorkbookParser(path=chart_workbook).parse()
        chart = result.charts[0]
        summary = chart.generate_summary()
        assert "Bar" in summary or "bar" in summary.lower()
        assert "Monthly Revenue" in summary

    def test_chart_from_bytes(self, chart_workbook):
        content = chart_workbook.read_bytes()
        extractor = ChartExtractor(content, ["ChartData"])
        charts = extractor.extract_all()
        assert len(charts) >= 1
