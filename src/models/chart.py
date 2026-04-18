"""
Chart DTOs for Excel chart extraction.

Since openpyxl has limited chart *reading* support, charts are extracted
by parsing the OOXML `/xl/charts/chart*.xml` files directly. This module
defines the structured representation of chart data for both machine
consumption and RAG text summaries.
"""

from __future__ import annotations

from pydantic import Field

from .common import CellRange, ChartType, StableModel, compute_hash


class ChartSeries(StableModel):
    """A single data series within a chart."""

    series_index: int
    name: str | None = None
    name_ref: str | None = None  # Cell reference for series name
    values_ref: str | None = None  # Range reference for values (e.g., "Sheet1!$B$2:$B$13")
    categories_ref: str | None = None  # Range reference for category labels
    values_range: CellRange | None = None  # Parsed range coordinates
    categories_range: CellRange | None = None


class ChartAxis(StableModel):
    """An axis definition within a chart."""

    axis_type: str  # "category", "value", "date", "series"
    title: str | None = None
    position: str | None = None  # "bottom", "left", "right", "top"
    number_format: str | None = None
    min_value: float | None = None
    max_value: float | None = None
    major_unit: float | None = None


class ChartAnchor(StableModel):
    """Position of a chart on the worksheet in cell coordinates."""

    from_col: int
    from_row: int
    from_col_offset: int = 0  # EMU offset within the cell
    from_row_offset: int = 0
    to_col: int | None = None
    to_row: int | None = None
    to_col_offset: int = 0
    to_row_offset: int = 0


class ChartDTO(StableModel):
    """
    Complete representation of an Excel chart.

    Extracted from OOXML chart parts. Contains chart type, series definitions
    with data range references, axis configuration, and a generated text
    summary for RAG retrieval.
    """

    model_config = {"frozen": False, "extra": "forbid"}

    # Identity
    chart_index: int  # Index within the sheet's chart collection
    sheet_name: str
    chart_id: str = Field(default="", description="Deterministic ID")

    # Type
    chart_type: ChartType = ChartType.UNKNOWN

    # Content
    title: str | None = None
    series: list[ChartSeries] = Field(default_factory=list)
    axes: list[ChartAxis] = Field(default_factory=list)
    legend_position: str | None = None  # "bottom", "right", "top", "left"

    # Position
    anchor: ChartAnchor | None = None

    # Style
    style_id: int | None = None

    # Generated summary
    summary_text: str = Field(default="", description="Human-readable chart summary for RAG")

    # Hash
    content_hash: str = Field(default="")

    def generate_summary(self) -> str:
        """Generate a human-readable text summary of the chart."""
        parts = []
        type_label = self.chart_type.value.replace("_", " ").title()
        if self.title:
            parts.append(f"{type_label} chart: \"{self.title}\"")
        else:
            parts.append(f"{type_label} chart (untitled)")

        if self.series:
            series_names = [s.name or f"Series {s.series_index + 1}" for s in self.series]
            parts.append(f"Series: {', '.join(series_names)}")

            # Include data range info
            for s in self.series:
                if s.values_ref:
                    parts.append(f"  {s.name or f'Series {s.series_index + 1}'} data: {s.values_ref}")
                if s.categories_ref:
                    parts.append(f"  Categories: {s.categories_ref}")

        for axis in self.axes:
            if axis.title:
                parts.append(f"{axis.axis_type.title()} axis: \"{axis.title}\"")

        return "\n".join(parts)

    def finalize(self, workbook_hash: str) -> None:
        """Compute stable ID, content hash, and summary."""
        self.chart_id = compute_hash(
            workbook_hash, self.sheet_name, str(self.chart_index)
        )
        series_sig = "|".join(
            f"{s.name or ''}:{s.values_ref or ''}" for s in self.series
        )
        self.content_hash = compute_hash(
            self.sheet_name,
            self.chart_type.value,
            self.title or "",
            series_sig,
        )
        self.summary_text = self.generate_summary()
