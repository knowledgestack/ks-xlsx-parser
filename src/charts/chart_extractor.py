"""
Chart extractor using direct OOXML XML parsing.

openpyxl can create charts but has limited support for reading existing
charts from workbooks. This module opens the .xlsx as a ZIP archive and
parses the chart XML parts directly using lxml/ElementTree to extract
chart type, series, axes, titles, and position anchors.

OOXML Chart Structure:
  /xl/charts/chart1.xml  — chart definition
  /xl/drawings/drawing1.xml  — drawing with chart anchors (positions)
  /xl/_rels/workbook.xml.rels  — relationships
  /xl/worksheets/_rels/sheet1.xml.rels  — sheet→drawing relationships
"""

from __future__ import annotations

import io
import logging
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from models.chart import ChartAnchor, ChartAxis, ChartDTO, ChartSeries
from models.common import ChartType

logger = logging.getLogger(__name__)

# OOXML namespaces
NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "c16r2": "http://schemas.microsoft.com/office/drawing/2015/06/chart",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

# Chart type element names → ChartType enum
CHART_TYPE_MAP = {
    "barChart": ChartType.BAR,
    "bar3DChart": ChartType.BAR,
    "lineChart": ChartType.LINE,
    "line3DChart": ChartType.LINE,
    "pieChart": ChartType.PIE,
    "pie3DChart": ChartType.PIE,
    "areaChart": ChartType.AREA,
    "area3DChart": ChartType.AREA,
    "scatterChart": ChartType.SCATTER,
    "doughnutChart": ChartType.DOUGHNUT,
    "radarChart": ChartType.RADAR,
    "bubbleChart": ChartType.BUBBLE,
    "surfaceChart": ChartType.SURFACE,
    "surface3DChart": ChartType.SURFACE,
    "stockChart": ChartType.STOCK,
}


class ChartExtractor:
    """
    Extracts chart data from .xlsx files via direct OOXML parsing.

    Opens the file as a ZIP archive and parses:
    - /xl/charts/chart*.xml for chart definitions
    - /xl/drawings/drawing*.xml for chart position anchors
    - Relationship files to map charts to sheets

    Usage:
        extractor = ChartExtractor("workbook.xlsx", ["Sheet1", "Sheet2"])
        charts = extractor.extract_all()
    """

    def __init__(
        self,
        source: str | Path | bytes,
        sheet_names: list[str],
    ):
        """
        Args:
            source: Path to .xlsx file or raw bytes.
            sheet_names: Ordered list of sheet names for index→name mapping.
        """
        if isinstance(source, bytes):
            self._zip_source = io.BytesIO(source)
        else:
            self._zip_source = str(source)
        self._sheet_names = sheet_names

    def extract_all(self) -> list[ChartDTO]:
        """Extract all charts from all sheets."""
        charts: list[ChartDTO] = []
        try:
            with zipfile.ZipFile(self._zip_source, "r") as zf:
                # Map chart parts to sheets via relationships
                chart_to_sheet = self._map_charts_to_sheets(zf)

                # Map chart parts to anchors via drawing files
                chart_to_anchor = self._extract_anchors(zf)

                # Parse each chart XML
                chart_files = [
                    n for n in zf.namelist()
                    if n.startswith("xl/charts/chart") and n.endswith(".xml")
                ]
                chart_files.sort()  # Deterministic order

                for idx, chart_path in enumerate(chart_files):
                    try:
                        chart_xml = zf.read(chart_path)
                        sheet_name = chart_to_sheet.get(chart_path, self._sheet_names[0] if self._sheet_names else "Sheet1")
                        anchor = chart_to_anchor.get(chart_path)

                        dto = self._parse_chart(chart_xml, idx, sheet_name, anchor)
                        charts.append(dto)
                    except Exception as e:
                        logger.warning("Failed to parse chart %s: %s", chart_path, e)

        except (zipfile.BadZipFile, OSError) as e:
            logger.error("Cannot open file as ZIP for chart extraction: %s", e)

        return charts

    def _map_charts_to_sheets(self, zf: zipfile.ZipFile) -> dict[str, str]:
        """
        Build a mapping from chart XML paths to sheet names.

        Traverses: sheet.xml.rels → drawing.xml.rels → chart.xml
        """
        chart_to_sheet: dict[str, str] = {}

        for sheet_idx, sheet_name in enumerate(self._sheet_names):
            sheet_rels_path = f"xl/worksheets/_rels/sheet{sheet_idx + 1}.xml.rels"
            if sheet_rels_path not in zf.namelist():
                continue

            try:
                rels_xml = zf.read(sheet_rels_path)
                rels_root = ET.fromstring(rels_xml)

                for rel in rels_root.findall("{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
                    rel_type = rel.get("Type", "")
                    if "drawing" in rel_type:
                        drawing_target = rel.get("Target", "")
                        drawing_path = self._resolve_path("xl/worksheets", drawing_target)

                        # Now find charts referenced from this drawing
                        drawing_rels_path = drawing_path.replace(
                            "xl/drawings/", "xl/drawings/_rels/"
                        ) + ".rels"

                        if drawing_rels_path in zf.namelist():
                            drawing_rels_xml = zf.read(drawing_rels_path)
                            drawing_rels_root = ET.fromstring(drawing_rels_xml)

                            for drel in drawing_rels_root.findall(
                                "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
                            ):
                                if "chart" in drel.get("Type", ""):
                                    chart_target = drel.get("Target", "")
                                    chart_path = self._resolve_path("xl/drawings", chart_target)
                                    chart_to_sheet[chart_path] = sheet_name

            except Exception as e:
                logger.debug("Failed to parse rels for sheet %d: %s", sheet_idx, e)

        return chart_to_sheet

    def _extract_anchors(self, zf: zipfile.ZipFile) -> dict[str, ChartAnchor]:
        """Extract chart position anchors from drawing XML files."""
        anchors: dict[str, ChartAnchor] = {}

        drawing_files = [
            n for n in zf.namelist()
            if n.startswith("xl/drawings/drawing") and n.endswith(".xml")
        ]

        for drawing_path in drawing_files:
            try:
                drawing_xml = zf.read(drawing_path)
                root = ET.fromstring(drawing_xml)

                # Get the relationships for this drawing
                rels_path = drawing_path.replace(
                    "xl/drawings/", "xl/drawings/_rels/"
                ) + ".rels"
                rid_to_chart: dict[str, str] = {}

                if rels_path in zf.namelist():
                    rels_xml = zf.read(rels_path)
                    rels_root = ET.fromstring(rels_xml)
                    for rel in rels_root.findall(
                        "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
                    ):
                        if "chart" in rel.get("Type", ""):
                            rid = rel.get("Id", "")
                            target = rel.get("Target", "")
                            chart_path = self._resolve_path("xl/drawings", target)
                            rid_to_chart[rid] = chart_path

                # Parse twoCellAnchor elements
                for anchor_elem in root.findall("xdr:twoCellAnchor", NS):
                    chart_frame = anchor_elem.find(".//xdr:graphicFrame", NS)
                    if chart_frame is None:
                        continue

                    # Get the chart relationship ID
                    graphic = chart_frame.find(".//a:graphic/a:graphicData", NS)
                    if graphic is None:
                        continue
                    chart_ref = graphic.find("c:chart", NS)
                    if chart_ref is None:
                        continue
                    rid = chart_ref.get(f"{{{NS['r']}}}id", "")
                    chart_path = rid_to_chart.get(rid)
                    if not chart_path:
                        continue

                    # Extract from/to positions
                    from_elem = anchor_elem.find("xdr:from", NS)
                    to_elem = anchor_elem.find("xdr:to", NS)
                    if from_elem is not None:
                        anchor = ChartAnchor(
                            from_col=int(self._text(from_elem, "xdr:col") or "0"),
                            from_row=int(self._text(from_elem, "xdr:row") or "0"),
                            from_col_offset=int(self._text(from_elem, "xdr:colOff") or "0"),
                            from_row_offset=int(self._text(from_elem, "xdr:rowOff") or "0"),
                            to_col=int(self._text(to_elem, "xdr:col") or "0") if to_elem is not None else None,
                            to_row=int(self._text(to_elem, "xdr:row") or "0") if to_elem is not None else None,
                        )
                        anchors[chart_path] = anchor

            except Exception as e:
                logger.debug("Failed to parse drawing %s: %s", drawing_path, e)

        return anchors

    def _parse_chart(
        self,
        chart_xml: bytes,
        chart_index: int,
        sheet_name: str,
        anchor: ChartAnchor | None,
    ) -> ChartDTO:
        """Parse a single chart XML into a ChartDTO."""
        root = ET.fromstring(chart_xml)

        # Find the chart type and plot area
        chart_type = ChartType.UNKNOWN
        series_list: list[ChartSeries] = []
        axes: list[ChartAxis] = []

        # Extract title
        title = self._extract_title(root)

        # Find the plot area
        plot_area = root.find(".//c:plotArea", NS)
        if plot_area is not None:
            # Determine chart type from child elements
            for type_name, chart_enum in CHART_TYPE_MAP.items():
                type_elem = plot_area.find(f"c:{type_name}", NS)
                if type_elem is not None:
                    chart_type = chart_enum
                    series_list = self._extract_series(type_elem)
                    break

            # Extract axes
            axes = self._extract_axes(plot_area)

        # Extract legend
        legend_pos = None
        legend = root.find(".//c:legend", NS)
        if legend is not None:
            pos_elem = legend.find("c:legendPos", NS)
            if pos_elem is not None:
                legend_pos = pos_elem.get("val")

        return ChartDTO(
            chart_index=chart_index,
            sheet_name=sheet_name,
            chart_type=chart_type,
            title=title,
            series=series_list,
            axes=axes,
            legend_position=legend_pos,
            anchor=anchor,
        )

    def _extract_title(self, root: ET.Element) -> str | None:
        """Extract chart title text."""
        title_elem = root.find(".//c:title", NS)
        if title_elem is None:
            return None

        # Title can be in c:tx/c:rich/a:p/a:r/a:t (rich text)
        texts = []
        for t in title_elem.findall(".//a:t", NS):
            if t.text:
                texts.append(t.text)
        if texts:
            return " ".join(texts)

        # Or in c:tx/c:strRef (reference to a cell)
        str_ref = title_elem.find(".//c:strRef/c:f", NS)
        if str_ref is not None and str_ref.text:
            return f"[{str_ref.text}]"

        return None

    def _extract_series(self, chart_type_elem: ET.Element) -> list[ChartSeries]:
        """Extract data series from a chart type element."""
        series_list = []
        for idx, ser_elem in enumerate(chart_type_elem.findall("c:ser", NS)):
            name = None
            name_ref = None
            values_ref = None
            categories_ref = None

            # Series name (c:tx)
            tx = ser_elem.find("c:tx", NS)
            if tx is not None:
                str_ref = tx.find("c:strRef/c:f", NS)
                if str_ref is not None and str_ref.text:
                    name_ref = str_ref.text
                v_elem = tx.find("c:v", NS)
                if v_elem is not None and v_elem.text:
                    name = v_elem.text
                # Check strCache for cached name
                if not name:
                    cached = tx.find(".//c:strCache/c:pt/c:v", NS)
                    if cached is not None and cached.text:
                        name = cached.text

            # Values (c:val or c:yVal for scatter)
            for val_tag in ("c:val", "c:yVal"):
                val = ser_elem.find(val_tag, NS)
                if val is not None:
                    num_ref = val.find("c:numRef/c:f", NS)
                    if num_ref is not None and num_ref.text:
                        values_ref = num_ref.text
                    break

            # Categories (c:cat or c:xVal for scatter)
            for cat_tag in ("c:cat", "c:xVal"):
                cat = ser_elem.find(cat_tag, NS)
                if cat is not None:
                    for ref_tag in ("c:strRef/c:f", "c:numRef/c:f"):
                        ref = cat.find(ref_tag, NS)
                        if ref is not None and ref.text:
                            categories_ref = ref.text
                            break
                    break

            series_list.append(ChartSeries(
                series_index=idx,
                name=name,
                name_ref=name_ref,
                values_ref=values_ref,
                categories_ref=categories_ref,
            ))

        return series_list

    def _extract_axes(self, plot_area: ET.Element) -> list[ChartAxis]:
        """Extract axis definitions from the plot area."""
        axes = []
        axis_tags = {
            "c:catAx": "category",
            "c:valAx": "value",
            "c:dateAx": "date",
            "c:serAx": "series",
        }

        for tag, axis_type in axis_tags.items():
            for ax_elem in plot_area.findall(tag, NS):
                title = None
                title_elem = ax_elem.find("c:title", NS)
                if title_elem is not None:
                    texts = [t.text for t in title_elem.findall(".//a:t", NS) if t.text]
                    if texts:
                        title = " ".join(texts)

                position = None
                pos_elem = ax_elem.find("c:axPos", NS)
                if pos_elem is not None:
                    position = pos_elem.get("val")

                num_fmt = None
                fmt_elem = ax_elem.find("c:numFmt", NS)
                if fmt_elem is not None:
                    num_fmt = fmt_elem.get("formatCode")

                axes.append(ChartAxis(
                    axis_type=axis_type,
                    title=title,
                    position=position,
                    number_format=num_fmt,
                ))

        return axes

    @staticmethod
    def _resolve_path(base: str, relative: str) -> str:
        """Resolve a relative path within the OOXML package."""
        if relative.startswith("/"):
            return relative.lstrip("/")
        # Handle ../charts/chart1.xml → xl/charts/chart1.xml
        base_parts = base.rstrip("/").split("/")
        rel_parts = relative.split("/")
        for part in rel_parts:
            if part == "..":
                base_parts.pop()
            else:
                base_parts.append(part)
        return "/".join(base_parts)

    @staticmethod
    def _text(parent: ET.Element, tag: str) -> str | None:
        """Get text content of a child element."""
        elem = parent.find(tag, NS)
        return elem.text if elem is not None else None
