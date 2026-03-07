"""
LLM-ready derived artifacts for workbook understanding.

Provides four analyzers that produce high-level semantic summaries:
  - SheetSummaryAnalyzer: detects sheet purpose and produces summaries
  - EntityIndexBuilder: extracts business entities from headers/names
  - KpiCatalogBuilder: identifies candidate KPI cells
  - ReadingOrderLinearizer: produces slide-like text linearization
"""

from __future__ import annotations

import logging
import re
from collections import Counter

from pydantic import Field

from ..models.block import BlockDTO
from ..models.chart import ChartDTO
from ..models.common import (
    BlockType,
    CellCoord,
    SheetPurpose,
    StableModel,
    compute_hash,
)
from ..models.dependency import DependencyGraph
from ..models.sheet import SheetDTO
from ..models.table import TableDTO
from ..models.workbook import KpiDTO, SheetSummaryDTO

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# 1. Sheet Summary Analyzer
# ---------------------------------------------------------------------------


class SheetSummaryAnalyzer:
    """
    Analyzes sheets to determine their purpose and produce concise summaries.

    Uses heuristics based on formula density, chart presence, data validation,
    cross-sheet dependency patterns, and formatting.
    """

    def __init__(
        self,
        sheet: SheetDTO,
        charts: list[ChartDTO] | None = None,
        tables: list[TableDTO] | None = None,
        blocks: list[BlockDTO] | None = None,
        dependency_graph: DependencyGraph | None = None,
    ):
        self._sheet = sheet
        self._charts = [c for c in (charts or []) if c.sheet_name == sheet.sheet_name]
        self._tables = [t for t in (tables or []) if t.sheet_name == sheet.sheet_name]
        self._blocks = blocks or []
        self._dep_graph = dependency_graph or DependencyGraph()

    def analyze(self) -> SheetSummaryDTO:
        total_cells = len(self._sheet.cells)
        formula_count = sum(
            1 for c in self._sheet.cells.values() if c.formula is not None
        )
        formula_density = formula_count / total_cells if total_cells > 0 else 0.0

        has_dv = len(self._sheet.data_validations) > 0
        has_charts = len(self._charts) > 0
        has_print_area = self._sheet.properties.print_area is not None

        purpose, confidence = self._detect_purpose(
            total_cells, formula_count, formula_density, has_dv, has_charts, has_print_area,
        )

        key_tables = [t.display_name or t.table_name for t in self._tables[:5]]
        key_outputs = self._identify_key_outputs()
        key_entities = self._extract_entities()

        summary_text = self._generate_summary(
            purpose, key_tables, key_outputs, total_cells, formula_count,
        )

        return SheetSummaryDTO(
            sheet_name=self._sheet.sheet_name,
            purpose=purpose,
            purpose_confidence=round(confidence, 2),
            total_cells=total_cells,
            formula_count=formula_count,
            formula_density=round(formula_density, 4),
            has_data_validation=has_dv,
            has_charts=has_charts,
            has_print_area=has_print_area,
            key_tables=key_tables,
            key_output_cells=[c.cell_ref for c in key_outputs[:5]],
            key_entities=key_entities[:10],
            summary_text=summary_text,
        )

    def _detect_purpose(
        self,
        total_cells: int,
        formula_count: int,
        formula_density: float,
        has_dv: bool,
        has_charts: bool,
        has_print_area: bool,
    ) -> tuple[SheetPurpose, float]:
        scores: dict[SheetPurpose, float] = {p: 0.0 for p in SheetPurpose}

        if has_charts and formula_density < 0.3:
            scores[SheetPurpose.DASHBOARD] += 0.6
        if has_charts:
            scores[SheetPurpose.DASHBOARD] += 0.2

        if formula_density < 0.05 and total_cells > 20:
            scores[SheetPurpose.RAW_DATA] += 0.5
        if formula_density < 0.01 and total_cells > 50:
            scores[SheetPurpose.RAW_DATA] += 0.3
        if formula_density < 0.15 and len(self._tables) > 0:
            scores[SheetPurpose.RAW_DATA] += 0.3

        cross_in = self._cross_sheet_in_degree()
        if cross_in > 3:
            scores[SheetPurpose.LOOKUP] += 0.7
        elif cross_in > 0:
            scores[SheetPurpose.LOOKUP] += 0.3

        if formula_density > 0.3:
            scores[SheetPurpose.CALCULATION] += 0.5
        if formula_density > 0.5 and not has_charts:
            scores[SheetPurpose.CALCULATION] += 0.3

        if has_dv:
            scores[SheetPurpose.INPUT] += 0.5
        if has_dv and formula_density < 0.2:
            scores[SheetPurpose.INPUT] += 0.3

        if has_print_area:
            scores[SheetPurpose.REPORT] += 0.4

        has_protection = self._sheet.properties.sheet_protection
        if has_protection and has_dv:
            scores[SheetPurpose.TEMPLATE] += 0.5

        if total_cells < 15:
            scores[SheetPurpose.CONFIG] += 0.3

        best = max(scores, key=lambda p: scores[p])
        best_score = scores[best]
        if best_score < 0.1:
            return SheetPurpose.UNKNOWN, 0.0
        return best, min(best_score, 1.0)

    def _cross_sheet_in_degree(self) -> int:
        count = 0
        for edge in self._dep_graph.edges:
            target_sheet = edge.target_sheet or edge.source_sheet
            if (
                target_sheet == self._sheet.sheet_name
                and edge.source_sheet != self._sheet.sheet_name
            ):
                count += 1
        return count

    def _identify_key_outputs(self) -> list[KpiDTO]:
        in_degree: Counter[str] = Counter()
        for edge in self._dep_graph.edges:
            target_sheet = edge.target_sheet or edge.source_sheet
            if target_sheet == self._sheet.sheet_name and edge.target_coord:
                key = f"{edge.target_coord.row},{edge.target_coord.col}"
                in_degree[key] += 1

        candidates: list[KpiDTO] = []
        for cell in self._sheet.cells.values():
            cell_key = f"{cell.coord.row},{cell.coord.col}"
            deg = in_degree.get(cell_key, 0)
            score = 0
            if deg >= 3:
                score += 3
            if cell.formula and cell.style and cell.style.font and cell.style.font.bold:
                score += 2
            if cell.style and cell.style.number_format:
                nf = cell.style.number_format
                if "$" in nf or "%" in nf:
                    score += 1
            if score >= 2:
                label = self._find_label(cell.coord)
                candidates.append(KpiDTO(
                    label=label,
                    cell_ref=f"{self._sheet.sheet_name}!{cell.coord.to_a1()}",
                    value_display=cell.display_value,
                    sheet_name=self._sheet.sheet_name,
                    in_degree=deg,
                ))

        return sorted(candidates, key=lambda c: c.in_degree, reverse=True)

    def _find_label(self, coord: CellCoord) -> str | None:
        left = self._sheet.get_cell(coord.row, coord.col - 1)
        if left and isinstance(left.raw_value, str):
            return left.raw_value
        above = self._sheet.get_cell(coord.row - 1, coord.col)
        if above and isinstance(above.raw_value, str):
            return above.raw_value
        return None

    def _extract_entities(self) -> list[str]:
        """Extract business entity names from header-like cells."""
        entities: set[str] = set()
        for cell in self._sheet.cells.values():
            if (
                cell.style
                and cell.style.font
                and cell.style.font.bold
                and isinstance(cell.raw_value, str)
                and len(cell.raw_value) < 50
            ):
                entities.add(cell.raw_value.strip())
        # Also from table column names
        for table in self._tables:
            for col in table.columns:
                if col.name:
                    entities.add(col.name)
        return sorted(entities)

    def _generate_summary(
        self,
        purpose: SheetPurpose,
        key_tables: list[str],
        key_outputs: list[KpiDTO],
        total_cells: int,
        formula_count: int,
    ) -> str:
        parts: list[str] = []
        label = purpose.value.replace("_", " ").title()
        parts.append(
            f'Sheet "{self._sheet.sheet_name}" ({label}): '
            f"{total_cells} cells, {formula_count} formulas."
        )
        if key_tables:
            parts.append(f"Tables: {', '.join(key_tables[:3])}.")
        if key_outputs:
            descs = []
            for kpi in key_outputs[:3]:
                d = kpi.cell_ref
                if kpi.label:
                    d = f"{kpi.label} ({kpi.cell_ref})"
                descs.append(d)
            parts.append(f"Key outputs: {'; '.join(descs)}.")
        return " ".join(parts)


# ---------------------------------------------------------------------------
# 2. Entity Index Builder
# ---------------------------------------------------------------------------


class EntityLocation(StableModel):
    """Where a business entity appears."""

    model_config = {"frozen": True, "extra": "forbid"}

    sheet_name: str
    range_a1: str | None = None
    source: str = "header"  # "header", "named_range", "table_column"


class EntityEntry(StableModel):
    """An extracted business entity and its locations."""

    model_config = {"frozen": False, "extra": "forbid"}

    name: str
    category: str = "unknown"  # "dimension", "measure", "metadata"
    locations: list[EntityLocation] = Field(default_factory=list)


class EntityIndexDTO(StableModel):
    """Index of business entities extracted from the workbook."""

    model_config = {"frozen": False, "extra": "forbid"}

    entities: list[EntityEntry] = Field(default_factory=list)
    entity_hash: str = Field(default="")

    def finalize(self, workbook_hash: str) -> None:
        names = "|".join(e.name for e in self.entities)
        self.entity_hash = compute_hash(workbook_hash, names)


class EntityIndexBuilder:
    """Extracts business entities from headers, named ranges, and table columns."""

    # Words that indicate a measure (numeric quantity)
    _MEASURE_KEYWORDS = {
        "revenue", "cost", "price", "amount", "total", "sum", "profit",
        "margin", "sales", "volume", "count", "qty", "quantity", "rate",
        "percent", "percentage", "value", "score", "budget", "forecast",
        "actual", "variance", "target", "balance", "tax", "fee", "income",
        "expense", "growth", "yield", "roi", "ebitda", "net", "gross",
    }

    def __init__(
        self,
        sheets: list[SheetDTO],
        tables: list[TableDTO] | None = None,
        named_ranges: list | None = None,
    ):
        self._sheets = sheets
        self._tables = tables or []
        self._named_ranges = named_ranges or []

    def build(self) -> EntityIndexDTO:
        entity_map: dict[str, EntityEntry] = {}

        # From table columns
        for table in self._tables:
            for col in table.columns:
                name = col.name.strip() if col.name else None
                if name and len(name) > 1:
                    key = name.lower()
                    if key not in entity_map:
                        entity_map[key] = EntityEntry(
                            name=name,
                            category=self._categorize(name),
                        )
                    entity_map[key].locations.append(EntityLocation(
                        sheet_name=table.sheet_name,
                        range_a1=table.ref_range.to_a1() if table.ref_range else None,
                        source="table_column",
                    ))

        # From bold header cells
        for sheet in self._sheets:
            for cell in sheet.cells.values():
                if (
                    cell.style
                    and cell.style.font
                    and cell.style.font.bold
                    and isinstance(cell.raw_value, str)
                    and 1 < len(cell.raw_value.strip()) < 50
                ):
                    name = cell.raw_value.strip()
                    key = name.lower()
                    if key not in entity_map:
                        entity_map[key] = EntityEntry(
                            name=name,
                            category=self._categorize(name),
                        )
                    entity_map[key].locations.append(EntityLocation(
                        sheet_name=sheet.sheet_name,
                        range_a1=cell.coord.to_a1(),
                        source="header",
                    ))

        # From named ranges
        for nr in self._named_ranges:
            name = nr.name.strip() if hasattr(nr, "name") else str(nr)
            if name and len(name) > 1 and not name.startswith("_"):
                key = name.lower()
                if key not in entity_map:
                    entity_map[key] = EntityEntry(
                        name=name,
                        category=self._categorize(name),
                    )
                entity_map[key].locations.append(EntityLocation(
                    sheet_name=nr.scope_sheet or "(workbook)",
                    range_a1=nr.ref_string if hasattr(nr, "ref_string") else None,
                    source="named_range",
                ))

        return EntityIndexDTO(entities=sorted(entity_map.values(), key=lambda e: e.name))

    def _categorize(self, name: str) -> str:
        lower = name.lower()
        for kw in self._MEASURE_KEYWORDS:
            if kw in lower:
                return "measure"
        return "dimension"


# ---------------------------------------------------------------------------
# 3. KPI Catalog Builder
# ---------------------------------------------------------------------------


class KpiCatalogBuilder:
    """
    Identifies candidate KPI cells based on formatting, dependency, and
    chart-reference signals.
    """

    def __init__(
        self,
        sheets: list[SheetDTO],
        charts: list[ChartDTO] | None = None,
        dependency_graph: DependencyGraph | None = None,
    ):
        self._sheets = sheets
        self._charts = charts or []
        self._dep_graph = dependency_graph or DependencyGraph()

    def build(self) -> list[KpiDTO]:
        # Pre-compute chart-referenced cells
        chart_ref_cells: set[str] = set()
        for chart in self._charts:
            for series in chart.series:
                if series.values_range:
                    rng = series.values_range
                    for r in range(rng.top_left.row, rng.bottom_right.row + 1):
                        for c in range(rng.top_left.col, rng.bottom_right.col + 1):
                            chart_ref_cells.add(f"{chart.sheet_name}|{r},{c}")

        # Compute in-degree per cell
        in_degree: Counter[str] = Counter()
        for edge in self._dep_graph.edges:
            ts = edge.target_sheet or edge.source_sheet
            if edge.target_coord:
                in_degree[f"{ts}|{edge.target_coord.row},{edge.target_coord.col}"] += 1

        candidates: list[KpiDTO] = []
        for sheet in self._sheets:
            for cell in sheet.cells.values():
                cell_key = f"{sheet.sheet_name}|{cell.coord.row},{cell.coord.col}"
                score = 0
                deg = in_degree.get(cell_key, 0)

                if deg >= 5:
                    score += 4
                elif deg >= 3:
                    score += 3
                elif deg >= 1:
                    score += 1

                if cell_key in chart_ref_cells:
                    score += 2

                if cell.formula and cell.style and cell.style.font and cell.style.font.bold:
                    score += 2

                if cell.style and cell.style.number_format:
                    nf = cell.style.number_format
                    if "$" in nf:
                        score += 2
                    elif "%" in nf:
                        score += 1

                if cell.style and cell.style.font and cell.style.font.size and cell.style.font.size >= 14:
                    score += 1

                if score >= 3:
                    label = self._find_label(sheet, cell.coord)
                    drivers = self._find_drivers(cell_key)
                    candidates.append(KpiDTO(
                        label=label,
                        cell_ref=f"{sheet.sheet_name}!{cell.coord.to_a1()}",
                        value_display=cell.display_value,
                        sheet_name=sheet.sheet_name,
                        in_degree=deg,
                        drivers=drivers[:5],
                    ))

        return sorted(candidates, key=lambda k: k.in_degree, reverse=True)

    @staticmethod
    def _find_label(sheet: SheetDTO, coord: CellCoord) -> str | None:
        left = sheet.get_cell(coord.row, coord.col - 1)
        if left and isinstance(left.raw_value, str):
            return left.raw_value
        above = sheet.get_cell(coord.row - 1, coord.col)
        if above and isinstance(above.raw_value, str):
            return above.raw_value
        return None

    def _find_drivers(self, cell_key: str) -> list[str]:
        """Find cells that this cell depends on (its inputs)."""
        parts = cell_key.split("|")
        if len(parts) != 2:
            return []
        sheet_name = parts[0]
        rc = parts[1].split(",")
        if len(rc) != 2:
            return []
        row, col = int(rc[0]), int(rc[1])

        drivers: list[str] = []
        for edge in self._dep_graph.edges:
            if (
                edge.source_sheet == sheet_name
                and edge.source_coord
                and edge.source_coord.row == row
                and edge.source_coord.col == col
            ):
                ts = edge.target_sheet or edge.source_sheet
                if edge.target_coord:
                    drivers.append(f"{ts}!{edge.target_coord.to_a1()}")
        return drivers


# ---------------------------------------------------------------------------
# 4. Reading-Order Linearizer
# ---------------------------------------------------------------------------


class ReadingOrderLinearizer:
    """
    Produces a slide-like text linearization of a sheet.

    Interleaves titles, block headers, key values, chart summaries,
    and notes in top-to-bottom, left-to-right reading order with
    stable anchors back to cell ranges.
    """

    def __init__(
        self,
        sheet: SheetDTO,
        charts: list[ChartDTO] | None = None,
        tables: list[TableDTO] | None = None,
        blocks: list[BlockDTO] | None = None,
    ):
        self._sheet = sheet
        self._charts = [c for c in (charts or []) if c.sheet_name == sheet.sheet_name]
        self._tables = [t for t in (tables or []) if t.sheet_name == sheet.sheet_name]
        self._blocks = blocks or []

    def linearize(self) -> str:
        """Produce a structured text representation of the sheet."""
        lines: list[str] = []
        lines.append(f"## Sheet: {self._sheet.sheet_name}")

        if not self._sheet.cells:
            lines.append("(empty sheet)")
            return "\n".join(lines)

        # Collect content items with row positions for ordering
        items: list[tuple[int, int, str]] = []  # (row, col, text)

        # Tables
        for table in self._tables:
            r = table.ref_range.top_left.row if table.ref_range else 0
            c = table.ref_range.top_left.col if table.ref_range else 0
            range_str = table.ref_range.to_a1() if table.ref_range else "?"
            name = table.display_name or table.table_name
            col_names = ", ".join(col.name for col in table.columns[:8])
            row_count = table.ref_range.row_count() - 1 if table.ref_range else 0
            items.append((r, c, f"### [{range_str}] Table: \"{name}\" ({row_count} rows)\nColumns: {col_names}"))

        # Charts
        for chart in self._charts:
            r = chart.anchor.from_row if chart.anchor else 0
            c = chart.anchor.from_col if chart.anchor else 0
            title = f'"{chart.title}"' if chart.title else "Untitled"
            series_count = len(chart.series)
            items.append((r, c, f"### Chart: {title} ({chart.chart_type.value}, {series_count} series)"))

        # Bold / large cells that look like titles or headers
        for cell in self._sheet.cells.values():
            if (
                cell.style
                and cell.style.font
                and isinstance(cell.raw_value, str)
                and len(cell.raw_value.strip()) > 0
            ):
                font = cell.style.font
                if font.size and font.size >= 14:
                    items.append((cell.coord.row, cell.coord.col,
                                  f"### [{cell.coord.to_a1()}] Title: \"{cell.raw_value}\""))
                elif font.bold and not cell.formula:
                    # Skip if already part of a table header
                    items.append((cell.coord.row, cell.coord.col,
                                  f"[{cell.coord.to_a1()}] {cell.raw_value}"))

        # Comments/notes
        for cell in self._sheet.cells.values():
            if cell.comment_text:
                items.append((cell.coord.row, cell.coord.col,
                              f"[{cell.coord.to_a1()}] Note: \"{cell.comment_text}\""))

        # Sort by reading order (top-to-bottom, left-to-right)
        items.sort(key=lambda x: (x[0], x[1]))

        for _, _, text in items:
            lines.append(text)

        return "\n".join(lines)
