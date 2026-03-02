"""
Top-level WorkbookDTO aggregating all parsed data from an Excel file.

This is the root object of the parse output. It contains all sheets,
tables, charts, shapes, named ranges, the dependency graph, and
workbook-level metadata. It is the entry point for serialization,
storage mapping, and RAG chunking.
"""

from __future__ import annotations

from datetime import datetime
from typing import Any

from pydantic import Field

from .chart import ChartDTO
from .common import CellRange, ParseError, StableModel, compute_hash
from .dependency import DependencyGraph
from .shape import ShapeDTO
from .sheet import SheetDTO
from .table import TableDTO
from .table_structure import TableStructure
from .template import GeneralizedTemplate, TemplateNode
from .tree import TreeNode


class NamedRangeDTO(StableModel):
    """
    A defined name (named range) in the workbook.

    Named ranges can be workbook-scoped or sheet-scoped. They map
    a human-readable name to a cell reference or range. They are
    first-class citation objects in the RAG system.
    """

    model_config = {"frozen": False, "extra": "forbid"}

    name: str
    ref_string: str  # e.g., "Sheet1!$A$1:$B$10"
    scope_sheet: str | None = None  # None = workbook scope
    parsed_range: CellRange | None = None  # Parsed cell range (if parseable)
    parsed_sheet: str | None = None  # Sheet name extracted from ref
    is_hidden: bool = False
    comment: str | None = None

    # ID
    named_range_id: str = Field(default="")

    def finalize(self, workbook_hash: str) -> None:
        self.named_range_id = compute_hash(
            workbook_hash, self.name, self.scope_sheet or "__workbook__"
        )


class ExternalLink(StableModel):
    """A reference to an external workbook."""

    link_index: int
    target_path: str  # File path or URL
    link_type: str = "workbook"  # "workbook", "dde", "ole"


class WorkbookProperties(StableModel):
    """Workbook-level metadata properties."""

    model_config = {"frozen": True, "extra": "forbid"}

    creator: str | None = None
    last_modified_by: str | None = None
    created: datetime | None = None
    modified: datetime | None = None
    title: str | None = None
    subject: str | None = None
    description: str | None = None
    keywords: str | None = None
    category: str | None = None
    content_status: str | None = None
    language: str | None = None
    revision: str | None = None

    # Calculation settings
    calc_mode: str | None = None  # "auto", "manual", "semiAutomatic"
    iterate_enabled: bool = False
    iterate_count: int | None = None

    # Security
    has_macros: bool = False
    has_vba_project: bool = False
    is_password_protected: bool = False


class WorkbookDTO(StableModel):
    """
    Root DTO for a fully parsed Excel workbook.

    Aggregates all sheets, tables, charts, shapes, named ranges,
    the dependency graph, and workbook-level metadata. This is the
    complete, structured output of the parsing pipeline.
    """

    model_config = {"frozen": False, "extra": "forbid"}

    # Identity
    filename: str
    file_path: str | None = None
    workbook_hash: str = Field(default="", description="xxhash64 of raw file bytes")
    workbook_id: str = Field(default="", description="Deterministic ID from hash + filename")

    # Content
    sheets: list[SheetDTO] = Field(default_factory=list)
    tables: list[TableDTO] = Field(default_factory=list)
    charts: list[ChartDTO] = Field(default_factory=list)
    shapes: list[ShapeDTO] = Field(default_factory=list)
    named_ranges: list[NamedRangeDTO] = Field(default_factory=list)

    # Dependency graph
    dependency_graph: DependencyGraph = Field(default_factory=DependencyGraph)

    # External references
    external_links: list[ExternalLink] = Field(default_factory=list)

    # Workbook metadata
    properties: WorkbookProperties = Field(default_factory=WorkbookProperties)

    # Stage 3: Table structures (header-body-footer assemblies)
    table_structures: list[TableStructure] = Field(default_factory=list)

    # Stage 7: Tree hierarchy nodes
    tree_nodes: list[TreeNode] = Field(default_factory=list)

    # Stage 8: Template nodes
    template_nodes: list[TemplateNode] = Field(default_factory=list)

    # Pivot table detection (ranges only in v1)
    pivot_table_ranges: list[dict[str, Any]] = Field(default_factory=list)

    # Errors
    errors: list[ParseError] = Field(default_factory=list)

    # Parse stats
    total_cells: int = 0
    total_formulas: int = 0
    total_sheets: int = 0
    parse_duration_ms: float | None = None

    def get_sheet(self, name: str) -> SheetDTO | None:
        """Look up a sheet by name."""
        for sheet in self.sheets:
            return sheet if sheet.sheet_name == name else None
        return None

    def get_sheet_by_index(self, index: int) -> SheetDTO | None:
        """Look up a sheet by 0-based index."""
        if 0 <= index < len(self.sheets):
            return self.sheets[index]
        return None

    def finalize(self) -> None:
        """Compute IDs, hashes, and stats for the entire workbook."""
        self.workbook_id = compute_hash(self.workbook_hash, self.filename)
        self.total_sheets = len(self.sheets)
        self.total_cells = sum(s.cell_count() for s in self.sheets)
        self.total_formulas = sum(
            1
            for s in self.sheets
            for c in s.cells.values()
            if c.formula is not None
        )

        for sheet in self.sheets:
            sheet.finalize(self.workbook_hash)
        for table in self.tables:
            table.finalize(self.workbook_hash)
        for chart in self.charts:
            chart.finalize(self.workbook_hash)
        for shape in self.shapes:
            shape.finalize(self.workbook_hash)
        for nr in self.named_ranges:
            nr.finalize(self.workbook_hash)

        self.dependency_graph.build_indexes()
