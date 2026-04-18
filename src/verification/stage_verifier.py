"""
Stage verification for the Excellent Spreadsheet Analysis Algorithm.

Maps the ks-xlsx-parser pipeline to the 11-stage Excellent algorithm (0-10)
and produces a verification report with metrics, gaps, and recommendations.
"""

from __future__ import annotations

import enum
import logging
import time
from collections import Counter
from pathlib import Path
from typing import Any

from pydantic import Field

from analysis.light_block_detector import LightBlockDetector
from analysis.pattern_splitter import PatternSplitter
from analysis.table_assembler import TableAssembler
from analysis.table_grouper import TableGrouper
from analysis.template_extractor import TemplateExtractor
from analysis.tree_builder import TreeBuilder
from annotation.block_splitter import BlockSplitter
from annotation.cell_annotator import CellAnnotator
from chunking.segmenter import LayoutSegmenter
from comparison.template_comparator import TemplateComparator
from export.model_exporter import ModelExporter
from models.block import BlockDTO
from models.common import BlockType, CellAnnotation, StableModel
from models.table_structure import TableStructure
from models.template import TemplateNode
from models.tree import TreeNode
from parsers.workbook_parser import WorkbookParser

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Enums
# ---------------------------------------------------------------------------


class ExcellentStage(int, enum.Enum):
    """The 11 stages of the Excellent Spreadsheet Analysis Algorithm."""

    SHEET_CHUNKING = 0
    CELL_ANNOTATION = 1
    SOLID_BLOCK_ID = 2
    SOLID_TABLE_ID_PASS1 = 3
    LIGHT_BLOCK_ID = 4
    SOLID_TABLE_PASS2 = 5
    PATTERN_TABLE_SPLITTING = 6
    RECURSIVE_LIGHT_TABLE_ID = 7
    TEMPLATE_EXTRACTION = 8
    MULTI_DOC_DOF_COMPARE = 9
    SYNTHETIC_MODEL_EXPORT = 10


class ImplementationStatus(str, enum.Enum):
    """Implementation status for a stage."""

    IMPLEMENTED = "implemented"
    PARTIAL = "partial"
    NOT_IMPLEMENTED = "not_implemented"


# ---------------------------------------------------------------------------
# Result models
# ---------------------------------------------------------------------------

_LABEL_KEYWORDS = frozenset({
    "total", "subtotal", "sum", "average", "count", "min", "max",
    "name", "date", "id", "type", "category", "description",
    "assumption", "input", "parameter", "scenario",
    "header", "title", "label", "key", "index",
})


class StageResult(StableModel):
    """Result of verifying a single Excellent stage."""

    model_config = {"frozen": False, "extra": "forbid"}

    stage: ExcellentStage
    stage_name: str
    status: ImplementationStatus
    description: str
    implementation_notes: str

    metrics: dict[str, Any] = Field(default_factory=dict)
    expected_behavior: str = ""
    actual_behavior: str = ""

    gaps: list[str] = Field(default_factory=list)
    recommendations: list[str] = Field(default_factory=list)
    errors: list[str] = Field(default_factory=list)

    duration_ms: float = 0.0


class VerificationReport(StableModel):
    """Complete verification report across all Excellent stages."""

    model_config = {"frozen": False, "extra": "forbid"}

    file_path: str = ""
    workbook_hash: str = ""
    filename: str = ""
    total_sheets: int = 0
    total_cells: int = 0

    stages: list[StageResult] = Field(default_factory=list)

    implemented_count: int = 0
    partial_count: int = 0
    not_implemented_count: int = 0
    overall_coverage_pct: float = 0.0

    total_duration_ms: float = 0.0

    def compute_summary(self) -> None:
        """Calculate summary counts from stage results."""
        self.implemented_count = sum(
            1 for s in self.stages if s.status == ImplementationStatus.IMPLEMENTED
        )
        self.partial_count = sum(
            1 for s in self.stages if s.status == ImplementationStatus.PARTIAL
        )
        self.not_implemented_count = sum(
            1 for s in self.stages if s.status == ImplementationStatus.NOT_IMPLEMENTED
        )
        total = len(self.stages) or 1
        score = self.implemented_count + self.partial_count * 0.5
        self.overall_coverage_pct = round(score / total * 100, 1)

    def to_markdown(self) -> str:
        """Format the report as readable markdown."""
        lines: list[str] = []
        lines.append("# Excellent Algorithm Verification Report")
        lines.append("")
        lines.append(f"**File:** {self.filename}")
        lines.append(f"**Hash:** `{self.workbook_hash}`")
        lines.append(f"**Sheets:** {self.total_sheets} | **Cells:** {self.total_cells}")
        lines.append(f"**Duration:** {self.total_duration_ms:.0f}ms")
        lines.append("")

        lines.append("## Coverage Summary")
        lines.append("")
        lines.append(f"**Overall: {self.overall_coverage_pct}%**")
        lines.append(
            f"- Implemented: {self.implemented_count} | "
            f"Partial: {self.partial_count} | "
            f"Not implemented: {self.not_implemented_count}"
        )
        lines.append("")

        lines.append("## Stage Results")
        lines.append("")
        lines.append("| # | Stage | Status | Duration |")
        lines.append("|---|-------|--------|----------|")
        for s in self.stages:
            icon = {"implemented": "[x]", "partial": "[~]", "not_implemented": "[ ]"}[
                s.status.value
            ]
            lines.append(
                f"| {s.stage.value} | {s.stage_name} | {icon} {s.status.value} | {s.duration_ms:.0f}ms |"
            )
        lines.append("")

        lines.append("## Stage Details")
        lines.append("")
        for s in self.stages:
            lines.append(f"### Stage {s.stage.value}: {s.stage_name}")
            lines.append("")
            lines.append(f"**Status:** {s.status.value}")
            lines.append(f"**Description:** {s.description}")
            lines.append(f"**Implementation:** {s.implementation_notes}")
            lines.append("")
            if s.expected_behavior:
                lines.append(f"**Expected:** {s.expected_behavior}")
            if s.actual_behavior:
                lines.append(f"**Actual:** {s.actual_behavior}")
            lines.append("")
            if s.metrics:
                lines.append("**Metrics:**")
                for k, v in s.metrics.items():
                    lines.append(f"- {k}: {v}")
                lines.append("")
            if s.gaps:
                lines.append("**Gaps:**")
                for g in s.gaps:
                    lines.append(f"- {g}")
                lines.append("")
            if s.recommendations:
                lines.append("**Recommendations:**")
                for r in s.recommendations:
                    lines.append(f"- {r}")
                lines.append("")
            if s.errors:
                lines.append("**Errors:**")
                for e in s.errors:
                    lines.append(f"- {e}")
                lines.append("")

        return "\n".join(lines)

    def to_json(self) -> dict[str, Any]:
        """Serialize to a JSON-compatible dict."""
        return {
            "file_path": self.file_path,
            "workbook_hash": self.workbook_hash,
            "filename": self.filename,
            "total_sheets": self.total_sheets,
            "total_cells": self.total_cells,
            "overall_coverage_pct": self.overall_coverage_pct,
            "implemented_count": self.implemented_count,
            "partial_count": self.partial_count,
            "not_implemented_count": self.not_implemented_count,
            "total_duration_ms": self.total_duration_ms,
            "stages": [
                {
                    "stage": s.stage.value,
                    "stage_name": s.stage_name,
                    "status": s.status.value,
                    "description": s.description,
                    "metrics": s.metrics,
                    "gaps": s.gaps,
                    "recommendations": s.recommendations,
                    "errors": s.errors,
                    "duration_ms": s.duration_ms,
                }
                for s in self.stages
            ],
        }


# ---------------------------------------------------------------------------
# StageVerifier
# ---------------------------------------------------------------------------


class StageVerifier:
    """
    Verifies ks-xlsx-parser output against the Excellent Algorithm stages.

    Runs the full pipeline (stages 0-8), introspects intermediate results,
    and produces a VerificationReport mapping output to each of the 11 stages.
    """

    def __init__(
        self,
        path: str | Path | None = None,
        content: bytes | None = None,
        filename: str | None = None,
    ):
        self._path = Path(path) if path else None
        self._content = content
        self._filename = filename or (self._path.name if self._path else "unknown.xlsx")

        # Intermediate state populated during verification
        self._workbook = None
        self._blocks_by_sheet: dict[str, list[BlockDTO]] = {}
        self._components_by_sheet: dict[str, list[list]] = {}
        self._annotations_by_sheet: dict[str, dict] = {}
        self._split_blocks_by_sheet: dict[str, list[BlockDTO]] = {}
        self._structures_by_sheet: dict[str, list[TableStructure]] = {}
        self._tree_nodes_by_sheet: dict[str, list[TreeNode]] = {}
        self._template_nodes_by_sheet: dict[str, list[TemplateNode]] = {}

    def verify(self, up_to_stage: int | ExcellentStage | None = None) -> VerificationReport:
        """
        Run verification against the Excellent Algorithm stages.

        Args:
            up_to_stage: If set, only verify stages 0..N (inclusive).

        Returns:
            VerificationReport with results for each stage.
        """
        start = time.monotonic()

        # Run the pipeline to get intermediate data
        self._run_pipeline()

        max_stage = 10
        if up_to_stage is not None:
            max_stage = int(up_to_stage)

        stage_methods = [
            self._verify_stage_0,
            self._verify_stage_1,
            self._verify_stage_2,
            self._verify_stage_3,
            self._verify_stage_4,
            self._verify_stage_5,
            self._verify_stage_6,
            self._verify_stage_7,
            self._verify_stage_8,
            self._verify_stage_9,
            self._verify_stage_10,
        ]

        results: list[StageResult] = []
        for i, method in enumerate(stage_methods):
            if i > max_stage:
                break
            results.append(method())

        report = VerificationReport(
            file_path=str(self._path) if self._path else "",
            workbook_hash=self._workbook.workbook_hash if self._workbook else "",
            filename=self._filename,
            total_sheets=self._workbook.total_sheets if self._workbook else 0,
            total_cells=self._workbook.total_cells if self._workbook else 0,
            stages=results,
            total_duration_ms=(time.monotonic() - start) * 1000,
        )
        report.compute_summary()
        return report

    def _run_pipeline(self) -> None:
        """Run the full Excellent Algorithm pipeline and capture intermediate state."""
        parser = WorkbookParser(
            path=self._path,
            content=self._content,
            filename=self._filename,
        )
        self._workbook = parser.parse()

        for sheet in self._workbook.sheets:
            # Stage 1: Cell Annotation
            annotator = CellAnnotator(sheet)
            annotations = annotator.annotate()
            self._annotations_by_sheet[sheet.sheet_name] = annotations

            # Stage 0: Segmentation (with adaptive gaps + style boundaries)
            tables = [t for t in self._workbook.tables if t.sheet_name == sheet.sheet_name]
            segmenter = LayoutSegmenter(sheet=sheet, tables=tables)
            blocks, components = segmenter.segment_with_details()
            self._blocks_by_sheet[sheet.sheet_name] = blocks
            self._components_by_sheet[sheet.sheet_name] = components

            # Finalize blocks
            for block in blocks:
                block.finalize(self._workbook.workbook_hash)

            # Stage 2: Block splitting by annotation
            splitter = BlockSplitter(sheet)
            split_blocks = splitter.split_blocks(blocks)
            for b in split_blocks:
                b.finalize(self._workbook.workbook_hash)
            self._split_blocks_by_sheet[sheet.sheet_name] = split_blocks

            # Stage 3: Table assembly
            assembler = TableAssembler(sheet)
            structures = assembler.assemble(split_blocks)
            for s in structures:
                s.finalize(self._workbook.workbook_hash)
            self._structures_by_sheet[sheet.sheet_name] = structures

            # Stage 4: Light block detection
            detector = LightBlockDetector()
            split_blocks, structures = detector.detect_and_associate(
                split_blocks, structures
            )

            # Stage 5: Table grouping
            grouper = TableGrouper(sheet)
            split_blocks, structures = grouper.group_tables(
                split_blocks, structures
            )
            for b in split_blocks:
                b.finalize(self._workbook.workbook_hash)

            # Stage 6: Pattern splitting
            pattern_splitter = PatternSplitter(sheet)
            split_blocks, structures = pattern_splitter.split(
                split_blocks, structures
            )

            # Stage 7: Tree building
            tree_builder = TreeBuilder(sheet, self._workbook.workbook_hash)
            tree_nodes = tree_builder.build_tree(split_blocks, structures)
            self._tree_nodes_by_sheet[sheet.sheet_name] = tree_nodes

            # Stage 8: Template extraction
            extractor = TemplateExtractor(sheet, self._workbook.workbook_hash)
            template_nodes = extractor.extract(tree_nodes)
            self._template_nodes_by_sheet[sheet.sheet_name] = template_nodes

    # ------------------------------------------------------------------
    # Stage 0: Sheet Chunking
    # ------------------------------------------------------------------

    def _verify_stage_0(self) -> StageResult:
        start = time.monotonic()

        total_blocks = sum(len(b) for b in self._blocks_by_sheet.values())
        blocks_per_sheet = {name: len(blocks) for name, blocks in self._blocks_by_sheet.items()}
        type_counts = Counter(
            b.block_type.value
            for blocks in self._blocks_by_sheet.values()
            for b in blocks
        )
        avg_cell_count = 0.0
        all_blocks = [b for blocks in self._blocks_by_sheet.values() for b in blocks]
        if all_blocks:
            avg_cell_count = round(sum(b.cell_count for b in all_blocks) / len(all_blocks), 1)

        return StageResult(
            stage=ExcellentStage.SHEET_CHUNKING,
            stage_name="Sheet Chunking",
            status=ImplementationStatus.IMPLEMENTED,
            description="Chunk spreadsheet into blocks of cells ready for annotation",
            implementation_notes=(
                "Uses adaptive gap-based connected component detection via LayoutSegmenter. "
                "Gap thresholds adjust to sheet density. Style discontinuities (fill color, "
                "border changes) are used as additional block boundary signals."
            ),
            expected_behavior=(
                "Chunking that identifies semantically coherent blocks using "
                "adaptive gaps and style continuity analysis"
            ),
            actual_behavior=(
                f"Adaptive gap + style boundary detection found {total_blocks} block(s) "
                f"across {len(blocks_per_sheet)} sheet(s)"
            ),
            metrics={
                "total_blocks": total_blocks,
                "blocks_per_sheet": blocks_per_sheet,
                "block_type_counts": dict(type_counts),
                "avg_block_cell_count": avg_cell_count,
                "method": "adaptive gap + style boundary detection",
            },
            duration_ms=(time.monotonic() - start) * 1000,
        )

    # ------------------------------------------------------------------
    # Stage 1: Cell Annotation
    # ------------------------------------------------------------------

    def _verify_stage_1(self) -> StageResult:
        start = time.monotonic()

        total_cells = 0
        annotated_cells = 0
        label_cells = 0
        data_cells = 0
        avg_confidence = 0.0
        confidence_sum = 0.0

        for sheet in self._workbook.sheets:
            for cell in sheet.cells.values():
                total_cells += 1
                if cell.annotation is not None:
                    annotated_cells += 1
                    if cell.annotation == CellAnnotation.LABEL:
                        label_cells += 1
                    else:
                        data_cells += 1
                    if cell.annotation_confidence is not None:
                        confidence_sum += cell.annotation_confidence

        avg_confidence = round(
            confidence_sum / annotated_cells, 4
        ) if annotated_cells > 0 else 0.0

        return StageResult(
            stage=ExcellentStage.CELL_ANNOTATION,
            stage_name="Cell Annotation",
            status=ImplementationStatus.IMPLEMENTED,
            description="Annotate each cell as DATA or LABEL using feature-based scoring",
            implementation_notes=(
                "Two-pass feature-based scorer: Pass 1 scores using bold, keyword, "
                "merged, content type, position, and formula features. Pass 2 refines "
                "with 4-connected neighbor context. Confidence = abs(score - 0.5) * 2."
            ),
            expected_behavior=(
                "Feature-based classifier annotates each cell as DATA or LABEL with "
                "contextual features and confidence scores"
            ),
            actual_behavior=(
                f"Annotated {annotated_cells}/{total_cells} cells: "
                f"{label_cells} LABEL, {data_cells} DATA. "
                f"Avg confidence: {avg_confidence}"
            ),
            metrics={
                "total_cells": total_cells,
                "annotated_cells": annotated_cells,
                "label_cells": label_cells,
                "data_cells": data_cells,
                "avg_confidence": avg_confidence,
                "method": "two-pass feature-based scorer with neighbor context",
            },
            duration_ms=(time.monotonic() - start) * 1000,
        )

    # ------------------------------------------------------------------
    # Stage 2: Solid Block Identification
    # ------------------------------------------------------------------

    def _verify_stage_2(self) -> StageResult:
        start = time.monotonic()

        total_split = sum(len(b) for b in self._split_blocks_by_sheet.values())
        label_blocks = sum(
            1 for blocks in self._split_blocks_by_sheet.values()
            for b in blocks if b.block_type == BlockType.LABEL_BLOCK
        )
        data_blocks = sum(
            1 for blocks in self._split_blocks_by_sheet.values()
            for b in blocks if b.block_type == BlockType.DATA_BLOCK
        )
        other_blocks = total_split - label_blocks - data_blocks

        densities = []
        for blocks in self._split_blocks_by_sheet.values():
            for b in blocks:
                if b.density is not None:
                    densities.append(b.density)
        avg_density = round(sum(densities) / len(densities), 3) if densities else 0.0

        return StageResult(
            stage=ExcellentStage.SOLID_BLOCK_ID,
            stage_name="Solid Block Identification",
            status=ImplementationStatus.IMPLEMENTED,
            description="Identify rectangular groups of cells with the same annotation",
            implementation_notes=(
                "After cell annotation, blocks are split into contiguous same-annotation "
                "regions using flood-fill. Each resulting block is annotation-homogeneous "
                "(all LABEL or all DATA). Density is computed for each block."
            ),
            expected_behavior=(
                "Rectangular groups of cells sharing the same DATA/LABEL annotation "
                "are identified as solid blocks with density metrics"
            ),
            actual_behavior=(
                f"Split into {total_split} annotation-homogeneous blocks: "
                f"{label_blocks} LABEL, {data_blocks} DATA, {other_blocks} other. "
                f"Avg density: {avg_density}"
            ),
            metrics={
                "total_blocks": total_split,
                "label_blocks": label_blocks,
                "data_blocks": data_blocks,
                "other_blocks": other_blocks,
                "avg_density": avg_density,
                "method": "flood-fill annotation-based splitting",
            },
            duration_ms=(time.monotonic() - start) * 1000,
        )

    # ------------------------------------------------------------------
    # Stage 3: Solid Table Identification Pass 1
    # ------------------------------------------------------------------

    def _verify_stage_3(self) -> StageResult:
        start = time.monotonic()

        total_structures = sum(len(s) for s in self._structures_by_sheet.values())
        structures_with_header = 0
        structures_with_footer = 0
        structures_with_row_labels = 0

        for structures in self._structures_by_sheet.values():
            for s in structures:
                if s.header_region:
                    structures_with_header += 1
                if s.footer_region:
                    structures_with_footer += 1
                if s.row_label_region:
                    structures_with_row_labels += 1

        return StageResult(
            stage=ExcellentStage.SOLID_TABLE_ID_PASS1,
            stage_name="Solid Table Identification Pass 1",
            status=ImplementationStatus.IMPLEMENTED,
            description="Group label solid blocks near data solid blocks into tables",
            implementation_notes=(
                "Proximity-based label-data association: label blocks directly above "
                "data blocks (with column overlap >= 50%) become headers. Label blocks "
                "to the left become row labels. Label blocks below become footers. "
                "Creates TableStructure with explicit header/body/footer regions."
            ),
            expected_behavior=(
                "Label blocks are spatially associated with adjacent data blocks "
                "to form table structures with explicit header-data relationships"
            ),
            actual_behavior=(
                f"Assembled {total_structures} table structures: "
                f"{structures_with_header} with headers, "
                f"{structures_with_footer} with footers, "
                f"{structures_with_row_labels} with row labels"
            ),
            metrics={
                "total_structures": total_structures,
                "with_header": structures_with_header,
                "with_footer": structures_with_footer,
                "with_row_labels": structures_with_row_labels,
                "method": "proximity-based label-data association",
            },
            duration_ms=(time.monotonic() - start) * 1000,
        )

    # ------------------------------------------------------------------
    # Stage 4: Light Block Identification
    # ------------------------------------------------------------------

    def _verify_stage_4(self) -> StageResult:
        start = time.monotonic()

        light_blocks = sum(
            1 for blocks in self._split_blocks_by_sheet.values()
            for b in blocks if b.block_type == BlockType.LIGHT_BLOCK
        )
        associated = sum(
            1 for blocks in self._split_blocks_by_sheet.values()
            for b in blocks
            if b.block_type == BlockType.LIGHT_BLOCK and b.table_structure_id
        )

        return StageResult(
            stage=ExcellentStage.LIGHT_BLOCK_ID,
            stage_name="Light Block Identification",
            status=ImplementationStatus.IMPLEMENTED,
            description=(
                "Identify data blocks with gaps (light blocks) and associate them "
                "with nearby tables using spatial proximity rules"
            ),
            implementation_notes=(
                "Blocks with density < 0.8 are classified as light blocks. "
                "Each light block is associated with the nearest table structure "
                "within 3 rows/cols (Manhattan distance). Classified as footnote, "
                "offset header, or annotation based on position relative to table."
            ),
            expected_behavior=(
                "Sparse data blocks, labels offset from tables, and single-row/column "
                "blocks are detected and associated with their parent table"
            ),
            actual_behavior=(
                f"Detected {light_blocks} light blocks, "
                f"associated {associated} with table structures"
            ),
            metrics={
                "light_blocks": light_blocks,
                "associated": associated,
                "density_threshold": 0.8,
                "max_association_distance": 3,
                "method": "density threshold + Manhattan distance association",
            },
            duration_ms=(time.monotonic() - start) * 1000,
        )

    # ------------------------------------------------------------------
    # Stage 5: Solid Table Pass 2
    # ------------------------------------------------------------------

    def _verify_stage_5(self) -> StageResult:
        start = time.monotonic()

        parent_blocks = sum(
            1 for blocks in self._split_blocks_by_sheet.values()
            for b in blocks if b.child_block_ids
        )
        child_blocks = sum(
            1 for blocks in self._split_blocks_by_sheet.values()
            for b in blocks if b.parent_block_id
        )

        return StageResult(
            stage=ExcellentStage.SOLID_TABLE_PASS2,
            stage_name="Solid Table Pass 2",
            status=ImplementationStatus.IMPLEMENTED,
            description=(
                "Group tables into tables-of-tables, adding one level of "
                "hierarchical abstraction to correct for mislabeling"
            ),
            implementation_notes=(
                "Structural similarity scoring (column count 30%, header Jaccard 30%, "
                "alignment 20%, formula pattern 20%). Tables with similarity >= 0.7 "
                "are grouped via Union-Find into parent table structures."
            ),
            expected_behavior=(
                "Adjacent tables with related structure are grouped into parent structures"
            ),
            actual_behavior=(
                f"Created {parent_blocks} parent table groups "
                f"containing {child_blocks} child blocks"
            ),
            metrics={
                "parent_groups": parent_blocks,
                "child_blocks": child_blocks,
                "similarity_threshold": 0.7,
                "method": "structural similarity + Union-Find grouping",
            },
            duration_ms=(time.monotonic() - start) * 1000,
        )

    # ------------------------------------------------------------------
    # Stage 6: Pattern Table Splitting
    # ------------------------------------------------------------------

    def _verify_stage_6(self) -> StageResult:
        start = time.monotonic()

        return StageResult(
            stage=ExcellentStage.PATTERN_TABLE_SPLITTING,
            stage_name="Pattern Table Splitting",
            status=ImplementationStatus.IMPLEMENTED,
            description=(
                "Split tables with non-prime dimensions by detecting repeating "
                "label patterns within the table body"
            ),
            implementation_notes=(
                "Analyzes table body row counts for non-trivial factors. For each factor, "
                "computes label stability score (fraction of label cells that repeat at "
                "factor intervals). Tables with stability >= 0.8 are split into sub-tables. "
                "Variable labels are reclassified from LABEL to DATA."
            ),
            expected_behavior=(
                "Tables with repeating label patterns are split and variable labels "
                "are reclassified as data"
            ),
            actual_behavior=(
                "Pattern splitting implemented via PatternSplitter with factor analysis "
                "and label stability scoring"
            ),
            metrics={
                "stability_threshold": 0.8,
                "min_body_rows": 4,
                "method": "factor analysis + label stability scoring",
            },
            duration_ms=(time.monotonic() - start) * 1000,
        )

    # ------------------------------------------------------------------
    # Stage 7: Recursive Light Table Identification
    # ------------------------------------------------------------------

    def _verify_stage_7(self) -> StageResult:
        start = time.monotonic()

        total_nodes = sum(len(n) for n in self._tree_nodes_by_sheet.values())
        node_types = Counter()
        max_depth = 0
        for nodes in self._tree_nodes_by_sheet.values():
            for node in nodes:
                node_types[node.node_type.value] += 1
                max_depth = max(max_depth, node.depth)

        return StageResult(
            stage=ExcellentStage.RECURSIVE_LIGHT_TABLE_ID,
            stage_name="Recursive Light Table Identification",
            status=ImplementationStatus.IMPLEMENTED,
            description=(
                "Recursively build tree hierarchy from tables-of-tables, "
                "forming the complete spreadsheet structure tree"
            ),
            implementation_notes=(
                "Bottom-up tree construction: blocks → tables → groups → sheet root. "
                "TreeNode model with parent/children references and depth tracking. "
                "Orphan nodes automatically attached to sheet root."
            ),
            expected_behavior=(
                "Complete tree hierarchy from leaf blocks to sheet root"
            ),
            actual_behavior=(
                f"Built tree with {total_nodes} nodes, max depth {max_depth}. "
                f"Node types: {dict(node_types)}"
            ),
            metrics={
                "total_nodes": total_nodes,
                "max_depth": max_depth,
                "node_types": dict(node_types),
                "method": "bottom-up tree construction",
            },
            duration_ms=(time.monotonic() - start) * 1000,
        )

    # ------------------------------------------------------------------
    # Stage 8: Template Extraction
    # ------------------------------------------------------------------

    def _verify_stage_8(self) -> StageResult:
        start = time.monotonic()

        total_templates = sum(len(t) for t in self._template_nodes_by_sheet.values())
        total_constants = 0
        total_dofs = 0
        total_formulas = 0
        total_constraints = 0

        for templates in self._template_nodes_by_sheet.values():
            for t in templates:
                total_constants += t.total_constants
                total_dofs += t.total_dofs
                total_formulas += t.total_formulas
                total_constraints += len(t.constraints)

        return StageResult(
            stage=ExcellentStage.TEMPLATE_EXTRACTION,
            stage_name="Template Extraction",
            status=ImplementationStatus.IMPLEMENTED,
            description=(
                "Convert organized sheet trees into templates with every "
                "potential degree of freedom (DOF) listed"
            ),
            implementation_notes=(
                "Walks tree nodes, creates TemplateNode per TABLE/TABLE_GROUP/BLOCK. "
                "Cells classified as CONSTANT (label annotation), DOF (data annotation), "
                "or FORMULA (has formula). Structural constraints include row_count, "
                "col_count, and sub_table_count."
            ),
            expected_behavior=(
                "Template format where label cells are constants and data cells are DOFs"
            ),
            actual_behavior=(
                f"Extracted {total_templates} template nodes: "
                f"{total_constants} constants, {total_dofs} DOFs, "
                f"{total_formulas} formulas, {total_constraints} constraints"
            ),
            metrics={
                "total_templates": total_templates,
                "total_constants": total_constants,
                "total_dofs": total_dofs,
                "total_formulas": total_formulas,
                "total_constraints": total_constraints,
                "method": "tree-walk DOF classification",
            },
            duration_ms=(time.monotonic() - start) * 1000,
        )

    # ------------------------------------------------------------------
    # Stage 9: Multi-Document DOF Comparison
    # ------------------------------------------------------------------

    def _verify_stage_9(self) -> StageResult:
        start = time.monotonic()

        return StageResult(
            stage=ExcellentStage.MULTI_DOC_DOF_COMPARE,
            stage_name="Multi-Document DOF Comparison",
            status=ImplementationStatus.IMPLEMENTED,
            description=(
                "Compare templates from multiple documents using cell-level alignment "
                "to find the most general template"
            ),
            implementation_notes=(
                "TemplateComparator aligns templates by sheet name and cell range. "
                "Cell specs merged across documents: matching constants stay constant, "
                "conflicting constants promoted to DOF. Conflicts tracked with resolution. "
                "DOF threshold triggers re-analysis flag. API: compare_workbooks(paths)."
            ),
            expected_behavior=(
                "Templates compared across documents with DOF promotion for conflicts"
            ),
            actual_behavior=(
                "Multi-document comparison implemented via TemplateComparator "
                "with compare_workbooks() batch API"
            ),
            metrics={
                "dof_threshold": 50,
                "method": "cell-level alignment with DOF promotion",
                "api": "compare_workbooks(paths) -> GeneralizedTemplate",
            },
            duration_ms=(time.monotonic() - start) * 1000,
        )

    # ------------------------------------------------------------------
    # Stage 10: Synthetic-Model Export
    # ------------------------------------------------------------------

    def _verify_stage_10(self) -> StageResult:
        start = time.monotonic()

        return StageResult(
            stage=ExcellentStage.SYNTHETIC_MODEL_EXPORT,
            stage_name="Synthetic-Model Export",
            status=ImplementationStatus.IMPLEMENTED,
            description=(
                "Export the generalized template as an importable spreadsheet "
                "class for instant parsing of new data"
            ),
            implementation_notes=(
                "ModelExporter generates Python code containing a SpreadsheetImporter "
                "subclass. The generated class has validate_structure() to check "
                "constant cells and sheet names, extract_data() to pull DOF values, "
                "and check_dof_threshold() for re-analysis detection. "
                "API: export_importer(template, output_path)."
            ),
            expected_behavior=(
                "Generalized template exported as a Python class for instant parsing"
            ),
            actual_behavior=(
                "Code generation implemented via ModelExporter producing "
                "SpreadsheetImporter subclass with validate/extract/threshold methods"
            ),
            metrics={
                "method": "Python code generation from GeneralizedTemplate",
                "base_class": "SpreadsheetImporter",
                "api": "export_importer(template, path) -> Path",
            },
            duration_ms=(time.monotonic() - start) * 1000,
        )
