"""
Tests for the Excellent Algorithm stage verification system.

Verifies that the StageVerifier correctly maps ks-xlsx-parser output
to each of the 11 Excellent Algorithm stages and produces accurate
reports with metrics, gaps, and recommendations.
"""

import json

import pytest

from verification import (
    ExcellentStage,
    ImplementationStatus,
    StageResult,
    StageVerifier,
    VerificationReport,
)


# ---------------------------------------------------------------------------
# Model tests
# ---------------------------------------------------------------------------


class TestStageModels:
    """Test the verification data models."""

    def test_excellent_stage_has_11_members(self):
        assert len(ExcellentStage) == 11
        assert ExcellentStage.SHEET_CHUNKING == 0
        assert ExcellentStage.SYNTHETIC_MODEL_EXPORT == 10

    def test_stage_result_creation(self):
        result = StageResult(
            stage=ExcellentStage.SHEET_CHUNKING,
            stage_name="Sheet Chunking",
            status=ImplementationStatus.PARTIAL,
            description="Chunk spreadsheet",
            implementation_notes="Gap-based",
        )
        assert result.stage == ExcellentStage.SHEET_CHUNKING
        assert result.status == ImplementationStatus.PARTIAL
        assert result.metrics == {}
        assert result.gaps == []

    def test_report_summary_computation(self):
        stages = [
            StageResult(
                stage=ExcellentStage(i),
                stage_name=f"Stage {i}",
                status=(
                    ImplementationStatus.IMPLEMENTED if i == 2
                    else ImplementationStatus.PARTIAL if i < 4
                    else ImplementationStatus.NOT_IMPLEMENTED
                ),
                description="",
                implementation_notes="",
            )
            for i in range(11)
        ]
        report = VerificationReport(stages=stages)
        report.compute_summary()
        assert report.implemented_count == 1
        assert report.partial_count == 3  # stages 0, 1, 3
        assert report.not_implemented_count == 7
        assert report.overall_coverage_pct > 0

    def test_report_to_json_serializable(self):
        report = VerificationReport(
            file_path="test.xlsx",
            filename="test.xlsx",
            workbook_hash="abc123",
            stages=[
                StageResult(
                    stage=ExcellentStage.SHEET_CHUNKING,
                    stage_name="Sheet Chunking",
                    status=ImplementationStatus.PARTIAL,
                    description="Test",
                    implementation_notes="Test",
                    metrics={"count": 5},
                )
            ],
        )
        report.compute_summary()
        j = report.to_json()
        # Must be JSON serializable
        json_str = json.dumps(j)
        assert "test.xlsx" in json_str
        assert j["stages"][0]["metrics"]["count"] == 5


# ---------------------------------------------------------------------------
# StageVerifier core tests
# ---------------------------------------------------------------------------


class TestStageVerifier:
    """Test the StageVerifier against various workbook fixtures."""

    def test_verify_returns_all_stages(self, simple_workbook):
        verifier = StageVerifier(path=simple_workbook)
        report = verifier.verify()
        assert len(report.stages) == 11

    def test_verify_up_to_stage(self, simple_workbook):
        verifier = StageVerifier(path=simple_workbook)
        report = verifier.verify(up_to_stage=3)
        assert len(report.stages) == 4
        assert report.stages[-1].stage == ExcellentStage.SOLID_TABLE_ID_PASS1

    def test_stage_0_finds_blocks(self, simple_workbook):
        verifier = StageVerifier(path=simple_workbook)
        report = verifier.verify(up_to_stage=0)
        s0 = report.stages[0]
        assert s0.stage == ExcellentStage.SHEET_CHUNKING
        assert s0.status == ImplementationStatus.IMPLEMENTED
        assert s0.metrics["total_blocks"] >= 1

    def test_stage_1_cell_counts(self, simple_workbook):
        verifier = StageVerifier(path=simple_workbook)
        report = verifier.verify(up_to_stage=1)
        s1 = report.stages[1]
        assert s1.stage == ExcellentStage.CELL_ANNOTATION
        assert s1.status == ImplementationStatus.IMPLEMENTED
        assert s1.metrics["total_cells"] > 0
        # Annotated cells: label + data should sum to annotated total
        assert (
            s1.metrics["label_cells"] + s1.metrics["data_cells"]
            == s1.metrics["annotated_cells"]
        )

    def test_stage_2_blocks_found(self, simple_workbook):
        verifier = StageVerifier(path=simple_workbook)
        report = verifier.verify(up_to_stage=2)
        s2 = report.stages[2]
        assert s2.stage == ExcellentStage.SOLID_BLOCK_ID
        assert s2.status == ImplementationStatus.IMPLEMENTED
        assert s2.metrics["total_blocks"] >= 1

    def test_stage_2_density_valid(self, simple_workbook):
        verifier = StageVerifier(path=simple_workbook)
        report = verifier.verify(up_to_stage=2)
        s2 = report.stages[2]
        assert 0.0 <= s2.metrics["avg_density"] <= 1.0

    def test_stage_3_table_classified(self, simple_workbook):
        verifier = StageVerifier(path=simple_workbook)
        report = verifier.verify(up_to_stage=3)
        s3 = report.stages[3]
        assert s3.stage == ExcellentStage.SOLID_TABLE_ID_PASS1
        assert s3.status == ImplementationStatus.IMPLEMENTED
        assert s3.metrics["total_structures"] >= 0

    def test_all_stages_implemented(self, simple_workbook):
        verifier = StageVerifier(path=simple_workbook)
        report = verifier.verify()
        for i in range(11):
            stage = report.stages[i]
            assert stage.status == ImplementationStatus.IMPLEMENTED, (
                f"Stage {i} ({stage.stage_name}) should be IMPLEMENTED"
            )

    def test_stage_0_multi_table(self, two_tables_vertical):
        verifier = StageVerifier(path=two_tables_vertical)
        report = verifier.verify(up_to_stage=0)
        s0 = report.stages[0]
        assert s0.metrics["total_blocks"] >= 2

    def test_formula_workbook_annotation(self, simple_formulas):
        verifier = StageVerifier(path=simple_formulas)
        report = verifier.verify(up_to_stage=1)
        s1 = report.stages[1]
        assert s1.metrics["annotated_cells"] > 0

    def test_deterministic_report(self, simple_workbook):
        v1 = StageVerifier(path=simple_workbook)
        v2 = StageVerifier(path=simple_workbook)
        r1 = v1.verify()
        r2 = v2.verify()
        assert r1.workbook_hash == r2.workbook_hash
        for s1, s2 in zip(r1.stages, r2.stages):
            assert s1.status == s2.status
            # Metrics should be identical (except timing)
            for key in s1.metrics:
                if key not in ("duration_ms",):
                    assert s1.metrics[key] == s2.metrics[key], (
                        f"Stage {s1.stage_name}, metric '{key}' differs"
                    )

    def test_verify_from_bytes(self, simple_workbook):
        with open(simple_workbook, "rb") as f:
            content = f.read()
        verifier = StageVerifier(content=content, filename="test.xlsx")
        report = verifier.verify(up_to_stage=0)
        assert report.workbook_hash != ""
        assert report.stages[0].metrics["total_blocks"] >= 1


# ---------------------------------------------------------------------------
# Report formatting tests
# ---------------------------------------------------------------------------


class TestVerificationReport:
    """Test report formatting (markdown and JSON)."""

    def test_markdown_has_all_stages(self, simple_workbook):
        verifier = StageVerifier(path=simple_workbook)
        report = verifier.verify()
        md = report.to_markdown()
        assert "Sheet Chunking" in md
        assert "Cell Annotation" in md
        assert "Solid Block Identification" in md
        assert "Solid Table Identification" in md
        assert "Light Block Identification" in md
        assert "Template Extraction" in md
        assert "Synthetic-Model Export" in md

    def test_markdown_has_status_indicators(self, simple_workbook):
        verifier = StageVerifier(path=simple_workbook)
        report = verifier.verify()
        md = report.to_markdown()
        assert "[x]" in md  # implemented

    def test_markdown_has_metrics(self, simple_workbook):
        verifier = StageVerifier(path=simple_workbook)
        report = verifier.verify()
        md = report.to_markdown()
        assert "total_blocks" in md
        assert "total_cells" in md

    def test_coverage_percentage(self, simple_workbook):
        verifier = StageVerifier(path=simple_workbook)
        report = verifier.verify()
        # All 11 stages implemented = 100%
        assert report.overall_coverage_pct == 100.0

    def test_json_has_all_fields(self, simple_workbook):
        verifier = StageVerifier(path=simple_workbook)
        report = verifier.verify()
        j = report.to_json()
        assert "file_path" in j
        assert "workbook_hash" in j
        assert "overall_coverage_pct" in j
        assert len(j["stages"]) == 11


# ---------------------------------------------------------------------------
# Edge case tests
# ---------------------------------------------------------------------------


class TestStageVerifierEdgeCases:
    """Test edge cases for the stage verifier."""

    def test_empty_workbook(self, tmp_dir):
        from openpyxl import Workbook

        path = tmp_dir / "empty.xlsx"
        wb = Workbook()
        wb.save(path)
        verifier = StageVerifier(path=path)
        report = verifier.verify()
        assert len(report.stages) == 11
        # Empty workbook should still produce a valid report
        assert report.stages[0].metrics["total_blocks"] == 0

    def test_single_cell_workbook(self, tmp_dir):
        from openpyxl import Workbook

        path = tmp_dir / "single_cell.xlsx"
        wb = Workbook()
        wb.active["A1"] = "Hello"
        wb.save(path)
        verifier = StageVerifier(path=path)
        report = verifier.verify()
        assert report.total_cells >= 1
        assert report.stages[1].metrics["total_cells"] >= 1

    def test_mixed_content_all_stages(self, mixed_content_layout):
        verifier = StageVerifier(path=mixed_content_layout)
        report = verifier.verify()
        # Mixed content should produce table structures
        s3 = report.stages[3]
        assert s3.status == ImplementationStatus.IMPLEMENTED

    def test_cross_sheet_formulas(self, cross_sheet_formulas):
        verifier = StageVerifier(path=cross_sheet_formulas)
        report = verifier.verify()
        # Should see multiple sheets
        assert report.total_sheets == 3
        # Stage 1 should detect annotated cells
        s1 = report.stages[1]
        assert s1.metrics["annotated_cells"] > 0
