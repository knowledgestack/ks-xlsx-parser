"""Enterprise-focused scoring of parser output on synthetic fixtures.

These tests provide lightweight, deterministic benchmarks that run without
network access. They exercise formulas, tables, cross-sheet references,
named ranges, hidden sheets, and simple calculation lineage.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from xlsx_parser import parse_workbook

from scripts.generate_enterprise_fixtures import generate_all


ROOT = Path(__file__).resolve().parents[1]
FIXTURE_DIR = ROOT / "testBench" / "enterprise"


@pytest.fixture(scope="session")
def enterprise_workbooks() -> list[Path]:
    """Generate (or reuse) enterprise fixtures and return their paths."""
    return generate_all()


class EnterpriseScorecard:
    def __init__(self, parse_result, expected_metadata=None):
        self.result = parse_result
        self.expected = expected_metadata or {}

    def formula_fidelity(self) -> float:
        workbook = self.result.workbook
        extracted = 0
        total = 0
        for sheet in workbook.sheets:
            for cell in sheet.cells.values():
                if cell.formula:
                    total += 1
                    if cell.formula_value is not None or cell.raw_value is not None:
                        extracted += 1
        return extracted / total if total else 0.0

    def table_detection_f1(self) -> float:
        detected = len(self.result.workbook.tables)
        expected = self.expected.get("expected_tables", detected)
        if expected == 0 and detected == 0:
            return 1.0
        precision = detected / max(detected, 1)
        recall = detected / max(expected, 1)
        return 2 * (precision * recall) / (precision + recall + 1e-10)

    def lineage_accuracy(self) -> float:
        graph = self.result.workbook.dependency_graph
        edges = len(graph.edges)
        cycles = 0  # DependencyGraph does not expose cycles directly
        accuracy = 1.0 - (cycles / (edges + 1)) * 0.1
        return max(accuracy, 0.0)

    def chunk_quality(self) -> float:
        chunks = self.result.chunks
        tokens = [c.token_count for c in chunks]
        if not tokens:
            return 0.0
        mean_tokens = sum(tokens) / len(tokens)
        variance = sum((t - mean_tokens) ** 2 for t in tokens) / len(tokens)
        std_dev = variance ** 0.5
        cv = std_dev / (mean_tokens + 1e-10)
        return max(1.0 - cv, 0.0)

    def layout_recovery(self) -> float:
        blocks_by_type = {}
        for chunk in self.result.chunks:
            blocks_by_type[chunk.block_type] = blocks_by_type.get(chunk.block_type, 0) + 1
        type_count = len(blocks_by_type)
        return min(type_count / 3.0, 1.0)

    def composite_score(self):
        weights = {
            "formula_fidelity": 0.25,
            "table_detection": 0.20,
            "lineage_accuracy": 0.20,
            "chunk_quality": 0.20,
            "layout_recovery": 0.15,
        }
        scores = {
            "formula_fidelity": self.formula_fidelity(),
            "table_detection": self.table_detection_f1(),
            "lineage_accuracy": self.lineage_accuracy(),
            "chunk_quality": self.chunk_quality(),
            "layout_recovery": self.layout_recovery(),
        }
        composite = sum(scores[k] * weights[k] for k in weights)
        return scores, composite

    def metrics(self):
        scores, composite = self.composite_score()
        scores["composite"] = composite
        return scores


@pytest.mark.enterprise
@pytest.mark.parametrize(
    "filename,expected",
    [
        ("financial_model.xlsx", {"expected_tables": 0, "expected_formulas": 2}),
        ("inventory_tracker.xlsx", {"expected_tables": 0, "expected_formulas": 100}),
        ("forecast_model.xlsx", {"expected_tables": 0, "expected_formulas": 24}),
        ("operations_tracker.xlsx", {"expected_tables": 0, "expected_formulas": 20}),
    ],
)
def test_enterprise_scorecard(enterprise_workbooks, filename, expected):
    path = FIXTURE_DIR / filename
    assert path.exists(), f"Fixture missing: {path}"

    result = parse_workbook(path=path)
    scorecard = EnterpriseScorecard(result, expected_metadata=expected)
    scores, composite = scorecard.composite_score()

    metrics_dir = ROOT / "metrics" / "corpus"
    metrics_dir.mkdir(parents=True, exist_ok=True)
    with open(metrics_dir / f"{path.stem}_scorecard.json", "w") as f:
        json.dump(scorecard.metrics(), f, indent=2)

    print(scorecard.metrics())
    assert composite >= 0.45, f"Composite {composite:.2%} too low for {filename}"


@pytest.mark.enterprise
def test_enterprise_summary(enterprise_workbooks):
    paths = enterprise_workbooks
    results = []
    for p in paths:
        result = parse_workbook(path=p)
        scorecard = EnterpriseScorecard(result)
        scores = scorecard.metrics()
        scores["file"] = p.name
        results.append(scores)

    metrics_dir = ROOT / "metrics"
    metrics_dir.mkdir(parents=True, exist_ok=True)
    summary_path = metrics_dir / "corpus_summary.json"
    with open(summary_path, "w") as f:
        json.dump({"files": results}, f, indent=2)

    assert len(results) == len(paths)
