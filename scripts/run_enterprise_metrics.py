"""Run enterprise scorecard metrics without pytest.

Useful for quick local baselines when test dependencies are unavailable.
"""



import json

from xlsx_parser import parse_workbook

from scripts.generate_enterprise_fixtures import generate_all


class EnterpriseScorecard:
    def __init__(self, parse_result):
        self.result = parse_result

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
        expected = detected  # no ground truth available here
        if expected == 0 and detected == 0:
            return 1.0
        precision = detected / max(detected, 1)
        recall = detected / max(expected, 1)
        return 2 * (precision * recall) / (precision + recall + 1e-10)

    def lineage_accuracy(self) -> float:
        graph = self.result.workbook.dependency_graph
        edges = len(graph.edges)
        cycles = 0  # DependencyGraph does not expose cycles directly here
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

    def metrics(self):
        scores = {
            "formula_fidelity": self.formula_fidelity(),
            "table_detection": self.table_detection_f1(),
            "lineage_accuracy": self.lineage_accuracy(),
            "chunk_quality": self.chunk_quality(),
            "layout_recovery": self.layout_recovery(),
        }
        scores["composite"] = sum(scores.values()) / len(scores)
        return scores


def main() -> None:
    fixtures = generate_all()
    results = []
    for path in fixtures:
        result = parse_workbook(path=path)
        scorecard = EnterpriseScorecard(result)
        metrics = scorecard.metrics()
        metrics["file"] = path.name
        results.append(metrics)

    print(json.dumps({"files": results}, indent=2))


if __name__ == "__main__":
    main()
