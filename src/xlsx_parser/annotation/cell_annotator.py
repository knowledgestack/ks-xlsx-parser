"""
Stage 1: Cell Annotation.

Annotates each cell as DATA or LABEL using a feature-based scoring system
with two passes: the first pass scores based on intrinsic cell features,
the second pass adds contextual features from neighboring cells.
"""

from __future__ import annotations

import logging
from collections import defaultdict

from ..models.cell import CellDTO
from ..models.common import CellAnnotation
from ..models.sheet import SheetDTO

logger = logging.getLogger(__name__)

# Keywords that suggest a cell is a label
_LABEL_KEYWORDS = frozenset({
    "total", "subtotal", "sum", "average", "count", "min", "max",
    "name", "date", "id", "type", "category", "description",
    "assumption", "input", "parameter", "scenario",
    "header", "title", "label", "key", "index",
    "revenue", "cost", "profit", "margin", "tax", "rate",
    "price", "quantity", "amount", "balance", "net", "gross",
    "year", "month", "quarter", "period", "fy", "q1", "q2", "q3", "q4",
})

# Feature weights for the scoring model
_WEIGHTS = {
    "bold": 0.25,
    "keyword": 0.20,
    "merged_master": 0.15,
    "content_type": 0.15,
    "position": 0.10,
    "formula": 0.05,
    "neighbor_context": 0.10,
}


class CellAnnotator:
    """
    Annotates cells as DATA or LABEL using a two-pass feature-based scorer.

    Pass 1: Score each cell using intrinsic features (bold, keyword, content type, etc.)
    Pass 2: Refine scores using neighboring cell annotations from Pass 1.

    The final annotation is LABEL if score >= 0.5, DATA otherwise.
    Confidence = abs(score - 0.5) * 2 (0.0 = uncertain, 1.0 = certain).
    """

    def __init__(self, sheet: SheetDTO):
        self._sheet = sheet

    def annotate(self) -> dict[str, CellAnnotation]:
        """
        Annotate all cells in the sheet.

        Returns:
            Dict mapping cell keys ("row,col") to their annotations.
        """
        if not self._sheet.cells:
            return {}

        # Build row/col index for neighbor lookups
        self._row_index: dict[int, list[CellDTO]] = defaultdict(list)
        self._col_index: dict[int, list[CellDTO]] = defaultdict(list)
        for cell in self._sheet.cells.values():
            self._row_index[cell.coord.row].append(cell)
            self._col_index[cell.coord.col].append(cell)

        # Compute used range for position features
        used = self._sheet.used_range or self._sheet.compute_used_range()
        self._min_row = used.top_left.row if used else 1
        self._max_row = used.bottom_right.row if used else 1
        self._min_col = used.top_left.col if used else 1
        self._max_col = used.bottom_right.col if used else 1

        # Pass 1: Score each cell using intrinsic features
        pass1_scores: dict[str, float] = {}
        for key, cell in self._sheet.cells.items():
            pass1_scores[key] = self._score_intrinsic(cell)

        # Assign preliminary annotations from Pass 1
        pass1_annotations: dict[str, CellAnnotation] = {}
        for key, score in pass1_scores.items():
            pass1_annotations[key] = (
                CellAnnotation.LABEL if score >= 0.5 else CellAnnotation.DATA
            )

        # Pass 2: Refine with neighbor context
        results: dict[str, CellAnnotation] = {}
        for key, cell in self._sheet.cells.items():
            base_score = pass1_scores[key]
            neighbor_score = self._score_neighbor_context(cell, pass1_annotations)

            # Blend: base features contribute (1 - neighbor_weight), neighbors contribute neighbor_weight
            neighbor_weight = _WEIGHTS["neighbor_context"]
            final_score = base_score * (1.0 - neighbor_weight) + neighbor_score * neighbor_weight

            annotation = CellAnnotation.LABEL if final_score >= 0.5 else CellAnnotation.DATA
            confidence = abs(final_score - 0.5) * 2.0

            cell.annotation = annotation
            cell.annotation_confidence = round(confidence, 4)
            results[key] = annotation

        label_count = sum(1 for a in results.values() if a == CellAnnotation.LABEL)
        logger.info(
            "Sheet '%s': annotated %d cells (%d LABEL, %d DATA)",
            self._sheet.sheet_name,
            len(results),
            label_count,
            len(results) - label_count,
        )
        return results

    def _score_intrinsic(self, cell: CellDTO) -> float:
        """Score a cell using intrinsic features. Higher = more likely LABEL."""
        scores: dict[str, float] = {}

        # Feature 1: Bold font
        is_bold = bool(cell.style and cell.style.font and cell.style.font.bold)
        scores["bold"] = 1.0 if is_bold else 0.0

        # Feature 2: Keyword match
        is_keyword = False
        if isinstance(cell.raw_value, str):
            lower = cell.raw_value.lower().strip()
            is_keyword = any(kw in lower for kw in _LABEL_KEYWORDS)
        scores["keyword"] = 1.0 if is_keyword else 0.0

        # Feature 3: Merged master cell
        scores["merged_master"] = 1.0 if cell.is_merged_master else 0.0

        # Feature 4: Content type (text = more likely label, number = more likely data)
        if cell.raw_value is None:
            scores["content_type"] = 0.5  # Neutral for empty
        elif isinstance(cell.raw_value, str):
            # Short text strings are more likely labels
            if len(cell.raw_value.strip()) <= 50:
                scores["content_type"] = 0.7
            else:
                scores["content_type"] = 0.4  # Long text could be data
        elif isinstance(cell.raw_value, bool):
            scores["content_type"] = 0.3
        elif isinstance(cell.raw_value, (int, float)):
            scores["content_type"] = 0.1  # Numbers are usually data
        else:
            scores["content_type"] = 0.3  # datetime, etc.

        # Feature 5: Position (first row/col more likely to be labels)
        row_position = 0.0
        col_position = 0.0
        total_rows = max(self._max_row - self._min_row + 1, 1)
        total_cols = max(self._max_col - self._min_col + 1, 1)
        row_frac = (cell.coord.row - self._min_row) / total_rows
        col_frac = (cell.coord.col - self._min_col) / total_cols

        if row_frac < 0.1:  # Top 10% of rows
            row_position = 0.8
        elif row_frac > 0.9:  # Bottom 10% (could be totals)
            row_position = 0.6
        else:
            row_position = 0.3

        if col_frac < 0.1:  # Leftmost 10% of columns
            col_position = 0.7
        else:
            col_position = 0.3

        scores["position"] = (row_position + col_position) / 2

        # Feature 6: Formula (formulas are usually data/calculations, not labels)
        scores["formula"] = 0.0 if cell.formula else 0.5

        # Weighted sum (excluding neighbor_context which is Pass 2)
        total = 0.0
        weight_sum = 0.0
        for feature, weight in _WEIGHTS.items():
            if feature == "neighbor_context":
                continue
            if feature in scores:
                total += scores[feature] * weight
                weight_sum += weight

        return total / weight_sum if weight_sum > 0 else 0.5

    def _score_neighbor_context(
        self,
        cell: CellDTO,
        annotations: dict[str, CellAnnotation],
    ) -> float:
        """
        Score based on neighboring cell annotations from Pass 1.

        If most neighbors are LABEL, this cell is more likely LABEL too.
        Looks at 4-connected neighbors (up, down, left, right).
        """
        neighbor_keys = [
            f"{cell.coord.row - 1},{cell.coord.col}",  # above
            f"{cell.coord.row + 1},{cell.coord.col}",  # below
            f"{cell.coord.row},{cell.coord.col - 1}",  # left
            f"{cell.coord.row},{cell.coord.col + 1}",  # right
        ]

        label_count = 0
        total = 0
        for key in neighbor_keys:
            if key in annotations:
                total += 1
                if annotations[key] == CellAnnotation.LABEL:
                    label_count += 1

        if total == 0:
            return 0.5  # No neighbors, neutral

        return label_count / total
