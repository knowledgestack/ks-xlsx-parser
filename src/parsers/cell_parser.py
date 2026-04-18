"""
Cell-level parsing: extract value, formula, style, and metadata from openpyxl cells.

This module converts openpyxl cell objects into CellDTO instances,
capturing raw values, display-formatted values, formulas, and
complete style snapshots. It handles merged cells, comments,
data validations, and hyperlinks.
"""

from __future__ import annotations

import logging
import re
from datetime import date, datetime, time
from typing import Any

from openpyxl.cell.cell import Cell as OpenpyxlCell
from openpyxl.cell.cell import MergedCell as OpenpyxlMergedCell

from models.cell import (
    AlignmentStyle,
    BorderSide,
    BorderStyle,
    CellDTO,
    CellStyle,
    FillStyle,
    FontStyle,
)
from models.common import CellCoord, RichTextRun

logger = logging.getLogger(__name__)

# Regex to extract cell and range references from formulas
# Matches: A1, $A$1, Sheet1!A1, Sheet1!$A$1:$B$10, 'Sheet Name'!A1
_REF_RE = re.compile(
    r"(?:'[^']+'\!|[A-Za-z_]\w*\!)?"  # Optional sheet prefix
    r"\$?[A-Z]{1,3}\$?\d+"            # Cell ref
    r"(?::\$?[A-Z]{1,3}\$?\d+)?",     # Optional :end range
)


class CellParser:
    """
    Parses individual openpyxl cell objects into CellDTO instances.

    Handles value extraction, type detection, formula capture,
    style serialization, and merge resolution. Designed to be
    reused across all cells in a sheet without state leakage.
    """

    def __init__(self, sheet_name: str):
        self._sheet_name = sheet_name
        # Cache for style objects to avoid re-serializing identical styles
        self._style_cache: dict[int, CellStyle] = {}

    def parse(
        self,
        cell: OpenpyxlCell | OpenpyxlMergedCell,
        computed_value: Any = None,
    ) -> CellDTO:
        """
        Parse a single openpyxl cell into a CellDTO.

        Args:
            cell: The openpyxl cell object.
            computed_value: The evaluated formula result from a data_only pass.

        Returns:
            A fully populated CellDTO (except cell_id/cell_hash which
            are set during finalize).
        """
        coord = CellCoord(row=cell.row, col=cell.column)

        # Handle merged slave cells
        if isinstance(cell, OpenpyxlMergedCell):
            return CellDTO(
                coord=coord,
                sheet_name=self._sheet_name,
                is_merged_slave=True,
                data_type="s",
            )

        # Extract raw value and formula
        raw_value = cell.value
        formula = None
        data_type = cell.data_type or "n"

        if isinstance(raw_value, str) and raw_value.startswith("="):
            formula = raw_value[1:]  # Strip leading '='
            raw_value = computed_value
            data_type = "f"
        elif cell.data_type == "f":
            formula = self._extract_formula_text(raw_value)
            raw_value = computed_value
            data_type = "f"

        # Determine display value
        display_value = self._format_display_value(raw_value, cell)
        # For formula cells with no computed value, show the formula as display
        if display_value is None and formula:
            display_value = f"={formula}"

        # Extract style
        style = self._extract_style(cell)

        # Extract comment
        comment_text = None
        comment_author = None
        if cell.comment:
            comment_text = cell.comment.text
            comment_author = cell.comment.author

        # Extract hyperlink
        hyperlink = None
        if cell.hyperlink:
            hyperlink = cell.hyperlink.target

        # Extract formula references
        formula_references: list[str] = []
        if formula:
            formula_references = _REF_RE.findall(formula)

        # Extract rich text runs
        rich_text_runs = self._extract_rich_text(cell)

        return CellDTO(
            coord=coord,
            sheet_name=self._sheet_name,
            raw_value=self._serialize_value(raw_value),
            display_value=display_value,
            data_type=data_type,
            formula=formula,
            formula_value=self._serialize_value(computed_value) if formula else None,
            formula_references=formula_references,
            rich_text_runs=rich_text_runs,
            style=style,
            comment_text=comment_text,
            comment_author=comment_author,
            hyperlink=hyperlink,
        )

    @staticmethod
    def _extract_formula_text(raw_value: Any) -> str | None:
        """Extract formula string from openpyxl cell value (handles ArrayFormula etc.)."""
        if raw_value is None:
            return None
        # ArrayFormula stores the formula in .text (includes leading '=')
        if hasattr(raw_value, "text") and raw_value.text is not None:
            text = str(raw_value.text).strip()
            return text[1:] if text.startswith("=") else text
        if isinstance(raw_value, str) and raw_value.startswith("="):
            return raw_value[1:]
        # Fallback: avoid str() on formula objects (produces <ArrayFormula ...>)
        if type(raw_value).__name__ in ("ArrayFormula", "DataTableFormula"):
            return None  # No readable formula text
        return str(raw_value) if raw_value else None

    @staticmethod
    def _extract_rich_text(cell: OpenpyxlCell) -> list[RichTextRun]:
        """Extract rich text runs if the cell contains mixed formatting."""
        try:
            from openpyxl.cell.rich_text import CellRichText
            if isinstance(cell.value, CellRichText):
                runs: list[RichTextRun] = []
                for part in cell.value:
                    if isinstance(part, str):
                        runs.append(RichTextRun(text=part))
                    elif hasattr(part, "text") and hasattr(part, "font"):
                        font = part.font
                        runs.append(RichTextRun(
                            text=part.text or "",
                            bold=bool(font.bold) if font and font.bold else False,
                            italic=bool(font.italic) if font and font.italic else False,
                            color=str(font.color.rgb)[2:] if font and font.color and font.color.rgb else None,
                            font_name=font.name if font else None,
                            font_size=font.size if font else None,
                        ))
                return runs
        except (ImportError, Exception):
            pass
        return []

    def _serialize_value(self, value: Any) -> Any:
        """Convert Python values to JSON-serializable types."""
        if value is None:
            return None
        if isinstance(value, datetime):
            return value.isoformat()
        if isinstance(value, date):
            return value.isoformat()
        if isinstance(value, time):
            return value.isoformat()
        if isinstance(value, (int, float, str, bool)):
            return value
        return str(value)

    def _format_display_value(self, value: Any, cell: OpenpyxlCell) -> str | None:
        """
        Format a cell value for display using its number format.

        This is a best-effort formatting; Excel's full format engine
        is complex. We handle common cases and fall back to str().
        """
        if value is None:
            return None

        number_format = getattr(cell, "number_format", None)

        if isinstance(value, bool):
            return "TRUE" if value else "FALSE"
        if isinstance(value, datetime):
            if number_format and "h" in number_format.lower():
                return value.strftime("%Y-%m-%d %H:%M:%S")
            return value.strftime("%Y-%m-%d")
        if isinstance(value, date):
            return value.strftime("%Y-%m-%d")
        if isinstance(value, time):
            return value.strftime("%H:%M:%S")
        if isinstance(value, float):
            if number_format and number_format != "General":
                return self._apply_number_format(value, number_format)
            if value == int(value):
                return str(int(value))
            return f"{value:g}"
        if isinstance(value, int):
            if number_format and number_format != "General":
                return self._apply_number_format(value, number_format)
            return str(value)

        return str(value)

    def _apply_number_format(self, value: int | float, fmt: str) -> str:
        """
        Best-effort application of Excel number formats.

        Handles common patterns like #,##0, 0.00, percentages.
        Falls back to str() for complex or locale-specific formats.
        """
        try:
            if "%" in fmt:
                return f"{value * 100:.{self._count_decimals(fmt)}f}%"
            if "#,##0" in fmt or ",0" in fmt:
                decimals = self._count_decimals(fmt)
                formatted = f"{value:,.{decimals}f}"
                return formatted
            if "0" in fmt:
                decimals = self._count_decimals(fmt)
                return f"{value:.{decimals}f}"
        except (ValueError, TypeError):
            pass
        return str(value)

    @staticmethod
    def _count_decimals(fmt: str) -> int:
        """Count decimal places in a number format string."""
        if "." not in fmt:
            return 0
        after_dot = fmt.split(".")[-1]
        return len([c for c in after_dot if c in ("0", "#")])

    def _extract_style(self, cell: OpenpyxlCell) -> CellStyle | None:
        """
        Extract the style of a cell as a CellStyle DTO.

        Uses a cache keyed by openpyxl's internal style ID to avoid
        re-serializing identical style objects.
        """
        # Use openpyxl's internal StyleArray as a collision-free cache key.
        # Note: id()-based keys on cell.font/fill/etc. are unreliable because
        # openpyxl properties return temporary objects that get garbage-collected,
        # causing id() reuse and false cache hits.
        style_key = tuple(cell._style)
        if style_key in self._style_cache:
            return self._style_cache[style_key]

        font = self._extract_font(cell)
        fill = self._extract_fill(cell)
        border = self._extract_border(cell)
        alignment = self._extract_alignment(cell)
        number_format = cell.number_format if cell.number_format != "General" else None

        # Check if any style is non-default
        has_style = any([
            font and (font.bold or font.italic or font.name or font.size or font.color),
            fill and fill.fg_color,
            border and any([border.left, border.right, border.top, border.bottom]),
            alignment and (alignment.horizontal or alignment.vertical or alignment.wrap_text),
            number_format,
        ])

        if not has_style:
            self._style_cache[style_key] = None
            return None

        style = CellStyle(
            font=font,
            fill=fill,
            border=border,
            alignment=alignment,
            number_format=number_format,
        )
        self._style_cache[style_key] = style
        return style

    def _extract_font(self, cell: OpenpyxlCell) -> FontStyle | None:
        """Extract font properties from a cell."""
        f = cell.font
        if not f:
            return None
        color = self._extract_color(f.color) if f.color else None
        return FontStyle(
            name=f.name,
            size=f.size,
            bold=bool(f.bold),
            italic=bool(f.italic),
            underline=f.underline if f.underline and f.underline != "none" else None,
            strikethrough=bool(f.strikethrough),
            color=color,
        )

    def _extract_fill(self, cell: OpenpyxlCell) -> FillStyle | None:
        """Extract fill/background properties from a cell."""
        f = cell.fill
        if not f or not f.patternType or f.patternType == "none":
            return None
        fg = self._extract_color(f.fgColor) if f.fgColor else None
        bg = self._extract_color(f.bgColor) if f.bgColor else None
        if not fg and not bg:
            return None
        return FillStyle(pattern_type=f.patternType, fg_color=fg, bg_color=bg)

    def _extract_border(self, cell: OpenpyxlCell) -> BorderStyle | None:
        """Extract border properties from a cell."""
        b = cell.border
        if not b:
            return None
        sides = {}
        for side_name in ("left", "right", "top", "bottom"):
            side = getattr(b, side_name, None)
            if side and side.style:
                color = self._extract_color(side.color) if side.color else None
                sides[side_name] = BorderSide(style=side.style, color=color)
        if not sides:
            return None
        return BorderStyle(**sides)

    def _extract_alignment(self, cell: OpenpyxlCell) -> AlignmentStyle | None:
        """Extract alignment properties from a cell."""
        a = cell.alignment
        if not a:
            return None
        if not any([a.horizontal, a.vertical, a.wrap_text, a.text_rotation, a.indent]):
            return None
        return AlignmentStyle(
            horizontal=a.horizontal,
            vertical=a.vertical,
            wrap_text=bool(a.wrap_text),
            text_rotation=a.textRotation if a.textRotation else None,
            indent=a.indent if a.indent else None,
        )

    @staticmethod
    def _extract_color(color) -> str | None:
        """Extract hex color string from openpyxl Color object."""
        if not color:
            return None
        if color.type == "rgb" and color.rgb:
            rgb = str(color.rgb)
            # Strip alpha channel if present (openpyxl stores as AARRGGBB)
            if len(rgb) == 8:
                return rgb[2:]
            return rgb
        if color.type == "theme":
            # Theme colors require the theme to resolve; store as theme index
            return f"theme:{color.theme}"
        if color.type == "indexed" and color.indexed is not None:
            return f"indexed:{color.indexed}"
        return None
