"""
Formula reference parser.

Parses Excel formula strings to extract all cell and range references,
including A1-style refs, cross-sheet refs, structured table refs,
and external workbook refs. Does NOT evaluate formulas — only extracts
the dependency references for graph construction.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from models.common import CellCoord, CellRange, col_letter_to_number

# Optional Rust fast-path: ks_xlsx_core.scan_formula returns reference tuples
# in the same emit order as the Python regex pipeline. When available, we
# build ParsedReference objects from those tuples instead of running three
# regex passes per formula. See rust/ks_xlsx_core/src/formula.rs.
try:
    from ks_xlsx_core import scan_formula as _rust_scan_formula  # type: ignore[import-not-found]
except ImportError:  # pragma: no cover
    _rust_scan_formula = None  # type: ignore[assignment]


def _parsed_refs_from_rust(tuples) -> list[ParsedReference]:
    """Convert raw (kind, sheet, col1, row1, col2, row2, workbook, table, ref_string)
    tuples from the Rust scanner into ParsedReference objects.

    Kept out of the class body so `parse()` is a single call + list comprehension
    on the hot path.
    """
    out: list[ParsedReference] = []
    for kind, sheet, col1, row1, col2, row2, workbook, table, ref_string in tuples:
        if kind == "structured":
            out.append(ParsedReference(
                ref_string=ref_string,
                is_structured=True,
                table_name=table,
            ))
            continue
        coord: CellCoord | None = None
        rng: CellRange | None = None
        if col2 is not None and row2 is not None:
            rng = CellRange(
                top_left=CellCoord(row=row1, col=col1),
                bottom_right=CellCoord(row=row2, col=col2),
            )
        else:
            coord = CellCoord(row=row1, col=col1)
        out.append(ParsedReference(
            ref_string=ref_string,
            sheet_name=sheet,
            coord=coord,
            range=rng,
            is_external=kind in ("external_cell", "external_range"),
            external_workbook=workbook,
        ))
    return out


@dataclass
class ParsedReference:
    """A single reference extracted from a formula."""

    ref_string: str  # Original reference text from the formula
    sheet_name: str | None = None  # Target sheet (None = same sheet)
    coord: CellCoord | None = None  # Single cell reference
    range: CellRange | None = None  # Range reference
    is_external: bool = False  # References an external workbook
    external_workbook: str | None = None
    is_structured: bool = False  # Table structured reference
    table_name: str | None = None
    named_range: str | None = None  # Named range reference


class FormulaParser:
    """
    Extracts cell and range references from Excel formula strings.

    Supports:
    - Simple A1 refs: A1, B5, $A$1
    - Range refs: A1:B10, $A$1:$B$10
    - Cross-sheet refs: Sheet1!A1, 'Sheet Name'!A1:B10
    - External refs: [Book1.xlsx]Sheet1!A1
    - Structured table refs: Table1[Column1], Table1[[#Headers],[Column1]]
    - Named ranges: detected as bare identifiers not matching A1 pattern

    Does NOT parse:
    - R1C1 references (rarely used in stored formulas)
    - Array formula syntax {=...}
    """

    # A1-style cell reference: optional sheet prefix, column letters, row number
    _CELL_RE = re.compile(
        r"(?:"
        r"(?:'([^']+)'|([A-Za-z0-9_]+))"  # Sheet name (quoted or unquoted)
        r"!)?"
        r"\$?([A-Z]{1,3})\$?(\d{1,7})"  # Column letters + row number
        r"(?::(?:"
        r"(?:'([^']+)'|([A-Za-z0-9_]+))"  # Optional sheet in range end (unusual)
        r"!)?"
        r"\$?([A-Z]{1,3})\$?(\d{1,7}))?"  # Optional range end
    )

    # External workbook reference: [BookName]SheetName!Ref
    _EXTERNAL_RE = re.compile(
        r"\[([^\]]+)\]"  # Workbook name in brackets
        r"(?:'([^']+)'|([A-Za-z0-9_]+))"  # Sheet name
        r"!"
        r"\$?([A-Z]{1,3})\$?(\d{1,7})"  # Cell ref
        r"(?::\$?([A-Z]{1,3})\$?(\d{1,7}))?"  # Optional range end
    )

    # Structured table reference: TableName[ColumnName] or TableName[[specifier],[column]]
    _STRUCTURED_RE = re.compile(
        r"([A-Za-z_][A-Za-z0-9_.]*)"  # Table name
        r"\["
        r"([^\]]*)"  # Column spec (may contain nested brackets)
        r"\]"
    )

    # Known Excel function names (to avoid treating them as named ranges)
    _FUNCTIONS = frozenset({
        "SUM", "AVERAGE", "COUNT", "COUNTA", "COUNTIF", "COUNTIFS",
        "SUMIF", "SUMIFS", "VLOOKUP", "HLOOKUP", "INDEX", "MATCH",
        "IF", "IFERROR", "IFNA", "IFS", "SWITCH",
        "LEFT", "RIGHT", "MID", "LEN", "TRIM", "UPPER", "LOWER",
        "CONCATENATE", "CONCAT", "TEXTJOIN", "TEXT", "VALUE",
        "MAX", "MIN", "LARGE", "SMALL", "RANK",
        "ABS", "ROUND", "ROUNDUP", "ROUNDDOWN", "CEILING", "FLOOR",
        "MOD", "POWER", "SQRT", "LOG", "LN", "EXP",
        "AND", "OR", "NOT", "TRUE", "FALSE",
        "DATE", "TODAY", "NOW", "YEAR", "MONTH", "DAY",
        "HOUR", "MINUTE", "SECOND", "DATEVALUE", "TIMEVALUE",
        "OFFSET", "INDIRECT", "ROW", "COLUMN", "ROWS", "COLUMNS",
        "TRANSPOSE", "SORT", "FILTER", "UNIQUE", "SEQUENCE",
        "XLOOKUP", "XMATCH", "LET", "LAMBDA", "MAP", "REDUCE",
        "CHOOSE", "HYPERLINK", "TYPE", "ISBLANK", "ISERROR", "ISNA",
        "ISNUMBER", "ISTEXT", "ISLOGICAL", "ISREF",
        "SUMPRODUCT", "MMULT", "MINVERSE",
        "NPV", "IRR", "PMT", "PV", "FV", "RATE", "NPER",
        "STDEV", "VAR", "MEDIAN", "MODE", "PERCENTILE",
        "SUBSTITUTE", "REPLACE", "FIND", "SEARCH", "REPT",
        "NUMBERVALUE", "FIXED", "DOLLAR",
        "ADDRESS", "CELL", "INFO", "N", "NA",
        "ARRAYTOTEXT", "VALUETOTEXT",
        "PI", "RAND", "RANDBETWEEN",
    })

    # Bare identifier that could be a named range
    _IDENT_RE = re.compile(r"\b([A-Za-z_][A-Za-z0-9_.]{1,})\b")

    def parse(self, formula: str, source_sheet: str) -> list[ParsedReference]:
        """
        Parse a formula string and return all references found.

        Args:
            formula: The formula string (without leading '=').
            source_sheet: The sheet containing this formula (for relative refs).

        Returns:
            List of ParsedReference objects.
        """
        # Rust fast path — byte-level scanner, ~10-20× faster than the three
        # regex passes for workbooks with thousands of formulas. The scanner
        # itself dedupes; emit order matches this Python implementation so
        # downstream classification in `dependency_builder` is identical.
        if _rust_scan_formula is not None:
            return _parsed_refs_from_rust(_rust_scan_formula(formula))

        refs: list[ParsedReference] = []
        seen: set[str] = set()  # Dedup references

        # 1. External references (must be matched first to avoid partial matches)
        for m in self._EXTERNAL_RE.finditer(formula):
            ref_str = m.group(0)
            if ref_str in seen:
                continue
            seen.add(ref_str)

            workbook = m.group(1)
            sheet = m.group(2) or m.group(3)
            col1, row1 = m.group(4), int(m.group(5))
            col2_str, row2_str = m.group(6), m.group(7)

            coord = CellCoord(row=row1, col=col_letter_to_number(col1))
            rng = None
            if col2_str and row2_str:
                rng = CellRange(
                    top_left=coord,
                    bottom_right=CellCoord(
                        row=int(row2_str), col=col_letter_to_number(col2_str)
                    ),
                )
                coord = None

            refs.append(ParsedReference(
                ref_string=ref_str,
                sheet_name=sheet,
                coord=coord,
                range=rng,
                is_external=True,
                external_workbook=workbook,
            ))

        # 2. Structured table references
        for m in self._STRUCTURED_RE.finditer(formula):
            ref_str = m.group(0)
            if ref_str in seen:
                continue
            table_name = m.group(1)
            # Skip if it looks like a function call
            if table_name.upper() in self._FUNCTIONS:
                continue
            seen.add(ref_str)
            refs.append(ParsedReference(
                ref_string=ref_str,
                is_structured=True,
                table_name=table_name,
            ))

        # 3. Standard A1-style references (with optional sheet prefix)
        for m in self._CELL_RE.finditer(formula):
            ref_str = m.group(0)
            if ref_str in seen:
                continue

            # Skip if this is part of an external ref already matched
            pos = m.start()
            if pos > 0 and formula[pos - 1] == "]":
                continue

            seen.add(ref_str)

            sheet = m.group(1) or m.group(2)  # Quoted or unquoted sheet name
            col1 = m.group(3)
            row1 = int(m.group(4))
            col2_str = m.group(7)
            row2_str = m.group(8)

            coord = CellCoord(row=row1, col=col_letter_to_number(col1))

            rng = None
            if col2_str and row2_str:
                rng = CellRange(
                    top_left=coord,
                    bottom_right=CellCoord(
                        row=int(row2_str), col=col_letter_to_number(col2_str)
                    ),
                )
                coord = None

            refs.append(ParsedReference(
                ref_string=ref_str,
                sheet_name=sheet,
                coord=coord,
                range=rng,
            ))

        return refs
