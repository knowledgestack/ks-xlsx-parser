"""
Type-aware value comparison for cross-validation between parsers.

Handles known divergences between openpyxl and calamine: int/float,
date/datetime, bool coercion, None vs empty string, etc.
"""

from __future__ import annotations

import datetime
from dataclasses import dataclass
from typing import Any


@dataclass
class Mismatch:
    """Record of a value mismatch between parsers."""

    sheet: str
    row: int
    col: int
    a1_ref: str
    parser_value: Any
    calamine_value: Any
    category: str  # "type", "value", "missing"


def is_empty(val: Any) -> bool:
    """Check if a value is effectively empty."""
    if val is None:
        return True
    if isinstance(val, str) and val.strip() == "":
        return True
    return False


def values_match(parser_val: Any, calamine_val: Any, tolerance: float = 1e-9) -> bool:
    """
    Compare a parser value against a calamine value with type normalization.

    Returns True if the values are considered equivalent.
    """
    # Both empty
    if is_empty(parser_val) and is_empty(calamine_val):
        return True

    # One empty, one not
    if is_empty(parser_val) or is_empty(calamine_val):
        return False

    # Bool comparison (before numeric, since bool is subclass of int)
    if isinstance(parser_val, bool) or isinstance(calamine_val, bool):
        return bool(parser_val) == bool(calamine_val)

    # Numeric comparison with tolerance
    if isinstance(parser_val, (int, float)) and isinstance(calamine_val, (int, float)):
        pf = float(parser_val)
        cf = float(calamine_val)
        if cf == 0 and pf == 0:
            return True
        if cf == 0:
            return abs(pf) < tolerance
        return abs(pf - cf) / max(abs(cf), 1.0) < tolerance

    # Date/datetime comparison (both are date types)
    if isinstance(parser_val, (datetime.date, datetime.datetime)) and isinstance(
        calamine_val, (datetime.date, datetime.datetime)
    ):
        return _dates_equal(parser_val, calamine_val)

    # Cross-type: parser has ISO string, calamine has date/datetime
    if isinstance(parser_val, str) and isinstance(
        calamine_val, (datetime.date, datetime.datetime)
    ):
        parsed = _try_parse_iso(parser_val)
        if parsed is not None:
            return _dates_equal(parsed, calamine_val)

    # Cross-type: calamine has ISO string, parser has date/datetime
    if isinstance(calamine_val, str) and isinstance(
        parser_val, (datetime.date, datetime.datetime)
    ):
        parsed = _try_parse_iso(calamine_val)
        if parsed is not None:
            return _dates_equal(parser_val, parsed)

    # Time comparison
    if isinstance(parser_val, datetime.time) and isinstance(
        calamine_val, datetime.time
    ):
        return parser_val == calamine_val

    # Timedelta comparison
    if isinstance(parser_val, datetime.timedelta) and isinstance(
        calamine_val, datetime.timedelta
    ):
        return abs(parser_val.total_seconds() - calamine_val.total_seconds()) < 1.0

    # String comparison (strip whitespace)
    return str(parser_val).strip() == str(calamine_val).strip()


def _dates_equal(a: datetime.date | datetime.datetime, b: datetime.date | datetime.datetime) -> bool:
    """Compare two date/datetime values, normalizing to datetime."""
    if isinstance(a, datetime.date) and not isinstance(a, datetime.datetime):
        a = datetime.datetime.combine(a, datetime.time())
    if isinstance(b, datetime.date) and not isinstance(b, datetime.datetime):
        b = datetime.datetime.combine(b, datetime.time())
    return a == b


def _try_parse_iso(s: str) -> datetime.datetime | None:
    """Try to parse an ISO datetime string."""
    s = s.strip()
    for fmt in (
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%dT%H:%M:%S.%f",
        "%Y-%m-%d",
    ):
        try:
            return datetime.datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def compare_cell_value(
    parser_cell,
    calamine_val: Any,
    tolerance: float = 1e-9,
) -> bool:
    """
    Compare a CellDTO against a calamine value.

    For formula cells, compare calamine's computed value against
    the parser's formula_value (cached result). For non-formula cells,
    compare raw_value.
    """
    if parser_cell is None:
        return is_empty(calamine_val)

    # Skip merged slave cells — calamine returns None for them
    if parser_cell.is_merged_slave:
        return True

    # Choose the right parser value
    if parser_cell.formula:
        parser_val = parser_cell.formula_value
        # If formula_value is None (no cached value), skip comparison
        if parser_val is None:
            return True
    else:
        parser_val = parser_cell.raw_value

    return values_match(parser_val, calamine_val, tolerance)
