"""
ks_xlsx_parser — public API entry point for the ks-xlsx-parser package.

The source tree is flat: top-level modules at ``src/`` (``pipeline``,
``models``, ``analysis``, ``verification``, etc.). This module re-exports
the stable, user-facing names so callers can do::

    from ks_xlsx_parser import parse_workbook, ParseResult

regardless of internal layout.
"""
from __future__ import annotations

__version__ = "0.1.1"

from pipeline import (  # noqa: F401
    ParseResult,
    compare_workbooks,
    export_importer,
    parse_workbook,
)
from verification import (  # noqa: F401
    ExcellentStage,
    StageVerifier,
    VerificationReport,
)

__all__ = [
    "parse_workbook",
    "compare_workbooks",
    "export_importer",
    "ParseResult",
    "StageVerifier",
    "VerificationReport",
    "ExcellentStage",
    "__version__",
]
