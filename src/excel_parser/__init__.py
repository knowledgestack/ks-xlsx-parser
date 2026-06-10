"""
excel_parser — public API entry point for the excel-parser package.

All modules live under the ``excel_parser`` package (``pipeline``,
``models``, ``analysis``, ``verification``, etc.). This module re-exports
the stable, user-facing names so callers can do::

    from excel_parser import parse_workbook, ParseResult

and submodules remain importable directly::

    from excel_parser.pipeline import parse_workbook
"""
from __future__ import annotations

__version__ = "0.2.1"

from .pipeline import (  # noqa: F401
    ParseResult,
    compare_workbooks,
    export_importer,
    parse_workbook,
)
from .verification import (  # noqa: F401
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
