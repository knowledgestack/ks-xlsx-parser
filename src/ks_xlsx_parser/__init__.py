"""
ks_xlsx_parser — import alias for ``xlsx_parser``.

The canonical Python module is ``xlsx_parser``; this alias exists so that
``import ks_xlsx_parser`` works, matching the PyPI package name
``ks-xlsx-parser`` exactly (PEP 503 normalisation converts dashes to
underscores for imports).

All names are re-exported from the upstream module; there is no
functional difference between ``from xlsx_parser import X`` and
``from ks_xlsx_parser import X``.
"""
from __future__ import annotations

from xlsx_parser import (  # noqa: F401
    ExcellentStage,
    ParseResult,
    StageVerifier,
    VerificationReport,
    __version__,
    compare_workbooks,
    export_importer,
    parse_workbook,
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
