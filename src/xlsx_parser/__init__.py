"""
xlsx_parser - Production-grade Excel Workflow Parser for RAG + auditability.

Parses .xlsx workbooks into structured, loss-minimizing representations
preserving cell values, formulas, formatting, tables, charts, and layout
with full lineage and citation support.

Usage:
    from xlsx_parser import parse_workbook

    result = parse_workbook(path="workbook.xlsx")
    for chunk in result.chunks:
        print(chunk.source_uri, chunk.render_text[:100])
"""

__version__ = "0.1.1"

from .pipeline import (
    ParseResult,
    compare_workbooks,
    export_importer,
    parse_workbook,
)
from .verification import ExcellentStage, StageVerifier, VerificationReport

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
