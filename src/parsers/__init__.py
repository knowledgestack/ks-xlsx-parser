"""
Parsers for extracting data from Excel workbooks.

Entry point: WorkbookParser.parse() → WorkbookDTO
"""

from .workbook_parser import WorkbookParser

__all__ = ["WorkbookParser"]
