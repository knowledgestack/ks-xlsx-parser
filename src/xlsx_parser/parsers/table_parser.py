"""
Table parser for Excel ListObject tables.

Extracts named table definitions from openpyxl worksheets, including
column names, structured references, totals rows, and banding styles.
Tables are first-class RAG retrieval units.
"""

from __future__ import annotations

import logging

from openpyxl.worksheet.worksheet import Worksheet as OpenpyxlWorksheet

from ..models.common import CellCoord, CellRange, col_letter_to_number
from ..models.table import TableColumn, TableDTO

logger = logging.getLogger(__name__)


class TableParser:
    """
    Extracts Excel ListObject table definitions from a worksheet.

    Parses table name, range, columns, totals row, and style info
    from openpyxl's table objects. Each table becomes a TableDTO.
    """

    def __init__(self, ws: OpenpyxlWorksheet, sheet_name: str):
        self._ws = ws
        self._sheet_name = sheet_name

    def parse_all(self) -> list[TableDTO]:
        """Extract all tables defined on this worksheet."""
        tables = []
        for table in self._ws.tables.values():
            try:
                dto = self._parse_table(table)
                tables.append(dto)
            except Exception as e:
                logger.warning(
                    "Failed to parse table '%s' on sheet '%s': %s",
                    getattr(table, "name", "unknown"),
                    self._sheet_name,
                    e,
                )
        return tables

    def _parse_table(self, table) -> TableDTO:
        """Parse a single openpyxl Table object into a TableDTO."""
        ref = table.ref  # e.g., "A1:D10"
        ref_range = self._parse_range(ref)

        # Parse columns
        columns = []
        if table.tableColumns:
            for i, col in enumerate(table.tableColumns):
                totals_fn = None
                totals_val = None
                if hasattr(col, "totalsRowFunction") and col.totalsRowFunction:
                    totals_fn = col.totalsRowFunction
                if hasattr(col, "totalsRowLabel") and col.totalsRowLabel:
                    totals_val = col.totalsRowLabel

                columns.append(TableColumn(
                    name=col.name,
                    column_index=i,
                    totals_function=totals_fn,
                    totals_value=totals_val,
                ))

        # Determine header and data ranges
        has_totals = bool(table.totalsRowCount)
        header_range = CellRange(
            top_left=ref_range.top_left,
            bottom_right=CellCoord(
                row=ref_range.top_left.row,
                col=ref_range.bottom_right.col,
            ),
        )

        data_start_row = ref_range.top_left.row + 1
        data_end_row = ref_range.bottom_right.row - (1 if has_totals else 0)

        data_range = None
        if data_start_row <= data_end_row:
            data_range = CellRange(
                top_left=CellCoord(row=data_start_row, col=ref_range.top_left.col),
                bottom_right=CellCoord(row=data_end_row, col=ref_range.bottom_right.col),
            )

        totals_range = None
        if has_totals:
            totals_range = CellRange(
                top_left=CellCoord(
                    row=ref_range.bottom_right.row,
                    col=ref_range.top_left.col,
                ),
                bottom_right=ref_range.bottom_right,
            )

        return TableDTO(
            table_name=table.name,
            display_name=table.displayName or table.name,
            sheet_name=self._sheet_name,
            ref_range=ref_range,
            header_range=header_range,
            data_range=data_range,
            totals_range=totals_range,
            columns=columns,
            style_name=table.tableStyleInfo.name if table.tableStyleInfo else None,
            show_first_column=bool(table.tableStyleInfo.showFirstColumn) if table.tableStyleInfo else False,
            show_last_column=bool(table.tableStyleInfo.showLastColumn) if table.tableStyleInfo else False,
            show_row_stripes=bool(table.tableStyleInfo.showRowStripes) if table.tableStyleInfo else True,
            show_column_stripes=bool(table.tableStyleInfo.showColumnStripes) if table.tableStyleInfo else False,
            has_auto_filter=bool(table.autoFilter),
            has_totals_row=has_totals,
        )

    @staticmethod
    def _parse_range(ref: str) -> CellRange:
        """Parse an A1-style range string like 'A1:D10' into a CellRange."""
        parts = ref.replace("$", "").split(":")
        if len(parts) != 2:
            raise ValueError(f"Invalid range reference: {ref}")
        return CellRange(
            top_left=TableParser._parse_coord(parts[0]),
            bottom_right=TableParser._parse_coord(parts[1]),
        )

    @staticmethod
    def _parse_coord(ref: str) -> CellCoord:
        """Parse an A1-style cell reference like 'B5' into a CellCoord."""
        ref = ref.replace("$", "")
        col_str = ""
        row_str = ""
        for ch in ref:
            if ch.isalpha():
                col_str += ch
            else:
                row_str += ch
        return CellCoord(
            row=int(row_str),
            col=col_letter_to_number(col_str),
        )
