"""
Structural invariant checker for WorkbookDTO output.

Runs a suite of invariants that must hold for any valid parse output,
regardless of the input file. Returns a list of violation messages.
"""

from __future__ import annotations

import re

from xlsx_parser.models.common import EdgeType


def check_invariants(workbook) -> list[str]:
    """
    Run all structural invariants on a parsed WorkbookDTO.

    Returns a list of violation description strings. An empty list means
    all invariants passed.
    """
    violations = []
    violations.extend(_check_merge_invariants(workbook))
    violations.extend(_check_used_range_invariants(workbook))
    violations.extend(_check_cell_identity_invariants(workbook))
    violations.extend(_check_dependency_invariants(workbook))
    violations.extend(_check_stats_invariants(workbook))
    violations.extend(_check_table_invariants(workbook))
    return violations


def _check_merge_invariants(workbook) -> list[str]:
    """M1-M4: Merge structure invariants."""
    violations = []
    for sheet in workbook.sheets:
        for cell in sheet.cells.values():
            # M1: slave has master
            if cell.is_merged_slave and cell.merge_master is None:
                violations.append(
                    f"M1: {cell.a1_ref} is merged slave with no merge_master"
                )

            # M2: master exists and is flagged
            if cell.is_merged_slave and cell.merge_master is not None:
                master = sheet.get_cell(
                    cell.merge_master.row, cell.merge_master.col
                )
                if master is None:
                    violations.append(
                        f"M2: {cell.a1_ref} merge_master "
                        f"{cell.merge_master.to_a1()} does not exist"
                    )
                elif not master.is_merged_master:
                    violations.append(
                        f"M2: {cell.merge_master.to_a1()} is referenced as "
                        f"master by {cell.a1_ref} but is_merged_master=False"
                    )

            # M4: not both master and slave
            if cell.is_merged_master and cell.is_merged_slave:
                violations.append(
                    f"M4: {cell.a1_ref} is both master and slave"
                )

        # M3: master is top_left of region
        for region in sheet.merged_regions:
            if region.master != region.range.top_left:
                violations.append(
                    f"M3: MergedRegion {region.range.to_a1()} master "
                    f"{region.master.to_a1()} != top_left "
                    f"{region.range.top_left.to_a1()}"
                )

    return violations


def _check_used_range_invariants(workbook) -> list[str]:
    """U1-U3: Used range invariants."""
    violations = []
    for sheet in workbook.sheets:
        if not sheet.cells:
            continue

        # U1: non-empty sheet has used_range
        if sheet.used_range is None:
            violations.append(
                f"U1: Sheet '{sheet.sheet_name}' has {len(sheet.cells)} "
                f"cells but used_range is None"
            )
            continue

        # U2: every cell within used_range
        for cell in sheet.cells.values():
            if not sheet.used_range.contains(cell.coord):
                violations.append(
                    f"U2: {cell.a1_ref} at ({cell.coord.row},{cell.coord.col}) "
                    f"is outside used_range {sheet.used_range.to_a1()}"
                )

        # U3: bounds are tight
        if sheet.cells:
            min_row = min(c.coord.row for c in sheet.cells.values())
            min_col = min(c.coord.col for c in sheet.cells.values())
            max_row = max(c.coord.row for c in sheet.cells.values())
            max_col = max(c.coord.col for c in sheet.cells.values())

            if sheet.used_range.top_left.row != min_row:
                violations.append(
                    f"U3: Sheet '{sheet.sheet_name}' used_range top row "
                    f"{sheet.used_range.top_left.row} != min cell row {min_row}"
                )
            if sheet.used_range.top_left.col != min_col:
                violations.append(
                    f"U3: Sheet '{sheet.sheet_name}' used_range left col "
                    f"{sheet.used_range.top_left.col} != min cell col {min_col}"
                )
            if sheet.used_range.bottom_right.row != max_row:
                violations.append(
                    f"U3: Sheet '{sheet.sheet_name}' used_range bottom row "
                    f"{sheet.used_range.bottom_right.row} != max cell row {max_row}"
                )
            if sheet.used_range.bottom_right.col != max_col:
                violations.append(
                    f"U3: Sheet '{sheet.sheet_name}' used_range right col "
                    f"{sheet.used_range.bottom_right.col} != max cell col {max_col}"
                )

    return violations


def _check_cell_identity_invariants(workbook) -> list[str]:
    """C1-C4: Cell identity invariants."""
    violations = []
    hex_pattern = re.compile(r"^[0-9a-f]{16}$")
    all_ids: list[str] = []

    for sheet in workbook.sheets:
        for cell in sheet.cells.values():
            # C1: cell_id non-empty
            if not cell.cell_id:
                violations.append(f"C1: Cell at {cell.a1_ref} has empty cell_id")
                continue

            all_ids.append(cell.cell_id)

            # C3: cell_hash is valid hex
            if not cell.cell_hash:
                violations.append(f"C3: {cell.a1_ref} has empty cell_hash")
            elif not hex_pattern.match(cell.cell_hash):
                violations.append(
                    f"C3: {cell.a1_ref} hash '{cell.cell_hash}' "
                    f"is not valid 16-char hex"
                )

            # C4: cell_id format
            expected_id = f"{cell.sheet_name}|{cell.coord.row}|{cell.coord.col}"
            if cell.cell_id != expected_id:
                violations.append(
                    f"C4: {cell.a1_ref} cell_id '{cell.cell_id}' != "
                    f"expected '{expected_id}'"
                )

    # C2: unique cell_ids
    if len(all_ids) != len(set(all_ids)):
        seen = set()
        dupes = set()
        for cid in all_ids:
            if cid in seen:
                dupes.add(cid)
            seen.add(cid)
        violations.append(
            f"C2: {len(dupes)} duplicate cell_ids: "
            f"{list(dupes)[:5]}"
        )

    return violations


def _check_dependency_invariants(workbook) -> list[str]:
    """F1-F4: Dependency graph invariants."""
    violations = []
    sheet_map = {s.sheet_name: s for s in workbook.sheets}

    edge_ids: list[str] = []
    for edge in workbook.dependency_graph.edges:
        # F1: edge source has a formula
        source_sheet = sheet_map.get(edge.source_sheet)
        if source_sheet is None:
            violations.append(
                f"F1: Edge source sheet '{edge.source_sheet}' not found"
            )
        else:
            source_cell = source_sheet.get_cell(
                edge.source_coord.row, edge.source_coord.col
            )
            if source_cell is None:
                violations.append(
                    f"F1: Edge source "
                    f"{edge.source_sheet}!{edge.source_coord.to_a1()} "
                    f"not found in cells"
                )
            elif source_cell.formula is None:
                violations.append(
                    f"F1: Edge source {source_cell.a1_ref} has no formula"
                )

        # F2: cross-sheet edges cross sheets
        if edge.edge_type == EdgeType.CROSS_SHEET:
            if edge.target_sheet is None or edge.target_sheet == edge.source_sheet:
                violations.append(
                    f"F2: Cross-sheet edge from "
                    f"{edge.source_sheet}!{edge.source_coord.to_a1()} "
                    f"has target_sheet={edge.target_sheet}"
                )

        # F4: collect edge_ids for uniqueness check
        if edge.edge_id:
            edge_ids.append(edge.edge_id)

    # F3: formula count matches stats
    actual_formulas = sum(
        1 for s in workbook.sheets for c in s.cells.values() if c.formula is not None
    )
    if workbook.total_formulas != actual_formulas:
        violations.append(
            f"F3: total_formulas={workbook.total_formulas} != "
            f"actual count={actual_formulas}"
        )

    # F4: unique edge_ids
    if len(edge_ids) != len(set(edge_ids)):
        violations.append(
            f"F4: {len(edge_ids) - len(set(edge_ids))} duplicate edge_ids"
        )

    return violations


def _check_stats_invariants(workbook) -> list[str]:
    """S1-S3: Stats consistency invariants."""
    violations = []

    # S1: total_sheets
    if workbook.total_sheets != len(workbook.sheets):
        violations.append(
            f"S1: total_sheets={workbook.total_sheets} != "
            f"len(sheets)={len(workbook.sheets)}"
        )

    # S2: total_cells
    actual_cells = sum(s.cell_count() for s in workbook.sheets)
    if workbook.total_cells != actual_cells:
        violations.append(
            f"S2: total_cells={workbook.total_cells} != "
            f"actual={actual_cells}"
        )

    return violations


def _check_table_invariants(workbook) -> list[str]:
    """T1-T3: Table invariants."""
    violations = []
    sheet_names = {s.sheet_name for s in workbook.sheets}
    table_names: list[str] = []

    for table in workbook.tables:
        # T1: table sheet exists
        if table.sheet_name not in sheet_names:
            violations.append(
                f"T1: Table '{table.table_name}' references "
                f"sheet '{table.sheet_name}' which doesn't exist"
            )

        # T2: column count matches range width
        if len(table.columns) != table.ref_range.col_count():
            violations.append(
                f"T2: Table '{table.table_name}' has "
                f"{len(table.columns)} columns but range "
                f"{table.ref_range.to_a1()} spans "
                f"{table.ref_range.col_count()} cols"
            )

        # T3: unique table names
        table_names.append(table.table_name)

    if len(table_names) != len(set(table_names)):
        seen = set()
        dupes = [n for n in table_names if n in seen or seen.add(n)]
        violations.append(f"T3: Duplicate table names: {dupes}")

    return violations
