"""
Dump everything the header detector could have seen for DECO failure records.

Reads failure records produced by analyze_deco_headers.py --dump-buckets, groups
them by workbook (one parse per file), and for each failure prints a row-by-row
grid of the top of the GT table region: cell values, types, styling, merges,
plus the detector's row-level verdicts and sheet-level header hints
(freeze pane / print title rows). Built for root-causing detector failures.

Usage:
    PYTHONPATH=src python scripts/diag_header_failure.py \
        --failures RUN/failures.json --bucket MR-UNDER --start 0 --count 20 \
        [--corpus data/corpora/deco/completed] [--rows 10]
"""

from __future__ import annotations

import argparse
import datetime as _dt
import json
from collections import defaultdict
from pathlib import Path

from excel_parser.analysis import header_detector as hd
from excel_parser.pipeline import parse_workbook


def _cell_brief(cell, row: int, col: int, merge_map: dict) -> str:
    if cell is None or cell.is_empty:
        mk = merge_map.get((row, col))
        if mk and (mk[0] != row or mk[1] != col):
            return "·M"  # merged slave (empty, owned by master)
        return "·"
    v = cell.raw_value
    if isinstance(v, str):
        s = v.strip().replace("\n", "\\n")
        tag = "s"
    elif isinstance(v, bool):
        s, tag = str(v), "b"
    elif isinstance(v, (int, float)):
        s, tag = repr(v), "n"
    elif isinstance(v, (_dt.date, _dt.datetime)):
        s, tag = v.isoformat()[:10], "d"
    else:
        s, tag = repr(v), "?"
    if len(s) > 14:
        s = s[:13] + "…"
    flags = ""
    st = cell.style
    if st is not None:
        font = st.font
        if font and font.bold:
            flags += "B"
        fill = st.fill
        if fill and fill.fg_color:
            flags += "F"
    if cell.formula is not None:
        flags += "="
    mk = merge_map.get((row, col))
    if mk:
        flags += "M" if (mk[0] == row and mk[1] == col) else "m"
    return f"{s}[{tag}{flags}]"


def dump_failure(sheet, fail: dict, rows_to_show: int) -> None:
    r0, c0, r1, c1 = fail["gt_range"]
    print(f"\n{'=' * 100}")
    print(f"{fail['bucket']}  {fail['file']} :: {fail['sheet']}  range=({r0},{c0})-({r1},{c1})")
    print(
        f"  GT header rows: {fail['gt_hrows']}   "
        f"ks iso: {fail['ks_hrows_iso']}   ks e2e: {fail['ks_hrows_e2e']}"
    )
    props = sheet.properties
    print(f"  sheet hints: freeze_pane={props.freeze_pane!r} print_title_rows={props.print_title_rows!r}")
    merge_map: dict[tuple[int, int], tuple[int, int]] = {}
    for mr in sheet.merged_regions:
        tl, br = mr.range.top_left, mr.range.bottom_right
        for rr in range(tl.row, br.row + 1):
            for cc in range(tl.col, br.col + 1):
                merge_map[(rr, cc)] = (mr.master.row, mr.master.col)
    show_cols = min(c1, c0 + 11)
    last = min(r1, r0 + rows_to_show - 1)
    from excel_parser.models.common import CellCoord, CellRange

    reg = hd._Region.build(
        sheet,
        CellRange(
            top_left=CellCoord(row=r0, col=c0),
            bottom_right=CellCoord(row=r1, col=c1),
        ),
    )
    for rr in range(r0, last + 1):
        nonempty = reg.cells(rr)
        styled = reg.is_styled_row(rr)
        is_data = reg.is_data_row(rr)
        is_hdr = hd._is_header_anchor(reg, rr)
        marks = []
        if rr in fail["gt_hrows"]:
            marks.append("GT")
        if rr in fail["ks_hrows_iso"]:
            marks.append("KS")
        verdict = f"ne={len(nonempty):>2} styled={int(styled)} data={int(is_data)} hdr={int(is_hdr)}"
        cells = "|".join(
            _cell_brief(sheet.get_cell(rr, cc), rr, cc, merge_map) for cc in range(c0, show_cols + 1)
        )
        print(f"  r{rr:<5} {','.join(marks) or '--':<6} {verdict}  {cells}")


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--failures", required=True)
    ap.add_argument("--corpus", default="data/corpora/deco/completed")
    ap.add_argument("--bucket", default=None, help="filter to one bucket")
    ap.add_argument("--start", type=int, default=0)
    ap.add_argument("--count", type=int, default=20)
    ap.add_argument("--rows", type=int, default=10, help="rows of the region to show")
    args = ap.parse_args()

    fails = json.loads(Path(args.failures).read_text())
    if args.bucket:
        fails = [f for f in fails if f["bucket"] == args.bucket]
    fails = fails[args.start : args.start + args.count]
    if not fails:
        print("no failures matched")
        return 0

    by_file: dict[str, list[dict]] = defaultdict(list)
    for f in fails:
        by_file[f["file"]].append(f)

    for fname, group in by_file.items():
        path = Path(args.corpus) / fname
        try:
            result = parse_workbook(str(path))
        except Exception as exc:  # noqa: BLE001
            print(f"\n!! parse failed for {fname}: {exc}")
            continue
        sheets = {s.sheet_name: s for s in result.workbook.sheets}
        for fail in group:
            sheet = sheets.get(fail["sheet"])
            if sheet is None:
                print(f"\n!! sheet {fail['sheet']!r} missing in {fname}")
                continue
            dump_failure(sheet, fail, args.rows)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
