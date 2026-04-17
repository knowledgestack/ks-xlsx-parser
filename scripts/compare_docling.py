"""
Docling vs ks-xlsx-parser head-to-head comparison.

Runs both parsers against the same .xlsx files and scores them across
the five dimensions that matter for RAG + citations:

  1. Table detection      – tables found vs raw-row dumping
  2. Header propagation   – column headers attached to data cells
  3. Formula preservation – formulas stored alongside computed values
  4. Cell lineage         – sheet/row/col/address on every chunk
  5. Chunking quality     – RAG-ready text that includes context

Usage:
    python scripts/compare_docling.py [path/to/file.xlsx ...]

Defaults to the examples/ directory files.
"""

from __future__ import annotations

import sys
import json
import textwrap
from pathlib import Path
from dataclasses import dataclass, field

# ── paths ──────────────────────────────────────────────────────────────────────
REPO_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(REPO_ROOT / "src"))

DEFAULT_FILES = [
    REPO_ROOT / "testBench" / "real_world" / "Financials Sample Data.xlsx",
    REPO_ROOT / "testBench" / "real_world" / "financial_model.xlsx",
    REPO_ROOT / "testBench" / "real_world" / "sales_dashboard.xlsx",
    REPO_ROOT / "testBench" / "real_world" / "Walbridge Coatings 8.9.23.xlsx",
]


# ─────────────────────────────────────────────────────────────────────────────
# Score card
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class ScoreCard:
    parser: str
    file: str
    tables_detected: int = 0
    header_cells_tagged: int = 0       # cells marked as column/row headers
    total_data_cells: int = 0
    header_propagation_pct: float = 0  # % data cells that carry a header label
    formulas_preserved: int = 0
    formula_cells_total: int = 0
    lineage_complete: bool = False      # every cell has sheet+row+col
    chunk_count: int = 0
    chunks_with_context: int = 0       # chunks whose text contains ≥1 "=" or header word
    notes: list[str] = field(default_factory=list)

    def score(self) -> float:
        """Weighted 0-100 composite score."""
        s = 0.0
        # 1. table detection (25 pts) – at least 1 table found
        s += 25.0 if self.tables_detected > 0 else 0.0
        # 2. header propagation (25 pts)
        s += 25.0 * min(self.header_propagation_pct, 1.0)
        # 3. formula preservation (20 pts)
        if self.formula_cells_total > 0:
            s += 20.0 * min(self.formulas_preserved / self.formula_cells_total, 1.0)
        else:
            s += 20.0  # no formulas to preserve → not penalised
        # 4. lineage (15 pts)
        s += 15.0 if self.lineage_complete else 0.0
        # 5. chunking quality (15 pts)
        if self.chunk_count > 0:
            s += 15.0 * min(self.chunks_with_context / self.chunk_count, 1.0)
        return round(s, 1)

    def summary(self) -> str:
        lines = [
            f"  Parser       : {self.parser}",
            f"  File         : {Path(self.file).name}",
            f"  Score        : {self.score()} / 100",
            f"  Tables found : {self.tables_detected}",
            f"  Header tags  : {self.header_cells_tagged} / {self.total_data_cells} data cells",
            f"  Header prop  : {self.header_propagation_pct*100:.1f}%",
            f"  Formulas     : {self.formulas_preserved} / {self.formula_cells_total}",
            f"  Lineage OK   : {self.lineage_complete}",
            f"  Chunks       : {self.chunk_count}  ({self.chunks_with_context} with context)",
        ]
        for n in self.notes:
            lines.append(f"  NOTE: {n}")
        return "\n".join(lines)


# ─────────────────────────────────────────────────────────────────────────────
# Docling runner
# ─────────────────────────────────────────────────────────────────────────────

def run_docling(path: Path) -> ScoreCard:
    from docling.document_converter import DocumentConverter

    card = ScoreCard(parser="Docling", file=str(path))

    try:
        conv = DocumentConverter()
        result = conv.convert(str(path))
        doc = result.document
    except Exception as exc:
        card.notes.append(f"Parse error: {exc}")
        return card

    card.tables_detected = len(doc.tables)

    header_tagged = 0
    data_cells = 0
    formula_cells = 0      # docling doesn't expose formulas
    chunks_with_ctx = 0

    for table in doc.tables:
        cells = table.data.table_cells if table.data else []
        for cell in cells:
            if cell.column_header or cell.row_header or cell.row_section:
                header_tagged += 1
            else:
                data_cells += 1

        # Chunk = table rendered as markdown
        md = table.export_to_dataframe().to_string() if hasattr(table, "export_to_dataframe") else ""
        if not md:
            try:
                md = "\n".join(
                    " | ".join(c.text for c in row)
                    for row in _table_rows(table)
                )
            except Exception:
                md = ""
        if md.strip():
            chunks_with_ctx += 1
            card.chunk_count += 1

    # Also count non-table text chunks
    for text in doc.texts:
        card.chunk_count += 1
        if text.text.strip():
            chunks_with_ctx += 1

    card.header_cells_tagged = header_tagged
    card.total_data_cells = data_cells
    if (header_tagged + data_cells) > 0:
        card.header_propagation_pct = header_tagged / (header_tagged + data_cells)

    card.formulas_preserved = 0           # docling does not extract formulas
    card.formula_cells_total = 0          # unknown at this point
    card.lineage_complete = False          # docling tracks row/col offsets but not A1 address or sheet name per cell

    card.chunks_with_context = chunks_with_ctx
    card.notes.append("Docling does not expose raw formulas")
    card.notes.append("Cell A1 address / sheet lineage not in Docling output")
    return card


def _table_rows(table):
    """Helper: group table cells into rows."""
    from itertools import groupby
    cells = sorted(table.data.table_cells, key=lambda c: c.start_row_offset_idx)
    for _, row_cells in groupby(cells, key=lambda c: c.start_row_offset_idx):
        yield list(row_cells)


# ─────────────────────────────────────────────────────────────────────────────
# ks-xlsx-parser runner
# ─────────────────────────────────────────────────────────────────────────────

def run_xlsx_parser(path: Path) -> ScoreCard:
    from xlsx_parser.pipeline import parse_workbook
    import openpyxl

    card = ScoreCard(parser="ks-xlsx-parser", file=str(path))

    try:
        result = parse_workbook(path=path)
    except Exception as exc:
        card.notes.append(f"Parse error: {exc}")
        return card

    # Count real formula cells via openpyxl (data_only=False to see formulas)
    formula_cell_addresses: set[str] = set()
    try:
        wb_raw = openpyxl.load_workbook(str(path), data_only=False)
        for ws in wb_raw.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        formula_cell_addresses.add(
                            f"{ws.title}!{cell.coordinate}"
                        )
        wb_raw.close()
    except Exception:
        pass

    card.formula_cells_total = len(formula_cell_addresses)

    # Score from parsed workbook
    workbook = result.workbook
    chunks = result.chunks

    # Table detection: number of table structures detected
    card.tables_detected = len(getattr(workbook, "tables", []))
    if card.tables_detected == 0 and getattr(workbook, "table_structures", None):
        card.tables_detected = len(workbook.table_structures)

    # Header propagation: walk chunks and look at render_text
    # A good chunk contains the column header alongside the value.
    header_tagged = 0
    data_cells_seen = 0
    formulas_found = 0
    chunks_with_ctx = 0

    for chunk in chunks:
        card.chunk_count += 1
        rt = chunk.render_text or ""

        # Does this chunk's text contain a header row separator (pipe table)?
        has_table_format = "|" in rt and "|-" in rt
        has_formula_marker = "[=]" in rt or "formula" in rt.lower()
        if has_table_format or has_formula_marker:
            chunks_with_ctx += 1

        # Count cells with formulas that appear in our chunks
        cells = chunk.cells if hasattr(chunk, "cells") else []

    # Count via to_json() cell list
    parsed_json = result.to_json()
    for ch in parsed_json.get("chunks", []):
        for cell in ch.get("cells", []):
            data_cells_seen += 1
            if cell.get("formula"):
                formulas_found += 1
                addr = f"{ch['sheet_name']}!{cell['address']}"
                # Check if this formula cell was in our ground-truth set
                if addr in formula_cell_addresses or not formula_cell_addresses:
                    header_tagged += 1  # formula-having cell → proxy for data cell with context

    # Header propagation proxy: if render_text includes pipe-table with ≥2 rows,
    # the first row is the header. We measure what fraction of chunks have this.
    table_chunks = [
        ch for ch in parsed_json.get("chunks", [])
        if "|" in (ch.get("render_text") or "") and "|-" in (ch.get("render_text") or "")
    ]
    all_data_chunks = [
        ch for ch in parsed_json.get("chunks", [])
        if ch.get("block_type") in ("table", "assumptions_table", "data", "mixed")
    ]
    if all_data_chunks:
        card.header_propagation_pct = len(table_chunks) / len(all_data_chunks)
    else:
        card.header_propagation_pct = len(table_chunks) / max(len(parsed_json.get("chunks", [])), 1)

    card.header_cells_tagged = len(table_chunks)
    card.total_data_cells = len(all_data_chunks) or len(parsed_json.get("chunks", []))

    # Formula preservation
    card.formulas_preserved = formulas_found
    if card.formula_cells_total == 0:
        card.formula_cells_total = formulas_found  # treat all found as ground truth

    # Lineage: every chunk has sheet_name, top_left_cell, bottom_right_cell
    lineage_ok = all(
        ch.get("sheet_name") and ch.get("top_left") and ch.get("bottom_right")
        for ch in parsed_json.get("chunks", [])
    )
    card.lineage_complete = lineage_ok

    card.chunks_with_context = chunks_with_ctx

    return card


# ─────────────────────────────────────────────────────────────────────────────
# Head-to-head display
# ─────────────────────────────────────────────────────────────────────────────

def print_comparison(docling_card: ScoreCard, ks_card: ScoreCard) -> None:
    name = Path(docling_card.file).name
    print(f"\n{'═'*60}")
    print(f"  FILE: {name}")
    print(f"{'═'*60}")

    dims = [
        ("Table detection",     "tables_detected",          lambda c: f"{c.tables_detected}"),
        ("Header propagation",  "header_propagation_pct",   lambda c: f"{c.header_propagation_pct*100:.1f}%"),
        ("Formula preservation","formulas_preserved",        lambda c: f"{c.formulas_preserved}/{c.formula_cells_total}"),
        ("Lineage complete",    "lineage_complete",          lambda c: "YES" if c.lineage_complete else "NO"),
        ("Chunks",              "chunk_count",               lambda c: f"{c.chunk_count} ({c.chunks_with_context} w/ctx)"),
        ("TOTAL SCORE",         "score",                     lambda c: f"{c.score()} / 100"),
    ]

    col1 = 22
    col2 = 20
    col3 = 20

    header = f"  {'Dimension':<{col1}} {'Docling':<{col2}} {'ks-xlsx-parser':<{col3}}"
    print(header)
    print(f"  {'-'*col1} {'-'*col2} {'-'*col3}")

    for label, attr, fmt in dims:
        dval = fmt(docling_card)
        kval = fmt(ks_card)
        winner_d = ""
        winner_k = ""
        try:
            # Simple numeric comparison for winner highlight
            dn = float(docling_card.score() if attr == "score" else getattr(docling_card, attr, 0) or 0)
            kn = float(ks_card.score() if attr == "score" else getattr(ks_card, attr, 0) or 0)
            if kn > dn:
                winner_k = " ✓"
            elif dn > kn:
                winner_d = " ✓"
        except (TypeError, ValueError):
            pass
        print(f"  {label:<{col1}} {dval+winner_d:<{col2}} {kval+winner_k:<{col3}}")

    print()
    # Notes
    all_notes = [(docling_card.parser, n) for n in docling_card.notes] + \
                [(ks_card.parser, n) for n in ks_card.notes]
    for parser, note in all_notes:
        print(f"  [{parser}] {note}")


def print_global_summary(all_docling: list[ScoreCard], all_ks: list[ScoreCard]) -> None:
    print(f"\n{'═'*60}")
    print("  GLOBAL SUMMARY")
    print(f"{'═'*60}")

    avg_d = sum(c.score() for c in all_docling) / len(all_docling)
    avg_k = sum(c.score() for c in all_ks) / len(all_ks)

    print(f"  Files tested     : {len(all_docling)}")
    print(f"  Docling avg      : {avg_d:.1f} / 100")
    print(f"  ks-xlsx-parser avg   : {avg_k:.1f} / 100")
    winner = "ks-xlsx-parser" if avg_k > avg_d else "Docling"
    print(f"  Overall winner   : {winner}")
    print()

    print("  Per-file winner:")
    for d, k in zip(all_docling, all_ks):
        name = Path(d.file).name
        if k.score() > d.score():
            w = f"ks-xlsx-parser (+{k.score()-d.score():.1f})"
        elif d.score() > k.score():
            w = f"Docling (+{d.score()-k.score():.1f})"
        else:
            w = "TIE"
        print(f"    {name:<45} {w}")
    print()


# ─────────────────────────────────────────────────────────────────────────────
# Sample chunk diff
# ─────────────────────────────────────────────────────────────────────────────

def print_sample_chunks(path: Path) -> None:
    """Print one sample chunk from each parser for qualitative comparison."""
    print(f"\n{'─'*60}")
    print(f"  SAMPLE CHUNK COMPARISON  –  {path.name}")
    print(f"{'─'*60}")

    # Docling
    try:
        from docling.document_converter import DocumentConverter
        conv = DocumentConverter()
        result = conv.convert(str(path))
        doc = result.document
        if doc.tables:
            cells = doc.tables[0].data.table_cells[:20]
            rows: dict[int, list] = {}
            for c in cells:
                rows.setdefault(c.start_row_offset_idx, []).append(c)
            sample = "\n".join(
                "  " + " | ".join(c.text for c in sorted(r, key=lambda x: x.start_col_offset_idx))
                for r in list(rows.values())[:4]
            )
            print(f"\n  [Docling] First table, first 4 rows:")
            print(sample or "  (empty)")
        else:
            print("\n  [Docling] No tables found")
    except Exception as e:
        print(f"\n  [Docling] Error: {e}")

    # ks-xlsx-parser
    try:
        from xlsx_parser.pipeline import parse_workbook
        result = parse_workbook(path=path)
        parsed = result.to_json()
        chunks = parsed.get("chunks", [])
        # Find first table chunk
        table_chunk = next(
            (ch for ch in chunks if "|" in (ch.get("render_text") or "") and "|-" in (ch.get("render_text") or "")),
            chunks[0] if chunks else None,
        )
        if table_chunk:
            rt = table_chunk.get("render_text", "")
            preview = "\n".join("  " + line for line in rt.splitlines()[:8])
            print(f"\n  [ks-xlsx-parser] First table chunk ({table_chunk.get('block_type')}) @ {table_chunk.get('source_uri', '')}:")
            print(preview or "  (empty)")
        else:
            print("\n  [ks-xlsx-parser] No chunks found")
    except Exception as e:
        print(f"\n  [ks-xlsx-parser] Error: {e}")


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────

def main() -> None:
    files = [Path(p) for p in sys.argv[1:]] if sys.argv[1:] else DEFAULT_FILES
    files = [f for f in files if f.exists()]

    if not files:
        print("No xlsx files found. Pass paths as arguments or populate examples/.")
        sys.exit(1)

    print(f"\nComparing Docling vs ks-xlsx-parser on {len(files)} file(s)…\n")

    all_docling: list[ScoreCard] = []
    all_ks: list[ScoreCard] = []

    for path in files:
        print(f"  Parsing: {path.name} …", flush=True)
        d_card = run_docling(path)
        k_card = run_xlsx_parser(path)
        all_docling.append(d_card)
        all_ks.append(k_card)
        print_comparison(d_card, k_card)
        print_sample_chunks(path)

    print_global_summary(all_docling, all_ks)


if __name__ == "__main__":
    main()
