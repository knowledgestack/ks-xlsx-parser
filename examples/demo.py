#!/usr/bin/env python3
"""
Demo script showing the xlsx_parser in action on example workbooks.

Run: python examples/demo.py
"""

import json
import sys
from pathlib import Path

# Add src to path for development
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from xlsx_parser.pipeline import parse_workbook
from xlsx_parser.utils.logging_config import configure_logging

EXAMPLES_DIR = Path(__file__).parent


def demo_financial_model():
    """Parse the financial model and display results."""
    print("=" * 70)
    print("DEMO: Financial Model")
    print("=" * 70)

    path = EXAMPLES_DIR / "financial_model.xlsx"
    if not path.exists():
        print("  Run generate_examples.py first!")
        return

    result = parse_workbook(path=path)
    wb = result.workbook

    print(f"  File: {wb.filename}")
    print(f"  Hash: {wb.workbook_hash}")
    print(f"  Sheets: {wb.total_sheets}")
    print(f"  Cells: {wb.total_cells}")
    print(f"  Formulas: {wb.total_formulas}")
    print(f"  Tables: {len(wb.tables)}")
    print(f"  Charts: {len(wb.charts)}")
    print(f"  Named Ranges: {len(wb.named_ranges)}")
    print(f"  Dependency Edges: {len(wb.dependency_graph.edges)}")
    print(f"  Parse Time: {wb.parse_duration_ms:.0f}ms")
    print(f"  Errors: {len(wb.errors)}")

    print(f"\n  Chunks: {result.total_chunks}")
    print(f"  Total Tokens: {result.total_tokens}")

    # Show named ranges
    if wb.named_ranges:
        print("\n  Named Ranges:")
        for nr in wb.named_ranges:
            print(f"    {nr.name} → {nr.ref_string}")

    # Show first chunk
    if result.chunks:
        c = result.chunks[0]
        print(f"\n  First Chunk:")
        print(f"    ID: {c.chunk_id}")
        print(f"    Source: {c.source_uri}")
        print(f"    Type: {c.block_type}")
        print(f"    Tokens: {c.token_count}")
        print(f"    Text preview:")
        for line in c.render_text.split("\n")[:8]:
            print(f"      {line}")

    # Show a chart chunk if present
    chart_chunks = [c for c in result.chunks if "chart" in str(c.block_type)]
    if chart_chunks:
        cc = chart_chunks[0]
        print(f"\n  Chart Chunk:")
        print(f"    {cc.render_text[:200]}")

    print()


def demo_sales_dashboard():
    """Parse the sales dashboard and display results."""
    print("=" * 70)
    print("DEMO: Sales Dashboard")
    print("=" * 70)

    path = EXAMPLES_DIR / "sales_dashboard.xlsx"
    if not path.exists():
        print("  Run generate_examples.py first!")
        return

    result = parse_workbook(path=path)
    wb = result.workbook

    print(f"  Sheets: {[s.sheet_name for s in wb.sheets]}")
    print(f"  Tables: {[t.table_name for t in wb.tables]}")
    print(f"  Charts: {len(wb.charts)}")

    # Show table info
    for table in wb.tables:
        print(f"\n  Table: {table.display_name}")
        print(f"    Range: {table.ref_range.to_a1()}")
        print(f"    Columns: {[c.name for c in table.columns]}")
        print(f"    Style: {table.style_name}")

    # Show chart info
    for chart in wb.charts:
        print(f"\n  Chart: {chart.title} ({chart.chart_type.value})")
        print(f"    Series: {len(chart.series)}")
        for s in chart.series:
            print(f"      {s.name}: {s.values_ref}")

    # Show conditional formatting
    for sheet in wb.sheets:
        if sheet.conditional_format_rules:
            print(f"\n  Conditional Formatting ({sheet.sheet_name}):")
            for rule in sheet.conditional_format_rules:
                print(f"    {rule.rule_type} on {rule.ranges}")

    print()


def demo_engineering_calcs():
    """Parse the engineering calculations and show dependency chain."""
    print("=" * 70)
    print("DEMO: Engineering Calculations")
    print("=" * 70)

    path = EXAMPLES_DIR / "engineering_calcs.xlsx"
    if not path.exists():
        print("  Run generate_examples.py first!")
        return

    result = parse_workbook(path=path)
    wb = result.workbook

    print(f"  Formulas: {wb.total_formulas}")
    print(f"  Dependency Edges: {len(wb.dependency_graph.edges)}")
    print(f"  Named Ranges: {[nr.name for nr in wb.named_ranges]}")

    # Show dependency chain for Design Moment (C15)
    from xlsx_parser.models import CellCoord
    upstream = wb.dependency_graph.get_upstream(
        "Beam Design", CellCoord(row=15, col=3), max_depth=3
    )
    if upstream:
        print(f"\n  Upstream dependencies of C15 (Design Moment):")
        for edge in upstream:
            print(f"    {edge.source_sheet}!{edge.source_coord.to_a1()} → {edge.target_ref_string} ({edge.edge_type.value})")

    # Show chunks with formulas
    formula_chunks = [
        c for c in result.chunks
        if c.dependency_summary.upstream_refs
    ]
    if formula_chunks:
        print(f"\n  Chunks with dependencies: {len(formula_chunks)}")
        for c in formula_chunks[:3]:
            print(f"    {c.source_uri}")
            print(f"      Upstream: {c.dependency_summary.upstream_refs[:5]}")

    print()


def demo_json_output():
    """Show JSON serialization output for a small workbook."""
    print("=" * 70)
    print("DEMO: JSON Output (Financial Model)")
    print("=" * 70)

    path = EXAMPLES_DIR / "financial_model.xlsx"
    if not path.exists():
        print("  Run generate_examples.py first!")
        return

    result = parse_workbook(path=path)
    j = result.to_json()

    # Truncate render_text in chunks for display
    for c in j["chunks"]:
        if len(c["render_text"]) > 100:
            c["render_text"] = c["render_text"][:100] + "..."

    print(json.dumps(j, indent=2, default=str)[:3000])
    print("  ... (truncated)")

    # Show storage records
    serializer = result.serializer
    print(f"\n  Workbook record keys: {list(serializer.to_workbook_record().keys())}")
    print(f"  Sheet records: {len(serializer.to_sheet_records())}")
    print(f"  Chunk records: {len(serializer.to_chunk_records())}")
    print(f"  Vector entries: {len(serializer.to_vector_store_entries())}")

    print()


if __name__ == "__main__":
    configure_logging(structured=False)

    demo_financial_model()
    demo_sales_dashboard()
    demo_engineering_calcs()
    demo_json_output()

    print("All demos complete!")
