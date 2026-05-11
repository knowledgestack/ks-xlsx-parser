"""
Docling worker: same NDJSON-on-stdio protocol as ks_adapter, but for the
IBM Docling DocumentConverter.

Docling treats xlsx as a document → markdown. We map its native objects
to the BenchmarkRecord schema as faithfully as we can; capabilities it
does NOT model are left as ``None`` (the schema's load-bearing distinction
between "feature absent in this file" and "feature not modeled at all").

What docling models for xlsx:
  - sheets        ← one table per sheet (its convention)
  - cells         ← sum of table cells across tables
  - tables        ← len(doc.tables)
  - merges        ← cells with row_span/col_span > 1

What it does NOT model:
  - formulas, formula_dependencies — computed values only, no `=…`
  - charts, chart_types
  - pivots
  - conditional formatting / data validation
  - named ranges
  - hyperlinks per cell, comments, sparklines
  - images

Chunks: docling has a chunker but we keep this adapter shallow. We export
the full document to markdown and emit it as a single chunk; the retrieval
metric splits on table boundaries and headings downstream.
"""

from __future__ import annotations

import json
import os
import sys
import time
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

_HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(_HERE.parent.parent.parent))  # repo root

from tests.benchmarks._mem import peak_rss_mb  # noqa: E402
from tests.benchmarks._schema import BenchmarkRecord, SCHEMA_VERSION  # noqa: E402

PARSER_NAME = "docling"
MAX_ERR_LEN = 500


def _write(obj: dict[str, Any]) -> None:
    sys.stdout.write(json.dumps(obj, separators=(",", ":")) + "\n")
    sys.stdout.flush()


def _build_converter():
    """Build a DocumentConverter once and reuse across files."""
    from docling.document_converter import DocumentConverter

    return DocumentConverter()


def _docling_version() -> str:
    try:
        import importlib.metadata

        return importlib.metadata.version("docling")
    except Exception:
        try:
            import docling

            return getattr(docling, "__version__", "unknown")
        except Exception:
            return "unknown"


def _count_features(doc, file_size: int, path: str, parse_time_ms: float,
                    peak_mb: float, commit: str, markdown: str) -> BenchmarkRecord:
    # Tables: docling produces one TableItem per detected table region.
    tables = list(doc.tables)
    n_tables = len(tables)

    # Cells: sum of table_cells. Docling does not surface cell counts
    # outside tables, so this undercounts free-floating text cells.
    n_cells = 0
    n_merges = 0
    for t in tables:
        cells = t.data.table_cells if t.data else []
        for c in cells:
            n_cells += 1
            row_span = (c.end_row_offset_idx or 0) - (c.start_row_offset_idx or 0)
            col_span = (c.end_col_offset_idx or 0) - (c.start_col_offset_idx or 0)
            if row_span > 1 or col_span > 1:
                n_merges += 1

    # Sheets: docling does not expose sheet count directly. The common
    # case is one table per sheet, but multi-table sheets exist. Best
    # available proxy: count unique pages — but docling for xlsx uses
    # one page per sheet, so len(doc.pages) ≈ sheet count.
    try:
        n_sheets = len(doc.pages) if doc.pages else None
    except Exception:
        n_sheets = None

    return BenchmarkRecord(
        file=path,
        file_size_bytes=file_size,
        parser=PARSER_NAME,
        parser_version=_docling_version(),
        status="ok",
        error=None,
        parse_time_ms=parse_time_ms,
        peak_memory_mb=peak_mb,
        sheets=n_sheets if n_sheets is not None else 1,
        cells=n_cells,
        formulas=None,                # not modeled
        formula_dependencies=None,    # not modeled
        charts=None,                  # not modeled for xlsx
        chart_types=None,
        tables=n_tables,
        pivots=None,                  # not modeled
        merges=n_merges,
        cf_rules=None,                # not modeled
        dv_rules=None,                # not modeled
        named_ranges=None,            # not modeled
        hyperlinks=None,              # not surfaced per-cell
        images=None,                  # not modeled for xlsx
        comments=None,                # not modeled
        sparklines=None,              # not modeled
        chunks=1,                     # single markdown blob — retrieval step re-chunks
        token_count=max(len(markdown) // 4, 1),  # crude estimate
        schema_version=SCHEMA_VERSION,
        timestamp=datetime.now(UTC).isoformat(),
        harness_commit=commit,
        extra={"markdown": markdown},
    )


def _error_record(path: str, file_size: int, status: str, error: str,
                  parse_time_ms: float | None, peak_mb: float | None,
                  commit: str) -> BenchmarkRecord:
    return BenchmarkRecord(
        file=path,
        file_size_bytes=file_size,
        parser=PARSER_NAME,
        parser_version=_docling_version(),
        status=status,
        error=(error or "")[:MAX_ERR_LEN] or None,
        parse_time_ms=parse_time_ms,
        peak_memory_mb=peak_mb,
        sheets=None,
        cells=None,
        formulas=None,
        formula_dependencies=None,
        charts=None,
        chart_types=None,
        tables=None,
        pivots=None,
        merges=None,
        cf_rules=None,
        dv_rules=None,
        named_ranges=None,
        hyperlinks=None,
        images=None,
        comments=None,
        sparklines=None,
        chunks=None,
        token_count=None,
        schema_version=SCHEMA_VERSION,
        timestamp=datetime.now(UTC).isoformat(),
        harness_commit=commit,
    )


def main() -> int:
    commit = os.environ.get("HARNESS_COMMIT", "")
    converter = _build_converter()
    _write({"event": "ready", "parser": PARSER_NAME, "version": _docling_version()})

    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            msg = json.loads(line)
            path = msg["path"]
        except Exception as exc:
            _write({"event": "error", "error": f"bad input line: {exc}"})
            continue

        try:
            file_size = os.path.getsize(path)
        except OSError:
            file_size = 0

        rss0 = peak_rss_mb()
        t0 = time.perf_counter()
        try:
            result = converter.convert(path)
            doc = result.document
            md = doc.export_to_markdown()
            t1 = time.perf_counter()
            rss1 = peak_rss_mb()
            rec = _count_features(
                doc=doc,
                file_size=file_size,
                path=path,
                parse_time_ms=(t1 - t0) * 1000.0,
                peak_mb=max(rss1 - rss0, 0.0),
                commit=commit,
                markdown=md,
            )
        except Exception as exc:  # noqa: BLE001
            t1 = time.perf_counter()
            rss1 = peak_rss_mb()
            rec = _error_record(
                path=path,
                file_size=file_size,
                status="error",
                error=f"{type(exc).__name__}: {exc}",
                parse_time_ms=(t1 - t0) * 1000.0,
                peak_mb=max(rss1 - rss0, 0.0),
                commit=commit,
            )

        sys.stdout.write(rec.to_json_line())
        sys.stdout.flush()

    _write({"event": "done"})
    return 0


if __name__ == "__main__":
    sys.exit(main())
