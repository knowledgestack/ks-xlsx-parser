"""
ks-xlsx-parser worker: long-running process that receives file paths on stdin
(one NDJSON line per file: {"path": "...", "request_id": "..."}) and emits
benchmark records on stdout (one NDJSON line per file).

Run directly: `python -m tests.benchmarks.adapters.ks_adapter`.

Handshake:
  out → `{"event":"ready","parser":"ks-xlsx-parser","version":"..."}`
  in  → one `{"path": ..., "request_id": ...}` per line
  out → one record per line (see _schema.BenchmarkRecord)
  in  → EOF
  out → `{"event":"done"}`
"""



import json
import os
import sys
import time
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

# Make the parent `tests/benchmarks` package importable when run as a script.
_HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(_HERE.parent.parent.parent))  # repo root
sys.path.insert(0, str(_HERE.parent.parent.parent / "src"))  # src layout

from tests.benchmarks._mem import peak_rss_mb  # noqa: E402
from tests.benchmarks._schema import BenchmarkRecord, SCHEMA_VERSION  # noqa: E402
from ks_xlsx_parser import __version__ as KS_VERSION  # noqa: E402
from ks_xlsx_parser import parse_workbook  # noqa: E402

MAX_ERR_LEN = 500

# Allow the benchmark harness to flip the parser into fast-mode via env var.
# "full" (default) keeps chunks + template/tree extraction; "fast" skips them.
# The parser name gets a suffix so summary.md and drift.md distinguish runs
# without needing a second adapter file.
_PARSE_MODE = os.environ.get("KS_PARSE_MODE", "full")
if _PARSE_MODE not in {"full", "fast"}:
    _PARSE_MODE = "full"
PARSER_NAME = "ks-xlsx-parser" if _PARSE_MODE == "full" else "ks-xlsx-parser-fast"


def _write(obj: dict[str, Any]) -> None:
    sys.stdout.write(json.dumps(obj, separators=(",", ":")) + "\n")
    sys.stdout.flush()


def _count_features(result: Any, file_size: int, path: str, parse_time_ms: float,
                    peak_mb: float, commit: str) -> BenchmarkRecord:
    wb = result.workbook
    chunks = result.chunks

    chart_types = sorted({c.chart_type.value for c in wb.charts})

    # Hyperlinks, comments, images
    hyperlinks = 0
    comments = 0
    for sheet in wb.sheets:
        for cell in sheet.cells.values():
            if getattr(cell, "hyperlink", None):
                hyperlinks += 1
            if getattr(cell, "comment_text", None):
                comments += 1
    images = sum(1 for sh in wb.shapes if sh.shape_type == "image")

    merges = sum(len(s.merged_regions) for s in wb.sheets)
    cf_rules = sum(len(s.conditional_format_rules) for s in wb.sheets)
    dv_rules = sum(len(s.data_validations) for s in wb.sheets)

    token_count = sum(getattr(c, "token_count", 0) or 0 for c in chunks)

    return BenchmarkRecord(
        file=path,
        file_size_bytes=file_size,
        parser=PARSER_NAME,
        parser_version=KS_VERSION,
        status="ok",
        error=None,
        parse_time_ms=parse_time_ms,
        peak_memory_mb=peak_mb,
        sheets=len(wb.sheets),
        cells=wb.total_cells,
        formulas=wb.total_formulas,
        formula_dependencies=len(wb.dependency_graph.edges),
        charts=len(wb.charts),
        chart_types=chart_types,
        tables=len(wb.tables),
        pivots=len(wb.pivot_tables),
        merges=merges,
        cf_rules=cf_rules,
        dv_rules=dv_rules,
        named_ranges=len(wb.named_ranges),
        hyperlinks=hyperlinks,
        images=images,
        comments=comments,
        sparklines=None,  # not modelled by ks-xlsx-parser
        chunks=len(chunks),
        token_count=token_count,
        schema_version=SCHEMA_VERSION,
        timestamp=datetime.now(UTC).isoformat(),
        harness_commit=commit,
    )


def _error_record(path: str, file_size: int, status: str, error: str,
                  parse_time_ms: float | None, peak_mb: float | None,
                  commit: str) -> BenchmarkRecord:
    return BenchmarkRecord(
        file=path,
        file_size_bytes=file_size,
        parser=PARSER_NAME,
        parser_version=KS_VERSION,
        status=status,
        error=error[:MAX_ERR_LEN] if error else None,
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
    _write({"event": "ready", "parser": PARSER_NAME, "version": KS_VERSION})

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
            result = parse_workbook(path=path, mode=_PARSE_MODE)
            t1 = time.perf_counter()
            rss1 = peak_rss_mb()
            rec = _count_features(
                result=result,
                file_size=file_size,
                path=path,
                parse_time_ms=(t1 - t0) * 1000.0,
                peak_mb=max(rss1 - rss0, 0.0),
                commit=commit,
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
