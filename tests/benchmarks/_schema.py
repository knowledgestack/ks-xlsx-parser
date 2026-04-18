"""
Canonical record schema for one (file, parser) benchmark row.

Null vs zero is load-bearing:
  - `None`  → parser does not model this feature
  - `0`     → parser models it and measured zero occurrences

The summary and drift generators treat them differently. Adapters MUST
respect this distinction.
"""
from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any

SCHEMA_VERSION = 1


@dataclass
class BenchmarkRecord:
    """One row per (file, parser) pair. Emitted as NDJSON, flattened to CSV."""

    file: str
    file_size_bytes: int
    parser: str
    parser_version: str
    status: str  # "ok" | "error" | "timeout" | "oom"
    error: str | None
    parse_time_ms: float | None
    peak_memory_mb: float | None

    sheets: int | None
    cells: int | None
    formulas: int | None
    formula_dependencies: int | None
    charts: int | None
    chart_types: list[str] | None
    tables: int | None
    pivots: int | None
    merges: int | None
    cf_rules: int | None
    dv_rules: int | None
    named_ranges: int | None
    hyperlinks: int | None
    images: int | None
    comments: int | None
    sparklines: int | None

    chunks: int | None
    token_count: int | None

    schema_version: int = SCHEMA_VERSION
    timestamp: str = ""
    harness_commit: str = ""
    extra: dict[str, Any] = field(default_factory=dict)

    def to_json_line(self) -> str:
        return json.dumps(asdict(self), separators=(",", ":"), default=_json_default) + "\n"

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> BenchmarkRecord:
        validate_record(d)
        known = {f for f in cls.__dataclass_fields__}
        extra = {k: v for k, v in d.items() if k not in known}
        kwargs = {k: v for k, v in d.items() if k in known}
        kwargs["extra"] = {**kwargs.get("extra", {}), **extra}
        return cls(**kwargs)


REQUIRED_FIELDS = {
    "file",
    "file_size_bytes",
    "parser",
    "parser_version",
    "status",
    "parse_time_ms",
    "peak_memory_mb",
    "sheets",
    "cells",
    "formulas",
    "formula_dependencies",
    "charts",
    "chart_types",
    "tables",
    "pivots",
    "merges",
    "cf_rules",
    "dv_rules",
    "named_ranges",
    "hyperlinks",
    "images",
    "comments",
    "sparklines",
    "chunks",
    "token_count",
}

NULLABLE_ON_OK = {
    "formula_dependencies",
    "chart_types",
    "sparklines",
    "chunks",
    "token_count",
}


def validate_record(d: dict[str, Any]) -> None:
    """Raise ValueError if the dict cannot be coerced into a BenchmarkRecord."""
    missing = REQUIRED_FIELDS - d.keys()
    if missing:
        raise ValueError(f"record missing required fields: {sorted(missing)}")

    status = d["status"]
    if status not in {"ok", "error", "timeout", "oom"}:
        raise ValueError(f"unknown status: {status!r}")

    if status == "ok":
        for numeric in ("sheets", "cells", "formulas", "parse_time_ms"):
            if d[numeric] is None:
                raise ValueError(f"status=ok but {numeric} is None")


CSV_FIELDS = [
    "file",
    "file_size_bytes",
    "parser",
    "parser_version",
    "status",
    "error",
    "parse_time_ms",
    "peak_memory_mb",
    "sheets",
    "cells",
    "formulas",
    "formula_dependencies",
    "charts",
    "chart_types",
    "tables",
    "pivots",
    "merges",
    "cf_rules",
    "dv_rules",
    "named_ranges",
    "hyperlinks",
    "images",
    "comments",
    "sparklines",
    "chunks",
    "token_count",
    "schema_version",
    "timestamp",
    "harness_commit",
]


def record_to_csv_row(rec: BenchmarkRecord) -> list[str]:
    """Flatten a record into the CSV row order. None → empty string."""
    d = asdict(rec)
    row: list[str] = []
    for col in CSV_FIELDS:
        v = d.get(col)
        if v is None:
            row.append("")
        elif isinstance(v, list):
            row.append("|".join(str(x) for x in v))
        else:
            row.append(str(v))
    return row


def _json_default(obj: Any) -> Any:
    if isinstance(obj, Path):
        return str(obj)
    if isinstance(obj, set):
        return sorted(obj)
    raise TypeError(f"not JSON-serializable: {type(obj).__name__}")
