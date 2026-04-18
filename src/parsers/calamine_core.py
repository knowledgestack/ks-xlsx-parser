"""
Adapter layer for the calamine-based fast cell reader.

This module is the *only* place in ``src/`` that imports ``ks_xlsx_core``
(our Rust/PyO3 crate). The rest of the parser calls into the functions
defined here. That isolation is deliberate — when python-calamine gains
formula support upstream, replacing the backend is a change to this file
only, nothing else in ``src/`` moves.

See ``rust/ks_xlsx_core/REMOVAL.md`` for the full removal procedure.

Public surface:
  read_cells(path, content) -> dict[str, dict[(row, col), CellData]]
    Keyed by sheet name, then (row_1idx, col_1idx). CellData is a
    ``_CellData`` namedtuple (value, formula, dtype). Returns ``None`` if
    the Rust backend is unavailable so the caller can fall back to openpyxl.
"""
from __future__ import annotations

import io
import logging
import os
import tempfile
from collections import namedtuple
from pathlib import Path

logger = logging.getLogger(__name__)

_CellData = namedtuple("_CellData", ["value", "formula", "dtype"])

try:
    import ks_xlsx_core  # type: ignore[import-not-found]
    _HAS_CORE = True
except ImportError:  # pragma: no cover - crate is an optional build artefact
    ks_xlsx_core = None  # type: ignore[assignment]
    _HAS_CORE = False


def available() -> bool:
    """True when the Rust core is importable in this process."""
    return _HAS_CORE


def read_cells(
    path: str | Path | None = None,
    content: bytes | None = None,
) -> dict[str, dict[tuple[int, int], _CellData]] | None:
    """Read values + formulas + dtypes for every non-empty cell in the workbook.

    Returns ``None`` if the Rust backend is unavailable or fails. Callers
    must handle that case (openpyxl fallback).

    Coordinates returned are **1-indexed** — identical to openpyxl's
    convention, so the downstream parser can address cells with the same
    keys it already uses.
    """
    if not _HAS_CORE:
        return None

    tmp_path: str | None = None
    try:
        if path is not None:
            target = str(path)
        else:
            # ks_xlsx_core's read_workbook takes a path today. Writing to
            # tempfile is cheap vs. the parse itself, and keeps the Rust
            # surface narrow — one entry point, path-based.
            fd, tmp_path = tempfile.mkstemp(suffix=".xlsx")
            with os.fdopen(fd, "wb") as f:
                f.write(content or b"")
            target = tmp_path

        raw = ks_xlsx_core.read_workbook(target)
    except Exception as exc:  # noqa: BLE001
        logger.warning("ks_xlsx_core read failed, falling back: %s", exc)
        return None
    finally:
        if tmp_path is not None:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass

    out: dict[str, dict[tuple[int, int], _CellData]] = {}
    for sheet_obj in raw:
        name = sheet_obj["name"]
        sheet_cells: dict[tuple[int, int], _CellData] = {}
        for row, col, value, formula, dtype in sheet_obj["cells"]:
            sheet_cells[(row, col)] = _CellData(
                value=value,
                formula=formula,
                dtype=dtype,
            )
        out[name] = sheet_cells
    return out


__all__ = ["available", "read_cells", "_CellData"]
