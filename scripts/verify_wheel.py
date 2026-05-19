#!/usr/bin/env python3
"""Verify the built wheel is installable and importable in a clean venv.

This is the regression guard for the v0.2.0 packaging bug: ``pipeline.py``
and ``api.py`` were top-level modules under ``src/`` and ``setuptools``
``packages.find`` only picks up *packages*, so they were silently dropped
from the wheel — ``from ks_xlsx_parser.pipeline import ...`` failed for
every installed user. The flat layout also leaked 13 generic top-level
packages (``models``, ``utils``, ``parsers`` ...) into ``site-packages``.

Run after ``python -m build --wheel``. Exits non-zero on any problem so it
can gate CI and ``make wheel-check``.
"""
from __future__ import annotations

import subprocess
import sys
import tempfile
import venv
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

# Imports a real downstream consumer relies on. Keep in sync with the
# public surface in ks_xlsx_parser/__init__.py.
SMOKE_IMPORTS = [
    "from ks_xlsx_parser import parse_workbook, ParseResult",
    "from ks_xlsx_parser.pipeline import parse_workbook",
    "from ks_xlsx_parser.verification import StageVerifier",
    "from ks_xlsx_parser.analysis.table_assembler import TableAssembler",
    "from ks_xlsx_parser.models.workbook import WorkbookDTO",
]


def find_wheel() -> Path:
    wheels = sorted((ROOT / "dist").glob("*.whl"))
    if not wheels:
        sys.exit("ERROR: no wheel in dist/ — run `python -m build --wheel` first")
    return wheels[-1]


def check_wheel_contents(wheel: Path) -> None:
    """Fail loudly if the wheel pollutes the global namespace or drops modules."""
    with zipfile.ZipFile(wheel) as zf:
        names = zf.namelist()
        top_level = next((n for n in names if n.endswith("top_level.txt")), None)
        if top_level:
            packages = zf.read(top_level).decode().split()
            if packages != ["ks_xlsx_parser"]:
                sys.exit(
                    f"ERROR: wheel exposes top-level packages {packages}; "
                    "expected only ['ks_xlsx_parser']. The flat src/ layout leaked."
                )
    required = ["ks_xlsx_parser/pipeline.py", "ks_xlsx_parser/api.py"]
    for req in required:
        if not any(n == req for n in names):
            sys.exit(f"ERROR: wheel is missing {req}")
    print(f"wheel contents OK ({len(names)} entries, top-level: ks_xlsx_parser)")


def check_install_and_import(wheel: Path) -> None:
    with tempfile.TemporaryDirectory() as tmp:
        env_dir = Path(tmp) / "venv"
        venv.create(env_dir, with_pip=True)
        py = env_dir / ("Scripts" if sys.platform == "win32" else "bin") / "python"
        subprocess.run([str(py), "-m", "pip", "install", "-q", str(wheel)], check=True)
        script = "; ".join(SMOKE_IMPORTS) + "; print('clean-venv import OK')"
        subprocess.run([str(py), "-c", script], check=True)


def main() -> None:
    wheel = find_wheel()
    print(f"verifying {wheel.name}")
    check_wheel_contents(wheel)
    check_install_and_import(wheel)
    print("wheel verification PASSED")


if __name__ == "__main__":
    main()
