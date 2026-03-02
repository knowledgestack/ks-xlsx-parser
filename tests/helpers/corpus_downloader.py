"""
Download real-world Excel corpora for robustness testing.

Supports downloading .xlsx files from:
- EUSES Spreadsheet Corpus (Zenodo)
- Enron Spreadsheet Corpus (SheetJS GitHub)
- Additional GitHub repositories with public xlsx samples
"""

from __future__ import annotations

import io
import logging
import zipfile
from pathlib import Path

import requests

logger = logging.getLogger(__name__)

# Timeout for HTTP requests
_TIMEOUT = 60


def download_and_extract_xlsx(
    url: str,
    target_dir: Path,
    max_files: int = 50,
) -> list[Path]:
    """
    Download a ZIP archive and extract .xlsx files.

    Returns list of extracted .xlsx paths.
    """
    target_dir.mkdir(parents=True, exist_ok=True)
    files: list[Path] = []

    try:
        logger.info("Downloading %s ...", url)
        resp = requests.get(url, timeout=_TIMEOUT, stream=True)
        resp.raise_for_status()
        content = resp.content

        if not zipfile.is_zipfile(io.BytesIO(content)):
            logger.warning("URL did not return a ZIP file: %s", url)
            return files

        with zipfile.ZipFile(io.BytesIO(content)) as zf:
            xlsx_names = [
                n for n in zf.namelist()
                if n.lower().endswith(".xlsx")
                and not n.startswith("__MACOSX")
                and not n.startswith(".")
            ]
            logger.info("Found %d .xlsx files in archive", len(xlsx_names))

            for name in xlsx_names[:max_files]:
                safe_name = Path(name).name
                if not safe_name:
                    continue
                dest = target_dir / safe_name
                if not dest.exists():
                    dest.write_bytes(zf.read(name))
                files.append(dest)

    except requests.RequestException as e:
        logger.warning("Failed to download %s: %s", url, e)
    except zipfile.BadZipFile as e:
        logger.warning("Bad ZIP from %s: %s", url, e)

    return files


def download_single_xlsx(
    url: str,
    target_dir: Path,
    filename: str | None = None,
) -> Path | None:
    """Download a single .xlsx file."""
    target_dir.mkdir(parents=True, exist_ok=True)
    fname = filename or url.rsplit("/", 1)[-1]
    dest = target_dir / fname

    if dest.exists():
        return dest

    try:
        resp = requests.get(url, timeout=_TIMEOUT)
        resp.raise_for_status()
        dest.write_bytes(resp.content)
        return dest
    except requests.RequestException as e:
        logger.warning("Failed to download %s: %s", url, e)
        return None


def download_github_xlsx_samples(
    target_dir: Path,
    max_files: int = 50,
) -> list[Path]:
    """
    Download diverse .xlsx samples from known public GitHub repos.

    These are individual files from repos known to contain xlsx samples.
    """
    target_dir.mkdir(parents=True, exist_ok=True)
    files: list[Path] = []

    # Known public xlsx sample URLs (raw GitHub links)
    sample_urls = [
        # SheetJS test files
        ("https://raw.githubusercontent.com/SheetJS/sheetjs/master/test_files/comments_stress_test.xlsx", "comments_stress.xlsx"),
        ("https://raw.githubusercontent.com/SheetJS/sheetjs/master/test_files/merge_cells.xlsx", "merge_cells.xlsx"),
        ("https://raw.githubusercontent.com/SheetJS/sheetjs/master/test_files/number_format.xlsx", "number_format.xlsx"),
        # openpyxl test fixtures
        ("https://raw.githubusercontent.com/openpyxl/openpyxl/master/openpyxl/tests/data/genuine/empty.xlsx", "genuine_empty.xlsx"),
    ]

    for url, fname in sample_urls:
        if len(files) >= max_files:
            break
        path = download_single_xlsx(url, target_dir, fname)
        if path:
            files.append(path)

    return files


def download_euses_corpus(
    target_dir: Path,
    max_files: int = 50,
) -> list[Path]:
    """
    Download EUSES corpus from Zenodo and extract .xlsx files.

    Note: EUSES is mostly .xls files. The .xlsx subset may be small.
    """
    url = "https://zenodo.org/records/581673/files/EUSES.zip"
    return download_and_extract_xlsx(url, target_dir, max_files)


def download_enron_corpus(
    target_dir: Path,
    max_files: int = 50,
) -> list[Path]:
    """
    Download Enron spreadsheets from SheetJS GitHub repo.

    Note: Enron files are almost entirely .xls (pre-2007).
    The .xlsx subset will likely be empty or very small.
    """
    url = "https://github.com/SheetJS/enron_xls/archive/refs/heads/master.zip"
    return download_and_extract_xlsx(url, target_dir, max_files)


def get_corpus_files(corpus_dir: Path) -> list[Path]:
    """Return all .xlsx files under a corpus directory."""
    if not corpus_dir.exists():
        return []
    return sorted(corpus_dir.glob("**/*.xlsx"))
