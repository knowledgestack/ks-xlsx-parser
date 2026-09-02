"""
Download real-world Excel corpora for robustness testing.

Sources are chosen for durability over volume:

- **PyPI sdists** of libraries that bundle .xlsx test fixtures. A published
  PyPI artifact is immutable, so these downloads cannot rot; only a yanked
  release would break them, and the version is resolved at run time.
- **Upstream test-fixture directories** on GitHub, enumerated through the
  contents API rather than a hand-written file list, so individual renames
  upstream do not silently reduce the corpus to nothing.

An earlier revision pointed at the SheetJS ``test_files`` submodule (now
blocked by GitHub), the EUSES archive on Zenodo (record withdrawn) and the
Enron ``.xls`` dump (321 MB, ~0 .xlsx). All four URLs 404'd or yielded
nothing, which is why every downloader here returns the paths it actually
obtained and the tests assert on that.
"""


import io
import logging
import tarfile
import zipfile
from pathlib import Path

import requests

logger = logging.getLogger(__name__)

# Timeout for HTTP requests
_TIMEOUT = 60

# Upstream fixture directories: (repo, ref, path). Enumerated via the GitHub
# contents API. These are deliberately parser test-suites — they concentrate
# the malformed, edge-case and adversarial workbooks that a robustness corpus
# wants (missing dimensions, empty shared strings, merged ranges, pivots,
# rich text, shared formulas, non-standard XML namespaces).
_GITHUB_FIXTURE_DIRS: tuple[tuple[str, str, str], ...] = (
    ("pandas-dev/pandas", "main", "pandas/tests/io/data/excel"),
    ("tafia/calamine", "master", "tests"),
)

# PyPI packages whose sdists ship .xlsx fixtures.
_PYPI_PACKAGES: tuple[str, ...] = (
    "python-calamine",
    "xlsx2csv",
    "pyexcel-xlsx",
    "excelrd",
    "tablib",
    "pyexcel",
)


def _is_safe_member(name: str) -> bool:
    """Reject archive members that are not plain, in-tree .xlsx files."""
    if not name.lower().endswith(".xlsx"):
        return False
    if name.startswith(("/", "__MACOSX", ".")):
        return False
    # Path traversal guard: no member may escape the extraction directory.
    return ".." not in Path(name).parts


def _write_member(target_dir: Path, name: str, data: bytes) -> Path | None:
    """
    Write one archive member into target_dir under its basename.

    Members that are not real XLSX containers are skipped. Upstream test
    suites ship negative fixtures under an .xlsx name — a zero-byte file, an
    encrypted OLE workbook — which belong in targeted unit tests, not in a
    corpus whose whole purpose is asserting that well-formed workbooks parse.
    """
    safe_name = Path(name).name
    if not safe_name:
        return None
    if not zipfile.is_zipfile(io.BytesIO(data)):
        logger.debug("Skipping %s: not an XLSX (ZIP) container", name)
        return None
    dest = target_dir / safe_name
    if not dest.exists():
        dest.write_bytes(data)
    return dest


def download_and_extract_xlsx(
    url: str,
    target_dir: Path,
    max_files: int = 50,
) -> list[Path]:
    """
    Download a ZIP or tar.gz archive and extract the .xlsx files it contains.

    Returns the list of extracted .xlsx paths (empty if the download failed).
    """
    target_dir.mkdir(parents=True, exist_ok=True)
    files: list[Path] = []

    try:
        logger.info("Downloading %s ...", url)
        resp = requests.get(url, timeout=_TIMEOUT)
        resp.raise_for_status()
        content = resp.content
    except requests.RequestException as e:
        logger.warning("Failed to download %s: %s", url, e)
        return files

    buf = io.BytesIO(content)
    try:
        if zipfile.is_zipfile(buf):
            buf.seek(0)
            with zipfile.ZipFile(buf) as zf:
                names = [n for n in zf.namelist() if _is_safe_member(n)]
                logger.info("Found %d .xlsx files in %s", len(names), url)
                for name in names[:max_files]:
                    dest = _write_member(target_dir, name, zf.read(name))
                    if dest is not None:
                        files.append(dest)
        else:
            buf.seek(0)
            with tarfile.open(fileobj=buf, mode="r:*") as tf:
                members = [m for m in tf.getmembers() if m.isfile() and _is_safe_member(m.name)]
                logger.info("Found %d .xlsx files in %s", len(members), url)
                for member in members[:max_files]:
                    handle = tf.extractfile(member)
                    if handle is None:
                        continue
                    dest = _write_member(target_dir, member.name, handle.read())
                    if dest is not None:
                        files.append(dest)
    except (zipfile.BadZipFile, tarfile.TarError) as e:
        logger.warning("Unreadable archive from %s: %s", url, e)

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
    except requests.RequestException as e:
        logger.warning("Failed to download %s: %s", url, e)
        return None

    if not zipfile.is_zipfile(io.BytesIO(resp.content)):
        logger.debug("Skipping %s: not an XLSX (ZIP) container", url)
        return None
    dest.write_bytes(resp.content)
    return dest


def _list_github_xlsx(repo: str, ref: str, path: str) -> list[tuple[str, str]]:
    """
    List .xlsx files in a GitHub directory as (raw_url, filename) pairs.

    Uses the contents API so an upstream rename shrinks the corpus by one file
    instead of breaking a hard-coded URL list.
    """
    api = f"https://api.github.com/repos/{repo}/contents/{path}?ref={ref}"
    try:
        resp = requests.get(api, timeout=_TIMEOUT, headers={"Accept": "application/vnd.github+json"})
        resp.raise_for_status()
        entries = resp.json()
    except (requests.RequestException, ValueError) as e:
        logger.warning("Could not list %s/%s: %s", repo, path, e)
        return []

    if not isinstance(entries, list):
        # A submodule or a file resolves to a dict, not a listing.
        logger.warning("%s/%s is not a directory listing", repo, path)
        return []

    out: list[tuple[str, str]] = []
    for entry in entries:
        name = entry.get("name", "")
        download_url = entry.get("download_url")
        if entry.get("type") == "file" and name.lower().endswith(".xlsx") and download_url:
            out.append((download_url, name))
    return out


def download_github_xlsx_samples(
    target_dir: Path,
    max_files: int = 50,
) -> list[Path]:
    """Download .xlsx fixtures from upstream parser test suites on GitHub."""
    target_dir.mkdir(parents=True, exist_ok=True)
    files: list[Path] = []

    for repo, ref, path in _GITHUB_FIXTURE_DIRS:
        if len(files) >= max_files:
            break
        for url, fname in _list_github_xlsx(repo, ref, path):
            if len(files) >= max_files:
                break
            # Namespace by repo: pandas and calamine both ship merge_cells.xlsx.
            prefix = repo.split("/")[-1]
            dest = download_single_xlsx(url, target_dir, f"{prefix}_{fname}")
            if dest is not None:
                files.append(dest)

    return files


def download_pypi_corpus(
    target_dir: Path,
    max_files: int = 50,
) -> list[Path]:
    """
    Download .xlsx fixtures from the sdists of Excel-handling PyPI packages.

    PyPI artifacts are immutable, making this the most durable of the sources.
    """
    target_dir.mkdir(parents=True, exist_ok=True)
    files: list[Path] = []

    for package in _PYPI_PACKAGES:
        if len(files) >= max_files:
            break
        sdist_url = _resolve_pypi_sdist(package)
        if sdist_url is None:
            continue
        files.extend(
            download_and_extract_xlsx(sdist_url, target_dir, max_files=max_files - len(files))
        )

    return files


def _resolve_pypi_sdist(package: str) -> str | None:
    """Resolve the sdist download URL for a package's current release."""
    try:
        resp = requests.get(f"https://pypi.org/pypi/{package}/json", timeout=_TIMEOUT)
        resp.raise_for_status()
        payload = resp.json()
    except (requests.RequestException, ValueError) as e:
        logger.warning("Could not resolve sdist for %s: %s", package, e)
        return None

    for url_entry in payload.get("urls", []):
        if url_entry.get("packagetype") == "sdist":
            return url_entry.get("url")

    logger.warning("No sdist published for %s", package)
    return None


def get_corpus_files(corpus_dir: Path) -> list[Path]:
    """Return all .xlsx files under a corpus directory."""
    if not corpus_dir.exists():
        return []
    return sorted(corpus_dir.glob("**/*.xlsx"))
