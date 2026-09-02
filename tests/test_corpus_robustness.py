"""
Corpus robustness tests for the excel_parser.

Tests the parser against large sets of real-world .xlsx files to catch
crashes, invariant violations, and regression. Corpus tests are skipped
by default — run with: pytest -m corpus

Corpus download tests require network — run with: pytest -m corpus -k download
"""



import json
import zipfile
from pathlib import Path

import openpyxl
import pytest
import requests

from excel_parser.models.common import Severity
from excel_parser.pipeline import parse_workbook
from tests.helpers.corpus_downloader import (
    download_github_xlsx_samples,
    download_pypi_corpus,
    get_corpus_files,
)
from tests.helpers.invariant_checker import check_invariants

CORPUS_DIR = Path(__file__).parent / "fixtures" / "corpus"


def _collect_corpus_files() -> list[Path]:
    """Collect all .xlsx files from all corpus subdirectories."""
    files = []
    if CORPUS_DIR.exists():
        for subdir in sorted(CORPUS_DIR.iterdir()):
            if subdir.is_dir():
                files.extend(get_corpus_files(subdir))
    return files


corpus_files = _collect_corpus_files()

_loadable_cache: dict[Path, bool] = {}


def _is_loadable(path: Path) -> bool:
    """
    True when openpyxl can open the file at all.

    Used to separate "the parser mishandled a workbook" from "this file is not
    a workbook any reader can open". Cached because the corpus tests re-parse
    each file per assertion.
    """
    cached = _loadable_cache.get(path)
    if cached is not None:
        return cached

    loadable = True
    try:
        wb = openpyxl.load_workbook(path, read_only=True)
        wb.close()
    except Exception:
        loadable = False
    _loadable_cache[path] = loadable
    return loadable


# ---------------------------------------------------------------------------
# Corpus download tests (require network)
# ---------------------------------------------------------------------------


@pytest.mark.corpus
class TestCorpusDownload:
    """
    Download corpus files. Run once, then corpus tests use the files.

    These assert that files actually arrived. An earlier revision asserted
    only ``isinstance(files, list)``, which stayed green for months while
    every upstream URL 404'd and the corpus directory sat empty — silently
    skipping the eight robustness tests that depend on it. A genuinely
    offline machine is a skip; a reachable source returning nothing usable
    is a failure.
    """

    @staticmethod
    def _require_network() -> None:
        try:
            requests.head("https://pypi.org/simple/", timeout=15)
        except requests.RequestException as e:
            pytest.skip(f"network unavailable: {e}")

    def test_download_pypi_corpus(self):
        self._require_network()
        target = CORPUS_DIR / "pypi"
        files = download_pypi_corpus(target, max_files=40)
        assert files, "no .xlsx fixtures extracted from any PyPI sdist"
        assert all(f.exists() and f.stat().st_size > 0 for f in files)

    def test_download_github_samples(self):
        self._require_network()
        target = CORPUS_DIR / "github_samples"
        files = download_github_xlsx_samples(target, max_files=40)
        if not files:
            pytest.skip("GitHub fixture listing unavailable (rate limit or upstream move)")
        assert all(f.exists() and f.stat().st_size > 0 for f in files)

    def test_downloaded_files_are_parseable_xlsx(self):
        """A download that yields unopenable bytes is worse than no download."""
        self._require_network()
        files = download_pypi_corpus(CORPUS_DIR / "pypi", max_files=5)
        assert files, "no .xlsx fixtures extracted from any PyPI sdist"
        for path in files[:5]:
            assert zipfile.is_zipfile(path), f"{path.name} is not a valid xlsx container"


# ---------------------------------------------------------------------------
# Corpus robustness tests (require downloaded files)
# ---------------------------------------------------------------------------


@pytest.mark.corpus
@pytest.mark.skipif(not corpus_files, reason="No corpus files downloaded")
@pytest.mark.parametrize(
    "xlsx_path",
    corpus_files,
    ids=[f.stem for f in corpus_files],
)
class TestCorpusParseRobustness:
    """Basic robustness: parse every corpus file without crashing."""

    def test_no_unhandled_exception(self, xlsx_path):
        """Parser must complete without unhandled exception."""
        try:
            result = parse_workbook(path=xlsx_path)
            assert result.workbook is not None
        except Exception as e:
            pytest.fail(f"Parser crashed on {xlsx_path.name}: {e}")

    def test_has_sheets(self, xlsx_path):
        """
        A workbook that can be opened at all must yield at least one sheet.

        Upstream parser test-suites contribute files that are valid ZIPs but
        not valid OOXML packages (a missing ``[Content_Types].xml``, a missing
        ``sharedStrings.xml``, a style attribute openpyxl rejects). No reader
        can open those, so demanding a sheet from them asserts the impossible;
        the contract that actually matters for them — degrade with a recorded
        error instead of crashing — is covered by
        ``test_unloadable_files_report_an_error``.
        """
        if not _is_loadable(xlsx_path):
            pytest.skip(f"{xlsx_path.name} is not a loadable OOXML package")
        result = parse_workbook(path=xlsx_path)
        assert len(result.workbook.sheets) >= 1, (
            f"{xlsx_path.name}: no sheets parsed"
        )

    def test_unloadable_files_report_an_error(self, xlsx_path):
        """An unopenable workbook must say so, not come back silently empty."""
        if _is_loadable(xlsx_path):
            pytest.skip(f"{xlsx_path.name} loads fine")
        result = parse_workbook(path=xlsx_path)
        assert result.workbook.errors, (
            f"{xlsx_path.name}: failed to load but recorded no error"
        )
        assert any(e.severity == Severity.ERROR for e in result.workbook.errors)

    def test_workbook_hash_present(self, xlsx_path):
        result = parse_workbook(path=xlsx_path)
        assert result.workbook.workbook_hash, (
            f"{xlsx_path.name}: empty workbook_hash"
        )

    def test_structural_invariants_hold(self, xlsx_path):
        result = parse_workbook(path=xlsx_path)
        violations = check_invariants(result.workbook)
        assert len(violations) == 0, (
            f"{xlsx_path.name}: {len(violations)} violations:\n"
            + "\n".join(violations[:10])
        )

    def test_json_serializable(self, xlsx_path):
        result = parse_workbook(path=xlsx_path)
        data = result.to_json()
        json.dumps(data)  # must not raise

    def test_deterministic_hash(self, xlsx_path):
        r1 = parse_workbook(path=xlsx_path)
        r2 = parse_workbook(path=xlsx_path)
        assert r1.workbook.workbook_hash == r2.workbook.workbook_hash


# ---------------------------------------------------------------------------
# Aggregate statistics
# ---------------------------------------------------------------------------


@pytest.mark.corpus
@pytest.mark.skipif(not corpus_files, reason="No corpus files downloaded")
class TestCorpusAggregateStats:
    """Aggregate statistics across the whole corpus."""

    def test_success_rate(self):
        """
        At least 95% of *loadable* corpus files parse without ERROR-level errors.

        The denominator is deliberately the loadable subset. Sources include
        upstream parser test-suites, which seed the corpus with files that are
        intentionally broken; counting those as parser failures would make the
        threshold a measure of how many negative fixtures upstream happens to
        ship rather than of this parser's health.
        """
        loadable = [p for p in corpus_files if _is_loadable(p)]
        total = len(loadable)
        if total == 0:
            pytest.skip("no loadable corpus files")

        failures = []
        for path in loadable:
            try:
                result = parse_workbook(path=path)
                if any(e.severity == Severity.ERROR for e in result.workbook.errors):
                    failures.append(path.name)
            except Exception as e:
                failures.append(f"{path.name} ({type(e).__name__}: {e})")

        rate = (total - len(failures)) / total
        assert rate >= 0.95, (
            f"Success rate {rate:.1%} ({total - len(failures)}/{total}) "
            f"below 95% threshold; failures: {failures[:10]}"
        )

    def test_aggregate_stats(self):
        """Log aggregate statistics (informational, always passes)."""
        total_sheets = 0
        total_cells = 0
        total_formulas = 0
        for path in corpus_files:
            try:
                result = parse_workbook(path=path)
                total_sheets += result.workbook.total_sheets
                total_cells += result.workbook.total_cells
                total_formulas += result.workbook.total_formulas
            except Exception:
                pass

        # Just log — this test is informational
        print(
            f"\nCorpus stats: {len(corpus_files)} files, "
            f"{total_sheets} sheets, {total_cells} cells, "
            f"{total_formulas} formulas"
        )
