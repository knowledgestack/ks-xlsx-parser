"""
Corpus robustness tests for the xlsx_parser.

Tests the parser against large sets of real-world .xlsx files to catch
crashes, invariant violations, and regression. Corpus tests are skipped
by default — run with: pytest -m corpus

Corpus download tests require network — run with: pytest -m corpus -k download
"""



import json
from pathlib import Path

import pytest

from pipeline import parse_workbook
from models.common import Severity

from tests.helpers.invariant_checker import check_invariants
from tests.helpers.corpus_downloader import (
    download_euses_corpus,
    download_enron_corpus,
    download_github_xlsx_samples,
    get_corpus_files,
)


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


# ---------------------------------------------------------------------------
# Corpus download tests (require network)
# ---------------------------------------------------------------------------


@pytest.mark.corpus
class TestCorpusDownload:
    """Download corpus files. Run once, then corpus tests use the files."""

    def test_download_github_samples(self):
        target = CORPUS_DIR / "github_samples"
        files = download_github_xlsx_samples(target, max_files=20)
        # Some URLs may fail; just ensure we tried
        assert isinstance(files, list)

    def test_download_euses(self):
        target = CORPUS_DIR / "euses"
        files = download_euses_corpus(target, max_files=50)
        assert isinstance(files, list)

    def test_download_enron(self):
        target = CORPUS_DIR / "enron"
        files = download_enron_corpus(target, max_files=50)
        assert isinstance(files, list)


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
        result = parse_workbook(path=xlsx_path)
        assert len(result.workbook.sheets) >= 1, (
            f"{xlsx_path.name}: no sheets parsed"
        )

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
        """At least 95% of corpus files should parse without ERROR-level errors."""
        total = len(corpus_files)
        errors = 0
        for path in corpus_files:
            try:
                result = parse_workbook(path=path)
                has_errors = any(
                    e.severity == Severity.ERROR
                    for e in result.workbook.errors
                )
                if has_errors:
                    errors += 1
            except Exception:
                errors += 1

        rate = (total - errors) / total if total > 0 else 1.0
        assert rate >= 0.95, (
            f"Success rate {rate:.1%} ({total - errors}/{total}) "
            f"below 95% threshold"
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
