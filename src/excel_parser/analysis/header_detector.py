"""
Shared header-row detection.

A single source of truth for "which row(s) of this block are the column header",
used by the segmenter (block classification + table stitching) and both renderers
(so a whole-table chunk and a windowed part render the header identically).

The detector does NOT assume the header is the block's first row. It scans the
top of the block, skipping title/caption/key-value preamble rows, and anchors
the header at the first row that *labels the data beneath it*. A header row is
recognised by independent signals:

  - *styled*    — a majority of the row's non-empty cells are bold or filled
                  (and the row is not itself data-like);
  - *contrast*  — columns where the row holds a text label while the cells
                  below are predominantly data (numbers / dates / date-like or
                  numeric-like strings / error literals).

From the anchor, the band extends downward over contiguous label rows — rows
that are text-majority and do *not* look like the table body. "Looks like the
body" is decided by a type-signature comparison against the rows further down
(the body-signature brake), which is what keeps a bold first data row out of
the band even in tables where every cell is styled. Vertical merges anchored in
the band, single-cell wrapped-label stubs, and hidden rows are handled
explicitly; punctuation divider rows and blank rows terminate the band.

Under-detection is still preferred over over-detection for *anchoring* (a row
is only an anchor if it clearly labels data beneath it), which keeps the
segmenter from gluing unrelated regions together.
"""

from __future__ import annotations

import datetime as _dt
import re
from dataclasses import dataclass, field

from excel_parser.models.common import CellRange
from excel_parser.models.sheet import SheetDTO

# How many *candidate* (multi-column, non-title) rows from the top of a block
# to consider for the header anchor. Title/caption/preamble/blank rows do not
# consume the window; _HARD_SCAN_LIMIT bounds the absolute depth.
DEFAULT_MAX_SCAN = 4
_HARD_SCAN_LIMIT = 10
# Minimum number of label-over-data columns for the "contrast" signal to fire.
_MIN_CONTRAST_COLS = 2
# Maximum number of rows a column header band may span. Bounds the downward
# extension (find_header_span) so a runaway scan can never swallow data rows.
MAX_HEADER_ROWS = 6
# Body-signature brake: how many rows below to sample, and the per-column
# agreement fraction above which a row is considered part of the body stripe.
# The loose fraction applies when the row's styling already mismatches the
# band (weaker prior that it belongs to the header).
_SIG_SAMPLE_ROWS = 4
_SIG_MATCH_FRAC = 0.8
_SIG_MATCH_FRAC_LOOSE = 0.6
# How far below a lone-cell candidate to look for a repeating section label.
_STUB_SIBLING_SCAN = 30
# Cap signature width so very wide sheets stay cheap.
_MAX_SIG_COLS = 64
# Per-column sampling depth for the data-majority test (contrast signal):
# capped both by non-empty hits and by rows walked (sparse columns).
_COL_SAMPLE_CELLS = 15
_COL_SCAN_ROWS = 400
# How deep body_sample may walk looking for its sample rows.
_BODY_SCAN_ROWS = 200
# How deep merge rectangles are expanded; merge facts are only consulted near
# the region top (scan window + band + body sample).
_MERGE_EXPAND_ROWS = 64
# A styled-only anchor candidate must cover at least this fraction of the
# body's data columns, otherwise it is treated as a title/preamble row.
_MIN_COVERAGE_FRAC = 0.3
# A horizontal merge spanning more than this fraction of the body columns
# marks the row as a banner/title.
_BANNER_MERGE_FRAC = 0.5

_ERROR_LITERALS = frozenset(
    {
        "#N/A",
        "#REF!",
        "#DIV/0!",
        "#VALUE!",
        "#NAME?",
        "#NUM!",
        "#NULL!",
        "#GETTING_DATA",
    }
)
# Date-like strings: ISO dates/datetimes (the cell parser serialises datetimes
# to isoformat strings) and the common slashed/dashed numeric forms.
_DATE_STR_RE = re.compile(
    r"^\s*(\d{4}-\d{2}-\d{2}([T ][\d:.]+)?|\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{1,2}:\d{2}(:\d{2})?)\s*$"
)
# Numeric-looking strings ("200536.01", "1,234", "45%", "$1.50", "(2.5)").
_NUMERIC_STR_RE = re.compile(r"^\s*\(?[+-]?\$?[\d,]+(\.\d+)?\)?%?\s*$")
# Punctuation-only divider rows ("------", "=====", "____").
_DIVIDER_RE = re.compile(r"^[\s\-=_*.·~]+$")

_YEAR_MIN, _YEAR_MAX = 1900, 2100


@dataclass(frozen=True)
class HeaderSpan:
    """Inclusive Excel row range of a block's column header."""

    top: int
    bottom: int


# ──────────────────────────────────────────────────────────── cell predicates
def _is_data_value(cell) -> bool:
    """True if the cell holds data (the stuff that sits under labels)."""
    if cell is None:
        return False
    raw = cell.raw_value
    if isinstance(raw, bool):  # bool is a subclass of int; treat as data anyway
        return True
    if isinstance(raw, (int, float, _dt.date, _dt.datetime)):
        return True
    if isinstance(raw, str):
        s = raw.strip()
        if s in _ERROR_LITERALS:
            return True
        if _DATE_STR_RE.match(s) or _NUMERIC_STR_RE.match(s):
            return bool(any(ch.isdigit() for ch in s))
    return False


def _is_year_label(value) -> bool:
    """An integer (or integral float / digit string) that reads as a year."""
    if isinstance(value, bool):
        return False
    if isinstance(value, str) and value.strip().isdigit():
        value = int(value.strip())
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return isinstance(value, int) and _YEAR_MIN <= value <= _YEAR_MAX


def _is_styled(cell) -> bool:
    if cell is None or cell.style is None:
        return False
    font = cell.style.font
    if font and font.bold:
        return True
    fill = cell.style.fill
    return bool(fill and fill.fg_color)


def _is_formula_blank(cell) -> bool:
    """A formula cell with no cached value — neither label nor data evidence."""
    return cell is not None and cell.formula is not None and cell.raw_value is None


# ──────────────────────────────────────────────────────────── region context
@dataclass
class _Region:
    """Per-call context: bounds, merge lookup, and memoised row facts."""

    sheet: SheetDTO
    top: int
    bot: int
    c0: int
    c1: int
    # (row, col) -> (master_row, master_col) for every merged cell in range
    merge_master: dict[tuple[int, int], tuple[int, int]] = field(default_factory=dict)
    # row -> list of (master_col, col_span) for horizontal merge masters there
    hmerge_spans: dict[int, list[tuple[int, int]]] = field(default_factory=dict)
    _sig_cache: dict[int, tuple] = field(default_factory=dict)
    _cells_cache: dict[int, list] = field(default_factory=dict)
    # (col, below_row) -> (n_total, n_data, n_date) sampled column profile
    _col_cache: dict[tuple[int, int], tuple[int, int, int]] = field(default_factory=dict)
    _contrast_cache: dict[int, int] = field(default_factory=dict)
    _body_cols_cache: dict[int, list[int]] = field(default_factory=dict)

    @classmethod
    def build(cls, sheet: SheetDTO, cell_range: CellRange) -> _Region:
        reg = cls(
            sheet=sheet,
            top=cell_range.top_left.row,
            bot=cell_range.bottom_right.row,
            c0=cell_range.top_left.col,
            c1=cell_range.bottom_right.col,
        )
        # Merge facts are only consulted near the top of the region (scan
        # window + band extension + body sample); clipping the expansion there
        # keeps a full-height merged label column from costing O(rows) per call.
        merge_depth = min(reg.bot, reg.top + _MERGE_EXPAND_ROWS - 1)
        for mr in sheet.merged_regions:
            tl, br = mr.range.top_left, mr.range.bottom_right
            if br.row < reg.top or tl.row > merge_depth or br.col < reg.c0 or tl.col > reg.c1:
                continue
            master = (mr.master.row, mr.master.col)
            for r in range(max(tl.row, reg.top), min(br.row, merge_depth) + 1):
                for c in range(max(tl.col, reg.c0), min(br.col, reg.c1) + 1):
                    reg.merge_master[(r, c)] = master
            in_span = min(br.col, reg.c1) - max(tl.col, reg.c0) + 1
            if in_span > 1 and tl.row >= reg.top:
                reg.hmerge_spans.setdefault(tl.row, []).append((max(tl.col, reg.c0), in_span))
        return reg

    # ── row facts -----------------------------------------------------------
    def is_hidden(self, row: int) -> bool:
        return row in self.sheet.hidden_rows

    def cells(self, row: int) -> list:
        """Non-empty cells of the row (value or formula), within the region."""
        cached = self._cells_cache.get(row)
        if cached is None:
            cached = [
                cell
                for col in range(self.c0, self.c1 + 1)
                if (cell := self.sheet.get_cell(row, col)) and not cell.is_empty
            ]
            self._cells_cache[row] = cached
        return cached

    def occupied_cols(self, row: int) -> set[int]:
        """Columns covered by a value — directly or via a merged master."""
        cols: set[int] = set()
        for col in range(self.c0, self.c1 + 1):
            cell = self.sheet.get_cell(row, col)
            if cell and not cell.is_empty:
                cols.add(col)
                continue
            master = self.merge_master.get((row, col))
            if master and master != (row, col):
                mcell = self.sheet.get_cell(*master)
                if mcell and not mcell.is_empty:
                    cols.add(col)
        return cols

    def vmerge_cover(self, row: int, band_top: int, band_bottom: int) -> int:
        """Columns of ``row`` covered by merges whose master row is *inside the
        accepted band* — vertical merges hanging down from the header. Masters
        in rows below the band (e.g. a merged body label column) must not pull
        their slave rows into the header."""
        n = 0
        for col in range(self.c0, self.c1 + 1):
            master = self.merge_master.get((row, col))
            if master and band_top <= master[0] <= band_bottom:
                n += 1
        return n

    def is_divider(self, row: int) -> bool:
        cells = self.cells(row)
        if not cells:
            return False
        return all(isinstance(c.raw_value, str) and _DIVIDER_RE.match(c.raw_value) for c in cells)

    def is_banner_row(self, row: int) -> bool:
        """A row that is essentially one wide horizontal merge — a section
        divider (merged "NORTH REGION" strip) or a sheet banner/title.

        A row with *several* merge masters is a group-header row tiling column
        bands ("Name | Payoff"), not a banner, so this requires exactly one
        merge and at most one other value cell.
        """
        if len(self.cells(row)) > 2:
            return False
        spans = self.hmerge_spans.get(row, ())
        if len(spans) != 1:
            return False
        width = max(self.c1 - self.c0 + 1, 2)
        return spans[0][1] / width > _BANNER_MERGE_FRAC

    def is_data_row(self, row: int, strict: bool = False) -> bool:
        """Majority of the row's value cells are data (numbers/dates/etc.).

        ``strict`` requires a strict majority — used when rejecting anchor
        candidates, so a 50/50 "label | value" row is not written off.
        """
        cells = [c for c in self.cells(row) if not _is_formula_blank(c)]
        if not cells:
            return False
        n_data = sum(1 for c in cells if _is_data_value(c))
        return n_data * 2 > len(cells) if strict else n_data * 2 >= len(cells)

    def is_styled_row(self, row: int) -> bool:
        cells = self.cells(row)
        if not cells:
            return False
        return sum(1 for c in cells if _is_styled(c)) * 2 >= len(cells)

    def is_axis_label_row(self, row: int) -> bool:
        """A row whose data cells are all axis labels: years (2019|2020|...)
        or dates/periods (Jan-01|Feb-01|... — month columns are ubiquitous).
        Such rows are column headers even though every cell reads as data."""
        data_vals = [c.raw_value for c in self.cells(row) if not _is_formula_blank(c) and _is_data_value(c)]
        if len(data_vals) < 2:
            return False
        if all(_is_year_label(v) for v in data_vals):
            return True
        return all(
            isinstance(v, (_dt.date, _dt.datetime)) or (isinstance(v, str) and _DATE_STR_RE.match(v.strip()))
            for v in data_vals
        )

    # ── type signatures ------------------------------------------------------
    def type_sig(self, row: int) -> tuple:
        """Per-column type classes for the row. Merge-aware: an empty slave
        cell inherits its master's class, so a merged body label column
        ("North" spanning five rows) gives every spanned row the same
        signature and the body brake can see the stripe."""
        sig = self._sig_cache.get(row)
        if sig is None:
            parts = []
            for col in range(self.c0, min(self.c1, self.c0 + _MAX_SIG_COLS - 1) + 1):
                cell = self.sheet.get_cell(row, col)
                if cell is None or cell.is_empty:
                    master = self.merge_master.get((row, col))
                    if master and master != (row, col):
                        mcell = self.sheet.get_cell(*master)
                        if mcell and not mcell.is_empty:
                            cell = mcell
                if cell is None or cell.is_empty or _is_formula_blank(cell):
                    parts.append("e")
                elif _is_data_value(cell):
                    # Year and date labels get their own classes so an axis
                    # header row never signature-matches a numeric body stripe.
                    raw = cell.raw_value
                    if _is_year_label(raw):
                        parts.append("y")
                    elif isinstance(raw, (_dt.date, _dt.datetime)) or (
                        isinstance(raw, str) and _DATE_STR_RE.match(raw.strip())
                    ):
                        parts.append("d")
                    else:
                        parts.append("n")
                else:
                    parts.append("s")
            sig = tuple(parts)
            self._sig_cache[row] = sig
        return sig

    def _sig_match(self, a: tuple, b: tuple, frac: float = _SIG_MATCH_FRAC) -> bool:
        denom = agree = 0
        for x, y in zip(a, b, strict=False):
            if x == "e" and y == "e":
                continue
            denom += 1
            if x == y:
                agree += 1
        return denom > 0 and agree / denom >= frac

    def body_sample(self, below_row: int, max_rows: int = _SIG_SAMPLE_ROWS) -> list[int]:
        """Up to ``max_rows`` visible, non-blank, non-divider rows below
        ``below_row``, looking at most :data:`_BODY_SCAN_ROWS` rows deep (a
        stitched range with a huge interior gap must not cost O(height))."""
        out: list[int] = []
        r = below_row + 1
        limit = min(self.bot, below_row + _BODY_SCAN_ROWS)
        while r <= limit and len(out) < max_rows:
            if not self.is_hidden(r) and self.cells(r) and not self.is_divider(r):
                out.append(r)
            r += 1
        return out

    def looks_like_body(self, row: int, frac: float = _SIG_MATCH_FRAC) -> bool:
        """The row's type signature matches the stripe of rows beneath it.

        ``frac`` is the per-column agreement threshold; pass a looser value
        when other evidence (e.g. a style mismatch with the band) already
        makes the row body-suspect.
        """
        sample = self.body_sample(row)
        if not sample:
            return False
        sig = self.type_sig(row)
        matches = sum(1 for r in sample if self._sig_match(sig, self.type_sig(r), frac))
        need = 2 if len(sample) >= 2 else 1
        return matches >= need

    def body_data_cols(self, below_row: int) -> list[int]:
        """Columns that are non-empty in at least half of the body sample."""
        cached = self._body_cols_cache.get(below_row)
        if cached is None:
            sample = self.body_sample(below_row, max_rows=6)
            counts: dict[int, int] = {}
            for r in sample:
                for c in self.cells(r):
                    counts[c.coord.col] = counts.get(c.coord.col, 0) + 1
            need = max(1, (len(sample) + 1) // 2)
            cached = sorted(c for c, n in counts.items() if n >= need)
            self._body_cols_cache[below_row] = cached
        return cached

    def col_profile(self, col: int, below_row: int) -> tuple[int, int, int]:
        """``(n_total, n_data, n_date)`` for the column below ``below_row``.

        Sampling is capped at the first :data:`_COL_SAMPLE_CELLS` non-empty
        cells — statistically equivalent for the majority tests and it keeps
        very tall regions cheap.
        """
        key = (col, below_row)
        cached = self._col_cache.get(key)
        if cached is not None:
            return cached
        n_total = n_data = n_date = 0
        # Bounded by rows as well as hits: a sparse column must not cost
        # O(region height) just because it never reaches the hit cap.
        for r in range(below_row + 1, min(self.bot, below_row + _COL_SCAN_ROWS) + 1):
            cell = self.sheet.get_cell(r, col)
            if cell is None or cell.is_empty or _is_formula_blank(cell):
                continue
            n_total += 1
            if _is_data_value(cell):
                n_data += 1
                raw = cell.raw_value
                if isinstance(raw, (_dt.date, _dt.datetime)) or (
                    isinstance(raw, str) and _DATE_STR_RE.match(raw.strip())
                ):
                    n_date += 1
            if n_total >= _COL_SAMPLE_CELLS:
                break
        prof = (n_total, n_data, n_date)
        self._col_cache[key] = prof
        return prof

    def col_data_majority(self, col: int, below_row: int) -> bool | None:
        """Whether the column below ``below_row`` is majority data values."""
        n_total, n_data, _ = self.col_profile(col, below_row)
        if not n_total:
            return None
        return n_data * 2 >= n_total

    def has_sibling_stub(self, row: int, col: int, styled: bool) -> bool:
        """Another lone-cell row in the same column further down — the mark of
        a repeating *section label* pattern ("Director" … "Manager"), as
        opposed to a one-off wrapped header stub."""
        for r in range(row + 1, min(row + _STUB_SIBLING_SCAN, self.bot) + 1):
            cells = self.cells(r)
            if len(cells) != 1:
                continue
            c = cells[0]
            if (
                c.coord.col == col
                and isinstance(c.raw_value, str)
                and not _is_data_value(c)
                and _is_styled(c) == styled
            ):
                return True
        return False

    # ── header hints from sheet metadata -------------------------------------
    def freeze_row(self) -> int | None:
        """First unfrozen row from the sheet's freeze pane, if any."""
        fp = self.sheet.properties.freeze_pane
        if not fp:
            return None
        m = re.match(r"^\$?[A-Za-z]*\$?(\d+)$", str(fp).strip())
        if not m:
            return None
        row = int(m.group(1))
        return row if row > 1 else None


# ──────────────────────────────────────────────────────────── row predicates
def _contrast_cols(reg: _Region, row: int) -> int:
    """Columns where this row labels the data column beneath it.

    Two label shapes qualify: a *text* head over a majority-data column, and a
    *date/year* head over a majority non-date data column (month/period axis
    labels over numbers — a date over more dates is just a data column).
    """
    cached = reg._contrast_cache.get(row)
    if cached is not None:
        return cached

    def _axis_head(value) -> bool:
        return _is_year_label(value) or (
            isinstance(value, (_dt.date, _dt.datetime))
            or (isinstance(value, str) and _DATE_STR_RE.match(value.strip()))
        )

    # Axis heads only count when the row carries a *run* of them: one lone
    # date is a value (e.g. the right half of a "Date: | 2024-01-01" pair),
    # not an axis label.
    axis_heads = sum(1 for c in reg.cells(row) if not _is_formula_blank(c) and _axis_head(c.raw_value))
    count = 0
    for col in range(reg.c0, reg.c1 + 1):
        head = reg.sheet.get_cell(row, col)
        if head is None or head.is_empty:
            continue
        raw = head.raw_value
        is_axis_head = axis_heads >= 2 and _axis_head(raw)
        is_text_head = isinstance(raw, str) and not _is_data_value(head)
        if not (is_text_head or is_axis_head):
            continue
        n_total, n_data, n_date = reg.col_profile(col, row)
        if not n_total:
            continue
        qualifies = (n_data - n_date) * 2 >= n_total if is_axis_head else n_data * 2 >= n_total
        if qualifies:
            count += 1
            if count >= _MIN_CONTRAST_COLS:
                break
    reg._contrast_cache[row] = count
    return count


def _is_header_anchor(reg: _Region, row: int) -> bool:
    """Whether ``row`` labels the data below it (anchor of the header band)."""
    cells = reg.cells(row)
    if len(cells) < 2:
        return False

    # Data check FIRST: a majority-data row is not a header even when styled
    # (bold "Totals" rows, numeric spill rows) — with one exception for axis
    # label rows (years 2019|2020|... or month/date columns), which are
    # genuine column headers despite every cell reading as data.
    if reg.is_data_row(row, strict=True) and not reg.is_axis_label_row(row):
        return False

    contrast = _contrast_cols(reg, row)
    if contrast >= _MIN_CONTRAST_COLS:
        return True

    styled = sum(1 for c in cells if _is_styled(c))
    if styled * 2 >= len(cells):
        return True

    # Adaptive contrast for narrow tables: when the body simply doesn't have
    # two data columns, one decisive label-over-data column is enough,
    # provided the row is all-text and distinct from the body stripe.
    if contrast >= 1:
        data_cols = [c for c in reg.body_data_cols(row) if reg.col_data_majority(c, row)]
        if len(data_cols) <= 1 and not reg.looks_like_body(row):
            return True
    return False


def _is_title_or_preamble(reg: _Region, row: int) -> bool:
    """Rows to skip during the scan without consuming the candidate window:
    banners (one wide horizontal merge), key-value preamble rows, and sparse
    styled rows that cover almost none of the body's data columns."""
    cells = reg.cells(row)
    if not cells:
        return False

    # Banner: a single wide horizontal merge (sheet title / section strip).
    if reg.is_banner_row(row):
        return True

    body_cols = reg.body_data_cols(row)
    width = max(len(body_cols), 2)

    # A row with a genuine label-over-data column is never a title.
    if _contrast_cols(reg, row) >= 1:
        return False

    value_cols = {c.coord.col for c in cells}
    coverage = len(value_cols) / width

    # Key-value preamble: "Label: value" pairs (text immediately followed by a
    # data cell to its right) or labels ending with ':', with low coverage.
    if coverage < 0.5:
        kv_pairs = 0
        text_cells = 0
        for c in cells:
            if isinstance(c.raw_value, str) and not _is_data_value(c):
                text_cells += 1
                if c.raw_value.rstrip().endswith(":"):
                    kv_pairs += 1
                    continue
                right = reg.sheet.get_cell(c.coord.row, c.coord.col + 1)
                if right and not right.is_empty and _is_data_value(right):
                    kv_pairs += 1
        if text_cells and kv_pairs >= max(1, text_cells - 1):
            return True

    # Sparse styled row covering a sliver of the data columns → preamble/title.
    return coverage < _MIN_COVERAGE_FRAC


# ──────────────────────────────────────────────────────────── band extension
def _extend_down(reg: _Region, anchor: int) -> int:
    """Walk downward from the anchor over contiguous header-band rows."""
    bottom = anchor
    rows_in_band = 1
    rr = anchor + 1
    # rows_in_band caps visible band rows; the absolute bound keeps interleaved
    # hidden rows from stretching the returned span past MAX_HEADER_ROWS.
    last = min(reg.bot - 1, anchor + MAX_HEADER_ROWS - 1)
    while rr <= last and rows_in_band < MAX_HEADER_ROWS:
        if reg.is_hidden(rr):
            rr += 1  # hidden rows are transparent: never joined, never break
            continue
        if reg.is_divider(rr) or reg.is_banner_row(rr):
            break  # punctuation rule-offs and merged section strips end the band
        occupied = reg.occupied_cols(rr)
        if not occupied:
            break  # blank row terminates the band

        # Rows covered by vertical merges hanging from the *accepted band* are
        # part of the stacked header by construction. (Masters below the band
        # — merged body label columns — must not qualify.)
        if reg.vmerge_cover(rr, anchor, bottom) >= 2:
            bottom = rr
            rows_in_band += 1
            rr += 1
            continue

        cells = reg.cells(rr)
        if len(occupied) >= 2 and len(cells) >= 1:
            if reg.is_data_row(rr) and not reg.is_axis_label_row(rr):
                break  # data has begun
            # Body-signature brake: the first row of a homogeneous stripe is
            # data even when styled. When the row's styling already mismatches
            # the anchor's, demand less signature agreement — an unstyled
            # all-text row under a styled header is body unless it clearly
            # differs from the stripe below (stacked sub-labels do).
            style_mismatch = reg.is_styled_row(rr) != reg.is_styled_row(anchor)
            frac = _SIG_MATCH_FRAC_LOOSE if style_mismatch else _SIG_MATCH_FRAC
            if reg.looks_like_body(rr, frac):
                break
            bottom = rr
            rows_in_band += 1
            rr += 1
            continue

        if len(cells) == 1:
            # Wrapped-label stub ("Country &" / "Commodity Group"): a single
            # short text cell, aligned with the anchor's columns and styled
            # like the band, directly joined to it. A stub with lone-cell
            # *siblings* further down is a repeating section label
            # ("Director" … "Manager"), not a wrapped header line.
            cell = cells[0]
            if (
                isinstance(cell.raw_value, str)
                and not _is_data_value(cell)
                and len(cell.raw_value.strip()) <= 40
                and cell.coord.col in reg.occupied_cols(anchor)
                and _is_styled(cell) == reg.is_styled_row(anchor)
                and not reg.has_sibling_stub(rr, cell.coord.col, _is_styled(cell))
            ):
                bottom = rr
                rows_in_band += 1
                rr += 1
                continue
        break
    # Freeze panes are the author's own declaration of which rows stay pinned
    # above the data — when the band runs past that boundary, trust the author
    # over our heuristics and clip back to it.
    f = reg.freeze_row()
    if f is not None and anchor < f <= bottom:
        return f - 1
    return bottom


# Deliberately NOT done: extending the band upward over single-cell caption
# rows that touch the anchor. DECO ground truth counts such captions (pivot
# corners like "Count of Buy/Sell") as header rows, and including them would
# raise the multi-row benchmark score — but the renderers use the band's rows
# for column naming, and a caption there replaces real column names with the
# title text. Product semantics win over the GT convention.


# ──────────────────────────────────────────────────────────── public API
def find_header_span(
    sheet: SheetDTO,
    cell_range: CellRange,
    max_scan: int = DEFAULT_MAX_SCAN,
) -> HeaderSpan | None:
    """
    Find the header row(s) of a block, or ``None`` if it has no recognisable
    header (e.g. a free-text block, a single-row region, or a key/value list).

    Scans the top of the block for the header *anchor*, skipping title /
    caption / key-value preamble / divider / hidden rows (these do not consume
    the ``max_scan`` candidate budget, bounded overall by an absolute depth
    limit). The anchor is the first multi-column row that labels the data
    below it; from there the band extends downward over contiguous label rows
    (see :func:`_extend_down`). A candidate row that instead matches the
    table's body stripe ends the scan — such a block has no header.
    """
    reg = _Region.build(sheet, cell_range)
    if reg.bot <= reg.top:
        return None  # need at least one data row beneath a header

    hard_limit = min(reg.top + _HARD_SCAN_LIMIT - 1, reg.bot - 1)
    # A freeze pane below the default window extends the scan: the author put
    # the header somewhere above the freeze boundary, so look that far down
    # (within reason — twice the normal depth).
    f = reg.freeze_row()
    if f is not None and reg.top < f <= reg.top + 2 * _HARD_SCAN_LIMIT:
        hard_limit = min(max(hard_limit, f - 1), reg.bot - 1)

    candidates_seen = 0
    found: HeaderSpan | None = None
    r = reg.top
    while r <= hard_limit and candidates_seen < max_scan:
        if reg.is_hidden(r) or reg.is_divider(r):
            r += 1
            continue
        cells = reg.cells(r)
        if len(cells) < 2:
            r += 1  # title/caption/section stub — keep looking below it
            continue
        if _is_title_or_preamble(reg, r):
            r += 1
            continue

        candidates_seen += 1
        if _is_header_anchor(reg, r):
            if found is not None:
                # A gapped anchor was already found. Only an anchor with real
                # label-over-data contrast may replace it — a styled body row
                # below the gap must not steal the header.
                if _contrast_cols(reg, r) >= _MIN_CONTRAST_COLS:
                    return HeaderSpan(top=r, bottom=_extend_down(reg, r))
                return found
            span = HeaderSpan(top=r, bottom=_extend_down(reg, r))
            if _touches_data(reg, span):
                return span
            # The band is followed by a blank gap. That is usually still the
            # header (blank spacer before data), but when a key-value/title
            # block sits over the table the real header is further down.
            # Remember the span and keep scanning; only a contrast-strong
            # anchor may replace it.
            found = span
            r = span.bottom + 1
            continue
        # Not a header: if this row is the top of the body stripe, the block
        # has no header above its data. Otherwise keep scanning below it.
        if reg.looks_like_body(r):
            return found
        r += 1
    return found


def _touches_data(reg: _Region, span: HeaderSpan) -> bool:
    """The first visible row under the band is non-blank (header sits on its
    table). Hidden and divider rows are transparent."""
    r = span.bottom + 1
    while r <= reg.bot:
        if not reg.is_hidden(r) and not reg.is_divider(r):
            return bool(reg.cells(r))
        r += 1
    return False


def has_header(sheet: SheetDTO, cell_range: CellRange) -> bool:
    """Convenience: whether the block has a recognisable column header."""
    return find_header_span(sheet, cell_range) is not None
