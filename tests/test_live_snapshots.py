"""
Live snapshot tests — hit real EDGAR API, ~12 min total for 8 tickers.

Run:   pytest -m slow                                            # all live tests
       pytest -m slow -v
       pytest -m "slow and b1"                                   # B1 overflow tests only
       pytest -m "slow and cf_overflow"                          # CF YTD overflow tests only
       pytest -m slow tests/test_live_snapshots.py::test_snapshot_is[AAPL]

Tickers: MSFT AMZN META GOOGL NVDA JPM GS JNJ
CF overflow tickers: COHR LITE AAPL NVDA GOOGL
Marks:
  @pytest.mark.slow       — excluded from default CI run (requires network)
  @pytest.mark.b1         — B1 overflow-specific tests (subset of slow)
  @pytest.mark.cf_overflow — CF YTD overflow correctness tests (subset of slow)

Pass definitions:
- Snapshot tests: IS/BS/CF key rows have ≥1 non-None value in recent 4Q
- B1 tests:
  - Each section in Data_Financials(Q) has ≥ template row count
  - Every overflow row has at least one non-None value across its quarters
  - If Data_Financials_NG(Q) exists, every non-header row has ≥1 non-None value
- CF overflow tests:
  - CF overflow rows present for ≥2 distinct quarters (verifies Q2/Q3 are filled in)
  - No CF overflow row is all-None after the fix (filtering is still applied)
"""
import pytest
from config import load_config
from fetcher_gaap import fetch_gaap_statements
from override_engine import check_key_rows

FINANCIAL_TICKERS = {"GS", "JPM"}
TICKERS = ["MSFT", "AMZN", "META", "GOOGL", "NVDA", "JPM", "GS", "JNJ"]
CF_OVERFLOW_TICKERS = ["COHR", "LITE", "AAPL", "NVDA", "GOOGL"]

# Template row counts per section (must stay in sync with IS/BS/CF_TEMPLATE in fetcher_gaap.py)
_IS_TEMPLATE_ROWS = 22
_BS_TEMPLATE_ROWS = 41
_CF_TEMPLATE_ROWS = 26


# ── Fixtures ──────────────────────────────────────────────────────────────

@pytest.fixture(scope="module")
def _cfg():
    return load_config()


@pytest.fixture(scope="module")
def identity(_cfg):
    ident = _cfg.get("identity", "").strip()
    if not ident:
        pytest.skip("SEC EDGAR identity not set in config.json — cannot run live tests")
    return ident


@pytest.fixture(scope="module")
def ai_config(_cfg):
    return _cfg.get("ai", {})


@pytest.fixture(scope="module")
def all_tables(identity, ai_config):
    """Fetch all 8 tickers once at module scope. Per-ticker errors are stored, not raised."""
    result: dict = {}
    for ticker in TICKERS:
        try:
            result[ticker] = fetch_gaap_statements(
                ticker, identity, max_filings=8, ai_config=ai_config
            )
        except Exception as exc:
            result[ticker] = exc
    return result


# ── Helpers ───────────────────────────────────────────────────────────────

def _quarterly_financials(tables):
    for t in tables:
        if t.sheet_name == "Data_Financials(Q)":
            return t
    return None


def _get_tbl(all_tables, ticker):
    result = all_tables[ticker]
    if isinstance(result, Exception):
        pytest.fail(f"{ticker}: fetch failed — {result}")
    tbl = _quarterly_financials(result)
    assert tbl is not None, f"{ticker}: Data_Financials(Q) not found in returned tables"
    return tbl


def _assert_no_missing(tbl, stmt, ticker, allowed_missing=()):
    """Fail if any key row for `stmt` is all-None in recent 4Q (except allowed_missing)."""
    missing = check_key_rows(tbl.concepts, tbl.values, stmt)
    unexpected = [r for r in missing if r not in allowed_missing]
    assert not unexpected, (
        f"[{ticker}/{stmt}] key rows all-None in recent 4Q: {unexpected}\n"
        f"(allowed_missing={list(allowed_missing)})"
    )


def _section_row_counts(tbl):
    """Return (is_rows, bs_rows, cf_rows) between section headers in a merged table.

    Structure produced by _merge_financials:
      "Income Statement" / [IS rows] / "" / "Balance Sheet" / [BS rows] / "" / "Cash Flow" / [CF rows]
    """
    c = tbl.concepts
    try:
        is_hdr = c.index("Income Statement")
        bs_hdr = c.index("Balance Sheet")
        cf_hdr = c.index("Cash Flow")
    except ValueError:
        return None, None, None
    is_rows = (bs_hdr - 1) - (is_hdr + 1)   # blank row before BS header
    bs_rows = (cf_hdr - 1) - (bs_hdr + 1)   # blank row before CF header
    cf_rows = len(c) - (cf_hdr + 1)
    return is_rows, bs_rows, cf_rows


# ── Live Tests ─────────────────────────────────────────────────────────────

@pytest.mark.slow
@pytest.mark.parametrize("ticker", TICKERS)
def test_snapshot_is(all_tables, ticker):
    """IS key rows (Revenue / Operating Income / Net Income / Diluted EPS) have data."""
    tbl = _get_tbl(all_tables, ticker)
    allowed = ("Operating Income",) if ticker in FINANCIAL_TICKERS else ()
    _assert_no_missing(tbl, "IS", ticker, allowed)


@pytest.mark.slow
@pytest.mark.parametrize("ticker", TICKERS)
def test_snapshot_bs(all_tables, ticker):
    """BS key rows (Total Assets / Total Liabilities / Total Equity — Parent) have data."""
    tbl = _get_tbl(all_tables, ticker)
    _assert_no_missing(tbl, "BS", ticker)


@pytest.mark.slow
@pytest.mark.parametrize("ticker", TICKERS)
def test_snapshot_cf(all_tables, ticker):
    """CF key rows (Operating Cash Flow / Capex) have data.

    Banks (GS/JPM) may not report Capex under PaymentsToAcquirePropertyPlantAndEquipment;
    they use different XBRL concepts (e.g. PurchasesOfPremisesAndEquipment) that our
    template doesn't catch yet — treat as structural_absence for financial tickers.
    """
    tbl = _get_tbl(all_tables, ticker)
    allowed = ("Capex",) if ticker in FINANCIAL_TICKERS else ()
    _assert_no_missing(tbl, "CF", ticker, allowed)


# ── B1 Overflow Row Tests ─────────────────────────────────────────────────
#
# Run with:  pytest -m "slow and b1" -v
#
# These tests verify the B1 overflow implementation:
# 1. Each section has at least as many rows as the fixed template
# 2. Every overflow row has at least one non-None value (all-None rows are filtered)
# 3. The NG sheet (when present) has correct structure and non-empty data rows

@pytest.mark.slow
@pytest.mark.b1
@pytest.mark.parametrize("ticker", TICKERS)
def test_b1_section_row_counts(all_tables, ticker):
    """Each section in Data_Financials(Q) must have ≥ template row count.

    Verifies that overflow rows don't replace template rows and that the
    section structure is intact.
    """
    tbl = _get_tbl(all_tables, ticker)
    is_rows, bs_rows, cf_rows = _section_row_counts(tbl)
    assert is_rows is not None, f"[{ticker}] Section headers missing from Data_Financials(Q)"
    assert is_rows >= _IS_TEMPLATE_ROWS, (
        f"[{ticker}] IS section: {is_rows} rows < template {_IS_TEMPLATE_ROWS}"
    )
    assert bs_rows >= _BS_TEMPLATE_ROWS, (
        f"[{ticker}] BS section: {bs_rows} rows < template {_BS_TEMPLATE_ROWS}"
    )
    assert cf_rows >= _CF_TEMPLATE_ROWS, (
        f"[{ticker}] CF section: {cf_rows} rows < template {_CF_TEMPLATE_ROWS}"
    )


@pytest.mark.slow
@pytest.mark.b1
@pytest.mark.parametrize("ticker", TICKERS)
def test_b1_overflow_rows_nonnull(all_tables, ticker):
    """Every overflow row must have at least one non-None value.

    The build functions already filter all-None overflow rows before appending,
    so any row that made it into the table must have data somewhere.
    """
    tbl = _get_tbl(all_tables, ticker)
    c = tbl.concepts
    is_rows, bs_rows, cf_rows = _section_row_counts(tbl)
    if is_rows is None:
        pytest.skip(f"[{ticker}] Section headers missing")

    is_hdr = c.index("Income Statement")
    bs_hdr = c.index("Balance Sheet")
    cf_hdr = c.index("Cash Flow")
    is_start, bs_start, cf_start = is_hdr + 1, bs_hdr + 1, cf_hdr + 1

    failures = []
    # IS overflow
    for i in range(is_start + _IS_TEMPLATE_ROWS, bs_hdr - 1):
        if all(v is None for v in tbl.values[i]):
            failures.append(f"IS overflow '{c[i]}' is all-None")
    # BS overflow
    for i in range(bs_start + _BS_TEMPLATE_ROWS, cf_hdr - 1):
        if all(v is None for v in tbl.values[i]):
            failures.append(f"BS overflow '{c[i]}' is all-None")
    # CF overflow
    for i in range(cf_start + _CF_TEMPLATE_ROWS, len(c)):
        if all(v is None for v in tbl.values[i]):
            failures.append(f"CF overflow '{c[i]}' is all-None")

    assert not failures, f"[{ticker}] all-None overflow rows found:\n" + "\n".join(failures)


@pytest.mark.slow
@pytest.mark.b1
@pytest.mark.parametrize("ticker", TICKERS)
def test_b1_ng_sheet_structure(all_tables, ticker):
    """If Data_Financials_NG(Q) exists, every non-header row must have ≥1 non-None value.

    Checks that the NG sheet has the expected section structure and that every
    data row carries actual values (all-None NG rows are also filtered at build time).
    """
    result = all_tables[ticker]
    if isinstance(result, Exception):
        pytest.skip(f"{ticker}: fetch failed")

    ng_tbl = next((t for t in result if t.sheet_name == "Data_Financials_NG(Q)"), None)
    if ng_tbl is None:
        pytest.skip(f"[{ticker}] No NG sheet (no Non-GAAP overflow rows found)")

    # Section headers must be present
    assert "Income Statement" in ng_tbl.concepts, f"[{ticker}] NG sheet missing IS header"
    assert "Balance Sheet"    in ng_tbl.concepts, f"[{ticker}] NG sheet missing BS header"
    assert "Cash Flow"        in ng_tbl.concepts, f"[{ticker}] NG sheet missing CF header"

    # Quarter/value dimensions must be consistent
    n_q = len(ng_tbl.quarter_labels)
    for i, row in enumerate(ng_tbl.values):
        assert len(row) == n_q, (
            f"[{ticker}] NG row {i} ('{ng_tbl.concepts[i]}'): "
            f"values length {len(row)} != {n_q} quarters"
        )

    # Non-header rows must have ≥1 non-None value
    _HEADER_LABELS = {"Income Statement", "Balance Sheet", "Cash Flow", ""}
    failures = []
    for i, concept in enumerate(ng_tbl.concepts):
        if concept in _HEADER_LABELS:
            continue
        if all(v is None for v in ng_tbl.values[i]):
            failures.append(f"NG row '{concept}' (labels='{ng_tbl.labels[i]}') is all-None")

    assert not failures, f"[{ticker}] all-None NG data rows:\n" + "\n".join(failures)


# ── CF Overflow YTD Correctness Tests ────────────────────────────────────────
#
# Run with:  pytest -m "slow and cf_overflow" -v
#
# Validates the CF YTD overflow fix: Q2 and Q3 overflow rows must be populated
# (not skipped) after the cross-filing subtraction was implemented.
# Before the fix, CF overflow only had data for Q1 and FY quarters.

@pytest.fixture(scope="module")
def cf_overflow_tables(identity, ai_config):
    """Fetch CF overflow tickers (COHR, LITE, AAPL, NVDA, GOOGL)."""
    result: dict = {}
    for ticker in CF_OVERFLOW_TICKERS:
        try:
            result[ticker] = fetch_gaap_statements(
                ticker, identity, max_filings=12, ai_config=ai_config
            )
        except Exception as exc:
            result[ticker] = exc
    return result


def _cf_overflow_rows(tbl):
    """Return overflow rows from the CF section of a merged financials table.

    Returns list of (concept_name, values_list) for rows beyond the template rows.
    """
    c = tbl.concepts
    try:
        cf_hdr = c.index("Cash Flow")
    except ValueError:
        return []
    cf_start = cf_hdr + 1
    overflow_start = cf_start + _CF_TEMPLATE_ROWS
    rows = []
    for i in range(overflow_start, len(c)):
        rows.append((c[i], tbl.values[i]))
    return rows


@pytest.mark.slow
@pytest.mark.cf_overflow
@pytest.mark.parametrize("ticker", CF_OVERFLOW_TICKERS)
def test_cf_overflow_rows_exist(cf_overflow_tables, ticker):
    """CF section must have at least one overflow row.

    If zero overflow rows, either the company has no unmatched CF XBRL items
    (acceptable) or the fix broke collection — fails only when we expect overflow
    based on known COHR/LITE behaviour.
    """
    result = cf_overflow_tables[ticker]
    if isinstance(result, Exception):
        pytest.fail(f"{ticker}: fetch failed — {result}")
    tbl = _quarterly_financials(result)
    assert tbl is not None, f"{ticker}: Data_Financials(Q) not found"
    overflow = _cf_overflow_rows(tbl)
    if ticker in ("COHR", "LITE"):
        assert len(overflow) > 0, (
            f"[{ticker}] Expected CF overflow rows but none found — "
            "check that _collect_overflow is being called for CF"
        )


@pytest.mark.slow
@pytest.mark.cf_overflow
@pytest.mark.parametrize("ticker", CF_OVERFLOW_TICKERS)
def test_cf_overflow_multi_quarter_coverage(cf_overflow_tables, ticker):
    """CF overflow rows should have data across ≥2 distinct quarters.

    Before the YTD fix, only Q1/FY filings contributed overflow → most rows
    had data in only 1 quarter.  After the fix, Q2 and Q3 are subtracted and
    populated, so a row with Q1 data should also show Q2/Q3.

    Test: at least one overflow row has non-None values in ≥2 quarters.
    Skipped if no CF overflow rows exist for this ticker.
    """
    result = cf_overflow_tables[ticker]
    if isinstance(result, Exception):
        pytest.skip(f"{ticker}: fetch failed")
    tbl = _quarterly_financials(result)
    if tbl is None:
        pytest.skip(f"{ticker}: no Data_Financials(Q)")
    overflow = _cf_overflow_rows(tbl)
    if not overflow:
        pytest.skip(f"[{ticker}] No CF overflow rows — nothing to test")

    # Count rows with ≥2 non-None values
    multi_q = sum(
        1 for _, vals in overflow
        if sum(v is not None for v in vals) >= 2
    )
    assert multi_q > 0, (
        f"[{ticker}] All CF overflow rows have data in only 1 quarter — "
        f"YTD subtraction may not be working. "
        f"Overflow rows: {[name for name, _ in overflow]}"
    )


@pytest.mark.slow
@pytest.mark.cf_overflow
@pytest.mark.parametrize("ticker", CF_OVERFLOW_TICKERS)
def test_cf_overflow_no_all_none_rows(cf_overflow_tables, ticker):
    """No CF overflow row should be all-None (filter applied at build time)."""
    result = cf_overflow_tables[ticker]
    if isinstance(result, Exception):
        pytest.skip(f"{ticker}: fetch failed")
    tbl = _quarterly_financials(result)
    if tbl is None:
        pytest.skip(f"{ticker}: no Data_Financials(Q)")
    overflow = _cf_overflow_rows(tbl)
    if not overflow:
        pytest.skip(f"[{ticker}] No CF overflow rows")

    failures = [
        name for name, vals in overflow if all(v is None for v in vals)
    ]
    assert not failures, (
        f"[{ticker}] All-None CF overflow rows (should have been filtered): {failures}"
    )
