"""Tests for fetcher_gaap.py — quarterly multi-filing approach."""
import pytest
from unittest.mock import MagicMock, patch
import pandas as pd
from fetcher_gaap import (
    fetch_gaap_statements,
    StatementTable,
    _col_to_quarter_label,
    _current_q_col,
    _match_is_row,
    _build_is_table,
)

# ── helpers ──────────────────────────────────────────────────────────────

META_COLS = {
    'concept', 'label', 'standard_concept', 'level', 'abstract',
    'dimension', 'is_breakdown', 'dimension_axis', 'dimension_member',
    'dimension_member_label', 'dimension_label', 'unit', 'point_in_time',
    'balance', 'weight', 'preferred_sign', 'parent_concept', 'parent_abstract_concept',
}


def _make_is_df(period_col="2025-12-27 (Q1)", val=100.0, prior_col="2024-12-28 (Q1)", prior_val=90.0):
    """Minimal IS DataFrame with Revenue, Gross Profit, Net Income rows."""
    return pd.DataFrame({
        "concept":               ["us-gaap_RevenueFromContractWithCustomer", "us-gaap_GrossProfit", "us-gaap_NetIncomeLoss"],
        "label":                 ["Net sales", "Gross margin", "Net income"],
        "standard_concept":      ["Revenue", "GrossProfit", "NetIncome"],
        "abstract":              [False, False, False],
        "is_breakdown":          [False, False, False],
        "level":                 [4, 3, 3],
        "dimension_member_label":[None, None, None],
        period_col:              [val * 10, val * 7, val * 2],
        prior_col:               [prior_val * 10, prior_val * 7, prior_val * 2],
    })


def _make_filing(period_col="2025-12-27 (Q1)", val=100.0,
                 prior_col="2024-12-28 (Q1)", prior_val=90.0,
                 filing_date="2026-01-30"):
    """Mock a single 10-Q Filing object."""
    df = _make_is_df(period_col, val, prior_col, prior_val)
    mock_stmt = MagicMock()
    mock_stmt.to_dataframe.return_value = df

    mock_financials = MagicMock()
    mock_financials.income_statement.return_value = mock_stmt
    mock_financials.balance_sheet.return_value = mock_stmt
    mock_financials.cashflow_statement.return_value = mock_stmt

    mock_tenq = MagicMock()
    mock_tenq.financials = mock_financials

    mock_filing = MagicMock()
    mock_filing.obj.return_value = mock_tenq
    mock_filing.filing_date = filing_date
    return mock_filing


# ── unit tests ────────────────────────────────────────────────────────────

def test_col_to_quarter_label_q1():
    assert _col_to_quarter_label("2023-03-31 (Q1)") == "FY2023Q1"

def test_col_to_quarter_label_fy():
    assert _col_to_quarter_label("2024-12-31 (FY)") == "FY2024"

def test_col_to_quarter_label_instant_passthrough():
    assert _col_to_quarter_label("2023-03-31") == "2023-03-31"


def test_current_q_col_picks_first_q_col():
    df = _make_is_df()  # has "2025-12-27 (Q1)" and "2024-12-28 (Q1)"
    col = _current_q_col(df)
    assert col == "2025-12-27 (Q1)"

def test_current_q_col_skips_ytd():
    df = pd.DataFrame({
        "concept": ["c"], "label": ["l"], "standard_concept": ["s"],
        "abstract": [False], "is_breakdown": [False], "level": [1],
        "dimension_member_label": [None],
        "2025-06-28 (YTD)": [1.0],
        "2025-06-28 (Q3)":  [2.0],
    })
    col = _current_q_col(df)
    assert col == "2025-06-28 (Q3)"

def test_current_q_col_returns_none_when_no_period():
    df = pd.DataFrame({"concept": ["c"], "label": ["l"]})
    assert _current_q_col(df) is None


def test_match_is_row_by_standard_concept():
    df = _make_is_df()
    idx = _match_is_row(df, std_concept="Revenue", fallback_suffix="RevenueFromContract")
    assert idx is not None
    assert df.loc[idx, "label"] == "Net sales"

def test_match_is_row_fallback_when_no_std_concept():
    df = _make_is_df()
    df.loc[0, "standard_concept"] = None
    idx = _match_is_row(df, std_concept="Revenue", fallback_suffix="RevenueFromContractWithCustomer")
    assert idx is not None

def test_match_is_row_returns_none_when_not_found():
    df = _make_is_df()
    idx = _match_is_row(df, std_concept="InterestExpense", fallback_suffix="InterestExpense")
    assert idx is None

def test_match_is_row_ignores_abstract_rows():
    df = _make_is_df()
    df.loc[0, "abstract"] = True
    idx = _match_is_row(df, std_concept="Revenue", fallback_suffix="RevenueFromContract")
    assert idx is None

def test_match_is_row_ignores_breakdown_rows():
    df = _make_is_df()
    df.loc[0, "is_breakdown"] = True
    idx = _match_is_row(df, std_concept="Revenue", fallback_suffix="RevenueFromContract")
    assert idx is None

def test_match_is_row_ignores_dimensional_rows():
    df = _make_is_df()
    df.loc[0, "dimension_member_label"] = "Products"
    idx = _match_is_row(df, std_concept="Revenue", fallback_suffix="RevenueFromContract")
    assert idx is None


def test_build_is_table_returns_statement_table():
    filing = _make_filing()
    tbl = _build_is_table([filing], max_filings=1)
    assert isinstance(tbl, StatementTable)
    assert tbl.sheet_name == "Data_IS"

def test_build_is_table_has_21_concept_rows():
    filing = _make_filing()
    tbl = _build_is_table([filing], max_filings=1)
    assert len(tbl.concepts) == 21

def test_build_is_table_quarter_labels_format():
    filing = _make_filing(period_col="2025-12-27 (Q1)")
    tbl = _build_is_table([filing], max_filings=1)
    assert tbl.quarter_labels == ["FY2025Q1"]

def test_build_is_table_filing_dates():
    filing = _make_filing(filing_date="2026-01-30")
    tbl = _build_is_table([filing], max_filings=1)
    assert tbl.filing_dates == ["2026-01-30"]

def test_build_is_table_revenue_value():
    filing = _make_filing(period_col="2025-12-27 (Q1)", val=100.0)
    tbl = _build_is_table([filing], max_filings=1)
    revenue_idx = tbl.concepts.index("Revenue")
    assert tbl.values[revenue_idx][0] == 1000.0  # val * 10

def test_build_is_table_missing_rows_are_none():
    filing = _make_filing()
    tbl = _build_is_table([filing], max_filings=1)
    interest_idx = tbl.concepts.index("Interest Expense")
    assert tbl.values[interest_idx][0] is None

def test_build_is_table_two_filings_oldest_to_newest():
    f1 = _make_filing(period_col="2025-12-27 (Q1)", val=100.0, filing_date="2026-01-30",
                       prior_col="2024-12-28 (Q1)", prior_val=90.0)
    f2 = _make_filing(period_col="2024-12-28 (Q1)", val=90.0, filing_date="2025-01-31",
                       prior_col="2023-12-30 (Q1)", prior_val=80.0)
    tbl = _build_is_table([f1, f2], max_filings=2)
    assert tbl.quarter_labels[0] == "FY2024Q1"
    assert tbl.quarter_labels[1] == "FY2025Q1"

def test_build_is_table_deduplicates_same_period():
    f1 = _make_filing(period_col="2025-12-27 (Q1)", val=100.0, filing_date="2026-01-30",
                       prior_col="2024-12-28 (Q1)", prior_val=90.0)
    f2 = _make_filing(period_col="2024-12-28 (Q1)", val=90.0, filing_date="2025-01-31",
                       prior_col="2023-12-30 (Q1)", prior_val=80.0)
    tbl = _build_is_table([f1, f2], max_filings=2)
    assert len(tbl.quarter_labels) == 2
    assert len(set(tbl.quarter_labels)) == 2

def test_build_is_table_respects_max_filings():
    filings = [_make_filing(period_col=f"202{i}-12-27 (Q1)", val=float(i),
                             prior_col=f"202{i-1}-12-28 (Q1)", prior_val=float(i-1),
                             filing_date=f"202{i+1}-01-30")
               for i in range(1, 6)]
    tbl = _build_is_table(filings, max_filings=3)
    assert len(tbl.quarter_labels) == 3


# ── integration tests ─────────────────────────────────────────────────────

def test_fetch_returns_list_of_statement_tables():
    with patch("fetcher_gaap.Company") as MockCo, patch("fetcher_gaap.set_identity"):
        MockCo.return_value = _make_mock_company()
        result = fetch_gaap_statements("AAPL", identity="Test test@test.com")
    assert isinstance(result, list)
    assert all(isinstance(t, StatementTable) for t in result)

def test_fetch_includes_required_sheets():
    with patch("fetcher_gaap.Company") as MockCo, patch("fetcher_gaap.set_identity"):
        MockCo.return_value = _make_mock_company()
        result = fetch_gaap_statements("AAPL", identity="Test test@test.com")
    sheet_names = [t.sheet_name for t in result]
    assert "Data_IS" in sheet_names
    assert "Data_BS" in sheet_names
    assert "Data_CF" in sheet_names
    assert "Data_Meta" in sheet_names

def test_fetch_consistent_row_col_lengths():
    with patch("fetcher_gaap.Company") as MockCo, patch("fetcher_gaap.set_identity"):
        MockCo.return_value = _make_mock_company()
        result = fetch_gaap_statements("AAPL", identity="Test test@test.com")
    for tbl in result:
        if tbl.sheet_name == "Data_Meta":
            continue
        n_q = len(tbl.quarter_labels)
        assert len(tbl.filing_dates) == n_q
        for row in tbl.values:
            assert len(row) == n_q, f"Sheet {tbl.sheet_name}: row length {len(row)} != {n_q}"

def test_fetch_raises_on_invalid_ticker():
    with patch("fetcher_gaap.Company") as MockCo, patch("fetcher_gaap.set_identity"):
        MockCo.return_value = MagicMock()
        MockCo.return_value.get_filings.return_value = []
        with pytest.raises(ValueError, match="No 10-Q"):
            fetch_gaap_statements("XXXX", identity="Test test@test.com")

def test_fetch_passes_max_filings():
    with patch("fetcher_gaap.Company") as MockCo, patch("fetcher_gaap.set_identity"):
        mock_co = _make_mock_company(n_filings=10)
        MockCo.return_value = mock_co
        result = fetch_gaap_statements("AAPL", identity="Test test@test.com", max_filings=3)
    is_tbl = next(t for t in result if t.sheet_name == "Data_IS")
    assert len(is_tbl.quarter_labels) <= 3


def test_fetch_sets_ticker_on_all_tables():
    with patch("fetcher_gaap.Company") as MockCo, patch("fetcher_gaap.set_identity"):
        MockCo.return_value = _make_mock_company()
        result = fetch_gaap_statements("AAPL", identity="Test test@test.com")
    assert all(t.ticker == "AAPL" for t in result)


# ── fixtures ──────────────────────────────────────────────────────────────

def _make_mock_company(n_filings=2):
    """Mock Company with n_filings 10-Q filings."""
    filings = [
        _make_filing(
            period_col=f"202{5 - i}-12-27 (Q1)",
            val=float(100 - i * 10),
            prior_col=f"202{4 - i}-12-28 (Q1)",
            prior_val=float(90 - i * 10),
            filing_date=f"202{6 - i}-01-30",
        )
        for i in range(n_filings)
    ]
    mock_filings_obj = MagicMock()
    mock_filings_obj.__iter__ = MagicMock(side_effect=lambda: iter(filings))
    mock_filings_obj.__len__ = MagicMock(return_value=len(filings))
    mock_filings_obj.__getitem__ = MagicMock(side_effect=lambda i: filings[i] if isinstance(i, int) else filings)

    mock_co = MagicMock()
    mock_co.name = "Apple Inc."
    mock_co.get_filings.return_value = mock_filings_obj
    return mock_co
