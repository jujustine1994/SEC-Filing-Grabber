"""Tests for fetcher_gaap.py — quarterly multi-filing approach.

New tests (overflow feature):
  - _is_nongaap_label: keyword-based Non-GAAP label detector
  - _collect_overflow: routing unmatched XBRL rows into GAAP / NG buckets
  - Build function return-type smoke tests (now return tuple[StatementTable, StatementTable])

Existing tests have been updated: _build_is_table / _build_cf_table now return
(gaap_tbl, ng_tbl) tuples.  All existing assertions target gaap_tbl (first element).
"""
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
    _build_bs_table,
    _build_cf_table,
    _merge_financials,
    _ytd_col,
    _prev_quarter_label,
    _is_nongaap_label,
    _collect_overflow,
)

# ── helpers ───────────────────────────────────────────────────────────────────

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


# ── _is_nongaap_label ─────────────────────────────────────────────────────────

def test_nongaap_label_non_gaap():
    assert _is_nongaap_label("Non-GAAP Revenue") is True

def test_nongaap_label_adjusted():
    assert _is_nongaap_label("Adjusted Operating Income") is True

def test_nongaap_label_excluding():
    assert _is_nongaap_label("Gross profit excluding discontinued ops") is True

def test_nongaap_label_excl_dot():
    assert _is_nongaap_label("Operating income, excl. SBC") is True

def test_nongaap_label_gaap_row():
    assert _is_nongaap_label("Revenue") is False

def test_nongaap_label_total_assets():
    assert _is_nongaap_label("Total assets") is False

def test_nongaap_label_case_insensitive():
    assert _is_nongaap_label("NON-GAAP EPS") is True

def test_nongaap_label_non_gaap_space():
    assert _is_nongaap_label("non gaap gross profit") is True


# ── _collect_overflow ─────────────────────────────────────────────────────────

def _make_overflow_df() -> pd.DataFrame:
    """
    Minimal DataFrame with 3 rows:
      index 0 — GrossProfit (will be marked consumed)
      index 1 — OperatingLeaseAsset (GAAP overflow)
      index 2 — AdjustedGrossProfit (NG overflow, label triggers NG routing)
    """
    return pd.DataFrame({
        "concept":                ["GrossProfit",  "OperatingLeaseAsset", "AdjustedGrossProfit"],
        "label":                  ["Gross profit", "Operating lease ROU", "Adjusted gross profit"],
        "standard_concept":       ["GrossProfit",  None,                  None],
        "abstract":               [False,          False,                 False],
        "is_breakdown":           [False,          False,                 False],
        "dimension_member_label": [None,           None,                  None],
        "2024-03-31 (Q1)":        [50_000,         10_000,                55_000],
    })


def test_collect_overflow_gaap_row_captured():
    df = _make_overflow_df()
    consumed = {0}
    gaap, ng = {}, {}
    _collect_overflow(df, consumed, "2024-03-31 (Q1)", "FY2024Q1", gaap, ng)
    assert "OperatingLeaseAsset" in gaap
    assert gaap["OperatingLeaseAsset"]["periods"]["FY2024Q1"] == 10_000


def test_collect_overflow_ng_row_routed():
    df = _make_overflow_df()
    consumed = {0}
    gaap, ng = {}, {}
    _collect_overflow(df, consumed, "2024-03-31 (Q1)", "FY2024Q1", gaap, ng)
    assert "AdjustedGrossProfit" in ng
    assert ng["AdjustedGrossProfit"]["periods"]["FY2024Q1"] == 55_000


def test_collect_overflow_consumed_excluded():
    df = _make_overflow_df()
    consumed = {0}
    gaap, ng = {}, {}
    _collect_overflow(df, consumed, "2024-03-31 (Q1)", "FY2024Q1", gaap, ng)
    assert "GrossProfit" not in gaap
    assert "GrossProfit" not in ng


def test_collect_overflow_none_value_not_stored():
    """None values are not stored in periods dict; key still created."""
    df = _make_overflow_df()
    df.loc[1, "2024-03-31 (Q1)"] = None
    consumed = {0}
    gaap, ng = {}, {}
    _collect_overflow(df, consumed, "2024-03-31 (Q1)", "FY2024Q1", gaap, ng)
    assert "OperatingLeaseAsset" in gaap
    assert gaap["OperatingLeaseAsset"]["periods"] == {}


def test_collect_overflow_accumulates_across_quarters():
    """Calling _collect_overflow twice with different quarters merges into same dict."""
    df1 = _make_overflow_df()
    df2 = _make_overflow_df().rename(columns={"2024-03-31 (Q1)": "2024-06-30 (Q2)"})
    consumed = {0}
    gaap, ng = {}, {}
    _collect_overflow(df1, consumed, "2024-03-31 (Q1)", "FY2024Q1", gaap, ng)
    _collect_overflow(df2, consumed, "2024-06-30 (Q2)", "FY2024Q2", gaap, ng)
    assert "FY2024Q1" in gaap["OperatingLeaseAsset"]["periods"]
    assert "FY2024Q2" in gaap["OperatingLeaseAsset"]["periods"]


def test_collect_overflow_abstract_rows_excluded():
    """Abstract rows must not appear in overflow."""
    df = _make_overflow_df()
    df.loc[1, "abstract"] = True
    consumed = {0}
    gaap, ng = {}, {}
    _collect_overflow(df, consumed, "2024-03-31 (Q1)", "FY2024Q1", gaap, ng)
    assert "OperatingLeaseAsset" not in gaap


def test_collect_overflow_dimension_rows_excluded():
    """Rows with dimension_member_label must not appear in overflow."""
    df = _make_overflow_df()
    df.loc[1, "dimension_member_label"] = "SomeSegment"
    consumed = {0}
    gaap, ng = {}, {}
    _collect_overflow(df, consumed, "2024-03-31 (Q1)", "FY2024Q1", gaap, ng)
    assert "OperatingLeaseAsset" not in gaap


def test_collect_overflow_empty_concept_skipped():
    """Rows with empty concept field are silently skipped."""
    df = _make_overflow_df()
    df.loc[1, "concept"] = ""
    consumed = {0}
    gaap, ng = {}, {}
    _collect_overflow(df, consumed, "2024-03-31 (Q1)", "FY2024Q1", gaap, ng)
    assert "" not in gaap


# ── Build function return-type smoke tests (empty filings) ────────────────────

def test_build_is_table_empty_filings_returns_tuple():
    """_build_is_table must return a 2-tuple of StatementTables even for empty input."""
    result = _build_is_table([], max_filings=8)
    assert isinstance(result, tuple) and len(result) == 2
    gaap_tbl, ng_tbl = result
    assert isinstance(gaap_tbl, StatementTable)
    assert isinstance(ng_tbl, StatementTable)
    assert len(gaap_tbl.concepts) > 0   # IS template concepts present
    assert ng_tbl.concepts == []


def test_build_bs_table_empty_filings_returns_tuple():
    result = _build_bs_table([], max_filings=8)
    assert isinstance(result, tuple) and len(result) == 2
    gaap_tbl, ng_tbl = result
    assert isinstance(gaap_tbl, StatementTable)
    assert len(gaap_tbl.concepts) > 0   # BS template concepts present
    assert ng_tbl.concepts == []


def test_build_cf_table_empty_filings_returns_tuple():
    result = _build_cf_table([], max_filings=8)
    assert isinstance(result, tuple) and len(result) == 2
    gaap_tbl, ng_tbl = result
    assert isinstance(gaap_tbl, StatementTable)
    assert len(gaap_tbl.concepts) > 0   # CF template concepts present
    assert ng_tbl.concepts == []


# ── unit tests ────────────────────────────────────────────────────────────────

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


# _build_is_table tests — unpack (gaap_tbl, ng_tbl) tuple

def test_build_is_table_returns_statement_table():
    filing = _make_filing()
    gaap_tbl, ng_tbl = _build_is_table([filing], max_filings=1)
    assert isinstance(gaap_tbl, StatementTable)
    assert gaap_tbl.sheet_name == "Data_IS"
    assert isinstance(ng_tbl, StatementTable)

def test_build_is_table_has_22_concept_rows():
    # Mock IS df has only 3 rows; all 3 are consumed by template → no overflow
    filing = _make_filing()
    gaap_tbl, _ = _build_is_table([filing], max_filings=1)
    assert len(gaap_tbl.concepts) == 22

def test_build_is_table_quarter_labels_format():
    filing = _make_filing(period_col="2025-12-27 (Q1)")
    gaap_tbl, _ = _build_is_table([filing], max_filings=1)
    assert gaap_tbl.quarter_labels == ["FY2025Q1"]

def test_build_is_table_filing_dates():
    filing = _make_filing(filing_date="2026-01-30")
    gaap_tbl, _ = _build_is_table([filing], max_filings=1)
    assert gaap_tbl.filing_dates == ["2026-01-30"]

def test_build_is_table_revenue_value():
    filing = _make_filing(period_col="2025-12-27 (Q1)", val=100.0)
    gaap_tbl, _ = _build_is_table([filing], max_filings=1)
    revenue_idx = gaap_tbl.concepts.index("Revenue")
    assert gaap_tbl.values[revenue_idx][0] == 1000.0  # val * 10

def test_build_is_table_missing_rows_are_none():
    filing = _make_filing()
    gaap_tbl, _ = _build_is_table([filing], max_filings=1)
    interest_idx = gaap_tbl.concepts.index("Interest Expense")
    assert gaap_tbl.values[interest_idx][0] is None


def test_match_is_row_label_fallback():
    """Third-tier: label column match when std_concept and concept suffix both miss."""
    df = pd.DataFrame({
        "concept":               ["co:CustomDepreciation"],
        "label":                 ["Depreciation, amortization and impairment"],
        "standard_concept":      [float("nan")],
        "abstract":              [False],
        "is_breakdown":          [False],
        "level":                 [3],
        "dimension_member_label":[None],
        "2025-12-27 (Q1)":       [50.0],
        "2024-12-28 (Q1)":       [45.0],
    })
    idx = _match_is_row(df, "DepreciationExpense", "DepreciationDepletion", label_fallback="depreciation")
    assert idx is not None
    assert df.loc[idx, "label"] == "Depreciation, amortization and impairment"


def test_build_is_table_net_income_profitloss_fallback():
    """Net Income uses ProfitLoss when NetIncome std_concept is absent (TSLA/BA/XOM/WMT)."""
    df = pd.DataFrame({
        "concept":               ["us-gaap_ProfitLoss"],
        "label":                 ["Net income"],
        "standard_concept":      ["ProfitLoss"],
        "abstract":              [False],
        "is_breakdown":          [False],
        "level":                 [3],
        "dimension_member_label":[None],
        "2025-12-27 (Q1)":       [200.0],
        "2024-12-28 (Q1)":       [180.0],
    })
    mock_stmt = MagicMock()
    mock_stmt.to_dataframe.return_value = df
    mock_financials = MagicMock()
    mock_financials.income_statement.return_value = mock_stmt
    mock_financials.cashflow_statement.return_value = mock_stmt
    mock_tenq = MagicMock()
    mock_tenq.financials = mock_financials
    mock_filing = MagicMock()
    mock_filing.obj.return_value = mock_tenq
    mock_filing.filing_date = "2026-01-30"

    gaap_tbl, _ = _build_is_table([mock_filing], max_filings=1)
    net_income_idx = gaap_tbl.concepts.index("Net Income")
    assert gaap_tbl.values[net_income_idx][0] == 200.0


def test_build_is_table_total_nonop_derived_from_pretax_minus_operating():
    """Total Non-op falls back to Pre-tax − Operating when XBRL row absent."""
    df = pd.DataFrame({
        "concept":               ["us-gaap_OperatingIncomeLoss", "us-gaap_PretaxIncomeLoss"],
        "label":                 ["Operating income", "Income before taxes"],
        "standard_concept":      ["OperatingIncomeLoss", "PretaxIncomeLoss"],
        "abstract":              [False, False],
        "is_breakdown":          [False, False],
        "level":                 [3, 3],
        "dimension_member_label":[None, None],
        "2025-12-27 (Q1)":       [100.0, 115.0],
        "2024-12-28 (Q1)":       [90.0, 103.0],
    })
    mock_stmt = MagicMock()
    mock_stmt.to_dataframe.return_value = df
    mock_financials = MagicMock()
    mock_financials.income_statement.return_value = mock_stmt
    mock_financials.cashflow_statement.return_value = mock_stmt
    mock_tenq = MagicMock()
    mock_tenq.financials = mock_financials
    mock_filing = MagicMock()
    mock_filing.obj.return_value = mock_tenq
    mock_filing.filing_date = "2026-01-30"

    gaap_tbl, _ = _build_is_table([mock_filing], max_filings=1)
    nonop_idx = gaap_tbl.concepts.index("Total Non-op Income/(Loss)")
    assert gaap_tbl.values[nonop_idx][0] == 15.0  # 115 − 100

def test_build_is_table_two_filings_oldest_to_newest():
    f1 = _make_filing(period_col="2025-12-27 (Q1)", val=100.0, filing_date="2026-01-30",
                       prior_col="2024-12-28 (Q1)", prior_val=90.0)
    f2 = _make_filing(period_col="2024-12-28 (Q1)", val=90.0, filing_date="2025-01-31",
                       prior_col="2023-12-30 (Q1)", prior_val=80.0)
    gaap_tbl, _ = _build_is_table([f1, f2], max_filings=2)
    assert gaap_tbl.quarter_labels[0] == "FY2024Q1"
    assert gaap_tbl.quarter_labels[1] == "FY2025Q1"

def test_build_is_table_deduplicates_same_period():
    f1 = _make_filing(period_col="2025-12-27 (Q1)", val=100.0, filing_date="2026-01-30",
                       prior_col="2024-12-28 (Q1)", prior_val=90.0)
    f2 = _make_filing(period_col="2024-12-28 (Q1)", val=90.0, filing_date="2025-01-31",
                       prior_col="2023-12-30 (Q1)", prior_val=80.0)
    gaap_tbl, _ = _build_is_table([f1, f2], max_filings=2)
    assert len(gaap_tbl.quarter_labels) == 2
    assert len(set(gaap_tbl.quarter_labels)) == 2

def test_build_is_table_respects_max_filings():
    filings = [_make_filing(period_col=f"202{i}-12-27 (Q1)", val=float(i),
                             prior_col=f"202{i-1}-12-28 (Q1)", prior_val=float(i-1),
                             filing_date=f"202{i+1}-01-30")
               for i in range(1, 6)]
    gaap_tbl, _ = _build_is_table(filings, max_filings=3)
    assert len(gaap_tbl.quarter_labels) == 3


# ── integration tests ─────────────────────────────────────────────────────────

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
    assert "Data_Financials(Q)" in sheet_names
    assert "Data_Financials(Y)" in sheet_names
    assert "Data_Meta" in sheet_names
    # Separate IS/BS/CF sheets are not produced (internal use only)
    assert "Data_IS" not in sheet_names
    assert "Data_BS" not in sheet_names
    assert "Data_CF" not in sheet_names

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
    fin_tbl = next(t for t in result if t.sheet_name == "Data_Financials(Q)")
    assert len(fin_tbl.quarter_labels) <= 3


def test_merge_financials_produces_data_financials_sheet():
    is_tbl = StatementTable(
        sheet_name="Data_IS", quarter_labels=["FY2024Q1"], filing_dates=["2024-02-01"],
        concepts=["Revenue"], values=[[100.0]], labels=["Net sales"],
    )
    bs_tbl = StatementTable(
        sheet_name="Data_BS", quarter_labels=["FY2024Q1"], filing_dates=["2024-02-01"],
        concepts=["Total Assets"], values=[[5000.0]], labels=["Total assets"],
    )
    cf_tbl = StatementTable(
        sheet_name="Data_CF", quarter_labels=["FY2024Q1"], filing_dates=["2024-02-01"],
        concepts=["Operating Cash Flow"], values=[[200.0]], labels=["Net cash from ops"],
    )
    merged = _merge_financials(is_tbl, bs_tbl, cf_tbl)
    assert merged.sheet_name == "Data_Financials(Q)"
    assert "Income Statement" in merged.concepts
    assert "Balance Sheet" in merged.concepts
    assert "Cash Flow" in merged.concepts
    assert "Revenue" in merged.concepts
    assert "Total Assets" in merged.concepts
    assert "Operating Cash Flow" in merged.concepts


def test_merge_financials_section_headers_have_none_values():
    is_tbl = StatementTable(
        sheet_name="Data_IS", quarter_labels=["FY2024Q1"], filing_dates=["2024-02-01"],
        concepts=["Revenue"], values=[[100.0]], labels=["Net sales"],
    )
    bs_tbl = StatementTable(
        sheet_name="Data_BS", quarter_labels=["FY2024Q1"], filing_dates=["2024-02-01"],
        concepts=["Assets"], values=[[5000.0]], labels=[""],
    )
    cf_tbl = StatementTable(
        sheet_name="Data_CF", quarter_labels=["FY2024Q1"], filing_dates=["2024-02-01"],
        concepts=["Operating Cash Flow"], values=[[200.0]], labels=[""],
    )
    merged = _merge_financials(is_tbl, bs_tbl, cf_tbl)
    header_idx = merged.concepts.index("Income Statement")
    assert merged.values[header_idx] == [None]
    assert merged.labels[header_idx] == ""


def test_build_is_table_populates_labels():
    """labels list should be populated with original XBRL labels."""
    filing = _make_filing()
    gaap_tbl, _ = _build_is_table([filing], max_filings=1)
    assert len(gaap_tbl.labels) == len(gaap_tbl.concepts)
    revenue_idx = gaap_tbl.concepts.index("Revenue")
    assert gaap_tbl.labels[revenue_idx] == "Net sales"


def test_fetch_sets_ticker_on_all_tables():
    with patch("fetcher_gaap.Company") as MockCo, patch("fetcher_gaap.set_identity"):
        MockCo.return_value = _make_mock_company()
        result = fetch_gaap_statements("AAPL", identity="Test test@test.com")
    assert all(t.ticker == "AAPL" for t in result)


# ── fixtures ──────────────────────────────────────────────────────────────────

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


# ── _ytd_col tests ────────────────────────────────────────────────────────────

def test_ytd_col_returns_ytd_column():
    df = pd.DataFrame({
        "concept": ["c"], "label": ["l"], "standard_concept": ["s"],
        "abstract": [False], "is_breakdown": [False], "level": [1],
        "dimension_member_label": [None],
        "2025-06-30 (YTD)": [100.0],
    })
    assert _ytd_col(df) == "2025-06-30 (YTD)"

def test_ytd_col_ignores_q_columns():
    df = pd.DataFrame({
        "concept": ["c"], "label": ["l"], "standard_concept": ["s"],
        "abstract": [False], "is_breakdown": [False], "level": [1],
        "dimension_member_label": [None],
        "2025-03-31 (Q1)": [100.0],
    })
    assert _ytd_col(df) is None

def test_ytd_col_returns_none_when_no_period_cols():
    df = pd.DataFrame({"concept": ["c"], "label": ["l"]})
    assert _ytd_col(df) is None

def test_ytd_col_returns_q3_ytd_column():
    df = pd.DataFrame({
        "concept": ["c"], "label": ["l"], "standard_concept": ["s"],
        "abstract": [False], "is_breakdown": [False], "level": [1],
        "dimension_member_label": [None],
        "2025-09-30 (YTD)": [300.0],
        "2024-09-30 (YTD)": [270.0],
    })
    assert _ytd_col(df) == "2025-09-30 (YTD)"


# ── _prev_quarter_label tests ─────────────────────────────────────────────────

def test_prev_quarter_label_q2_returns_q1():
    assert _prev_quarter_label("FY2025Q2") == "FY2025Q1"

def test_prev_quarter_label_q3_returns_q2():
    assert _prev_quarter_label("FY2025Q3") == "FY2025Q2"

def test_prev_quarter_label_q4_returns_q3():
    assert _prev_quarter_label("FY2025Q4") == "FY2025Q3"

def test_prev_quarter_label_q1_returns_none():
    assert _prev_quarter_label("FY2025Q1") is None

def test_prev_quarter_label_annual_returns_none():
    assert _prev_quarter_label("FY2025") is None


# ── CF YTD subtraction integration tests ─────────────────────────────────────

def _make_cf_df_minimal(period_col, net_income_val, ocf_val):
    """Minimal CF DataFrame with Net Income and OCF rows."""
    return pd.DataFrame({
        "concept":               ["us-gaap_NetIncomeLoss",  "us-gaap_NetCashProvidedByUsedInOperatingActivities"],
        "label":                 ["Net income",             "Net cash provided by operating activities"],
        "standard_concept":      ["NetIncome",              "NetCashFromOperatingActivities"],
        "abstract":              [False,                     False],
        "is_breakdown":          [False,                     False],
        "level":                 [3,                         3],
        "dimension_member_label":[None,                      None],
        period_col:              [net_income_val,            ocf_val],
    })

def _make_is_df_minimal(period_col):
    """Minimal IS DataFrame for quarter label derivation."""
    return pd.DataFrame({
        "concept": ["us-gaap_RevenueFromContractWithCustomer"],
        "label": ["Net sales"],
        "standard_concept": ["Revenue"],
        "abstract": [False], "is_breakdown": [False], "level": [4],
        "dimension_member_label": [None],
        period_col: [1000.0],
    })

def _make_cf_filing(is_period_col, cf_period_col, ni, ocf, filing_date):
    """Mock a 10-Q filing with given IS and CF columns."""
    is_df = _make_is_df_minimal(is_period_col)
    cf_df = _make_cf_df_minimal(cf_period_col, ni, ocf)
    mock_is = MagicMock(); mock_is.to_dataframe.return_value = is_df
    mock_cf = MagicMock(); mock_cf.to_dataframe.return_value = cf_df
    mock_fin = MagicMock()
    mock_fin.income_statement.return_value = mock_is
    mock_fin.cashflow_statement.return_value = mock_cf
    mock_tenq = MagicMock(); mock_tenq.financials = mock_fin
    mock_filing = MagicMock()
    mock_filing.obj.return_value = mock_tenq
    mock_filing.filing_date = filing_date
    return mock_filing


def test_build_cf_table_q1_standalone_unchanged():
    """Q1 has a standalone (Q1) CF column — value should pass through as-is."""
    q1 = _make_cf_filing("2025-03-31 (Q1)", "2025-03-31 (Q1)", ni=100.0, ocf=150.0,
                          filing_date="2025-04-30")
    gaap_tbl, _ = _build_cf_table([q1], max_filings=80)
    assert "FY2025Q1" in gaap_tbl.quarter_labels
    ni_idx = gaap_tbl.concepts.index("Net Income")
    assert gaap_tbl.values[ni_idx][0] == pytest.approx(100.0)


def test_build_cf_table_q2_ytd_subtracted_from_q1():
    """Q2 standalone = Q2 YTD − Q1."""
    q1_ni, q1_ocf = 100.0, 150.0
    q2_ni, q2_ocf = 130.0, 180.0
    q2_ytd_ni  = q1_ni  + q2_ni
    q2_ytd_ocf = q1_ocf + q2_ocf

    q1 = _make_cf_filing("2025-03-31 (Q1)", "2025-03-31 (Q1)", q1_ni, q1_ocf, "2025-04-30")
    q2 = _make_cf_filing("2025-06-30 (Q2)", "2025-06-30 (YTD)", q2_ytd_ni, q2_ytd_ocf, "2025-07-30")

    gaap_tbl, _ = _build_cf_table([q2, q1], max_filings=80)  # newest-first order

    assert "FY2025Q1" in gaap_tbl.quarter_labels
    assert "FY2025Q2" in gaap_tbl.quarter_labels
    ni_idx = gaap_tbl.concepts.index("Net Income")
    q2_col = gaap_tbl.quarter_labels.index("FY2025Q2")
    assert gaap_tbl.values[ni_idx][q2_col] == pytest.approx(q2_ni)


# ── Task 4: Revenue fallback expansion ────────────────────────────────────────

def _make_is_df_revenues_only(period_col="2024-03-31 (Q1)", val=1000.0):
    """IS df where Revenue uses us-gaap_Revenues (not RevenueFromContractWithCustomer)."""
    return pd.DataFrame({
        "concept":               ["us-gaap_Revenues", "us-gaap_NetIncomeLoss"],
        "label":                 ["Revenues",         "Net income"],
        "standard_concept":      ["TotalRevenues",    "NetIncome"],   # not "Revenue"
        "abstract":              [False, False],
        "is_breakdown":          [False, False],
        "level":                 [3, 3],
        "dimension_member_label":[None, None],
        period_col:              [val, val * 0.1],
        "2023-03-31 (Q1)":       [val * 0.9, val * 0.08],
    })


def _make_filing_revenues_only(**kwargs):
    df = _make_is_df_revenues_only(**kwargs)
    mock_stmt = MagicMock(); mock_stmt.to_dataframe.return_value = df
    mock_fin = MagicMock()
    mock_fin.income_statement.return_value = mock_stmt
    mock_fin.cashflow_statement.return_value = mock_stmt
    mock_tenq = MagicMock(); mock_tenq.financials = mock_fin
    filing = MagicMock(); filing.obj.return_value = mock_tenq; filing.filing_date = "2024-04-30"
    return filing


def test_revenue_fallback_picks_us_gaap_revenues():
    """Revenue must resolve when XBRL uses us-gaap_Revenues (concept ends with _Revenues)."""
    filing = _make_filing_revenues_only(val=2000.0)
    gaap_tbl, _ = _build_is_table([filing], max_filings=1)
    rev_idx = gaap_tbl.concepts.index("Revenue")
    assert gaap_tbl.values[rev_idx][0] == pytest.approx(2000.0), (
        f"Expected 2000.0, got {gaap_tbl.values[rev_idx][0]}"
    )


def test_revenue_fallback_picks_sales_revenue_net():
    """Revenue must resolve when XBRL uses us-gaap_SalesRevenueNet."""
    df = pd.DataFrame({
        "concept":               ["us-gaap_SalesRevenueNet", "us-gaap_NetIncomeLoss"],
        "label":                 ["Net sales",               "Net income"],
        "standard_concept":      ["SalesRevenueNet",          "NetIncome"],
        "abstract":              [False, False],
        "is_breakdown":          [False, False],
        "level":                 [3, 3],
        "dimension_member_label":[None, None],
        "2024-03-31 (Q1)":       [3000.0, 300.0],
        "2023-03-31 (Q1)":       [2700.0, 270.0],
    })
    mock_stmt = MagicMock(); mock_stmt.to_dataframe.return_value = df
    mock_fin = MagicMock()
    mock_fin.income_statement.return_value = mock_stmt
    mock_fin.cashflow_statement.return_value = mock_stmt
    mock_tenq = MagicMock(); mock_tenq.financials = mock_fin
    filing = MagicMock(); filing.obj.return_value = mock_tenq; filing.filing_date = "2024-04-30"

    gaap_tbl, _ = _build_is_table([filing], max_filings=1)
    rev_idx = gaap_tbl.concepts.index("Revenue")
    assert gaap_tbl.values[rev_idx][0] == pytest.approx(3000.0)


def test_build_cf_table_q3_ytd_subtracted_from_q2_ytd():
    """Q3 standalone = Q3 YTD − Q2 YTD."""
    q1_ni, q2_ni, q3_ni = 100.0, 130.0, 150.0
    q2_ytd_ni = q1_ni + q2_ni
    q3_ytd_ni = q1_ni + q2_ni + q3_ni

    q1 = _make_cf_filing("2025-03-31 (Q1)", "2025-03-31 (Q1)", q1_ni, q1_ni * 1.5, "2025-04-30")
    q2 = _make_cf_filing("2025-06-30 (Q2)", "2025-06-30 (YTD)", q2_ytd_ni, q2_ytd_ni * 1.5, "2025-07-30")
    q3 = _make_cf_filing("2025-09-30 (Q3)", "2025-09-30 (YTD)", q3_ytd_ni, q3_ytd_ni * 1.5, "2025-10-30")

    gaap_tbl, _ = _build_cf_table([q3, q2, q1], max_filings=80)

    ni_idx = gaap_tbl.concepts.index("Net Income")
    q3_col = gaap_tbl.quarter_labels.index("FY2025Q3")
    assert gaap_tbl.values[ni_idx][q3_col] == pytest.approx(q3_ni)


def test_build_cf_table_q2_ytd_without_q1_keeps_raw():
    """When Q1 is absent, Q2 YTD value is kept as-is (best-effort)."""
    ytd_ni = 230.0
    q2 = _make_cf_filing("2025-06-30 (Q2)", "2025-06-30 (YTD)", ytd_ni, ytd_ni * 1.5, "2025-07-30")

    gaap_tbl, _ = _build_cf_table([q2], max_filings=80)

    assert "FY2025Q2" in gaap_tbl.quarter_labels
    ni_idx = gaap_tbl.concepts.index("Net Income")
    q2_col = gaap_tbl.quarter_labels.index("FY2025Q2")
    assert gaap_tbl.values[ni_idx][q2_col] == pytest.approx(ytd_ni)


# ── Override integration tests ────────────────────────────────────────────────

def _make_filing_odd_concepts(period_col="2025-12-27 (Q1)", val=100.0, filing_date="2026-01-30"):
    """Filing where Revenue uses 'TotalRevenues' std_concept (not in IS_TEMPLATE priority 1)."""
    df = pd.DataFrame({
        "concept":               ["us-gaap_TotalRevenues", "us-gaap_GrossProfit", "us-gaap_NetIncomeLoss"],
        "label":                 ["Total revenues", "Gross margin", "Net income"],
        "standard_concept":      ["TotalRevenues", "GrossProfit", "NetIncome"],
        "abstract":              [False, False, False],
        "is_breakdown":          [False, False, False],
        "level":                 [4, 3, 3],
        "dimension_member_label":[None, None, None],
        period_col:              [val * 10, val * 7, val * 2],
        "2024-12-28 (Q1)":       [val * 9, val * 6, val * 1.5],
    })
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


def test_build_is_table_revenue_none_without_override():
    """Revenue is None when std_concept doesn't match and no fallback."""
    filing = _make_filing_odd_concepts(val=100.0)
    gaap_tbl, _ = _build_is_table([filing], max_filings=1)
    revenue_idx = gaap_tbl.concepts.index("Revenue")
    assert gaap_tbl.values[revenue_idx][0] is None


def test_build_is_table_concept_override_restores_revenue():
    """concept_override makes Revenue resolve via the override's std_concept."""
    filing = _make_filing_odd_concepts(val=100.0)
    overrides = {"Revenue": {"fix_type": "concept_override", "std_concept": "TotalRevenues"}}
    gaap_tbl, _ = _build_is_table([filing], max_filings=1, is_overrides=overrides)
    revenue_idx = gaap_tbl.concepts.index("Revenue")
    assert gaap_tbl.values[revenue_idx][0] == pytest.approx(1000.0)  # val * 10


def test_build_is_table_structural_absence_keeps_none():
    """structural_absence override skips lookup and keeps None."""
    filing = _make_filing_odd_concepts(val=100.0)
    overrides = {"Revenue": {"fix_type": "structural_absence", "confirmed_absent": True}}
    gaap_tbl, _ = _build_is_table([filing], max_filings=1, is_overrides=overrides)
    revenue_idx = gaap_tbl.concepts.index("Revenue")
    assert gaap_tbl.values[revenue_idx][0] is None


def test_build_is_table_override_applies_to_all_filings():
    """Override is applied to every filing in the loop, not just the first."""
    f1 = _make_filing_odd_concepts("2025-12-27 (Q1)", val=100.0, filing_date="2026-01-30")
    f2 = _make_filing_odd_concepts("2024-12-28 (Q1)", val=90.0, filing_date="2025-01-30")
    overrides = {"Revenue": {"fix_type": "concept_override", "std_concept": "TotalRevenues"}}
    gaap_tbl, _ = _build_is_table([f1, f2], max_filings=2, is_overrides=overrides)
    revenue_idx = gaap_tbl.concepts.index("Revenue")
    assert len(gaap_tbl.quarter_labels) == 2
    assert all(v is not None for v in gaap_tbl.values[revenue_idx])


# ── CF overflow YTD subtraction unit tests ────────────────────────────────────

def _make_cf_df_with_overflow(period_col, ni_val, ocf_val, overflow_val):
    """CF DataFrame with Net Income, OCF (template rows) and one overflow row."""
    return pd.DataFrame({
        "concept":               [
            "us-gaap_NetIncomeLoss",
            "us-gaap_NetCashProvidedByUsedInOperatingActivities",
            "us-gaap_SpecialItemCashFlow",    # overflow: not in template
        ],
        "label":                 ["Net income", "Net cash from operations", "Special item"],
        "standard_concept":      ["NetIncome",  "NetCashFromOperatingActivities", None],
        "abstract":              [False, False, False],
        "is_breakdown":          [False, False, False],
        "level":                 [3, 3, 4],
        "dimension_member_label":[None, None, None],
        period_col:              [ni_val, ocf_val, overflow_val],
    })


def _make_cf_filing_with_overflow(is_period_col, cf_period_col, ni, ocf, overflow, filing_date):
    is_df  = _make_is_df_minimal(is_period_col)
    cf_df  = _make_cf_df_with_overflow(cf_period_col, ni, ocf, overflow)
    mock_is = MagicMock(); mock_is.to_dataframe.return_value = is_df
    mock_cf = MagicMock(); mock_cf.to_dataframe.return_value = cf_df
    mock_fin = MagicMock()
    mock_fin.income_statement.return_value = mock_is
    mock_fin.cashflow_statement.return_value = mock_cf
    mock_tenq = MagicMock(); mock_tenq.financials = mock_fin
    mock_filing = MagicMock()
    mock_filing.obj.return_value = mock_tenq
    mock_filing.filing_date = filing_date
    return mock_filing


def test_cf_overflow_q1_standalone_captured():
    """Overflow concept from Q1 standalone filing is included in GAAP overflow."""
    q1 = _make_cf_filing_with_overflow(
        "2025-03-31 (Q1)", "2025-03-31 (Q1)",
        ni=100.0, ocf=150.0, overflow=20.0, filing_date="2025-04-30",
    )
    gaap_tbl, _ = _build_cf_table([q1], max_filings=80)
    assert "us-gaap_SpecialItemCashFlow" in gaap_tbl.labels
    idx = gaap_tbl.labels.index("us-gaap_SpecialItemCashFlow")
    q1_col = gaap_tbl.quarter_labels.index("FY2025Q1")
    assert gaap_tbl.values[idx][q1_col] == pytest.approx(20.0)


def test_cf_overflow_q2_ytd_subtracted():
    """Overflow Q2 standalone = Q2_YTD_overflow − Q1_overflow."""
    q1_ov, q2_ov = 20.0, 35.0         # standalone quarterly values
    q2_ytd_ov    = q1_ov + q2_ov      # = 55.0

    q1 = _make_cf_filing_with_overflow(
        "2025-03-31 (Q1)", "2025-03-31 (Q1)",
        ni=100.0, ocf=150.0, overflow=q1_ov, filing_date="2025-04-30",
    )
    q2 = _make_cf_filing_with_overflow(
        "2025-06-30 (Q2)", "2025-06-30 (YTD)",
        ni=230.0, ocf=330.0, overflow=q2_ytd_ov, filing_date="2025-07-30",
    )
    gaap_tbl, _ = _build_cf_table([q2, q1], max_filings=80)  # newest-first

    assert "us-gaap_SpecialItemCashFlow" in gaap_tbl.labels
    idx = gaap_tbl.labels.index("us-gaap_SpecialItemCashFlow")
    q2_col = gaap_tbl.quarter_labels.index("FY2025Q2")
    assert gaap_tbl.values[idx][q2_col] == pytest.approx(q2_ov)


def test_cf_overflow_q3_ytd_subtracted():
    """Overflow Q3 standalone = Q3_YTD_overflow − Q2_YTD_overflow."""
    q1_ov, q2_ov, q3_ov = 20.0, 35.0, 40.0
    q2_ytd_ov = q1_ov + q2_ov
    q3_ytd_ov = q1_ov + q2_ov + q3_ov

    q1 = _make_cf_filing_with_overflow(
        "2025-03-31 (Q1)", "2025-03-31 (Q1)",
        ni=100.0, ocf=150.0, overflow=q1_ov, filing_date="2025-04-30",
    )
    q2 = _make_cf_filing_with_overflow(
        "2025-06-30 (Q2)", "2025-06-30 (YTD)",
        ni=230.0, ocf=330.0, overflow=q2_ytd_ov, filing_date="2025-07-30",
    )
    q3 = _make_cf_filing_with_overflow(
        "2025-09-30 (Q3)", "2025-09-30 (YTD)",
        ni=380.0, ocf=480.0, overflow=q3_ytd_ov, filing_date="2025-10-30",
    )
    gaap_tbl, _ = _build_cf_table([q3, q2, q1], max_filings=80)

    idx = gaap_tbl.labels.index("us-gaap_SpecialItemCashFlow")
    q3_col = gaap_tbl.quarter_labels.index("FY2025Q3")
    assert gaap_tbl.values[idx][q3_col] == pytest.approx(q3_ov)


def test_cf_overflow_q2_without_q1_is_none():
    """When Q1 overflow is absent (None), Q2 YTD can't be subtracted → no entry."""
    q2_ytd_ov = 55.0

    q1_no_ov = _make_cf_filing(
        "2025-03-31 (Q1)", "2025-03-31 (Q1)",
        ni=100.0, ocf=150.0, filing_date="2025-04-30",
    )  # no overflow row in this filing
    q2 = _make_cf_filing_with_overflow(
        "2025-06-30 (Q2)", "2025-06-30 (YTD)",
        ni=230.0, ocf=330.0, overflow=q2_ytd_ov, filing_date="2025-07-30",
    )
    gaap_tbl, _ = _build_cf_table([q2, q1_no_ov], max_filings=80)

    # Concept may not appear at all, or Q2 value should be None
    if "us-gaap_SpecialItemCashFlow" in gaap_tbl.labels:
        idx = gaap_tbl.labels.index("us-gaap_SpecialItemCashFlow")
        q2_col = gaap_tbl.quarter_labels.index("FY2025Q2")
        assert gaap_tbl.values[idx][q2_col] is None


# ── Task 1: pre-XBRL early exit ───────────────────────────────────────────────

from datetime import date as _date

def _make_old_filing(filing_date_str: str):
    """Mock filing with given date but no XBRL (pre-2008)."""
    f = MagicMock()
    f.filing_date = _date.fromisoformat(filing_date_str)
    f.obj.side_effect = Exception("No XBRL")
    return f


def test_build_is_table_stops_before_pre_xbrl():
    """Filings dated before 2008-01-01 should cause early loop termination, not exception."""
    modern = _make_filing(period_col="2024-03-31 (Q1)", val=100.0, filing_date="2024-04-30")
    modern.filing_date = _date(2024, 4, 30)
    old = _make_old_filing("2007-04-30")

    gaap_tbl, _ = _build_is_table([modern, old], max_filings=80)
    assert len(gaap_tbl.quarter_labels) == 1   # only the modern one
    old.obj.assert_not_called()  # break must have fired, not the exception handler


def test_build_bs_table_stops_before_pre_xbrl():
    modern = _make_filing(period_col="2024-03-31 (Q1)", val=100.0, filing_date="2024-04-30")
    modern.filing_date = _date(2024, 4, 30)
    old = _make_old_filing("2007-04-30")
    gaap_tbl, _ = _build_bs_table([modern, old], max_filings=80)
    assert len(gaap_tbl.quarter_labels) == 1
    old.obj.assert_not_called()


def test_build_cf_table_stops_before_pre_xbrl():
    modern = _make_cf_filing("2024-03-31 (Q1)", "2024-03-31 (Q1)", 100.0, 150.0, "2024-04-30")
    modern.filing_date = _date(2024, 4, 30)
    old = _make_old_filing("2007-04-30")
    gaap_tbl, _ = _build_cf_table([modern, old], max_filings=80)
    assert len(gaap_tbl.quarter_labels) == 1
    old.obj.assert_not_called()


# ── Task 2: Dividends std_concept bug ─────────────────────────────────────────

def _make_cf_dividends_df():
    """CF df with NCI distribution row AND a real dividends row.

    NCI row label contains 'dividend' to ensure label_hint doesn't rescue the old
    std_concept=DistributionsToMinorityInterests bug (label_hint would allow the NCI
    row through priority-1 matching, picking 30.0 instead of 80.0).
    """
    return pd.DataFrame({
        "concept":               [
            "us-gaap_NetCashProvidedByUsedInOperatingActivities",
            "us-gaap_DistributionsToMinorityInterests",   # NCI — must NOT be picked
            "us-gaap_PaymentsOfDividendsCommonStock",     # real dividends — must be picked
        ],
        "label":                 ["Net cash from ops", "Dividend distributions to NCI", "Dividends paid"],
        "standard_concept":      ["NetCashFromOperatingActivities", "DistributionsToMinorityInterests", None],
        "abstract":              [False, False, False],
        "is_breakdown":          [False, False, False],
        "level":                 [3, 4, 4],
        "dimension_member_label":[None, None, None],
        "2024-03-31 (Q1)":       [500.0, 30.0, 80.0],
    })


def test_dividends_paid_does_not_pick_nci_distribution():
    """Dividends Paid must pick PaymentsOfDividends, not DistributionsToMinorityInterests."""
    df = _make_cf_dividends_df()
    mock_is = MagicMock(); mock_is.to_dataframe.return_value = _make_is_df_minimal("2024-03-31 (Q1)")
    mock_cf = MagicMock(); mock_cf.to_dataframe.return_value = df
    mock_fin = MagicMock()
    mock_fin.income_statement.return_value = mock_is
    mock_fin.cashflow_statement.return_value = mock_cf
    mock_tenq = MagicMock(); mock_tenq.financials = mock_fin
    filing = MagicMock(); filing.obj.return_value = mock_tenq; filing.filing_date = "2024-04-30"

    gaap_tbl, _ = _build_cf_table([filing], max_filings=1)
    div_idx = gaap_tbl.concepts.index("Dividends Paid")
    assert gaap_tbl.values[div_idx][0] == pytest.approx(80.0), (
        f"Expected 80.0 (real dividends), got {gaap_tbl.values[div_idx][0]}"
    )


# ── Task 3: Net Income fallback chain ─────────────────────────────────────────

def test_build_is_table_prefers_attributable_to_parent_over_profitloss():
    """NetIncomeLossAttributableToParent should be picked before ProfitLoss.

    Concept name deliberately avoids 'NetIncomeLoss' substring so the IS_TEMPLATE
    fallback_suffix doesn't match it — only the post-processing 2a block (which
    matches by standard_concept) can pick it up.
    """
    df = pd.DataFrame({
        "concept":               ["us-gaap_ProfitLoss",   "us-gaap_ParentCompanyNetIncome"],
        "label":                 ["Net income incl. NCI", "Net income attributable to common"],
        "standard_concept":      ["ProfitLoss",           "NetIncomeLossAttributableToParent"],
        "abstract":              [False,                                        False],
        "is_breakdown":          [False,                                        False],
        "level":                 [3,                                            3],
        "dimension_member_label":[None,                                         None],
        "2024-03-31 (Q1)":       [300.0,                                        280.0],   # 300 = 280 parent + 20 NCI
        "2023-03-31 (Q1)":       [270.0,                                        255.0],
    })
    mock_stmt = MagicMock(); mock_stmt.to_dataframe.return_value = df
    mock_fin = MagicMock()
    mock_fin.income_statement.return_value = mock_stmt
    mock_fin.cashflow_statement.return_value = mock_stmt
    mock_tenq = MagicMock(); mock_tenq.financials = mock_fin
    filing = MagicMock(); filing.obj.return_value = mock_tenq; filing.filing_date = "2024-04-30"

    gaap_tbl, _ = _build_is_table([filing], max_filings=1)
    ni_idx = gaap_tbl.concepts.index("Net Income")
    assert gaap_tbl.values[ni_idx][0] == pytest.approx(280.0), (
        f"Expected 280.0 (parent-only), got {gaap_tbl.values[ni_idx][0]}"
    )


# ── Task 5: Total Non-op derived guard ────────────────────────────────────────

def test_total_nonop_not_derived_when_discontinued_ops_present():
    """When discontinued operations exist, Total Non-op should NOT be derived (return None)."""
    df = pd.DataFrame({
        "concept":               [
            "us-gaap_OperatingIncomeLoss",
            "us-gaap_IncomeLossFromDiscontinuedOperationsNetOfTax",
            "us-gaap_IncomeLossFromContinuingOperationsBeforeIncomeTax",
        ],
        "label":                 ["Operating income", "Discont. ops income", "Income before taxes"],
        "standard_concept":      ["OperatingIncomeLoss", None, "PretaxIncomeLoss"],
        "abstract":              [False, False, False],
        "is_breakdown":          [False, False, False],
        "level":                 [3, 3, 3],
        "dimension_member_label":[None, None, None],
        "2024-03-31 (Q1)":       [100.0, 20.0, 120.0],   # Non-op = 0, Discont = 20
        "2023-03-31 (Q1)":       [90.0, 15.0, 105.0],
    })
    mock_stmt = MagicMock(); mock_stmt.to_dataframe.return_value = df
    mock_fin = MagicMock()
    mock_fin.income_statement.return_value = mock_stmt
    mock_fin.cashflow_statement.return_value = mock_stmt
    mock_tenq = MagicMock(); mock_tenq.financials = mock_fin
    filing = MagicMock(); filing.obj.return_value = mock_tenq; filing.filing_date = "2024-04-30"

    gaap_tbl, _ = _build_is_table([filing], max_filings=1)
    nonop_idx = gaap_tbl.concepts.index("Total Non-op Income/(Loss)")
    assert gaap_tbl.values[nonop_idx][0] is None, (
        f"Expected None (discontinued ops present), got {gaap_tbl.values[nonop_idx][0]}"
    )


def test_total_nonop_derived_normally_without_discontinued_ops():
    """When no discontinued ops, Total Non-op should still be derived as Pretax - Operating."""
    df = pd.DataFrame({
        "concept":               ["us-gaap_OperatingIncomeLoss", "us-gaap_IncomeLossFromContinuingOperationsBeforeIncomeTax"],
        "label":                 ["Operating income",            "Income before taxes"],
        "standard_concept":      ["OperatingIncomeLoss",         "PretaxIncomeLoss"],
        "abstract":              [False, False],
        "is_breakdown":          [False, False],
        "level":                 [3, 3],
        "dimension_member_label":[None, None],
        "2024-03-31 (Q1)":       [100.0, 115.0],
        "2023-03-31 (Q1)":       [90.0, 103.0],
    })
    mock_stmt = MagicMock(); mock_stmt.to_dataframe.return_value = df
    mock_fin = MagicMock()
    mock_fin.income_statement.return_value = mock_stmt
    mock_fin.cashflow_statement.return_value = mock_stmt
    mock_tenq = MagicMock(); mock_tenq.financials = mock_fin
    filing = MagicMock(); filing.obj.return_value = mock_tenq; filing.filing_date = "2024-04-30"

    gaap_tbl, _ = _build_is_table([filing], max_filings=1)
    nonop_idx = gaap_tbl.concepts.index("Total Non-op Income/(Loss)")
    assert gaap_tbl.values[nonop_idx][0] == pytest.approx(15.0)  # 115 - 100
