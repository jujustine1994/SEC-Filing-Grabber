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
    preview_sheets,
    StatementTable,
    _col_to_quarter_label,
    _current_q_col,
    _match_is_row,
    _build_is_table,
    _build_bs_table,
    _build_cf_table,
    _merge_financials,
    _synthesize_q4,
    _ytd_col,
    _prev_quarter_label,
    _is_nongaap_label,
    _collect_overflow,
    _filter_filings_by_year,
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

def test_build_is_table_has_all_template_concept_rows():
    # Mock IS df has only 3 rows; all 3 are consumed by template → no overflow
    filing = _make_filing()
    gaap_tbl, _ = _build_is_table([filing], max_filings=1)
    from fetcher_gaap import IS_TEMPLATE
    # 2026-08-03 依 49 家實測新增 Total Costs and Expenses、Net Income incl. NCI
    assert len(gaap_tbl.concepts) == len(IS_TEMPLATE) == 24

def test_build_is_table_quarter_labels_format():
    # fy_end_month 預設 12，期末日 2025-12-27 就是 Q4——季編號由日期反推，
    # 不採信欄名裡的 (Q1)（見 _col_to_quarter_label 的說明）
    filing = _make_filing(period_col="2025-12-27 (Q1)")
    gaap_tbl, _ = _build_is_table([filing], max_filings=1)
    assert gaap_tbl.quarter_labels == ["FY2025Q4"]

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
    assert gaap_tbl.quarter_labels[0] == "FY2024Q4"
    assert gaap_tbl.quarter_labels[1] == "FY2025Q4"

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
    # 數「真的抓到的期間」——G6（2026-08-25）之後欄位清單還會補上缺口欄
    # （這份 mock 的三筆 filing 相隔一年，中間的季度全部補成空白欄），
    # max_filings 限制的是抓幾份 filing，不是最後有幾欄
    assert len([e for e in fin_tbl.period_ends if e]) <= 3


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


# ── _synthesize_q4 ────────────────────────────────────────────────────────────

def test_synthesize_q4_flow_derives_annual_minus_q1q2q3():
    """IS/CF (flow) rows: Q4 = annual − Q1 − Q2 − Q3."""
    q_tbl = StatementTable(
        sheet_name="Data_IS",
        quarter_labels=["FY2024Q1", "FY2024Q2", "FY2024Q3"],
        filing_dates=["2024-05-01", "2024-08-01", "2024-11-01"],
        concepts=["Revenue"], values=[[100.0, 110.0, 120.0]],
    )
    ann_tbl = StatementTable(
        sheet_name="Data_IS", quarter_labels=["FY2024"], filing_dates=["2025-02-01"],
        concepts=["Revenue"], values=[[500.0]],
    )
    result = _synthesize_q4(q_tbl, ann_tbl, n_template_rows=1, is_balance=False)
    assert "FY2024Q4" in result.quarter_labels
    idx = result.quarter_labels.index("FY2024Q4")
    assert result.values[0][idx] == 500.0 - 100.0 - 110.0 - 120.0


def test_synthesize_q4_balance_uses_annual_value_directly():
    """BS (point-in-time) rows: Q4 = annual value as-is, no subtraction."""
    q_tbl = StatementTable(
        sheet_name="Data_BS", quarter_labels=["FY2024Q1"], filing_dates=["2024-05-01"],
        concepts=["Total Assets"], values=[[1000.0]],
    )
    ann_tbl = StatementTable(
        sheet_name="Data_BS", quarter_labels=["FY2024"], filing_dates=["2025-02-01"],
        concepts=["Total Assets"], values=[[5000.0]],
    )
    result = _synthesize_q4(q_tbl, ann_tbl, n_template_rows=1, is_balance=True)
    idx = result.quarter_labels.index("FY2024Q4")
    assert result.values[0][idx] == 5000.0


def test_synthesize_q4_skipped_when_quarters_incomplete():
    """Missing Q3 (or any of Q1-Q3) means Q4 can't be derived for flow statements."""
    q_tbl = StatementTable(
        sheet_name="Data_IS", quarter_labels=["FY2024Q1", "FY2024Q2"],
        filing_dates=["2024-05-01", "2024-08-01"],
        concepts=["Revenue"], values=[[100.0, 110.0]],
    )
    ann_tbl = StatementTable(
        sheet_name="Data_IS", quarter_labels=["FY2024"], filing_dates=["2025-02-01"],
        concepts=["Revenue"], values=[[500.0]],
    )
    result = _synthesize_q4(q_tbl, ann_tbl, n_template_rows=1, is_balance=False)
    assert "FY2024Q4" not in result.quarter_labels


def test_synthesize_q4_does_not_overwrite_existing_q4():
    """If a Q4 column somehow already exists, leave it untouched."""
    q_tbl = StatementTable(
        sheet_name="Data_IS", quarter_labels=["FY2024Q4"], filing_dates=["2025-01-15"],
        concepts=["Revenue"], values=[[999.0]],
    )
    ann_tbl = StatementTable(
        sheet_name="Data_IS", quarter_labels=["FY2024"], filing_dates=["2025-02-01"],
        concepts=["Revenue"], values=[[500.0]],
    )
    result = _synthesize_q4(q_tbl, ann_tbl, n_template_rows=1, is_balance=True)
    idx = result.quarter_labels.index("FY2024Q4")
    assert result.values[0][idx] == 999.0
    assert result.quarter_labels.count("FY2024Q4") == 1


def test_synthesize_q4_no_annual_data_returns_unchanged():
    q_tbl = StatementTable(
        sheet_name="Data_IS", quarter_labels=["FY2024Q1"], filing_dates=["2024-05-01"],
        concepts=["Revenue"], values=[[100.0]],
    )
    ann_tbl = StatementTable(
        sheet_name="Data_IS", quarter_labels=[], filing_dates=[],
        concepts=[], values=[],
    )
    result = _synthesize_q4(q_tbl, ann_tbl, n_template_rows=1, is_balance=False)
    assert result.quarter_labels == ["FY2024Q1"]


def test_synthesize_q4_overflow_rows_get_none():
    """Rows beyond n_template_rows (overflow) are extended with None, not derived."""
    q_tbl = StatementTable(
        sheet_name="Data_IS", quarter_labels=["FY2024Q1", "FY2024Q2", "FY2024Q3"],
        filing_dates=["2024-05-01", "2024-08-01", "2024-11-01"],
        concepts=["Revenue", "SomeOverflowConcept"],
        values=[[100.0, 110.0, 120.0], [1.0, 2.0, 3.0]],
    )
    ann_tbl = StatementTable(
        sheet_name="Data_IS", quarter_labels=["FY2024"], filing_dates=["2025-02-01"],
        concepts=["Revenue"], values=[[500.0]],
    )
    result = _synthesize_q4(q_tbl, ann_tbl, n_template_rows=1, is_balance=False)
    idx = result.quarter_labels.index("FY2024Q4")
    assert result.values[1][idx] is None


@patch("fetcher_gaap.Company")
@patch("fetcher_gaap.set_identity")
@patch("fetcher_gaap.load_overrides", return_value={})
def test_fetch_gaap_synthesizes_q4_in_quarterly_sheet(mock_ov, mock_id, mock_co):
    """Data_Financials(Q) should gain a Q4 column derived from the 10-K annual filing."""
    q1 = _make_filing(period_col="2025-03-29 (Q1)", val=100.0, filing_date="2025-04-30")
    q2 = _make_filing(period_col="2025-06-28 (Q2)", val=110.0, filing_date="2025-07-30")
    q3 = _make_filing(period_col="2025-09-27 (Q3)", val=120.0, filing_date="2025-10-30")
    k = _make_filing(period_col="2025-12-27 (FY)", val=500.0, filing_date="2026-02-01")
    mock_co.return_value = _make_mock_company_fgs(q_filings=[q1, q2, q3], k_filings=[k])

    tables = fetch_gaap_statements("TEST", "Test test@test.com")
    fin_tbl = next(t for t in tables if t.sheet_name == "Data_Financials(Q)")

    assert "FY2025Q4" in fin_tbl.quarter_labels
    idx = fin_tbl.quarter_labels.index("FY2025Q4")
    rev_row = fin_tbl.concepts.index("Revenue")
    assert fin_tbl.values[rev_row][idx] == 500.0 * 10 - 100.0 * 10 - 110.0 * 10 - 120.0 * 10


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


# ── Task 6: Investment Proceeds multi-sum ─────────────────────────────────────

def _make_cf_df_investment_proceeds():
    """CF df with multiple investment proceeds lines (no single sum row)."""
    return pd.DataFrame({
        "concept":               [
            "us-gaap_NetCashProvidedByUsedInOperatingActivities",
            "us-gaap_ProceedsFromSaleOfAvailableForSaleSecurities",
            "us-gaap_ProceedsFromMaturitiesPrepaymentsAndCallsOfAvailableForSaleSecurities",
            "us-gaap_ProceedsFromSaleOfShortTermInvestments",
        ],
        "label":                 [
            "Net cash from ops",
            "Proceeds from sale of AFS securities",
            "Proceeds from maturities of AFS securities",
            "Proceeds from sale of ST investments",
        ],
        "standard_concept":      ["NetCashFromOperatingActivities", None, None, None],
        "abstract":              [False, False, False, False],
        "is_breakdown":          [False, False, False, False],
        "level":                 [3, 4, 4, 4],
        "dimension_member_label":[None, None, None, None],
        "2024-03-31 (Q1)":       [500.0, 200.0, 150.0, 100.0],
    })


def test_investment_proceeds_sums_multiple_concepts():
    """Investment Proceeds must sum all ProceedsFrom*AFS*|ShortTerm* rows."""
    is_df = _make_is_df_minimal("2024-03-31 (Q1)")
    cf_df = _make_cf_df_investment_proceeds()
    mock_is = MagicMock(); mock_is.to_dataframe.return_value = is_df
    mock_cf = MagicMock(); mock_cf.to_dataframe.return_value = cf_df
    mock_fin = MagicMock()
    mock_fin.income_statement.return_value = mock_is
    mock_fin.cashflow_statement.return_value = mock_cf
    mock_tenq = MagicMock(); mock_tenq.financials = mock_fin
    filing = MagicMock(); filing.obj.return_value = mock_tenq; filing.filing_date = "2024-04-30"

    gaap_tbl, _ = _build_cf_table([filing], max_filings=1)
    proc_idx = gaap_tbl.concepts.index("Investment Proceeds")
    # 200 + 150 + 100 = 450
    assert gaap_tbl.values[proc_idx][0] == pytest.approx(450.0), (
        f"Expected 450.0 (sum), got {gaap_tbl.values[proc_idx][0]}"
    )


# ── Task 7: Debt Proceeds/Repayments multi-sum ────────────────────────────────

def _make_cf_df_debt_lines():
    """CF df with separate LT and ST debt proceeds/repayments, no summary row."""
    return pd.DataFrame({
        "concept":               [
            "us-gaap_NetCashProvidedByUsedInOperatingActivities",
            "us-gaap_ProceedsFromIssuanceOfLongTermDebt",
            "us-gaap_ProceedsFromShortTermBorrowings",
            "us-gaap_RepaymentsOfLongTermDebt",
            "us-gaap_RepaymentsOfShortTermDebt",
        ],
        "label":                 [
            "Net cash from ops",
            "Proceeds from LT debt",
            "Proceeds from ST borrowings",
            "Repayments of LT debt",
            "Repayments of ST debt",
        ],
        "standard_concept":      ["NetCashFromOperatingActivities", None, None, None, None],
        "abstract":              [False, False, False, False, False],
        "is_breakdown":          [False, False, False, False, False],
        "level":                 [3, 4, 4, 4, 4],
        "dimension_member_label":[None, None, None, None, None],
        "2024-03-31 (Q1)":       [500.0, 1000.0, 200.0, 800.0, 100.0],
    })


def test_debt_proceeds_sums_lt_and_st():
    """Debt Proceeds must sum LT + ST debt issuance proceeds."""
    is_df = _make_is_df_minimal("2024-03-31 (Q1)")
    cf_df = _make_cf_df_debt_lines()
    mock_is = MagicMock(); mock_is.to_dataframe.return_value = is_df
    mock_cf = MagicMock(); mock_cf.to_dataframe.return_value = cf_df
    mock_fin = MagicMock()
    mock_fin.income_statement.return_value = mock_is
    mock_fin.cashflow_statement.return_value = mock_cf
    mock_tenq = MagicMock(); mock_tenq.financials = mock_fin
    filing = MagicMock(); filing.obj.return_value = mock_tenq; filing.filing_date = "2024-04-30"

    gaap_tbl, _ = _build_cf_table([filing], max_filings=1)
    proc_idx = gaap_tbl.concepts.index("Debt Proceeds")
    assert gaap_tbl.values[proc_idx][0] == pytest.approx(1200.0), (  # 1000 + 200
        f"Expected 1200.0, got {gaap_tbl.values[proc_idx][0]}"
    )


def test_debt_repayments_sums_lt_and_st():
    """Debt Repayments must sum LT + ST repayments."""
    is_df = _make_is_df_minimal("2024-03-31 (Q1)")
    cf_df = _make_cf_df_debt_lines()
    mock_is = MagicMock(); mock_is.to_dataframe.return_value = is_df
    mock_cf = MagicMock(); mock_cf.to_dataframe.return_value = cf_df
    mock_fin = MagicMock()
    mock_fin.income_statement.return_value = mock_is
    mock_fin.cashflow_statement.return_value = mock_cf
    mock_tenq = MagicMock(); mock_tenq.financials = mock_fin
    filing = MagicMock(); filing.obj.return_value = mock_tenq; filing.filing_date = "2024-04-30"

    gaap_tbl, _ = _build_cf_table([filing], max_filings=1)
    rep_idx = gaap_tbl.concepts.index("Debt Repayments")
    assert gaap_tbl.values[rep_idx][0] == pytest.approx(900.0), (  # 800 + 100
        f"Expected 900.0, got {gaap_tbl.values[rep_idx][0]}"
    )


# ── Task 8: FY label fiscal year alignment ────────────────────────────────────

def test_col_to_quarter_label_default_december_fy_unchanged():
    """Default fy_end_month=12：財年不進位，季編號跟日曆季一致。

    2023-12-30 這個期末日對 12 月結算的公司就是 Q4——舊版採信 edgartools 的
    `(Q1)` 標記所以回 FY2023Q1，現在一律由日期反推，回 FY2023Q4。
    """
    assert _col_to_quarter_label("2023-03-31 (Q1)") == "FY2023Q1"
    assert _col_to_quarter_label("2023-12-30 (Q1)") == "FY2023Q4"
    assert _col_to_quarter_label("2023-09-30 (FY)") == "FY2023"


def test_col_to_quarter_label_sep_fy_q1_increments_year():
    """AAPL-style Sep FY: Q1 ends Dec 2023 → label FY2024Q1 (company's FY2024)."""
    assert _col_to_quarter_label("2023-12-30 (Q1)", fy_end_month=9) == "FY2024Q1"


def test_col_to_quarter_label_sep_fy_q2_same_year():
    """AAPL Q2 ends Mar 2024 → label FY2024Q2 (same FY2024)."""
    assert _col_to_quarter_label("2024-03-30 (Q2)", fy_end_month=9) == "FY2024Q2"


def test_col_to_quarter_label_sep_fy_q3_same_year():
    """AAPL Q3 ends Jun 2024 → label FY2024Q3 (same FY2024)."""
    assert _col_to_quarter_label("2024-06-29 (Q3)", fy_end_month=9) == "FY2024Q3"


def test_col_to_quarter_label_sep_fy_annual_unchanged():
    """AAPL annual FY ends Sep 2024 → label FY2024 (unchanged, period is FY not Q)."""
    assert _col_to_quarter_label("2024-09-28 (FY)", fy_end_month=9) == "FY2024"


def test_col_to_quarter_label_jun_fy_q1_increments_year():
    """MSFT-style Jun FY: Q1 ends Sep 2024 → label FY2025Q1."""
    assert _col_to_quarter_label("2024-09-30 (Q1)", fy_end_month=6) == "FY2025Q1"


def test_col_to_quarter_label_jun_fy_q2_increments_year():
    """MSFT Q2 ends Dec 2024 → label FY2025Q2."""
    assert _col_to_quarter_label("2024-12-31 (Q2)", fy_end_month=6) == "FY2025Q2"


def test_col_to_quarter_label_jun_fy_q3_same_year():
    """MSFT Q3 ends Mar 2025 → label FY2025Q3."""
    assert _col_to_quarter_label("2025-03-31 (Q3)", fy_end_month=6) == "FY2025Q3"


def test_detect_fy_end_month_returns_9_for_sep_end():
    """_detect_fy_end_month should return 9 when the 10-K FY column ends in September."""
    from fetcher_gaap import _detect_fy_end_month
    df = pd.DataFrame({
        "concept": ["us-gaap_RevenueFromContractWithCustomer"],
        "label":   ["Revenue"],
        "standard_concept": ["Revenue"],
        "abstract": [False], "is_breakdown": [False], "level": [3],
        "dimension_member_label": [None],
        "2024-09-28 (FY)": [1000.0],
    })
    mock_stmt = MagicMock(); mock_stmt.to_dataframe.return_value = df
    mock_fin = MagicMock(); mock_fin.income_statement.return_value = mock_stmt
    mock_tenq = MagicMock(); mock_tenq.financials = mock_fin
    filing = MagicMock(); filing.obj.return_value = mock_tenq
    assert _detect_fy_end_month([filing]) == 9


def test_detect_fy_end_month_defaults_to_12_on_failure():
    """_detect_fy_end_month should return 12 when no filings given."""
    from fetcher_gaap import _detect_fy_end_month
    assert _detect_fy_end_month([]) == 12


# ── _filter_filings_by_year ───────────────────────────────────────────────────

def _make_dated_filing(year: int, month: int = 1, day: int = 15):
    """Mock filing with a date object as filing_date."""
    from datetime import date
    f = MagicMock()
    f.filing_date = date(year, month, day)
    return f


def test_filter_no_bounds_returns_all():
    filings = [_make_dated_filing(2020), _make_dated_filing(2022)]
    assert _filter_filings_by_year(filings, None, None) == filings


def test_filter_start_year_excludes_older():
    f2019 = _make_dated_filing(2019)
    f2021 = _make_dated_filing(2021)
    result = _filter_filings_by_year([f2019, f2021], start_year=2020, end_year=None)
    assert result == [f2021]


def test_filter_end_year_excludes_newer():
    f2021 = _make_dated_filing(2021)
    f2023 = _make_dated_filing(2023)
    result = _filter_filings_by_year([f2021, f2023], start_year=None, end_year=2022)
    assert result == [f2021]


def test_filter_both_bounds():
    filings = [_make_dated_filing(y) for y in [2018, 2020, 2022, 2024]]
    result = _filter_filings_by_year(filings, start_year=2019, end_year=2022)
    assert [f.filing_date.year for f in result] == [2020, 2022]


def test_filter_string_filing_date():
    """Mock filings that return filing_date as a string (unit-test pattern)."""
    f = MagicMock()
    f.filing_date = "2021-06-15"
    result = _filter_filings_by_year([f], start_year=2020, end_year=2022)
    assert result == [f]


def test_filter_string_filing_date_excluded():
    f = MagicMock()
    f.filing_date = "2019-03-01"
    result = _filter_filings_by_year([f], start_year=2020, end_year=None)
    assert result == []


def test_filter_empty_list():
    assert _filter_filings_by_year([], start_year=2020, end_year=2022) == []


# ── fetch_gaap_statements new params ─────────────────────────────────────────

def _make_mock_company_fgs(q_filings=None, k_filings=None):
    """Mock edgartools Company returning given filing lists (for fetch_gaap_statements tests)."""
    company = MagicMock()
    q_filings = q_filings or []
    k_filings = k_filings or []

    def get_filings(form, amendments=False):
        if form == "10-Q":
            return q_filings
        if form == "10-K":
            return k_filings
        return []

    company.get_filings.side_effect = get_filings
    company.name = "Test Corp"
    return company


def _make_k_filing(period_col="2024-12-28", val=100.0, filing_date="2025-02-01"):
    """Mock a 10-K filing (same structure as 10-Q mock)."""
    return _make_filing(period_col=period_col, val=val, filing_date=filing_date)


@patch("fetcher_gaap.Company")
@patch("fetcher_gaap.set_identity")
@patch("fetcher_gaap.load_overrides", return_value={})
def test_fetch_gaap_annual_only_no_q_table(mock_ov, mock_id, mock_co):
    """fetch_quarterly=False should produce Data_Financials(Y) but NOT Data_Financials(Q)."""
    k = _make_k_filing()
    mock_co.return_value = _make_mock_company_fgs(q_filings=[], k_filings=[k])

    tables = fetch_gaap_statements("TEST", "Test test@test.com",
                                   fetch_quarterly=False, fetch_annual=True)
    sheet_names = [t.sheet_name for t in tables]
    assert "Data_Financials(Q)" not in sheet_names
    assert "Data_Financials(Y)" in sheet_names


@patch("fetcher_gaap.Company")
@patch("fetcher_gaap.set_identity")
@patch("fetcher_gaap.load_overrides", return_value={})
def test_fetch_gaap_quarterly_only_no_y_table(mock_ov, mock_id, mock_co):
    """fetch_annual=False should produce Data_Financials(Q) but NOT Data_Financials(Y)."""
    q = _make_filing()
    mock_co.return_value = _make_mock_company_fgs(q_filings=[q], k_filings=[])

    tables = fetch_gaap_statements("TEST", "Test test@test.com",
                                   fetch_quarterly=True, fetch_annual=False)
    sheet_names = [t.sheet_name for t in tables]
    assert "Data_Financials(Q)" in sheet_names
    assert "Data_Financials(Y)" not in sheet_names


@patch("fetcher_gaap.Company")
@patch("fetcher_gaap.set_identity")
@patch("fetcher_gaap.load_overrides", return_value={})
def test_fetch_gaap_annual_only_raises_when_no_k(mock_ov, mock_id, mock_co):
    """fetch_quarterly=False with no 10-K filings should raise ValueError."""
    mock_co.return_value = _make_mock_company_fgs(q_filings=[], k_filings=[])
    with pytest.raises(ValueError, match="10-K"):
        fetch_gaap_statements("TEST", "Test test@test.com",
                               fetch_quarterly=False, fetch_annual=True)


@patch("fetcher_gaap.Company")
@patch("fetcher_gaap.set_identity")
@patch("fetcher_gaap.load_overrides", return_value={})
def test_fetch_gaap_excluded_sheets_removes_seg(mock_ov, mock_id, mock_co):
    """excluded_sheets should skip matching sheet names in the result."""
    q = _make_filing()
    mock_co.return_value = _make_mock_company_fgs(q_filings=[q], k_filings=[])

    tables = fetch_gaap_statements("TEST", "Test test@test.com",
                                   fetch_quarterly=True, fetch_annual=False,
                                   excluded_sheets={"Data_Seg_Revenue"})
    sheet_names = [t.sheet_name for t in tables]
    assert "Data_Seg_Revenue" not in sheet_names


# ── preview_sheets ────────────────────────────────────────────────────────────

@patch("fetcher_gaap.Company")
@patch("fetcher_gaap.set_identity")
def test_preview_sheets_fixed_always_present(mock_id, mock_co):
    """Fixed sheets should always appear regardless of what the company has."""
    q = _make_filing()
    company = _make_mock_company_fgs(q_filings=[q])
    mock_co.return_value = company

    result = preview_sheets("AAPL", "Test test@test.com")

    assert "Data_Financials(Q)" in result["sheets"]
    assert "Data_Financials(Y)" in result["sheets"]
    assert "Data_Meta" in result["sheets"]


@patch("fetcher_gaap.Company")
@patch("fetcher_gaap.set_identity")
def test_preview_sheets_no_q_filings(mock_id, mock_co):
    """When no 10-Q filings exist, only fixed sheets are returned."""
    company = _make_mock_company_fgs(q_filings=[])
    mock_co.return_value = company

    result = preview_sheets("NOFILINGS", "Test test@test.com")

    assert result["sheets"] == ["Data_Financials(Q)", "Data_Financials(Y)", "Data_Meta"]
    assert result["latest_label"] == ""
    assert result["latest_period_end"] == ""
    assert result["filing_date"] == ""


@patch("fetcher_gaap.Company")
@patch("fetcher_gaap.set_identity")
def test_preview_sheets_returns_list_of_strings(mock_id, mock_co):
    """result["sheets"] should be list[str]."""
    q = _make_filing()
    mock_co.return_value = _make_mock_company_fgs(q_filings=[q])

    result = preview_sheets("TEST", "Test test@test.com")

    assert isinstance(result["sheets"], list)
    assert all(isinstance(s, str) for s in result["sheets"])


@patch("fetcher_gaap.Company")
@patch("fetcher_gaap.set_identity")
def test_preview_sheets_computes_latest_quarter(mock_id, mock_co):
    """latest_label / latest_period_end / filing_date should come from the newest 10-Q."""
    q = _make_filing(filing_date="2026-02-01")
    q.period_of_report = "2025-12-27"
    company = _make_mock_company_fgs(q_filings=[q])
    company.fiscal_year_end = "0930"  # AAPL: FY 結束於 9 月
    mock_co.return_value = company

    result = preview_sheets("AAPL", "Test test@test.com")

    assert result["latest_label"] == "FY2026Q1"
    assert result["latest_period_end"] == "2025-12-27"
    assert result["filing_date"] == "2026-02-01"


# ── FY month probe when fetch_annual=False ────────────────────────────────────

@patch("fetcher_gaap.Company")
@patch("fetcher_gaap.set_identity")
@patch("fetcher_gaap.load_overrides", return_value={})
def test_fy_month_probed_when_quarterly_only(mock_ov, mock_id, mock_co):
    """When fetch_annual=False, one 10-K filing is still fetched for FY month detection."""
    q = _make_filing()
    company = _make_mock_company_fgs(q_filings=[q], k_filings=[])
    mock_co.return_value = company

    fetch_gaap_statements("TEST", "Test test@test.com",
                          fetch_quarterly=True, fetch_annual=False)

    all_forms = [c.kwargs.get("form") for c in company.get_filings.call_args_list]
    assert "10-K" in all_forms


# ── 期末流通股數（2026-08-01 新增）──────────────────────────────────────────
#
# IS 模板原有的 Basic/Diluted Shares 是**加權平均**股數（算 EPS 用）。
# 分析要看的「現在在外流通幾股」是期末時點值，兩者在有買回或增發的季度差很多。
# XBRL 概念：us-gaap:CommonStockSharesOutstanding（時點值，屬 BS）。

def test_bs_template_has_shares_outstanding():
    from fetcher_gaap import BS_TEMPLATE
    names = [row[0] for row in BS_TEMPLATE]
    assert "Shares Outstanding" in names


def test_shares_outstanding_maps_to_point_in_time_concept():
    from fetcher_gaap import BS_TEMPLATE
    row = next(r for r in BS_TEMPLATE if r[0] == "Shares Outstanding")
    _std, _concept, fallback, source = row[0], row[1], row[2], row[3]
    assert "CommonStockSharesOutstanding" in fallback
    assert source == "BS"


def test_shares_outstanding_is_not_weighted_average():
    """不可誤用加權平均的概念——那是 IS 的 Basic/Diluted Shares。"""
    from fetcher_gaap import BS_TEMPLATE
    row = next(r for r in BS_TEMPLATE if r[0] == "Shares Outstanding")
    assert "WeightedAverage" not in (row[2] or "")


def test_shares_outstanding_sits_in_equity_section():
    """位置要在權益段，不可跑到資產或負債段。"""
    from fetcher_gaap import BS_TEMPLATE
    names = [r[0] for r in BS_TEMPLATE]
    assert names.index("Total Liabilities") < names.index("Shares Outstanding")


# ═════════════════════════════════════════════════════════════════════════════
# 期末流通股數改走封面頁 dei fact（2026-08-02）
#
# 實測 ARLO/AAPL/NVDA/MSFT/COHR 五家都沒有在資產負債表 tag
# us-gaap:CommonStockSharesOutstanding，股數只寫在 CommonStockValue 的 label
# 文字裡。真正拿得到的是封面頁的 dei:EntityCommonStockSharesOutstanding，
# 走 Company.get_facts()，ARLO 有 32 筆、AAPL 70 筆，2009 年起逐季都有。
#
# ⚠ 這個 fact 的日期是**封面頁「最近可行日期」**，比財季結束晚幾週
#    （ARLO FY2025Q1 財季結束 2025-03-30，股數是 2025-05-02 的 103,400,957）。
#    它是能拿到的最接近的時點股數，但不是財季結束當天的數字。
# ═════════════════════════════════════════════════════════════════════════════

def test_shares_label_for_quarter():
    from fetcher_gaap import _shares_label
    assert _shares_label(2025, "Q1") == "FY2025Q1"


def test_shares_label_for_annual():
    """10-K 的 fiscal_period 是 FY，對到年報表的標籤。"""
    from fetcher_gaap import _shares_label
    assert _shares_label(2025, "FY") == "FY2025"


def test_shares_label_rejects_garbage():
    from fetcher_gaap import _shares_label
    assert _shares_label(2025, "") is None
    assert _shares_label(None, "Q1") is None


def test_shares_map_from_records():
    from fetcher_gaap import _shares_map_from_records
    recs = [
        {"fiscal_year": 2025, "fiscal_period": "Q1", "numeric_value": 103400957.0},
        {"fiscal_year": 2025, "fiscal_period": "Q2", "numeric_value": 104370654.0},
        {"fiscal_year": 2025, "fiscal_period": "FY", "numeric_value": 106855416.0},
    ]
    m = _shares_map_from_records(recs)
    assert m["FY2025Q1"] == 103400957.0
    assert m["FY2025"] == 106855416.0


def test_shares_map_keeps_latest_when_duplicated():
    """同一季重複申報（10-Q/A）時取最後一筆——後送的才是更正後的。"""
    from fetcher_gaap import _shares_map_from_records
    recs = [
        {"fiscal_year": 2025, "fiscal_period": "Q1", "numeric_value": 100.0},
        {"fiscal_year": 2025, "fiscal_period": "Q1", "numeric_value": 111.0},
    ]
    assert _shares_map_from_records(recs)["FY2025Q1"] == 111.0


def test_shares_map_skips_unusable_records():
    from fetcher_gaap import _shares_map_from_records
    recs = [
        {"fiscal_year": 2025, "fiscal_period": "Q1", "numeric_value": None},
        {"fiscal_year": None, "fiscal_period": "Q2", "numeric_value": 5.0},
        {"fiscal_year": 2025, "fiscal_period": "Q3", "numeric_value": 7.0},
    ]
    assert _shares_map_from_records(recs) == {"FY2025Q3": 7.0}


def test_apply_shares_fills_the_template_row():
    from fetcher_gaap import _apply_shares_outstanding, StatementTable
    tbl = StatementTable(
        sheet_name="Data_Financials(Q)",
        quarter_labels=["FY2025Q1", "FY2025Q2"],
        filing_dates=["", ""],
        concepts=["Revenue", "Shares Outstanding"],
        values=[[10.0, 20.0], [None, None]],
        ticker="ARLO", labels=["", ""],
    )
    _apply_shares_outstanding([tbl], {"FY2025Q1": 103.0, "FY2025Q2": 104.0})
    assert tbl.values[1] == [103.0, 104.0]


def test_apply_shares_leaves_other_rows_alone():
    from fetcher_gaap import _apply_shares_outstanding, StatementTable
    tbl = StatementTable(
        sheet_name="Data_Financials(Q)",
        quarter_labels=["FY2025Q1"],
        filing_dates=[""],
        concepts=["Revenue", "Shares Outstanding"],
        values=[[10.0], [None]],
        ticker="ARLO", labels=["", ""],
    )
    _apply_shares_outstanding([tbl], {"FY2025Q1": 103.0})
    assert tbl.values[0] == [10.0]


def test_apply_shares_leaves_missing_quarters_blank():
    from fetcher_gaap import _apply_shares_outstanding, StatementTable
    tbl = StatementTable(
        sheet_name="Data_Financials(Q)",
        quarter_labels=["FY2025Q1", "FY2025Q2"],
        filing_dates=["", ""],
        concepts=["Shares Outstanding"],
        values=[[None, None]],
        ticker="ARLO", labels=[""],
    )
    _apply_shares_outstanding([tbl], {"FY2025Q1": 103.0})
    assert tbl.values[0] == [103.0, None]


def test_apply_shares_no_row_is_a_noop():
    """沒有那一列的表（Data_Seg_* 等）不可炸。"""
    from fetcher_gaap import _apply_shares_outstanding, StatementTable
    tbl = StatementTable(
        sheet_name="Data_Seg_X", quarter_labels=["FY2025Q1"], filing_dates=[""],
        concepts=["Americas"], values=[[1.0]], ticker="ARLO", labels=[""],
    )
    _apply_shares_outstanding([tbl], {"FY2025Q1": 103.0})
    assert tbl.values[0] == [1.0]


# ═════════════════════════════════════════════════════════════════════════════
# 依 49 家實測補進模板的列（2026-08-03，docs/docs_statement_template_proposal.md）
#
# 判準：多數公司都有的就進固定模板，某些公司沒有就留空白列。
# 這五列的跨公司覆蓋率（非金融 46 家為分母）：
#   CF Change in Accounts Payable        52%  ← 營運資金四大項唯一漏掉的
#   IS Net Income incl. NCI (ProfitLoss) 46%  ← 有 NCI 結構的公司分開報
#   CF Change in Prepaid & Other Assets  35%
#   IS Total Costs and Expenses          33%
#   CF Change in Other Operating Assets  26%
# 刻意不收 BS CommitmentsAndContingencies（70%）——法定揭露列，幾乎永遠無值。
# ═════════════════════════════════════════════════════════════════════════════

def _names(template):
    return [r[0] for r in template]


def test_is_has_net_income_incl_nci():
    from fetcher_gaap import IS_TEMPLATE
    assert "Net Income incl. NCI" in _names(IS_TEMPLATE)


def test_net_income_incl_nci_maps_to_profitloss():
    from fetcher_gaap import IS_TEMPLATE
    row = next(r for r in IS_TEMPLATE if r[0] == "Net Income incl. NCI")
    assert "ProfitLoss" in r[2] if (r := row) else False


def test_net_income_incl_nci_sits_after_minority_interest():
    from fetcher_gaap import IS_TEMPLATE
    n = _names(IS_TEMPLATE)
    assert n.index("Minority Interest") < n.index("Net Income incl. NCI")


def test_is_has_total_costs_and_expenses():
    from fetcher_gaap import IS_TEMPLATE
    assert "Total Costs and Expenses" in _names(IS_TEMPLATE)


def test_cf_has_change_in_accounts_payable():
    """營運資金四大項：應收、存貨、應付、預付。應付原本漏了。"""
    from fetcher_gaap import CF_TEMPLATE
    assert "Change in Accounts Payable" in _names(CF_TEMPLATE)


def test_change_in_ap_maps_to_the_right_concept():
    from fetcher_gaap import CF_TEMPLATE
    row = next(r for r in CF_TEMPLATE if r[0] == "Change in Accounts Payable")
    assert "IncreaseDecreaseInAccountsPayable" in row[2]


def test_change_in_ap_sits_with_the_other_working_capital_rows():
    from fetcher_gaap import CF_TEMPLATE
    n = _names(CF_TEMPLATE)
    assert n.index("Change in Inventories") < n.index("Change in Accounts Payable")
    assert n.index("Change in Accounts Payable") < n.index("Operating Cash Flow")


def test_cf_has_prepaid_and_other_asset_changes():
    from fetcher_gaap import CF_TEMPLATE
    n = _names(CF_TEMPLATE)
    assert "Change in Prepaid & Other Assets" in n
    assert "Change in Other Operating Assets" in n


def test_commitments_and_contingencies_deliberately_absent():
    """70% 的公司有這個 tag，但它是法定揭露列、幾乎永遠無值，收進來只是多一列空白。"""
    from fetcher_gaap import BS_TEMPLATE
    joined = " ".join(f"{r[0]} {r[1] or ''} {r[2] or ''}" for r in BS_TEMPLATE)
    assert "CommitmentsAndContingencies" not in joined


def test_no_template_row_was_removed():
    """使用者明確要求「不要砍東西」——低命中列一律保留。"""
    from fetcher_gaap import IS_TEMPLATE, BS_TEMPLATE, CF_TEMPLATE
    for name in ("Other Operating Expense", "Interest Income", "Other Non-op Inc/(Exp)"):
        assert name in _names(IS_TEMPLATE)
    for name in ("Finance Lease Liabilities, LT", "Treasury Stock",
                 "Deferred Revenue, LT", "Pension & Retirement Oblig."):
        assert name in _names(BS_TEMPLATE)
    assert "Amortization of Intangibles" in _names(CF_TEMPLATE)


# ═════════════════════════════════════════════════════════════════════════════
# overflow 移到 sheet 最底部（2026-08-03）
#
# 原本 overflow 接在每個 section 的模板列之後，所以 IS 多幾行 overflow，
# BS 整段就往下推——實測 `Cash` 在 11 個輸出檔裡落在第 28~56 列之間。
# 移到底部之後模板列號跨公司固定，跨檔案公式才寫得出來。
# ═════════════════════════════════════════════════════════════════════════════

def _mk(sheet, concepts, quarters=("FY2025Q1",), labels=None):
    from fetcher_gaap import StatementTable
    return StatementTable(
        sheet_name=sheet, quarter_labels=list(quarters),
        filing_dates=[""] * len(quarters), concepts=list(concepts),
        values=[[1.0] * len(quarters) for _ in concepts],
        ticker="T", labels=list(labels or [""] * len(concepts)),
    )


def _merged_with_overflow():
    from fetcher_gaap import _merge_financials, IS_TEMPLATE, BS_TEMPLATE, CF_TEMPLATE
    is_names = [r[0] for r in IS_TEMPLATE] + ["公司特有的 IS 科目"]
    bs_names = [r[0] for r in BS_TEMPLATE] + ["公司特有的 BS 科目"]
    cf_names = [r[0] for r in CF_TEMPLATE] + ["公司特有的 CF 科目"]
    return _merge_financials(_mk("IS", is_names), _mk("BS", bs_names), _mk("CF", cf_names))


def test_overflow_rows_moved_to_the_bottom():
    from fetcher_gaap import OVERFLOW_SECTION
    t = _merged_with_overflow()
    assert OVERFLOW_SECTION in t.concepts
    head = t.concepts.index(OVERFLOW_SECTION)
    for name in ("公司特有的 IS 科目", "公司特有的 BS 科目", "公司特有的 CF 科目"):
        assert t.concepts.index(name) > head


def test_template_row_positions_are_independent_of_overflow_count():
    """同樣的模板，overflow 多寡不可影響 BS/CF 模板列的位置。"""
    from fetcher_gaap import _merge_financials, IS_TEMPLATE, BS_TEMPLATE, CF_TEMPLATE
    is_base = [r[0] for r in IS_TEMPLATE]
    bs_names = [r[0] for r in BS_TEMPLATE]
    cf_names = [r[0] for r in CF_TEMPLATE]

    few = _merge_financials(_mk("IS", is_base + ["X1"]), _mk("BS", bs_names), _mk("CF", cf_names))
    many = _merge_financials(_mk("IS", is_base + [f"X{i}" for i in range(1, 12)]),
                             _mk("BS", bs_names), _mk("CF", cf_names))
    for probe in ("Cash", "Total Assets", "Operating Cash Flow", "Capex"):
        assert few.concepts.index(probe) == many.concepts.index(probe), probe


def test_overflow_row_keeps_its_original_label():
    t = _merged_with_overflow()
    i = t.concepts.index("公司特有的 BS 科目")
    assert t.values[i] is not None


def test_overflow_section_absent_when_no_overflow():
    from fetcher_gaap import _merge_financials, OVERFLOW_SECTION, IS_TEMPLATE, BS_TEMPLATE, CF_TEMPLATE
    t = _merge_financials(_mk("IS", [r[0] for r in IS_TEMPLATE]),
                          _mk("BS", [r[0] for r in BS_TEMPLATE]),
                          _mk("CF", [r[0] for r in CF_TEMPLATE]))
    assert OVERFLOW_SECTION not in t.concepts


def test_three_sections_still_present_and_ordered():
    t = _merged_with_overflow()
    a = t.concepts.index("Income Statement")
    b = t.concepts.index("Balance Sheet")
    c = t.concepts.index("Cash Flow")
    assert a < b < c


# ═════════════════════════════════════════════════════════════════════════════
# 品質檢查併進 Data_Meta（2026-08-03，原本在已移除的 Index sheet）
#
# 判斷方式（override_engine.check_key_rows）：9 個關鍵科目，每個檢查
# 「最近 4 期是否全部為空」——全空才算缺。只要有任一期有值就算通過，
# 所以 9/9 的意思是「九個核心科目都至少抓到一期」，不代表每期都完整。
# ═════════════════════════════════════════════════════════════════════════════

def _q_table_for_meta(missing=()):
    from fetcher_gaap import StatementTable
    from excel_formatter import ALL_KEY_ROWS
    concepts, values = [], []
    for name in ALL_KEY_ROWS:
        concepts.append(name)
        values.append([None, None] if name in missing else [1.0, 2.0])
    return StatementTable(
        sheet_name="Data_Financials(Q)", quarter_labels=["FY2025Q1", "FY2025Q2"],
        filing_dates=["", ""], concepts=concepts, values=values,
        ticker="T", labels=[""] * len(concepts),
    )


def test_meta_has_quality_rows():
    from fetcher_gaap import _build_meta_table
    m = _build_meta_table("T", "Test Inc", [_q_table_for_meta()])
    assert "Key Rows Complete" in m.concepts


def test_meta_quality_all_present():
    from fetcher_gaap import _build_meta_table
    from excel_formatter import ALL_KEY_ROWS
    m = _build_meta_table("T", "Test Inc", [_q_table_for_meta()])
    assert m.values[m.concepts.index("Key Rows Complete")][0] == f"{len(ALL_KEY_ROWS)}/{len(ALL_KEY_ROWS)}"


def test_meta_quality_lists_missing_rows():
    from fetcher_gaap import _build_meta_table
    m = _build_meta_table("T", "Test Inc", [_q_table_for_meta(missing={"Capex", "Diluted EPS"})])
    score = m.values[m.concepts.index("Key Rows Complete")][0]
    missing = m.values[m.concepts.index("Key Rows Missing")][0]
    assert score == "7/9"
    assert "Capex" in missing and "Diluted EPS" in missing


def test_meta_quality_blank_when_no_quarterly_table():
    """只抓年報時沒有季表，品質列留空而不是報 0/9（那會被誤讀成全缺）。"""
    from fetcher_gaap import _build_meta_table
    m = _build_meta_table("T", "Test Inc", [])
    idx = m.concepts.index("Key Rows Complete")
    assert m.values[idx] in ([], [""], [None]) or all(not v for v in m.values[idx])


# ── Data_Meta 再補：最新期間與財年起訖（2026-08-03）──────────────────────

def _q_table_with_periods():
    from fetcher_gaap import StatementTable
    return StatementTable(
        sheet_name="Data_Financials(Q)",
        quarter_labels=["FY2025Q3", "FY2026Q1"],
        filing_dates=["2025-11-06", "2026-05-07"],
        period_ends=["2025-09-28", "2026-03-29"],
        concepts=["Revenue"], values=[[1.0, 2.0]], ticker="T", labels=[""],
    )


def test_meta_reports_latest_period():
    """要一眼看出「這份檔案的資料抓到哪一季」。"""
    from fetcher_gaap import _build_meta_table
    m = _build_meta_table("T", "Test Inc", [_q_table_with_periods()], fy_end_month=12)
    assert m.values[m.concepts.index("Latest Period")][0] == "FY2026Q1"


def test_meta_reports_latest_period_end_date():
    from fetcher_gaap import _build_meta_table
    m = _build_meta_table("T", "Test Inc", [_q_table_with_periods()], fy_end_month=12)
    assert m.values[m.concepts.index("Latest Period End")][0] == "2026-03-29"


def test_meta_reports_fiscal_year_span_for_december_fye():
    from fetcher_gaap import _build_meta_table
    m = _build_meta_table("T", "Test Inc", [_q_table_with_periods()], fy_end_month=12)
    assert m.values[m.concepts.index("Fiscal Year Span")][0] == "1 月 – 12 月"


def test_meta_reports_fiscal_year_span_for_september_fye():
    """AAPL 9 月結算 → 財年從 10 月開始。"""
    from fetcher_gaap import _build_meta_table
    m = _build_meta_table("T", "Apple", [_q_table_with_periods()], fy_end_month=9)
    assert m.values[m.concepts.index("Fiscal Year Span")][0] == "10 月 – 9 月"


def test_meta_latest_period_blank_without_quarterly_table():
    from fetcher_gaap import _build_meta_table
    m = _build_meta_table("T", "Test Inc", [])
    assert all(not v for v in m.values[m.concepts.index("Latest Period")]) or            m.values[m.concepts.index("Latest Period")] == []


# ── 財季編號改用期末日反推（2026-08-22）─────────────────────────────────────
#
# edgartools 欄名裡的 `(Qn)` 對 52/53 週財年制的公司會標錯（實測 NVDA、INTC），
# 兩份不同期間的 10-Q 因此算出同一個 label，`_build_*_table` 的 dedup
# （`if label in periods: continue`）就把舊的那一季靜默丟掉，連帶讓
# `_synthesize_q4()` 缺 Q1/Q2/Q3 而合成不出 Q4。
# 這幾個測試釘住「財季編號一律由期末日 + 財年結束月反推，不採信 `(Qn)`」。

def test_col_to_quarter_label_ignores_wrong_edgartools_q_marker():
    """NVDA 2016-05-01（FY2017 Q1）被 edgartools 標成 (Q2)，必須以日期為準。"""
    assert _col_to_quarter_label("2016-05-01 (Q2)", fy_end_month=1) == "FY2017Q1"


def test_col_to_quarter_label_nvda_q2_not_collapsed_into_q1():
    """同一財年的下一季（2016-07-31，也被標成 (Q2)）必須算出 FY2017Q2。"""
    assert _col_to_quarter_label("2016-07-31 (Q2)", fy_end_month=1) == "FY2017Q2"


def test_col_to_quarter_label_intc_q1_ending_in_april():
    """INTC 2023-04-01 是 2023 Q1（13 週制溢出到 4 月），edgartools 標 (Q2)。"""
    assert _col_to_quarter_label("2023-04-01 (Q2)", fy_end_month=12) == "FY2023Q1"


def test_col_to_quarter_label_no_collision_across_one_fiscal_year():
    """NVDA FY2011 三季各自算出不同 label——這正是資料被丟棄的根因。"""
    labels = [
        _col_to_quarter_label("2010-05-02 (Q2)", fy_end_month=1),
        _col_to_quarter_label("2010-08-01 (Q3)", fy_end_month=1),
        _col_to_quarter_label("2010-10-31 (Q3)", fy_end_month=1),
    ]
    assert labels == ["FY2011Q1", "FY2011Q2", "FY2011Q3"]


# ── G1：第 4 列的日曆季不可以自己算一套（2026-08-22）──────────────────────
#
# `_calendar_quarter()` 原本直接取期末日的月份，完全不內縮——INTC 結束在
# 2023-04-01 的那一季（實際涵蓋 1~3 月）會被算成 2023Q2。平常這個值會被
# `fiscal_input._apply_to_sheet()` 的公式蓋掉看不到，但第 5 列不是完整 ISO
# 日期的殘留格（合成 Q4 的年報期末日有時只有 `2010-01`）會保留它，那就是錯值。
# 改成一律委派給 fiscal_input 的 `basis="end"`，不要有第二套實作。

def test_calendar_quarter_row_uses_the_shared_end_basis():
    """52/53 週制溢出到 4 月初的季，要算 Q1 不是 Q2。"""
    from fetcher_gaap import _calendar_quarter
    assert _calendar_quarter("FY2023Q1", 12, "2023-04-01") == "2023Q1"


def test_calendar_quarter_row_matches_fiscal_input_exactly():
    """釘住「只有一份實作」——逐格比對，不可以各自演化。"""
    from fetcher_gaap import _calendar_quarter
    from fiscal_input import calendar_quarter_of
    for pe in ["2023-04-01", "2025-07-27", "2026-01-25", "2025-06-28", "2026-01-02"]:
        assert _calendar_quarter("FY2025Q1", 12, pe) == calendar_quarter_of(pe, basis="end")


def test_calendar_quarter_row_falls_back_to_label_when_no_period_end():
    """抓不到期末日時仍走財季標籤反推的退路，不可以因為改實作就退化成空字串。"""
    from fetcher_gaap import _calendar_quarter
    assert _calendar_quarter("FY2026Q1", 12, "") == "2026Q1"
    assert _calendar_quarter("FY2026", 12, "") == ""


# ── G3：IS 的 CF-sourced 列改成從 cf_tbl 回填（2026-08-22）─────────────────
#
# IS_TEMPLATE 裡 source=="CF" 的兩列（`SBC`、`D&A (CF memo)`）原本在
# _build_is_table() 走 `_current_q_col(cf_df)` 直接找現金流量表的單季欄。但
# 10-Q 的現金流量表是 YTD 累計，Q2/Q3 的 filing 根本沒有單季欄，那兩列就整片
# 空白（實測 NVDA 缺 51/68），連帶 _synthesize_q4() 也算不出 Q4。
# _build_cf_table() 已經做了 YTD 拆算（本季 YTD − 上季 YTD），所以 CF 區同名
# 兩列是好的（缺 1/68）。修法是**共用 CF 已經算好的單季值**，不要在 IS 再寫
# 一份 YTD 拆算——那就是第二份會漂移的實作。

def _tbl(sheet, labels, concepts, values):
    return StatementTable(
        sheet_name=sheet, quarter_labels=labels,
        filing_dates=[""] * len(labels), concepts=concepts, values=values,
        ticker="T", labels=[""] * len(concepts),
        period_ends=[""] * len(labels),
    )


def test_backfill_cf_sourced_rows_fills_is_from_cf():
    from fetcher_gaap import _backfill_cf_sourced_rows
    is_tbl = _tbl("Data_IS", ["FY2025Q1", "FY2025Q2", "FY2025Q3"],
                  ["Revenue", "SBC", "D&A (CF memo)"],
                  [[10.0, 20.0, 30.0], [1.0, None, None], [2.0, None, None]])
    cf_tbl = _tbl("Data_CF", ["FY2025Q1", "FY2025Q2", "FY2025Q3"],
                  ["SBC", "D&A"],
                  [[1.0, 1.5, 1.8], [2.0, 2.5, 2.8]])

    out = _backfill_cf_sourced_rows(is_tbl, cf_tbl)

    assert out.values[out.concepts.index("SBC")] == [1.0, 1.5, 1.8]
    assert out.values[out.concepts.index("D&A (CF memo)")] == [2.0, 2.5, 2.8]
    assert out.values[out.concepts.index("Revenue")] == [10.0, 20.0, 30.0]


def test_backfill_cf_sourced_rows_matches_by_label_not_position():
    """兩張表的欄位順序/期數不保證一樣，一定要依 quarter_labels 對照。"""
    from fetcher_gaap import _backfill_cf_sourced_rows
    is_tbl = _tbl("Data_IS", ["FY2025Q1", "FY2025Q2", "FY2025Q3"],
                  ["SBC"], [[None, None, None]])
    cf_tbl = _tbl("Data_CF", ["FY2025Q2", "FY2025Q3"], ["SBC"], [[1.5, 1.8]])

    out = _backfill_cf_sourced_rows(is_tbl, cf_tbl)
    assert out.values[0] == [None, 1.5, 1.8]


def test_backfill_cf_sourced_rows_keeps_is_value_when_cf_has_none():
    """CF 那格沒值就別動 IS 已有的值，不要用 None 蓋掉真資料。"""
    from fetcher_gaap import _backfill_cf_sourced_rows
    is_tbl = _tbl("Data_IS", ["FY2025Q1"], ["SBC"], [[9.9]])
    cf_tbl = _tbl("Data_CF", ["FY2025Q1"], ["SBC"], [[None]])

    out = _backfill_cf_sourced_rows(is_tbl, cf_tbl)
    assert out.values[0] == [9.9]


def test_backfill_cf_sourced_rows_survives_missing_rows():
    """CF 表沒有那一列（或 IS 沒有）時安靜跳過，不可以拋例外。"""
    from fetcher_gaap import _backfill_cf_sourced_rows
    is_tbl = _tbl("Data_IS", ["FY2025Q1"], ["Revenue"], [[10.0]])
    cf_tbl = _tbl("Data_CF", ["FY2025Q1"], ["Operating Cash Flow"], [[5.0]])
    assert _backfill_cf_sourced_rows(is_tbl, cf_tbl).values == [[10.0]]


# ── G12：現金流量表裡的時點值不可以做 YTD 相減（2026-08-22）────────────────
#
# `Ending Cash`（期末現金餘額）排在 CF_TEMPLATE 裡，但它是**餘額**不是本期
# 發生額。`_build_cf_table()` 對整張表做「本季 YTD − 上季 YTD」，減出來變成
# 「現金變動額」。實測 AAPL：
#     2026-03-28   現行     255,000,000   正確 45,572,000,000
#     2026-06-27   現行  -6,028,000,000   正確 39,544,000,000
# 52 家裡 50 家中招（逐格命中率只有 32.8%）。

def _make_cf_df_with_cash(period_col, ni, ocf, ending_cash):
    df = _make_cf_df_minimal(period_col, ni, ocf)
    extra = pd.DataFrame({
        "concept": ["us-gaap_CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents"],
        "label": ["Cash and cash equivalents, end of period"],
        "standard_concept": ["CashAndCashEquivalents"],
        "abstract": [False], "is_breakdown": [False], "level": [3],
        "dimension_member_label": [None],
        period_col: [ending_cash],
    })
    return pd.concat([df, extra], ignore_index=True)


def _make_cf_filing_with_cash(is_col, cf_col, ni, ocf, cash, filing_date):
    is_df = _make_is_df_minimal(is_col)
    cf_df = _make_cf_df_with_cash(cf_col, ni, ocf, cash)
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


def test_ending_cash_is_a_balance_and_must_not_be_ytd_subtracted():
    """Q2 的期末現金餘額就是那個餘額，不是「Q2 餘額 − Q1 餘額」。"""
    q1 = _make_cf_filing_with_cash("2025-03-31 (Q1)", "2025-03-31 (Q1)",
                                    ni=100.0, ocf=150.0, cash=900.0,
                                    filing_date="2025-04-30")
    q2 = _make_cf_filing_with_cash("2025-06-30 (Q2)", "2025-06-30 (YTD)",
                                    ni=230.0, ocf=330.0, cash=950.0,
                                    filing_date="2025-07-30")
    gaap_tbl, _ = _build_cf_table([q2, q1], max_filings=80)

    cash_idx = gaap_tbl.concepts.index("Ending Cash")
    q2_pos = gaap_tbl.quarter_labels.index("FY2025Q2")
    # 錯誤行為會給 950 − 900 = 50
    assert gaap_tbl.values[cash_idx][q2_pos] == pytest.approx(950.0)


def test_flow_rows_are_still_ytd_subtracted_after_the_fix():
    """修時點值不可以順手把流量項的 YTD 拆算也關掉。"""
    q1 = _make_cf_filing_with_cash("2025-03-31 (Q1)", "2025-03-31 (Q1)",
                                    ni=100.0, ocf=150.0, cash=900.0,
                                    filing_date="2025-04-30")
    q2 = _make_cf_filing_with_cash("2025-06-30 (Q2)", "2025-06-30 (YTD)",
                                    ni=230.0, ocf=330.0, cash=950.0,
                                    filing_date="2025-07-30")
    gaap_tbl, _ = _build_cf_table([q2, q1], max_filings=80)

    q2_pos = gaap_tbl.quarter_labels.index("FY2025Q2")
    ocf_idx = gaap_tbl.concepts.index("Operating Cash Flow")
    assert gaap_tbl.values[ocf_idx][q2_pos] == pytest.approx(180.0)   # 330 − 150


def test_synthesize_q4_takes_balance_rows_from_annual_not_subtraction():
    """同一個 bug 的第二個實例：合成 Q4 也不可以對餘額做「年報 − Q1 − Q2 − Q3」。

    實測 AAPL 2025-09-27 修完第一處後仍是 -58,796,000,000（那是合成 Q4 欄）。
    餘額的 Q4 就是年報上的期末餘額，直接取用。
    """
    from fetcher_gaap import _CF_IDX, _CF_POINT_IN_TIME_IDX, CF_TEMPLATE
    n = len(CF_TEMPLATE)
    cash_i, ocf_i = _CF_IDX["Ending Cash"], _CF_IDX["Operating Cash Flow"]

    q = StatementTable(
        sheet_name="Data_CF",
        quarter_labels=["FY2025Q1", "FY2025Q2", "FY2025Q3"],
        filing_dates=[""] * 3,
        concepts=[r[0] for r in CF_TEMPLATE],
        values=[[10.0, 20.0, 30.0] if i == ocf_i else
                [100.0, 200.0, 300.0] if i == cash_i else [None] * 3
                for i in range(n)],
        ticker="T", labels=[""] * n, period_ends=["2025-03-31", "2025-06-30", "2025-09-30"],
    )
    ann = StatementTable(
        sheet_name="Data_CF", quarter_labels=["FY2025"], filing_dates=[""],
        concepts=[r[0] for r in CF_TEMPLATE],
        values=[[100.0] if i == ocf_i else [400.0] if i == cash_i else [None]
                for i in range(n)],
        ticker="T", labels=[""] * n, period_ends=["2025-12-31"],
    )
    out = _synthesize_q4(q, ann, n, is_balance=False,
                         point_in_time_idx=_CF_POINT_IN_TIME_IDX)
    p = out.quarter_labels.index("FY2025Q4")
    # 流量項照減：100 − 10 − 20 − 30 = 40
    assert out.values[ocf_i][p] == pytest.approx(40.0)
    # 餘額直接取年報值 400，不是 400 − 100 − 200 − 300 = −200
    assert out.values[cash_i][p] == pytest.approx(400.0)


# ── G9：一次執行內同一份 filing 只解析一次（2026-08-22）─────────────────────
#
# 實測 ARLO（25 份 filing，66 秒）：`_filing_obj` 被呼叫 96 次（每份 3.8 次），
# `financials`（XBRL 解析）花 19.9s、`to_dataframe` 花 28.4s。IS/BS/CF/segments
# 四個 build pass 各自對同一批 filing 重解析一次。edgartools 不會跨呼叫快取
# （同一支 ticker 連跑兩次：64.5s vs 67.3s，完全沒有變快）。

def test_filing_obj_parses_each_accession_only_once_per_run():
    from fetcher_gaap import _filing_obj, _parse_cache_scope
    calls = {"n": 0}

    def _obj():
        calls["n"] += 1
        return MagicMock()

    f = MagicMock()
    f.obj.side_effect = _obj
    f.accession_no = "0001045810-25-000116"

    with _parse_cache_scope():
        a = _filing_obj(f)
        b = _filing_obj(f)
    assert a is b
    assert calls["n"] == 1


def test_parse_cache_does_not_leak_between_runs():
    """快取只能活在一次抓取內——跨 ticker 或跨執行殘留會拿到過期資料。"""
    from fetcher_gaap import _filing_obj, _parse_cache_scope
    calls = {"n": 0}

    def _obj():
        calls["n"] += 1
        return MagicMock()

    f = MagicMock()
    f.obj.side_effect = _obj
    f.accession_no = "0001045810-25-000116"

    with _parse_cache_scope():
        _filing_obj(f)
    with _parse_cache_scope():
        _filing_obj(f)
    assert calls["n"] == 2


def test_filing_obj_still_works_without_a_cache_scope():
    """沒有開快取範圍時要照常運作（cli.py 之類的路徑可能直接呼叫）。"""
    from fetcher_gaap import _filing_obj
    f = MagicMock()
    f.obj.return_value = "X"
    f.accession_no = "acc-1"
    assert _filing_obj(f) == "X"


# ── 列 filing 清單也要有退避重試（2026-08-25）────────────────────────────────
#
# `_filing_obj()` 抓單份 filing 內容早就有 `with_retry` 保護瞬斷；但「列出這家
# 公司有哪些 filing」（`company.get_filings()`）比那更早一步，完全沒被任何重試
# 機制蓋到——2026-08-25 實測 201 家重建撞到 6 家逾時（JCI/JPM/MCD/MDLZ/MDT/META），
# 全部發生在這一步，直接讓整趟 `fetch_gaap_statements()` 拋例外，連 D11-B 的
# 缺漏帳本都還沒開始記，重試不到。

def test_list_filings_retries_transient_network_error():
    from fetcher_gaap import _list_filings

    class ReadTimeout(Exception):
        pass

    calls = {"n": 0}

    def _get_filings(form, amendments):
        calls["n"] += 1
        if calls["n"] < 2:
            raise ReadTimeout("timed out")
        return ["filing1", "filing2"]

    company = MagicMock()
    company.get_filings.side_effect = _get_filings

    result = _list_filings(company, "10-Q", sleep=lambda _s: None)
    assert result == ["filing1", "filing2"]
    assert calls["n"] == 2


def test_list_filings_gives_up_after_exhausting_retries():
    from fetcher_gaap import _list_filings
    from net_retry import NetworkDownError

    class ReadTimeout(Exception):
        pass

    company = MagicMock()
    company.get_filings.side_effect = ReadTimeout("timed out")

    with pytest.raises(NetworkDownError):
        _list_filings(company, "10-Q", sleep=lambda _s: None)


# ── H3 系統性 concept 對照修復（2026-08-23）────────────────────────────────
#
# 每個測試都釘住「哪一家公司、哪個 concept / label」——這些 concept 與 label
# 字串是從真實的 `to_dataframe()` 逐字抄下來的，不是編的（來源：2026-08-23
# 對 22 家最新 10-Q 的實測掃描）。起因見
# `docs/template-coverage-baseline-2026-08-23.md`：同一個「`label_hint` 太窄」
# 或「fallback 正則太窄」的問題在多家公司重複出現，修一次多家一起變好。

def _row_df(concept: str, label: str, std_concept, *, extra=None):
    """單列（或多列）的最小 statement dataframe。extra 是 [(concept, label, std)]。"""
    rows = [(concept, label, std_concept)] + list(extra or [])
    return pd.DataFrame({
        "concept":                [r[0] for r in rows],
        "label":                  [r[1] for r in rows],
        "standard_concept":       [r[2] for r in rows],
        "abstract":               [False] * len(rows),
        "is_breakdown":           [False] * len(rows),
        "level":                  [3] * len(rows),
        "dimension_member_label": [None] * len(rows),
        "2026-06-30 (Q2)":        [float(i + 1) for i in range(len(rows))],
    })


def _template_entry(row_name: str):
    from fetcher_gaap import BS_TEMPLATE, CF_TEMPLATE, IS_TEMPLATE
    for T in (IS_TEMPLATE, BS_TEMPLATE, CF_TEMPLATE):
        for r in T:
            if r[0] == row_name:
                return r
    raise AssertionError("模板裡沒有這一列：" + row_name)


def _match_template_row(row_name: str, df):
    """照模板設定跑一次 _match_is_row，回傳命中的 label（沒命中回 None）。"""
    _, std, fallback, _src, match, hint, lbl_fb = _template_entry(row_name)
    idx = _match_is_row(df, std, fallback, label_fallback=lbl_fb,
                        match=match, label_hint=hint)
    return None if idx is None else df.loc[idx, "label"]


@pytest.mark.parametrize("label", [
    "Treasury stock purchases",                 # PG / MCD / LRCX
    "Purchases of stock for treasury",          # KO
    "Purchases of common stock for treasury",   # GE
    "Common stock acquired",                    # XOM
    "Purchase of Company stock",                # WMT
    "Payments to purchase common stock",        # CAT
    "Purchase of treasury shares and restricted stock unit withholdings",  # NXPI
    "Repurchases of common stock",              # AMD（原本就過，防止改壞）
])
def test_share_repurchases_matches_regardless_of_label_wording(label):
    """`label_hint='repurchas'` 把 22 家裡的 9 家過濾掉——多數公司寫 treasury 不寫 repurchase。"""
    df = _row_df("us-gaap_PaymentsForRepurchaseOfCommonStock", label,
                 "EquityExpenseIncome(BuybackIssued)")
    assert _match_template_row("Share Repurchases", df) == label


@pytest.mark.parametrize("concept,label", [
    ("us-gaap_IncreaseDecreaseInInventories", "Inventory"),                 # GOOGL/AVGO/SWKS/TSLA
    ("us-gaap_IncreaseDecreaseInRetailRelatedInventories", "Inventories"),  # WMT
])
def test_change_in_inventories_matches_singular_label_and_retail_concept(concept, label):
    """hint 是複數 inventories，但多數公司的 label 是單數 Inventory。"""
    df = _row_df(concept, label, float("nan"))
    assert _match_template_row("Change in Inventories", df) == label


def _sum_debt_row(kind: str, df):
    """跑一次借款／還款的加總後處理，回傳金額（沒有任何列命中回 None）。"""
    from fetcher_gaap import (_DEBT_PROCEEDS_PATTERNS, _DEBT_REPAYMENTS_PATTERNS,
                              _sum_matching_rows)
    pats = _DEBT_PROCEEDS_PATTERNS if kind == "proceeds" else _DEBT_REPAYMENTS_PATTERNS
    val, _idx = _sum_matching_rows(df, "2026-06-30 (Q2)", pats, set())
    return val


@pytest.mark.parametrize("concept", [
    "us-gaap_ProceedsFromIssuanceOfLongTermDebt",
    "us-gaap_ProceedsFromIssuanceOfSeniorLongTermDebt",     # NOW
    "us-gaap_ProceedsFromIssuanceOfCommercialPaper",        # AMAT / INTC / NOW
    "us-gaap_ProceedsFromDebtNetOfIssuanceCosts",           # GOOGL
    "us-gaap_ProceedsFromDebtMaturingInMoreThanThreeMonths",  # CAT
    "us-gaap_ProceedsFromConvertibleDebt",                  # PANW
])
def test_debt_proceeds_sums_real_world_concepts(concept):
    """原本的 pattern 清單漏掉商業本票、可轉債、GOOGL 與 CAT 的寫法。"""
    df = _row_df(concept, "Proceeds from borrowings", float("nan"))
    assert _sum_debt_row("proceeds", df) == 1.0


def test_debt_proceeds_adds_up_every_borrowing_line():
    """公司常常同時有長期借款與商業本票兩條以上，要加總不是挑一條。"""
    df = _row_df("us-gaap_ProceedsFromIssuanceOfLongTermDebt", "Proceeds from long-term debt",
                 float("nan"),
                 extra=[("us-gaap_ProceedsFromIssuanceOfCommercialPaper",
                         "Proceeds from commercial paper", float("nan"))])
    assert _sum_debt_row("proceeds", df) == 3.0    # 1.0 + 2.0


def test_debt_proceeds_ignores_net_proceeds_from_repayments_rows():
    """`ProceedsFromRepaymentsOf...` 是淨額（借款減還款），塞進 Debt Proceeds 會失真。"""
    df = _row_df("us-gaap_ProceedsFromRepaymentsOfShortTermDebtMaturingInThreeMonthsOrLess",
                 "Short-term borrowings, net", float("nan"))   # CAT / GE
    assert _sum_debt_row("proceeds", df) is None


@pytest.mark.parametrize("concept", [
    "us-gaap_RepaymentsOfLongTermDebt",
    "us-gaap_RepaymentsOfCommercialPaper",                    # AMAT
    "us-gaap_RepaymentsOfDebtMaturingInMoreThanThreeMonths",  # MSFT / CAT
    "us-gaap_RepaymentsOfDebtAndCapitalLeaseObligations",     # GOOGL / LRCX
    "us-gaap_RepaymentsOfNotesPayable",                       # NXPI
])
def test_debt_repayments_sums_real_world_concepts(concept):
    df = _row_df(concept, "Repayments of borrowings", float("nan"))
    assert _sum_debt_row("repayments", df) == 1.0


def test_debt_repayments_ignores_net_proceeds_from_repayments_rows():
    """`RepaymentsOfShortTermDebt` 這個舊 pattern 會誤吃 WMT 的
    `ProceedsFromRepaymentsOfShortTermDebt`（淨額，可正可負）。"""
    df = _row_df("us-gaap_ProceedsFromRepaymentsOfShortTermDebt",
                 "Net change in short-term borrowings", float("nan"))   # WMT
    assert _sum_debt_row("repayments", df) is None


@pytest.mark.parametrize("concept,label", [
    ("us-gaap_IncomeTaxesPaid",    "Income taxes, net"),                  # COST
    ("us-gaap_IncomeTaxesPaidNet", "Income taxes, net of tax refunds"),   # CRM
])
def test_cash_taxes_paid_matches_when_label_omits_the_word_paid(concept, label):
    df = _row_df(concept, label, float("nan"))
    assert _match_template_row("Cash Taxes Paid", df) == label


def test_cash_taxes_paid_never_returns_deferred_tax_expense():
    """std_concept `IncomeTaxes` 會命中遞延所得稅費用，那不是付出去的現金稅。
    22 家裡有 11 家的 CF 有這一列（NVDA/AMZN/PG/KO/CAT/ORCL/TSLA/LRCX/MCD/NXPI/PFE）。"""
    df = _row_df("us-gaap_DeferredIncomeTaxExpenseBenefit", "Deferred income taxes",
                 "IncomeTaxes")
    assert _match_template_row("Cash Taxes Paid", df) is None


def test_cash_interest_paid_matches_label_without_the_word_paid():
    df = _row_df("us-gaap_InterestPaidNet", "Interest", float("nan"))   # COST / CRM
    assert _match_template_row("Cash Interest Paid", df) == "Interest"


def test_cash_interest_paid_picks_the_total_not_the_operating_portion():
    """XOM 把利息拆成三列：營運活動內、資本化、以及合計。三列的 concept 都以
    `InterestPaid` 開頭，取第一列會只拿到營運那一段（402M 而不是 910M）。"""
    df = _row_df("us-gaap_InterestPaidNet", "Included in cash flows from operating activities",
                 "InterestExpense",
                 extra=[("us-gaap_InterestPaidCapitalized",
                         "Capitalized, included in cash flows from investing activities", "InterestExpense"),
                        ("us-gaap_InterestPaid", "Total cash interest paid", "InterestExpense")])
    assert _match_template_row("Cash Interest Paid", df) == "Total cash interest paid"


def test_cash_interest_paid_never_returns_debt_extinguishment_loss():
    """std_concept `InterestExpense` 在 AVGO 命中 `GainsLossesOnExtinguishmentOfDebt`。"""
    df = _row_df("us-gaap_GainsLossesOnExtinguishmentOfDebt", "Loss on debt extinguishment",
                 "InterestExpense")
    assert _match_template_row("Cash Interest Paid", df) is None


@pytest.mark.parametrize("row_name,concept,label", [
    ("Operating Cash Flow", "us-gaap_NetCashProvidedByUsedInOperatingActivities", "TOTAL OPERATING ACTIVITIES"),
    ("Investing Cash Flow", "us-gaap_NetCashProvidedByUsedInInvestingActivities", "TOTAL INVESTING ACTIVITIES"),
    ("Financing Cash Flow", "us-gaap_NetCashProvidedByUsedInFinancingActivities", "TOTAL FINANCING ACTIVITIES"),
])
def test_cash_flow_subtotals_match_pg_total_wording(row_name, concept, label):
    """PG 三個小計都寫 TOTAL ... ACTIVITIES，hint `^net cash|^cash` 把整層濾掉。"""
    df = _row_df(concept, label, float("nan"))
    assert _match_template_row(row_name, df) == label


def test_operating_cash_flow_still_rejects_supplemental_lease_rows():
    """hint 放寬之後仍要擋掉補充揭露列——`NetCashProvidedByUsedInOperatingActivities`
    這個 fallback 用 match='last'，不擋的話 XOM/COST/MU/TSLA/SWKS/LRCX 會抓到
    現金流量表最下面的租賃補充揭露列。"""
    df = _row_df("us-gaap_RightOfUseAssetObtainedInExchangeForFinanceLeaseLiability",
                 "Right-of-use assets obtained in exchange for finance lease liabilities",
                 "NetCashFromOperatingActivities")
    assert _match_template_row("Operating Cash Flow", df) is None


@pytest.mark.parametrize("concept,label", [
    ("us-gaap_PaymentsToAcquirePropertyPlantAndEquipment", "Capital expenditures"),   # PG/ORCL/CRM/SWKS/MCD
    ("us-gaap_PaymentsToAcquireProductiveAssets", "Capital expenditures and intangible assets"),  # LRCX
])
def test_capex_matches_capital_expenditures_wording(concept, label):
    """hint 'property' 把「Capital expenditures」這個最常見的寫法整層濾掉。"""
    df = _row_df(concept, label, "CapitalExpenses")
    assert _match_template_row("Capex", df) == label


@pytest.mark.parametrize("concept,label", [
    ("us-gaap_ReceivablesNetCurrent", "Receivables, net"),                       # COST/WMT
    ("us-gaap_ReceivablesNetCurrent", "Receivables"),                            # MU
    ("us-gaap_AccountsReceivableNetCurrent", "Receivables - trade and other"),   # CAT
    ("us-gaap_AccountsNotesAndLoansReceivableNetCurrent", "Accounts and notes receivable"),  # MCD
    ("us-gaap_AccountsReceivableNetCurrent", "Trade receivables, net of allowances"),  # ORCL
])
def test_accounts_receivable_matches_shorter_receivable_wording(concept, label):
    df = _row_df(concept, label, "TradeReceivables")
    assert _match_template_row("Accounts Receivable", df) == label


def test_cash_matches_cash_and_equivalents_without_the_second_cash():
    df = _row_df("us-gaap_CashAndCashEquivalentsAtCarryingValue", "Cash and equivalents",
                 "CashAndMarketableSecurities")   # MCD
    assert _match_template_row("Cash", df) == "Cash and equivalents"


def test_other_current_assets_matches_prepaid_and_other_wording():
    df = _row_df("us-gaap_PrepaidExpenseAndOtherAssetsCurrent", "Prepaid expenses and other",
                 "OtherNonOperatingCurrentAssets")   # WMT
    assert _match_template_row("Other Current Assets", df) == "Prepaid expenses and other"


def test_other_current_assets_still_rejects_nontrade_receivables():
    """AAPL 的 `NontradeReceivablesCurrent` 也掛在 `OtherNonOperatingCurrentAssets`
    這個 std_concept 下，hint 放寬之後仍不能讓它冒充其他流動資產。"""
    df = _row_df("us-gaap_NontradeReceivablesCurrent", "Non-trade receivables",
                 "OtherNonOperatingCurrentAssets")
    assert _match_template_row("Other Current Assets", df) is None


def test_other_non_current_assets_matches_miscellaneous():
    df = _row_df("us-gaap_OtherAssetsNoncurrent", "Miscellaneous",
                 "OtherNonOperatingNonCurrentAssets")   # MCD
    assert _match_template_row("Other Non-current Assets", df) == "Miscellaneous"


def test_common_stock_and_apic_matches_googl_multi_class_label():
    """GOOGL 的 label 是「Class A, Class B, and Class C stock and additional paid-in capital」，
    沒有「common stock」這個字。"""
    label = ("Class A, Class B, and Class C stock and additional paid-in capital, "
             "$0.001 par value per share")
    df = _row_df("us-gaap_CommonStocksIncludingAdditionalPaidInCapital", label, "CommonEquity")
    assert _match_template_row("Common Stock & APIC", df) == label


@pytest.mark.parametrize("concept,label", [
    ("us-gaap_LongTermDebtNoncurrent",                 "Term debt"),          # AAPL
    ("us-gaap_LongTermDebtNoncurrent",                 "Noncurrent debt"),    # CRM
    ("us-gaap_LongTermDebtAndCapitalLeaseObligations", "Long-term borrowings"),  # GE
    ("us-gaap_LongTermNotesAndLoans", "Notes payable and other borrowings, non-current"),  # ORCL
])
def test_long_term_debt_matches_when_label_omits_long_term(concept, label):
    df = _row_df(concept, label, float("nan"))
    assert _match_template_row("Long-term Debt", df) == label


def test_long_term_debt_never_matches_a_current_portion_row():
    """TSLA 的 `tsla_LongTermDebtAndFinanceLeasesCurrent` 是流動部分，
    不能塞進長期負債（拿掉 label_hint 之後就會撞到這一列）。"""
    df = _row_df("tsla_LongTermDebtAndFinanceLeasesCurrent",
                 "Current portion of debt and finance leases", float("nan"))
    assert _match_template_row("Long-term Debt", df) is None


@pytest.mark.parametrize("concept,label", [
    ("us-gaap_ContractWithCustomerLiabilityCurrent", "Deferred revenue"),    # AAPL/TSLA
    ("us-gaap_ContractWithCustomerLiabilityCurrent", "Deferred revenues"),   # ORCL
    ("us-gaap_ContractWithCustomerLiabilityCurrent", "Contract liabilities and deferred income"),  # GE
])
def test_deferred_revenue_current_matches_deferred_revenue_wording(concept, label):
    """hint 'unearned revenue' 幾乎沒有公司在用——52 家只抓到 3 家。"""
    df = _row_df(concept, label, "OtherOperatingCurrentLiabilities")
    assert _match_template_row("Deferred Revenue, current", df) == label


def test_deferred_revenue_current_never_returns_accrued_liabilities():
    """std_concept `OtherOperatingCurrentLiabilities` 會命中應計負債，那不是遞延收入。
    22 家裡 NVDA/GOOGL/PG/CAT/COST/WMT/LRCX 都是這一列。"""
    df = _row_df("us-gaap_AccruedLiabilitiesCurrent", "Accrued expenses and other current liabilities",
                 "OtherOperatingCurrentLiabilities")
    assert _match_template_row("Deferred Revenue, current", df) is None


@pytest.mark.parametrize("concept,std,label", [
    ("us-gaap_OtherCostAndExpenseOperating",   "OtherExpenseIS", "Other operating charges"),          # KO
    ("us-gaap_OtherOperatingIncomeExpenseNet", "OtherIncomeIS",  "Other operating (income) expense, net"),  # MCD / CAT
])
def test_other_operating_expense_matches_the_concepts_companies_actually_tag(concept, std, label):
    """這一列 52 家**一家都沒抓到**。模板猜的 `OtherOperatingExpenses`（std）與
    `OtherOperatingExpense`（fallback）沒有任何公司在用，實際的 tag 是這兩個。"""
    df = _row_df(concept, label, std)
    assert _match_template_row("Other Operating Expense", df) == label


# ── H5：Interest Income fallback 太窄，201 家有 137 家整列空白 ────────────────
#
# 模板的 fallback 是 `InterestIncome`，而 `InvestmentIncomeInterest`（字序相反）
# 兩層都比不到。實測 201 家：`InvestmentIncomeInterest` 84 家、`InterestIncomeOther`
# 21 家、`InvestmentIncomeInterestAndDividend` 14 家。三個都是純收入 concept。
@pytest.mark.parametrize("concept,label", [
    ("us-gaap_InvestmentIncomeInterest", "Interest income"),                       # KO（實測確認在 IS 表面）
    ("us-gaap_InvestmentIncomeInterestAndDividend", "Interest and dividend income"),
    ("us-gaap_InterestIncomeOther", "Other interest income"),
])
def test_interest_income_matches_concepts_companies_actually_tag(concept, label):
    """`InvestmentIncomeInterest` 字序與模板的 fallback `InterestIncome` 相反，
    substring 比對兩層都比不到——即使 std_concept 命名相似，也要靠 fallback 正則抓到。"""
    df = _row_df(concept, label, "InterestAndDividendIncome")
    assert _match_template_row("Interest Income", df) == label


@pytest.mark.parametrize("concept,label", [
    ("us-gaap_InterestIncomeExpenseNet", "Interest income (expense), net"),
    ("us-gaap_InterestIncomeExpenseNonoperatingNet", "Interest expense, net"),  # 50 家抽樣：CHTR/KR/LOW/NEM
])
def test_interest_income_does_not_match_net_concept(concept, label):
    """`InterestIncomeExpense...Net` 這整族都是利息收入減支出的淨額，跟
    `Interest Expense` 那一列可能重複計算——刻意不收，只收純收入的 concept。"""
    df = _row_df(concept, label, "NonoperatingIncomeExpense")
    assert _match_template_row("Interest Income", df) is None


def test_interest_income_does_not_match_noninterest_income():
    """`NoninterestIncomeOtherOperatingIncome`（銀行股常見的「Other income」）字串
    裡剛好包著 `...nterestIncome...`，substring 比對會誤命中——201 家逐格回歸測出來的
    真實案例（BAC）：新 fallback 排除 `ExpenseNet` 之後，比對往後掉到這個完全不相關
    的科目。"""
    df = _row_df("us-gaap_NoninterestIncomeOtherOperatingIncome", "Other income (loss)",
                 "OtherOperatingExpense")
    assert _match_template_row("Interest Income", df) is None


# ── H4 第一步：模板 tuple 第七欄 label_fallback ─────────────────────────────
#
# `_match_is_row()` 一直有第三層（label 比對），但模板的 6-tuple 餵不進去，
# 等於死碼。公司自訂的延伸 tag（`nvda_...`）concept 名字每家不同，只有 label
# 對得上——不接這一層就永遠抓不到。設計見
# `docs/superpowers/specs/2026-08-23-concept-rename-linking-design.md`。

def test_every_template_row_is_a_seven_tuple():
    """三張模板的每一列都要有第七欄，不然 build 函式解包會炸。"""
    from fetcher_gaap import BS_TEMPLATE, CF_TEMPLATE, IS_TEMPLATE
    for T in (IS_TEMPLATE, BS_TEMPLATE, CF_TEMPLATE):
        for row in T:
            assert len(row) == 7, f"{row[0]} 不是 7-tuple（{len(row)} 欄）"


def test_template_label_fallback_is_passed_through_to_the_matcher():
    """第七欄要真的接到 `_match_is_row` 的第三層，不能只是加了欄位沒接線。"""
    df = _row_df("nvda_SomethingCompletelyCustom", "Widget purchases", float("nan"))
    assert _match_is_row(df, "NoSuchStd", "NoSuchConcept") is None
    assert _match_is_row(df, "NoSuchStd", "NoSuchConcept",
                         label_fallback="widget purchases") is not None


def test_capex_matches_nvda_custom_extension_tag_by_label():
    """NVDA 的 10-K 從 FY2013 到 FY2023 共 11 年用自訂 tag
    `nvda_PurchasesOfPropertyAndEquipmentAndIntangibleAssets`，concept 兩層都比不到。
    label 這十一年只有兩種寫法，都含「purchases ... property and equipment」。"""
    for label in ("Purchases of property and equipment and intangible assets",
                  "Purchases related to property and equipment and intangible assets"):
        df = _row_df("nvda_PurchasesOfPropertyAndEquipmentAndIntangibleAssets", label,
                     float("nan"))
        assert _match_template_row("Capex", df) == label


def test_capex_label_fallback_does_not_grab_proceeds_from_selling_property():
    """第三層很寬，處分不動產的**流入**不能被當成資本支出。"""
    df = _row_df("us-gaap_ProceedsFromSaleOfPropertyPlantAndEquipment",
                 "Proceeds from sales of property and equipment", float("nan"))
    assert _match_template_row("Capex", df) is None


def test_capex_label_fallback_does_not_grab_depreciation_of_property():
    df = _row_df("us-gaap_DepreciationDepletionAndAmortization",
                 "Depreciation of property and equipment", float("nan"))
    assert _match_template_row("Capex", df) is None


# ── 期末流通股數：10-K 的封面頁 fact 標 fp='FY'，要對到該財年的 Q4 ──────────

def test_shares_outstanding_maps_annual_cover_fact_to_q4():
    """43/52 家「中間有洞」的成因：10-K 的 dei fact 標 fp='FY'，對出來的標籤是
    `FY2025`，季表要的是 `FY2025Q4` → **每一年的 Q4 都是洞**。
    實測 AAPL/NVDA/WMT/COST/MU/ADBE 六家 Q1~Q3 全中、Q4 全空。
    數字取自 AAPL：FY2025Q3 的 14,840,390,000 與 10-K 封面的 14,776,353,000。"""
    from fetcher_gaap import _shares_map_from_records
    records = [
        {"fiscal_year": 2025, "fiscal_period": "Q3", "numeric_value": 14840390000},
        {"fiscal_year": 2025, "fiscal_period": "FY", "numeric_value": 14776353000},
    ]
    out = _shares_map_from_records(records)
    assert out["FY2025Q3"] == 14840390000
    assert out["FY2025"] == 14776353000      # 年表照舊
    assert out["FY2025Q4"] == 14776353000    # 季表的 Q4 補上


def test_shares_outstanding_explicit_q4_wins_over_annual_fact():
    """公司若真的另外標了 fp='Q4'，那筆比 10-K 封面頁的 FY 更貼近季末，不能被蓋掉。"""
    from fetcher_gaap import _shares_map_from_records
    records = [
        {"fiscal_year": 2025, "fiscal_period": "FY", "numeric_value": 222},
        {"fiscal_year": 2025, "fiscal_period": "Q4", "numeric_value": 111},
    ]
    assert _shares_map_from_records(records)["FY2025Q4"] == 111


# ═════════════════════════════════════════════════════════════════════════════
# D11-B：偵測到帳本有網路造成的缺漏，自動重試一次並合併結果
# ═════════════════════════════════════════════════════════════════════════════
#
# `_fetch_with_retry` / `_merge_retry_tables` / `_patch_meta_gap_note` 都是純函式
# （不碰網路），注入點讓測試把「跑一次抓取」換成假的。真正打網路的那層
# （`fetch_gaap_statements` 頂層的 `_ledger() is None` 分支）只是接線，
# 靠既有的 1120+ 個離線測試守住「沒有缺漏時行為完全不變」。

def _meta_table(quarter_labels, gap_notes):
    return StatementTable(sheet_name="Data_Meta", quarter_labels=quarter_labels,
                           filing_dates=[""] * len(quarter_labels),
                           concepts=["Ticker", "Fetch Gaps"],
                           values=[["AAPL"] * len(quarter_labels), gap_notes])


class _Boom(Exception):
    pass


def test_merge_retry_tables_fills_none_cells_from_retry():
    from fetcher_gaap import _merge_retry_tables
    orig = [StatementTable("Data_Financials(Q)", ["FY2025Q1", "FY2025Q2"], ["", ""],
                            ["Revenue"], [[100.0, None]])]
    retry = [StatementTable("Data_Financials(Q)", ["FY2025Q1", "FY2025Q2"], ["", ""],
                             ["Revenue"], [[999.0, 200.0]])]
    merged = _merge_retry_tables(orig, retry)
    assert merged[0].values == [[100.0, 200.0]], "orig 有值的保留，None 才用 retry 補"


def test_merge_retry_tables_skips_mismatched_shape():
    """期別對不上（理論上不該發生，但求穩）就不合併，保留原樣，不硬湊。"""
    from fetcher_gaap import _merge_retry_tables
    orig = [StatementTable("Data_Financials(Q)", ["FY2025Q1"], [""], ["Revenue"], [[None]])]
    retry = [StatementTable("Data_Financials(Q)", ["FY2025Q1", "FY2025Q2"], ["", ""],
                             ["Revenue"], [[1.0, 2.0]])]
    merged = _merge_retry_tables(orig, retry)
    assert merged[0].values == [[None]]


def test_merge_retry_tables_keeps_sheet_missing_from_retry_untouched():
    from fetcher_gaap import _merge_retry_tables
    orig = [StatementTable("Data_Meta", ["FY2025Q1"], [""], ["Ticker"], [["AAPL"]])]
    merged = _merge_retry_tables(orig, [])
    assert merged[0].values == [["AAPL"]]


def test_patch_meta_gap_note_reflects_absorbed_ledger():
    """重試吸收帳本之後，Data_Meta 的 Fetch Gaps 欄位要跟著更新，
    不然 Index 還是會顯示重試前的舊缺漏數。"""
    from fetcher_gaap import _patch_meta_gap_note
    from fetch_ledger import FetchLedger
    from i18n import t
    led = FetchLedger()   # 吸收後乾淨，沒有缺漏
    tbl = _meta_table(["FY2025Q1", "FY2025Q2"], ["stale note", "stale note"])
    _patch_meta_gap_note([tbl], led)
    assert tbl.values[1] == [t("xls.meta.none")] * 2


def test_fetch_with_retry_skips_retry_when_no_network_gaps():
    """沒有網路造成的缺漏就不重試——不浪費時間，也不該去 sleep、不該呼叫 retry_once。

    `led` 是呼叫端已經在用的那本帳本（不管是它自己開的還是外層 caller 開的）
    ——這個函式不負責開第一輪的帳本，只負責「要不要重試、重試完怎麼合併」。
    """
    from fetcher_gaap import _fetch_with_retry
    from fetch_ledger import FetchLedger

    led = FetchLedger(probe=lambda: True)   # 沒有任何 record，乾淨

    def _no_retry():
        raise AssertionError("不該呼叫 retry_once")

    def _no_sleep(_secs):
        raise AssertionError("不該 sleep")

    tables = _fetch_with_retry([_meta_table(["FY2025Q1"], [""])], led,
                                _no_retry, sleep=_no_sleep)
    assert not led.has_gaps


def test_fetch_with_retry_retries_once_on_network_gap_and_merges():
    from fetcher_gaap import _fetch_with_retry
    from fetch_ledger import FetchLedger

    led = FetchLedger(probe=lambda: False)   # 連不上 = network
    led.record("FY2025Q2", _Boom("x"))
    first_tables = [StatementTable("Data_Financials(Q)", ["FY2025Q1", "FY2025Q2"], ["", ""],
                                    ["Revenue"], [[100.0, None]])]

    def retry_once():
        retry_led = FetchLedger(probe=lambda: False)   # 這輪沒有失敗
        tbl = StatementTable("Data_Financials(Q)", ["FY2025Q1", "FY2025Q2"], ["", ""],
                              ["Revenue"], [[100.0, 200.0]])
        return [tbl], retry_led

    sleeps = []
    tables = _fetch_with_retry(first_tables, led, retry_once, sleep=sleeps.append)
    assert tables[0].values == [[100.0, 200.0]]
    assert not led.has_gaps, "救回來了，帳本要吸收掉這筆缺漏"
    assert sleeps, "重試前要退避一下，不要立刻連打"


def test_fetch_with_retry_only_retries_once():
    """重試那輪還是有網路缺漏，也不再重試第三次——避免網路真的斷線時拖很久。"""
    from fetcher_gaap import _fetch_with_retry
    from fetch_ledger import FetchLedger

    led = FetchLedger(probe=lambda: False)
    led.record("FY2025Q2", _Boom("x"))
    first_tables = [StatementTable("Data_Financials(Q)", ["FY2025Q1", "FY2025Q2"], ["", ""],
                                    ["Revenue"], [[100.0, None]])]

    calls = []

    def retry_once():
        calls.append(1)
        retry_led = FetchLedger(probe=lambda: False)
        retry_led.record("FY2025Q2", _Boom("x"))   # 重試還是失敗
        tbl = StatementTable("Data_Financials(Q)", ["FY2025Q1", "FY2025Q2"], ["", ""],
                              ["Revenue"], [[100.0, None]])
        return [tbl], retry_led

    _fetch_with_retry(first_tables, led, retry_once, sleep=lambda _s: None)
    assert len(calls) == 1, "重試只呼叫一次 retry_once，不會再多"
    assert led.has_gaps, "兩輪都失敗，帳本仍要記著"


def test_fetch_with_retry_falls_back_to_first_pass_if_retry_itself_blows_up():
    """重試本身意外整個炸掉（不是「這期又沒抓到」那種正常缺漏，是重試路徑
    自己出錯）不該把「第一輪已經抓到大半資料」變成整趟失敗——保留第一輪
    的結果，帳本維持原本記的缺漏，不要讓 D11-B 反而讓事情變更糟。"""
    from fetcher_gaap import _fetch_with_retry
    from fetch_ledger import FetchLedger

    led = FetchLedger(probe=lambda: False)
    led.record("FY2025Q2", _Boom("x"))
    first_tables = [StatementTable("Data_Financials(Q)", ["FY2025Q1", "FY2025Q2"], ["", ""],
                                    ["Revenue"], [[100.0, None]])]

    def retry_once():
        raise RuntimeError("重試路徑本身炸了")

    tables = _fetch_with_retry(first_tables, led, retry_once, sleep=lambda _s: None)
    assert tables[0].values == [[100.0, None]], "退回第一輪的結果，不憑空丟資料"
    assert led.has_gaps, "帳本維持原本記的缺漏，不要假裝救回來了"


def test_fetch_with_retry_does_not_retry_data_kind_gaps():
    """資料類缺漏（SEC 有回應，是這份資料本身的性質）重試沒用，不該白白多打一輪。"""
    from fetcher_gaap import _fetch_with_retry
    from fetch_ledger import FetchLedger

    led = FetchLedger(probe=lambda: True)   # 連得上 = data
    led.record("FY2025Q2", _Boom("x"))
    first_tables = [_meta_table(["FY2025Q1", "FY2025Q2"], ["", ""])]

    def _no_retry():
        raise AssertionError("data 類不該觸發重試")

    def _no_sleep(_secs):
        raise AssertionError("data 類不該 sleep")

    tables = _fetch_with_retry(first_tables, led, _no_retry, sleep=_no_sleep)
    assert led.has_gaps


# ── G6：單一公司輸出，抓不到的季度留一整欄空白（2026-08-25）───────────────

def _q_tbl(sheet, labels, ends, values):
    return StatementTable(
        sheet_name=sheet, quarter_labels=list(labels),
        filing_dates=[""] * len(labels),
        concepts=["Revenue"], values=[list(values)], labels=[""],
        period_ends=list(ends),
    )


def _merged_labels(labels, ends, values=None):
    values = values if values is not None else [1.0] * len(labels)
    merged = _merge_financials(
        _q_tbl("Data_IS", labels, ends, values),
        _q_tbl("Data_BS", labels, ends, values),
        _q_tbl("Data_CF", labels, ends, values),
    )
    return merged


def test_merge_financials_keeps_a_column_for_a_quarter_that_was_never_fetched():
    """某一季掛掉整欄消失，畫面上 FY2025Q1 直接跳到 FY2025Q3，使用者與 AI 都
    看不出中間漏了一季。改成保留欄位、內容全空，讓「有漏」看得見。"""
    merged = _merged_labels(["FY2025Q1", "FY2025Q3"],
                            ["2025-03-29", "2025-09-27"])

    assert merged.quarter_labels == ["FY2025Q1", "FY2025Q2", "FY2025Q3"]
    rev = merged.values[merged.concepts.index("Revenue")]
    assert rev[1] is None


def test_merge_financials_gap_column_falls_back_to_a_derived_period_end():
    """補出來那一欄沒有真實期末日，只能用財季標籤反推年月（`2025-06`）。"""
    merged = _merged_labels(["FY2025Q1", "FY2025Q3"],
                            ["2025-03-29", "2025-09-27"])
    end_row = merged.values[merged.concepts.index("Period End")]

    assert end_row[0] == "2025-03-29"
    assert end_row[1] == "2025-06"          # 反推，不是編一個假的完整日期
    assert end_row[2] == "2025-09-27"


def test_merge_financials_gap_column_rolls_over_into_the_next_fiscal_year():
    merged = _merged_labels(["FY2025Q4", "FY2026Q2"],
                            ["2024-12-28", "2025-06-28"])
    assert merged.quarter_labels == ["FY2025Q4", "FY2026Q1", "FY2026Q2"]


def test_merge_financials_does_not_invent_a_gap_for_a_sixteen_week_quarter():
    """COSTCO 的第四季是 16 週（112~119 天）。固定門檻會把它誤判成缺一季，
    `round(112/91) = 1` 才是對的——52 家實測 111~150 天那 16 筆全是 COSTCO。"""
    merged = _merged_labels(
        ["FY2024Q1", "FY2024Q2", "FY2024Q3", "FY2024Q4"],
        ["2023-11-26", "2024-02-18", "2024-05-12", "2024-09-01"])
    assert merged.quarter_labels == ["FY2024Q1", "FY2024Q2", "FY2024Q3", "FY2024Q4"]


def test_merge_financials_never_generates_more_than_four_gap_columns():
    """實測沒有任何 >210 天的缺口，真的出現就是資料異常，不該無限生欄。"""
    merged = _merged_labels(["FY2016Q1", "FY2025Q3"],
                            ["2016-03-26", "2025-09-27"])
    assert merged.quarter_labels == ["FY2016Q1", "FY2025Q3"]


def test_merge_financials_leaves_annual_tables_alone():
    """年報欄位是 FY2025 這種年度標籤，91 天那套季度算法不適用。"""
    merged = _merge_financials(
        _q_tbl("Data_IS", ["FY2022", "FY2024"], ["2022-12-31", "2024-12-28"], [1.0, 2.0]),
        _q_tbl("Data_BS", ["FY2022", "FY2024"], ["2022-12-31", "2024-12-28"], [1.0, 2.0]),
        _q_tbl("Data_CF", ["FY2022", "FY2024"], ["2022-12-31", "2024-12-28"], [1.0, 2.0]),
        sheet_name="Data_Financials(Y)",
    )
    assert merged.quarter_labels == ["FY2022", "FY2024"]


def test_merge_financials_does_not_touch_a_complete_quarter_sequence():
    """沒有缺口的公司，輸出必須一格都不變。"""
    labels = ["FY2025Q1", "FY2025Q2", "FY2025Q3"]
    ends = ["2025-03-29", "2025-06-28", "2025-09-27"]
    merged = _merged_labels(labels, ends, [1.0, 2.0, 3.0])
    assert merged.quarter_labels == labels
    assert merged.values[merged.concepts.index("Revenue")] == [1.0, 2.0, 3.0]


# ── H6 label_hint 擴充（2026-08-25，201 家掃描的三項真缺口 + Cost of Revenue 子集）──
#
# 每個 label 都是從 `output/_hintsweep_201/hintsweep_201_result.txt` 逐字抄的真實
# 措辭（201 家最新 10-Q 實測），不是編的。四類各抽一家用 `scripts/diag_rowprobe.py`
# 回頭核對過原始 10-Q 的 dataframe（UNP / IP / LIN / EXC）。
# 分類與證據見 `output/_hintsweep_201/classification.md` 與 TODO H6。


@pytest.mark.parametrize("concept,label", [
    ("us-gaap_PaymentsToAcquireProductiveAssets",          "Capital spending"),                    # F / KMB / PEP
    ("us-gaap_PaymentsToAcquirePropertyPlantAndEquipment", "Capital investments"),                 # UNP
    ("us-gaap_PaymentsToAcquirePropertyPlantAndEquipment", "Capital additions (including software)"),   # HSY
    ("us-gaap_PaymentsToAcquirePropertyPlantAndEquipment", "Capital and technology expenditures"),  # MAR
    ("us-gaap_PaymentsToAcquirePropertyPlantAndEquipment", "Purchase of premises and equipment, net of sales"),  # AXP
    ("us-gaap_PaymentsToAcquirePropertyPlantAndEquipment", "Purchases of premises and equipment/capitalized software"),  # BK
    ("us-gaap_PaymentsToAcquirePropertyPlantAndEquipment", "Changes in premises and equipment"),    # COF
    ("us-gaap_PaymentsToAcquirePropertyPlantAndEquipment", "Additions to plant and equipment, including long-term deposits"),  # APD
    ("us-gaap_PaymentsToAcquirePropertyPlantAndEquipment", "Additions to plant and equipment"),     # ITW
    ("us-gaap_PaymentsToAcquirePropertyPlantAndEquipment", "Purchases of land, buildings, and equipment"),  # GIS
    ("us-gaap_PaymentsToAcquireProductiveAssets",          "Purchase of land, buildings, equipment and software"),  # AMP
    ("us-gaap_PaymentsToAcquireProductiveAssets",          "Acquisitions of Generation Facilities"),  # AEP
    ("us-gaap_PaymentsToAcquirePropertyPlantAndEquipment", "Purchases of property and equipment"),  # AMD（原本就過，防改壞）
])
def test_capex_matches_regardless_of_label_wording(concept, label):
    """H3 把 hint 改成 `propert|capital expenditure` 後，201 家裡 14 家全損——
    這些公司寫 Capital spending／investments／premises and equipment，兩個詞根一個都不含。
    concept 層本來就對得上（`PaymentsToAcquire(Productive|PropertyPlantAndEquipment)`）。"""
    df = _row_df(concept, label, "CapitalExpenses")
    assert _match_template_row("Capex", df) == label


def test_capex_skips_the_accrued_but_not_paid_row_unp():
    """UNP／AMD 的現金流量表有兩列 `std_concept=CapitalExpenses`，第二列是
    「已發生但尚未付款」的非現金揭露，**加總會重複計算**（TODO G10 記過）。
    實測 UNP 2026-07-23：[17] Capital investments、[32] Capital investments accrued
    but not yet paid。hint 放寬後仍然不可以挑到後者。"""
    df = _row_df("us-gaap_PaymentsToAcquirePropertyPlantAndEquipment", "Capital investments",
                 "CapitalExpenses",
                 extra=[("us-gaap_CapitalExpendituresIncurredButNotYetPaid",
                         "Capital investments accrued but not yet paid", "CapitalExpenses")])
    assert _match_template_row("Capex", df) == "Capital investments"


def test_capex_does_not_match_an_accrued_only_row():
    """整張表只有「accrued but not yet paid」那列時要留空，不可以拿它當 Capex。"""
    df = _row_df("us-gaap_CapitalExpendituresIncurredButNotYetPaid",
                 "Capital investments accrued but not yet paid", "CapitalExpenses")
    assert _match_template_row("Capex", df) is None


@pytest.mark.parametrize("concept,label", [
    ("us-gaap_CashAndCashEquivalentsAtCarryingValue", "Cash"),                              # ETN
    ("us-gaap_Cash",                                  "Cash"),                              # SLB
    ("us-gaap_CashAndCashEquivalentsAtCarryingValue", "Cash and cash items"),               # APD
    ("us-gaap_CashAndCashEquivalentsAtCarryingValue", "Cash and temporary investments"),    # IP
    ("us-gaap_CashAndCashEquivalentsAtCarryingValue", "Cash and temporary cash investments"),  # KR
    ("us-gaap_CashAndCashEquivalentsAtCarryingValue", "Cash and cash equivalents"),         # 多數公司（防改壞）
])
def test_cash_matches_non_standard_wording(concept, label):
    """std_concept 已經正確命中，卻被 hint 的字面要求濾掉——hint 在這裡幫倒忙。

    SLB 那筆的 std_concept 也是 `CashAndMarketableSecurities`（2026-07-29 那份 BS
    的 [2] 列實測）——TODO H6 原本記「SLB 退到 fallback_suffix 層」是錯的，
    `us-gaap_Cash` 這個 concept 一樣被 edgartools 標成標準 std_concept。
    """
    df = _row_df(concept, label, "CashAndMarketableSecurities")
    assert _match_template_row("Cash", df) == label


def test_cash_still_excludes_cash_and_due_from_banks():
    """銀行的 `CashAndDueFromBanks`（AXP/BAC/BK/C/COF/JPM/WFC 7 家）是**概念取捨題**，
    CTH 還沒決定要不要算進 Cash——H6 刻意不動它，放寬 hint 時不可以順手吃進來。"""
    df = _row_df("us-gaap_CashAndDueFromBanks", "Cash and due from banks", float("nan"))
    assert _match_template_row("Cash", df) is None


@pytest.mark.parametrize("label", [
    "Ordinary shares, value",                                          # ACN
    "Ordinary shares - $0.01 nominal value",                           # AON
    "Common Shares (CHF 0.50 par value; 435,331,832 shares issued)",   # CB
    "Ordinary shares (388.4 million outstanding)",                     # ETN
    "Ordinary shares, $0.01 par value",                                # JCI
    "Ordinary shares, €0.001 par value, authorized 1,750,000,000 shares",  # LIN
    "Ordinary shares— par value $0.0001",                              # MDT
    "Common stock, $0.00001 par value",                                # 多數美國公司（防改壞）
])
def test_common_stock_apic_matches_ordinary_and_common_shares(label):
    """愛爾蘭／英國／瑞士註冊（或改遷冊）的公司寫 Ordinary shares，
    現行 hint 只認 common stock / paid-in capital，7 家全損。concept 全部是
    `CommonStockValue`，本來就對得上。"""
    df = _row_df("us-gaap_CommonStockValue", label, "CommonEquity")
    assert _match_template_row("Common Stock & APIC", df) == label


@pytest.mark.parametrize("concept,label", [
    # NSC 2026 Q2：普通股那一列自己就寫「net of treasury shares」。排除 treasury 時
    # 用「label 含 treasury 就踢掉」會誤傷它——201 家重掃時實際踩到的回歸。
    ("us-gaap_CommonStockValueOutstanding", "Common stock, net of treasury shares"),
    ("us-gaap_CommonStockValue", "Common shares, net of treasury"),
])
def test_common_stock_apic_keeps_rows_that_merely_mention_treasury(concept, label):
    df = _row_df(concept, label, "CommonEquity")
    assert _match_template_row("Common Stock & APIC", df) == label


@pytest.mark.parametrize("label", [
    "Less: Treasury shares, at cost (2026 - 29,786,809 shares)",   # LIN
    "Treasury shares, at cost (249,886,450 shares)",               # AMP
    "Common shares held in treasury, at cost - Shares: 267,907,258",  # ABT
    "Common shares in treasury, at cost, 1,305 shares",            # KR
    "Treasury stock (at cost: 2026-1,054,626,440 shares)",         # COP
    # 下面兩個沒有「at cost」字樣，靠 word boundary 那幾條分支擋。2026-08-25 踩過：
    # 用 heredoc 寫 regex 時 `` 被當跳脫字元寫成真的 backspace 位元組，那幾條
    # 分支整個失效，而上面帶 at cost 的案例照樣過——這兩筆才擋得住那種壞法。
    "Common shares in treasury",
    "Treasury common shares",
])
def test_common_stock_apic_does_not_pick_real_treasury_rows(label):
    """這幾家的庫藏股列 std_concept 也是 `CommonEquity`（實測 LIN [28]、ABT [51]、
    AMP [77]、KR [37]），放寬 hint 後不可以挑到它們。"""
    df = _row_df("us-gaap_TreasuryStockCommonValue", label, "CommonEquity")
    assert _match_template_row("Common Stock & APIC", df) is None


def test_common_stock_apic_does_not_pick_treasury_shares():
    """LIN 實測：同一張 BS 上 `TreasuryStockCommonValue` 的 std_concept 也是
    `CommonEquity`（[28] Less: Treasury shares, at cost）。放寬 hint 後不可以挑到它。"""
    df = _row_df("us-gaap_TreasuryStockCommonValue",
                 "Less: Treasury common shares, at cost (29,786,809 shares)", "CommonEquity")
    assert _match_template_row("Common Stock & APIC", df) is None


@pytest.mark.parametrize("concept,label", [
    ("us-gaap_CostOfGoodsAndServicesSold", "Purchased crude oil and products"),   # CVX / PSX
    ("us-gaap_CostOfGoodsAndServicesSold", "Purchased commodities"),              # COP
    ("us-gaap_CostOfGoodsAndServiceExcludingDepreciationDepletionAndAmortization",
     "Purchased Electricity, Fuel and Other Consumables Used for Electric Generation"),  # AEP
    ("us-gaap_CostDirectMaterial",         "Purchased power and/or fuel"),        # EXC
    ("us-gaap_CostDirectMaterial",         "Food, beverage and packaging"),       # CMG
    ("us-gaap_CostOfGoodsAndServicesSold", "Cost of sales"),                      # 多數公司（防改壞）
])
def test_cost_of_revenue_matches_energy_utility_and_restaurant_wording(concept, label):
    """採購原油／電力燃料／食材包材是這幾家真實的 COGS 對應項，目前被 hint 誤擋。"""
    df = _row_df(concept, label, "CostOfGoodsAndServicesSold")
    assert _match_template_row("Cost of Revenue", df) == label


@pytest.mark.parametrize("concept,label", [
    ("us-gaap_LaborAndRelatedExpense", "Compensation and benefits"),        # AXP/BAC/C/BK/BLK/CME/COF
    ("us-gaap_LaborAndRelatedExpense", "Salaries and employee benefits"),
    ("us-gaap_LaborAndRelatedExpense", "Labor and Fringe"),                 # CSX
    ("us-gaap_DirectCostsOfLeasedAndRentedPropertyOrEquipment", "Rental expense"),  # AMT/CCI
])
def test_cost_of_revenue_still_excludes_labour_and_rent(concept, label):
    """銀行／保險／鐵路／REIT 概念上沒有 Cost of Revenue（D8 同一類），維持空白才對。
    放寬 hint 吃進人事費就是**製造錯誤數字**，比留空更糟。"""
    df = _row_df(concept, label, "CostOfGoodsAndServicesSold")
    assert _match_template_row("Cost of Revenue", df) is None


def test_cost_of_revenue_prefers_purchased_fuel_over_operating_and_maintenance_exc():
    """EXC 實測：同一個 std_concept 底下有 [208] Purchased power and/or fuel 與
    [246] Operating and maintenance 兩列。hint 要挑前者——這正是這一列不能乾脆
    拿掉 hint 的理由。"""
    df = _row_df("us-gaap_CostDirectMaterial", "Purchased power and/or fuel",
                 "CostOfGoodsAndServicesSold",
                 extra=[("us-gaap_UtilitiesOperatingExpenseMaintenanceOperationsAndOtherCostsAndExpenses",
                         "Operating and maintenance", "CostOfGoodsAndServicesSold")])
    assert _match_template_row("Cost of Revenue", df) == "Purchased power and/or fuel"
