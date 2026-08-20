"""Tests for comparison.py — 跨公司比較的資料重組。"""
from unittest.mock import patch

import pytest

from fetcher_gaap import StatementTable
from comparison import build_comparison, ComparisonResult, CompanyFetchError


def _fake_q_table(ticker, revenue, gross_profit, period_ends):
    labels = [f"FY2024Q{i+1}" for i in range(len(revenue))]
    return StatementTable(
        sheet_name="Data_Financials(Q)",
        quarter_labels=labels,
        filing_dates=[""] * len(revenue),
        concepts=["Revenue", "Gross Profit"],
        values=[revenue, gross_profit],
        ticker=ticker,
        labels=["", ""],
        period_ends=period_ends,
    )


def test_build_comparison_extracts_raw_concept_across_companies():
    def fake_fetch(ticker, identity, **kwargs):
        data = {
            "NVDA": _fake_q_table("NVDA", [100.0, 110.0], [50.0, 60.0],
                                  ["2024-03-31", "2024-06-30"]),
            "AMD": _fake_q_table("AMD", [80.0, 90.0], [30.0, 35.0],
                                 ["2024-03-31", "2024-06-30"]),
        }
        return [data[ticker]]

    with patch("comparison.fetch_gaap_statements", side_effect=fake_fetch):
        result = build_comparison(
            ["NVDA", "AMD"], "test@example.com", ["Revenue"],
            frequency="quarterly", start_year=None, end_year=None,
        )

    assert isinstance(result, ComparisonResult)
    assert result.metrics["Revenue"]["NVDA"]["FY2024Q1"] == 100.0
    assert result.metrics["Revenue"]["AMD"]["FY2024Q2"] == 90.0
    assert result.period_ends["NVDA"]["FY2024Q1"] == "2024-03-31"
    assert result.failures == []


def test_build_comparison_extracts_ratio_metric():
    def fake_fetch(ticker, identity, **kwargs):
        return [_fake_q_table(ticker, [100.0, 200.0], [50.0, 80.0],
                              ["2024-03-31", "2024-06-30"])]

    with patch("comparison.fetch_gaap_statements", side_effect=fake_fetch):
        result = build_comparison(
            ["NVDA"], "test@example.com", ["Gross Margin (%)"],
            frequency="quarterly", start_year=None, end_year=None,
        )

    # Gross Margin = Gross Profit / Revenue: 50/100=50%, 80/200=40%
    assert result.metrics["Gross Margin (%)"]["NVDA"]["FY2024Q1"] == pytest.approx(50.0)
    assert result.metrics["Gross Margin (%)"]["NVDA"]["FY2024Q2"] == pytest.approx(40.0)


def test_build_comparison_skips_failed_company_and_continues():
    def fake_fetch(ticker, identity, **kwargs):
        if ticker == "BADTICKER":
            raise ValueError("No 10-Q filings found for ticker 'BADTICKER'.")
        return [_fake_q_table(ticker, [100.0], [50.0], ["2024-03-31"])]

    with patch("comparison.fetch_gaap_statements", side_effect=fake_fetch):
        result = build_comparison(
            ["NVDA", "BADTICKER"], "test@example.com", ["Revenue"],
            frequency="quarterly", start_year=None, end_year=None,
        )

    assert result.metrics["Revenue"]["NVDA"]["FY2024Q1"] == 100.0
    assert "BADTICKER" not in result.metrics["Revenue"]
    assert len(result.failures) == 1
    assert result.failures[0] == CompanyFetchError(ticker="BADTICKER", error_type="ValueError")


def test_build_comparison_annual_frequency_reads_data_financials_y():
    def fake_fetch(ticker, identity, **kwargs):
        q_tbl = _fake_q_table(ticker, [100.0], [50.0], ["2024-03-31"])
        y_tbl = StatementTable(
            sheet_name="Data_Financials(Y)", quarter_labels=["FY2024"],
            filing_dates=[""], concepts=["Revenue"], values=[[400.0]],
            ticker=ticker, labels=[""], period_ends=["2024-12-31"],
        )
        return [q_tbl, y_tbl]

    with patch("comparison.fetch_gaap_statements", side_effect=fake_fetch):
        result = build_comparison(
            ["NVDA"], "test@example.com", ["Revenue"],
            frequency="annual", start_year=None, end_year=None,
        )

    assert result.metrics["Revenue"]["NVDA"]["FY2024"] == 400.0


def test_build_comparison_filters_by_year_range():
    def fake_fetch(ticker, identity, **kwargs):
        labels = ["FY2022Q4", "FY2023Q1", "FY2024Q1"]
        return [StatementTable(
            sheet_name="Data_Financials(Q)", quarter_labels=labels,
            filing_dates=[""] * 3, concepts=["Revenue"], values=[[10.0, 20.0, 30.0]],
            ticker=ticker, labels=[""], period_ends=["2022-12-31", "2023-03-31", "2024-03-31"],
        )]

    with patch("comparison.fetch_gaap_statements", side_effect=fake_fetch):
        result = build_comparison(
            ["NVDA"], "test@example.com", ["Revenue"],
            frequency="quarterly", start_year=2023, end_year=2023,
        )

    assert list(result.metrics["Revenue"]["NVDA"].keys()) == ["FY2023Q1"]


def test_build_comparison_calls_on_company_start_for_every_ticker_in_order():
    """先前抓取跨公司比較時看不到進度像卡死——每家公司開始抓之前要先回報
    一次，就算後面失敗了也要回報過（使用者才知道正在試哪一家）。"""
    def fake_fetch(ticker, identity, **kwargs):
        if ticker == "BADTICKER":
            raise ValueError("boom")
        return [_fake_q_table(ticker, [100.0], [50.0], ["2024-03-31"])]

    calls = []

    def on_company_start(ticker, current, total):
        calls.append((ticker, current, total))

    with patch("comparison.fetch_gaap_statements", side_effect=fake_fetch):
        build_comparison(
            ["NVDA", "BADTICKER"], "test@example.com", ["Revenue"],
            frequency="quarterly", start_year=None, end_year=None,
            on_company_start=on_company_start,
        )

    assert calls == [("NVDA", 1, 2), ("BADTICKER", 2, 2)]


def test_build_comparison_forwards_filing_level_progress_from_fetcher_gaap():
    """單一公司抓好幾份 filing 時要看得到細部進度，不是整段空白直到抓完。"""
    import fetcher_gaap

    def fake_fetch(ticker, identity, **kwargs):
        fetcher_gaap._set_progress_total(2)
        fetcher_gaap._tick_progress(f"{ticker} 第一份")
        fetcher_gaap._tick_progress(f"{ticker} 第二份")
        return [_fake_q_table(ticker, [100.0], [50.0], ["2024-03-31"])]

    ticks = []

    def progress_cb(current, total, label):
        ticks.append((current, total, label))

    with patch("comparison.fetch_gaap_statements", side_effect=fake_fetch):
        build_comparison(
            ["NVDA"], "test@example.com", ["Revenue"],
            frequency="quarterly", start_year=None, end_year=None,
            progress_cb=progress_cb,
        )

    assert ticks == [(1, 2, "NVDA 第一份"), (2, 2, "NVDA 第二份")]
