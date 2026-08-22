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
    assert result.metrics["Revenue"]["NVDA"]["2024Q1"] == 100.0
    assert result.metrics["Revenue"]["AMD"]["2024Q2"] == 90.0
    assert result.period_ends["NVDA"]["2024Q1"] == "2024-03-31"
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
    assert result.metrics["Gross Margin (%)"]["NVDA"]["2024Q1"] == pytest.approx(50.0)
    assert result.metrics["Gross Margin (%)"]["NVDA"]["2024Q2"] == pytest.approx(40.0)


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

    assert result.metrics["Revenue"]["NVDA"]["2024Q1"] == 100.0
    assert "BADTICKER" not in result.metrics["Revenue"]
    assert len(result.failures) == 1
    assert result.failures[0] == CompanyFetchError(ticker="BADTICKER", error_type="ValueError")


def test_build_comparison_quarterly_frequency_still_fetches_annual_for_q4_synthesis():
    """2026-08-21 CTH 回報：跨公司比較的季度資料中間會有一大截缺 Q4。根因是
    frequency='quarterly' 時 fetch_annual 被算成 False，年報完全不抓，
    D0-1 的 Q4 合成（fetcher_gaap._synthesize_q4()，靠「年報－Q1－Q2－Q3」
    湊出 Q4）沒有材料可用，整個跳過。單一公司 Tab1 抓取用兩個獨立勾選框
    （fetch_q／fetch_k 預設都 True），季度模式年報照樣抓；跨公司比較誤把
    「顯示頻率」跟「該不該抓年報」綁成互斥單選，才會漏了這步。修法：季度
    模式也要把 fetch_annual 傳 True，年報只當 Q4 合成材料用，不影響輸出頻率。"""
    captured_kwargs = {}

    def fake_fetch(ticker, identity, **kwargs):
        captured_kwargs.update(kwargs)
        return [_fake_q_table(ticker, [100.0], [50.0], ["2024-03-31"])]

    with patch("comparison.fetch_gaap_statements", side_effect=fake_fetch):
        build_comparison(
            ["NVDA"], "test@example.com", ["Revenue"],
            frequency="quarterly", start_year=None, end_year=None,
        )

    assert captured_kwargs["fetch_quarterly"] is True
    assert captured_kwargs["fetch_annual"] is True


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

    assert result.metrics["Revenue"]["NVDA"]["2024"] == 400.0


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

    assert list(result.metrics["Revenue"]["NVDA"].keys()) == ["2023Q1"]


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


# ── 跨公司改用日曆季對齊（2026-08-22）──────────────────────────────────────
#
# 舊版直接拿各公司自己的財季標籤（FY2026Q2）當共同欄位。NVDA 一月結算，
# 它的 FY2026Q2 實際結束在 2025-07-27，卻跟 AMD 結束在 2026-06-27 的
# FY2026Q2 疊在同一欄——NVDA 整條線在日曆時間上偏移約一年。
# 改成用期中點算出的日曆季當欄位鍵。

def _q_table(ticker, revenue, period_ends, labels):
    return StatementTable(
        sheet_name="Data_Financials(Q)",
        quarter_labels=labels,
        filing_dates=[""] * len(revenue),
        concepts=["Revenue"],
        values=[revenue],
        ticker=ticker,
        labels=[""],
        period_ends=period_ends,
    )


def test_build_comparison_aligns_companies_by_calendar_quarter():
    """NVDA 7 月底那一季要跟 AMD 6 月底那一季同欄（同一波財報）。"""
    def fake_fetch(ticker, identity, **kwargs):
        data = {
            # NVDA 財年二月起算：這兩季是 FY2026 的 Q2、Q3
            "NVDA": _q_table("NVDA", [46743.0, 57006.0],
                             ["2025-07-27", "2025-10-26"],
                             ["FY2026Q2", "FY2026Q3"]),
            # AMD 日曆年：同一波財報是 2025 的 Q2、Q3
            "AMD": _q_table("AMD", [7685.0, 9246.0],
                            ["2025-06-28", "2025-09-27"],
                            ["FY2025Q2", "FY2025Q3"]),
        }
        return [data[ticker]]

    with patch("comparison.fetch_gaap_statements", side_effect=fake_fetch):
        result = build_comparison(
            ["NVDA", "AMD"], "test@example.com", ["Revenue"],
            frequency="quarterly", start_year=None, end_year=None,
        )

    rev = result.metrics["Revenue"]
    assert rev["NVDA"]["2025Q2"] == 46743.0
    assert rev["AMD"]["2025Q2"] == 7685.0
    assert rev["NVDA"]["2025Q3"] == 57006.0
    assert rev["AMD"]["2025Q3"] == 9246.0
    assert result.period_ends["NVDA"]["2025Q2"] == "2025-07-27"
    assert result.period_ends["AMD"]["2025Q2"] == "2025-06-28"


def test_build_comparison_annual_aligns_by_calendar_year():
    """NVDA FY2026（結束 2026-01）內容是日曆 2025 年，要跟 AMD FY2025 同欄。"""
    def fake_fetch(ticker, identity, **kwargs):
        data = {
            "NVDA": StatementTable(
                sheet_name="Data_Financials(Y)", quarter_labels=["FY2026"],
                filing_dates=[""], concepts=["Revenue"], values=[[130497.0]],
                ticker="NVDA", labels=[""], period_ends=["2026-01-25"]),
            "AMD": StatementTable(
                sheet_name="Data_Financials(Y)", quarter_labels=["FY2025"],
                filing_dates=[""], concepts=["Revenue"], values=[[32639.0]],
                ticker="AMD", labels=[""], period_ends=["2025-12-27"]),
        }
        return [data[ticker]]

    with patch("comparison.fetch_gaap_statements", side_effect=fake_fetch):
        result = build_comparison(
            ["NVDA", "AMD"], "test@example.com", ["Revenue"],
            frequency="annual", start_year=None, end_year=None,
        )

    assert result.metrics["Revenue"]["NVDA"]["2025"] == 130497.0
    assert result.metrics["Revenue"]["AMD"]["2025"] == 32639.0


def test_build_comparison_keeps_fiscal_label_when_period_end_missing():
    """期末日抓不到就退回原本的財季標籤，不因為算不出日曆季就整欄丟掉。"""
    def fake_fetch(ticker, identity, **kwargs):
        return [_q_table("NVDA", [100.0], [""], ["FY2026Q2"])]

    with patch("comparison.fetch_gaap_statements", side_effect=fake_fetch):
        result = build_comparison(
            ["NVDA"], "test@example.com", ["Revenue"],
            frequency="quarterly", start_year=None, end_year=None,
        )

    assert result.metrics["Revenue"]["NVDA"]["FY2026Q2"] == 100.0
