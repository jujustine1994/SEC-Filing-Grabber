"""scripts/audit_8k_period_labels.py 的抽取規則測試。

docs/8k-period-off-by-one.md 的所有數字都是這幾個函式算出來的，規則一改
結論就變，所以把實測過的措辭釘成測試。字串全部取自 scratchpad 的原文快取。
"""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from audit_8k_period_labels import (  # noqa: E402
    fy_end_month_from,
    gaap_style_label,
    stated_fiscal_quarter,
    stated_period_end,
)


@pytest.mark.parametrize("text, expected", [
    # 非 12 月結算：公司自己會寫 fiscal
    ("Broadcom Inc. Announces Third Quarter Fiscal Year 2024 Financial Results",
     "FY2024Q3"),
    ("Palo Alto Networks Reports Fiscal Fourth Quarter and Fiscal Year 2024",
     "FY2024Q4"),
    # 12 月結算：不寫 fiscal
    ("Arlo Reports Third Quarter 2024 Results", "FY2024Q3"),
    ("ServiceNow Reports Third Quarter 2024 Financial Results", "FY2024Q3"),
    # INTC / AMD 用連字號
    ("Intel Reports Third-Quarter 2024 Financial Results", "FY2024Q3"),
    ("Intel Reports Fourth-Quarter and Full-Year 2024 Financial Results",
     "FY2024Q4"),
])
def test_stated_fiscal_quarter(text, expected):
    assert stated_fiscal_quarter(text) == expected


def test_stated_fiscal_quarter_none_when_absent():
    assert stated_fiscal_quarter("Company announces dividend") is None


@pytest.mark.parametrize("text, expected", [
    ("results for the quarter ended September 29, 2024", "2024-09-29"),
    ("the 52-week fiscal year ended September 1, 2024", "2024-09-01"),
    ("fiscal 2024, which ended August 29, 2024", "2024-08-29"),
    ("fiscal fourth quarter and fiscal year ended July 31, 2024", "2024-07-31"),
])
def test_stated_period_end(text, expected):
    assert stated_period_end(text) == expected


def test_period_end_takes_latest_not_first():
    """新聞稿常把去年同期排在前面，取第一個會抓到比較期間。"""
    text = ("per diluted share, for the quarter ended June 30, 2024 ... "
            "financial results for the quarter ended September 29, 2024")
    assert stated_period_end(text) == "2024-09-29"


def test_period_end_ignores_dates_after_the_release():
    """ARLO 的 Q3 新聞稿裡有「fiscal year ended December 31」這種講法，
    比本季期末晚，會把財年結束月推成 3 月（實際 12 月），整家比對全錯。"""
    text = ("third quarter ended September 29, 2024 ... "
            "for the fiscal year ended December 31, 2024")
    assert stated_period_end(text, not_after="2024-11-07") == "2024-09-29"


@pytest.mark.parametrize("period_end, stated, expected", [
    ("2026-04-26", "FY2027Q1", 1),    # NVDA：1 月結算
    ("2025-12-27", "FY2026Q1", 9),    # AAPL：9 月結算
    ("2026-05-10", "FY2026Q3", 8),    # COST：8 月結算
    ("2026-06-28", "FY2026Q2", 12),   # ARLO：12 月結算
])
def test_fy_end_month_from(period_end, stated, expected):
    assert fy_end_month_from(period_end, stated) == expected


@pytest.mark.parametrize("period_end, fy_end, expected", [
    ("2026-04-26", 1, "FY2027Q1"),    # NVDA
    ("2025-12-27", 9, "FY2026Q1"),    # AAPL
    ("2026-06-28", 12, "FY2026Q2"),   # ARLO
    ("2026-06-30", 6, "FY2026Q4"),    # MSFT
])
def test_gaap_style_label(period_end, fy_end, expected):
    assert gaap_style_label(period_end, fy_end) == expected


def test_gaap_style_label_matches_fetcher_gaap_convention():
    """與 fetcher_gaap._col_to_quarter_label() 同慣例——這是整份報告的比較基準。"""
    from fetcher_gaap import _col_to_quarter_label
    assert gaap_style_label("2025-12-27", 9) == _col_to_quarter_label(
        "2025-12-27 (Q1)", fy_end_month=9)
    assert gaap_style_label("2026-06-28", 12) == _col_to_quarter_label(
        "2026-06-28 (Q2)", fy_end_month=12)
