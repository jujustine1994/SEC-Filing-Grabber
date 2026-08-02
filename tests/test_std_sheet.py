"""Tests for std_sheet.py — Data_Std 跨公司固定版面表。"""
import pytest
from fetcher_gaap import StatementTable
from std_sheet import (
    build_std_table, STD_ROWS, FROZEN_ROW_NUMBERS,
    _calendar_quarter, ROW_CALENDAR, ROW_PERIOD_END, ROW_SCHEMA, SCHEMA_VERSION,
)


def _merged(quarters, sections):
    """組一張 Data_Financials(Q) 樣子的合併表。

    sections = {"Income Statement": {concept: [values]}, ...}
    刻意讓 Net Income 在 IS 與 CF 都出現——那正是 VLOOKUP 會抓錯的情形。
    """
    concepts, values = [], []
    for section, rows in sections.items():
        concepts.append(section)
        values.append([None] * len(quarters))
        for name, vals in rows.items():
            concepts.append(name)
            values.append(list(vals))
    return StatementTable(
        sheet_name="Data_Financials(Q)",
        quarter_labels=list(quarters),
        filing_dates=["2026-05-07"] * len(quarters),
        concepts=concepts,
        values=values,
        ticker="TEST",
        labels=[""] * len(concepts),
    )


def _row_by_key(tbl, key):
    assert key in tbl.labels, f"缺機器鍵 {key}"
    return tbl.values[tbl.labels.index(key)]


# ── 日曆季換算 ─────────────────────────────────────────────────────────────

def test_calendar_quarter_for_december_fye():
    """ARLO 12 月結算：FY2026Q1 = 2026 年 3 月底 = 日曆 2026Q1。"""
    assert _calendar_quarter("FY2026Q1", 12) == "2026Q1"


def test_calendar_quarter_for_september_fye():
    """AAPL 9 月結算：FY2026Q1 結束於 2025 年 12 月 = 日曆 2025Q4。"""
    assert _calendar_quarter("FY2026Q1", 9) == "2025Q4"


def test_calendar_quarter_for_january_fye():
    """NVDA 1 月結算：FY2026Q1 結束於 2025 年 4 月 = 日曆 2025Q2。"""
    assert _calendar_quarter("FY2026Q1", 1) == "2025Q2"


def test_calendar_quarter_q4():
    assert _calendar_quarter("FY2026Q4", 12) == "2026Q4"


def test_calendar_quarter_of_annual_label():
    """年度標籤 FY2026 沒有季，回空字串而不是亂猜。"""
    assert _calendar_quarter("FY2026", 12) == ""


def test_calendar_quarter_of_garbage_label():
    assert _calendar_quarter("nonsense", 12) == ""


# ── 標籤列 ─────────────────────────────────────────────────────────────────

def test_fiscal_labels_stay_in_row_one():
    tbl = build_std_table(_merged(["FY2026Q1"], {}), None, fy_end_month=12)
    assert tbl.quarter_labels == ["FY2026Q1"]


def test_calendar_row_present():
    tbl = build_std_table(_merged(["FY2026Q1"], {}), None, fy_end_month=9)
    assert _row_by_key(tbl, ROW_CALENDAR) == ["2025Q4"]


def test_period_end_row_present():
    tbl = build_std_table(_merged(["FY2026Q1"], {}), None, fy_end_month=9)
    assert _row_by_key(tbl, ROW_PERIOD_END) == ["2025-12"]


def test_two_companies_with_different_fye_align_on_calendar_row():
    """同一個日曆季在兩家不同結算月的公司上，日曆列的值必須相同。"""
    arlo = build_std_table(_merged(["FY2026Q1"], {}), None, fy_end_month=12)
    nvda = build_std_table(_merged(["FY2026Q4"], {}), None, fy_end_month=1)
    assert _row_by_key(arlo, ROW_CALENDAR) == _row_by_key(nvda, ROW_CALENDAR) == ["2026Q1"]


# ── 欄位順序 ───────────────────────────────────────────────────────────────

def test_columns_stay_oldest_to_newest():
    src = _merged(["FY2025Q3", "FY2025Q4", "FY2026Q1"], {})
    tbl = build_std_table(src, None, fy_end_month=12)
    assert tbl.quarter_labels == ["FY2025Q3", "FY2025Q4", "FY2026Q1"]


def test_values_follow_the_same_order():
    src = _merged(["FY2025Q3", "FY2026Q1"],
                  {"Income Statement": {"Revenue": [100.0, 150.0]}})
    tbl = build_std_table(src, None, fy_end_month=12)
    assert _row_by_key(tbl, "IS.REVENUE") == [100.0, 150.0]


# ── 固定列位 ───────────────────────────────────────────────────────────────

def test_row_numbers_are_frozen():
    """列號寫死在測試裡。任何人（包括我）插入一列都會立刻紅掉——
    沒有這條，「固定列位」三個月後會悄悄失效，而使用者的公式會靜默抓錯。"""
    tbl = build_std_table(_merged(["FY2026Q1"], {}), None, fy_end_month=12)
    for key, expected_row in FROZEN_ROW_NUMBERS.items():
        assert key in tbl.labels, f"缺機器鍵 {key}"
        actual = tbl.labels.index(key) + 3       # sheet 上第 3 列開始放資料
        assert actual == expected_row, f"{key} 應在第 {expected_row} 列，實際 {actual}"


def test_row_count_identical_regardless_of_company():
    a = build_std_table(_merged(["FY2026Q1"], {"Income Statement": {"Revenue": [1.0]}}),
                        None, fy_end_month=12)
    b = build_std_table(_merged(["FY2026Q1"], {"Balance Sheet": {"Cash": [2.0]}}),
                        None, fy_end_month=12)
    assert len(a.concepts) == len(b.concepts)


def test_overflow_rows_are_not_included():
    """來源表的 overflow 行（公司特有的 XBRL 科目）不可進 Data_Std——
    它們正是害列位浮動的原因。"""
    src = _merged(["FY2026Q1"], {"Income Statement": {
        "Revenue": [100.0],
        "某公司特有的奇怪科目": [5.0],
    }})
    tbl = build_std_table(src, None, fy_end_month=12)
    assert "某公司特有的奇怪科目" not in tbl.concepts


def test_machine_keys_are_unique():
    tbl = build_std_table(_merged(["FY2026Q1"], {}), None, fy_end_month=12)
    keys = [k for k in tbl.labels if k]
    assert len(keys) == len(set(keys))


def test_schema_version_is_a_row():
    """版本號放成一列而不是 dataclass 欄位——寫檔器只認得 concepts/labels/values，
    多開一個欄位還要改寫檔器，放一列就能直接出現在 sheet 上。"""
    tbl = build_std_table(_merged(["FY2026Q1"], {}), None, fy_end_month=12)
    assert _row_by_key(tbl, ROW_SCHEMA) == [SCHEMA_VERSION]


# ── 同名科目分區取值 ────────────────────────────────────────────────────────

def test_net_income_from_is_and_cf_are_separate_rows():
    """Net Income 在 IS 與 CF 都有。VLOOKUP 只會抓到第一個，
    機器鍵必須把兩者分開。"""
    src = _merged(["FY2026Q1"], {
        "Income Statement": {"Net Income": [10.0]},
        "Cash Flow":        {"Net Income": [99.0]},
    })
    tbl = build_std_table(src, None, fy_end_month=12)
    assert _row_by_key(tbl, "IS.NET_INCOME") == [10.0]
    assert _row_by_key(tbl, "CF.NET_INCOME") == [99.0]


def test_missing_concept_leaves_blank_row():
    tbl = build_std_table(_merged(["FY2026Q1"], {}), None, fy_end_month=12)
    assert _row_by_key(tbl, "BS.CASH") == [None]


# ── 比率併入 ───────────────────────────────────────────────────────────────

def test_ratio_rows_merged_in():
    ratios = StatementTable(
        sheet_name="Data_Ratios",
        quarter_labels=["FY2026Q1"],
        filing_dates=[""],
        concepts=["毛利率 (%)"],
        values=[[38.0]],
        ticker="TEST",
        labels=["Gross Profit / Revenue"],
    )
    tbl = build_std_table(_merged(["FY2026Q1"], {}), ratios, fy_end_month=12)
    assert _row_by_key(tbl, "RATIO.毛利率") == [38.0]


def test_works_without_ratio_table():
    assert build_std_table(_merged(["FY2026Q1"], {}), None, fy_end_month=12) is not None


def test_returns_none_without_source():
    assert build_std_table(None, None) is None
