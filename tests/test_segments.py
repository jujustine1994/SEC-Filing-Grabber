"""Tests for segments.py — Data_Segments 長格式彙總表。"""
import pytest
from fetcher_gaap import StatementTable
from segments import build_segments_long


def _seg(sheet_name, quarters, members):
    """members = {顯示名: [各季值]}"""
    return StatementTable(
        sheet_name=sheet_name,
        quarter_labels=list(quarters),
        filing_dates=["2025-01-01"] * len(quarters),
        concepts=list(members),
        values=[list(v) for v in members.values()],
        ticker="TEST",
        labels=[""] * len(members),
    )


def _row(tbl, name):
    assert name in tbl.concepts, f"缺列 {name}；現有 {tbl.concepts}"
    return tbl.values[tbl.concepts.index(name)]


def test_returns_none_without_segment_tables():
    assert build_segments_long([]) is None


def test_sheet_name_is_fixed():
    tbl = build_segments_long([_seg("Data_Seg_Revenue", ["FY2025Q1"], {"US": [10.0]})])
    assert tbl.sheet_name == "Data_Segments"


def test_row_name_combines_metric_and_member():
    """列名要自我描述——skill 讀 A 欄就知道這是哪個指標的哪個分項。"""
    tbl = build_segments_long([_seg("Data_Seg_Revenue", ["FY2025Q1"], {"US": [10.0]})])
    assert "Revenue — US" in tbl.concepts


def test_source_sheet_recorded_in_column_b():
    tbl = build_segments_long([_seg("Data_Seg_Revenue", ["FY2025Q1"], {"US": [10.0]})])
    i = tbl.concepts.index("Revenue — US")
    assert tbl.labels[i] == "Data_Seg_Revenue"


def test_multiple_axes_land_in_one_sheet():
    """這就是長格式的重點：不管公司有幾個分類軸都是同一張表。"""
    tbl = build_segments_long([
        _seg("Data_Seg_Revenue", ["FY2025Q1"], {"Data Center": [100.0], "Gaming": [20.0]}),
        _seg("Data_Seg_RevenueGeo", ["FY2025Q1"], {"United States": [70.0], "Taiwan": [50.0]}),
    ])
    for name in ("Revenue — Data Center", "Revenue — Gaming",
                 "RevenueGeo — United States", "RevenueGeo — Taiwan"):
        assert name in tbl.concepts


def test_columns_are_the_union_of_all_periods():
    """各軸的季度範圍可能不同，欄位要取聯集並排序。"""
    tbl = build_segments_long([
        _seg("Data_Seg_A", ["FY2025Q1", "FY2025Q2"], {"x": [1.0, 2.0]}),
        _seg("Data_Seg_B", ["FY2025Q2", "FY2025Q3"], {"y": [3.0, 4.0]}),
    ])
    assert tbl.quarter_labels == ["FY2025Q1", "FY2025Q2", "FY2025Q3"]


def test_values_align_to_the_right_period_after_union():
    """聯集後不可整列左移——那會讓每個數字都掛到錯的季。"""
    tbl = build_segments_long([
        _seg("Data_Seg_A", ["FY2025Q1", "FY2025Q2"], {"x": [1.0, 2.0]}),
        _seg("Data_Seg_B", ["FY2025Q2", "FY2025Q3"], {"y": [3.0, 4.0]}),
    ])
    assert _row(tbl, "A — x") == [1.0, 2.0, None]
    assert _row(tbl, "B — y") == [None, 3.0, 4.0]


def test_duplicate_member_names_across_axes_stay_separate():
    """兩個軸都有「Other」時不可互相覆蓋。"""
    tbl = build_segments_long([
        _seg("Data_Seg_A", ["FY2025Q1"], {"Other": [1.0]}),
        _seg("Data_Seg_B", ["FY2025Q1"], {"Other": [9.0]}),
    ])
    assert _row(tbl, "A — Other") == [1.0]
    assert _row(tbl, "B — Other") == [9.0]


def test_ticker_carried_over():
    tbl = build_segments_long([_seg("Data_Seg_Revenue", ["FY2025Q1"], {"US": [10.0]})])
    assert tbl.ticker == "TEST"


def test_non_segment_tables_ignored():
    """只吃 Data_Seg_*，三表不可被捲進來。"""
    other = _seg("Data_Financials(Q)", ["FY2025Q1"], {"Revenue": [1.0]})
    assert build_segments_long([other]) is None


def test_empty_segment_table_skipped():
    tbl = build_segments_long([
        _seg("Data_Seg_Empty", ["FY2025Q1"], {}),
        _seg("Data_Seg_Revenue", ["FY2025Q1"], {"US": [10.0]}),
    ])
    assert tbl.concepts == ["Revenue — US"]
