"""Tests for ratios.py — Data_Ratios sheet."""
import pytest
from fetcher_gaap import StatementTable
from ratios import build_ratio_table, RATIO_DEFS, _safe_div, _ttm, _quarter_ordinal, _lag_index


def _consecutive_labels(n, start_year=2024, start_q=1):
    """產生連續季度標籤，跨年會進位。

    不可寫成 f"FY2025Q{i+1}"——n > 4 時會產生不存在的 FY2025Q5，
    比率模組依標籤對齊期間，看到非法標籤會整列算不出來。
    """
    labels = []
    for k in range(n):
        total = (start_year * 4 + start_q - 1) + k
        labels.append(f"FY{total // 4}Q{total % 4 + 1}")
    return labels


def _q_table(**rows):
    """組一張 Data_Financials(Q)。每個 kwarg 是一列，值是各季（舊→新）。"""
    n = len(next(iter(rows.values())))
    concepts = list(rows)
    return StatementTable(
        sheet_name="Data_Financials(Q)",
        quarter_labels=_consecutive_labels(n),
        filing_dates=[""] * n,
        concepts=concepts,
        values=[list(rows[c]) for c in concepts],
        ticker="TEST",
        labels=[""] * len(concepts),
    )


def _row(tbl, name):
    return tbl.values[tbl.concepts.index(name)]


def _find(tbl, keyword):
    """用關鍵字找列名（列名帶單位後綴，例如「營收 YoY (%)」）。"""
    hits = [c for c in tbl.concepts if c.startswith(keyword)]
    assert hits, f"找不到以 {keyword} 開頭的列：{tbl.concepts}"
    return _row(tbl, hits[0])


# ── 基礎工具 ───────────────────────────────────────────────────────────────

def test_safe_div_normal():
    assert _safe_div(50.0, 200.0) == pytest.approx(0.25)

def test_safe_div_by_zero_returns_none():
    assert _safe_div(50.0, 0.0) is None

def test_safe_div_none_operand_returns_none():
    assert _safe_div(None, 200.0) is None
    assert _safe_div(50.0, None) is None

def test_ttm_needs_four_quarters():
    assert _ttm([10.0, 20.0, 30.0], 2) is None          # 只有 3 季
    assert _ttm([10.0, 20.0, 30.0, 40.0], 3) == 100.0

def test_ttm_with_hole_returns_none():
    """中間有缺值時不可當 0 加總——會低估。"""
    assert _ttm([10.0, None, 30.0, 40.0], 3) is None


# ── 利潤率 ─────────────────────────────────────────────────────────────────

def test_gross_margin():
    tbl = build_ratio_table(_q_table(**{
        "Revenue": [1000.0], "Gross Profit": [380.0],
    }))
    assert _find(tbl, "毛利率")[0] == pytest.approx(38.0)

def test_margins_stored_as_percentage_numbers():
    """存 38.0 不是 0.38——excel_formatter 會再 ÷100 套 0.0% 格式。"""
    tbl = build_ratio_table(_q_table(**{
        "Revenue": [1000.0], "Operating Income": [150.0],
    }))
    assert _find(tbl, "營利率")[0] == pytest.approx(15.0)

def test_opex_ratio():
    tbl = build_ratio_table(_q_table(**{
        "Revenue": [1000.0], "Total Operating Expense": [230.0],
    }))
    assert _find(tbl, "營業費用率")[0] == pytest.approx(23.0)

def test_net_margin():
    tbl = build_ratio_table(_q_table(**{
        "Revenue": [1000.0], "Net Income": [120.0],
    }))
    assert _find(tbl, "稅後淨利率")[0] == pytest.approx(12.0)

def test_effective_tax_rate():
    tbl = build_ratio_table(_q_table(**{
        "Pre-tax Income": [200.0], "Income Tax": [42.0],
    }))
    assert _find(tbl, "有效稅率")[0] == pytest.approx(21.0)

def test_depreciation_over_cost_and_expense():
    """折舊 / (成本 + 費用)。"""
    tbl = build_ratio_table(_q_table(**{
        "D&A": [50.0], "Cost of Revenue": [600.0], "Total Operating Expense": [400.0],
    }))
    assert _find(tbl, "折舊/(成本+費用)")[0] == pytest.approx(5.0)

def test_nonop_over_pretax():
    tbl = build_ratio_table(_q_table(**{
        "Total Non-op Income/(Loss)": [20.0], "Pre-tax Income": [200.0],
    }))
    assert _find(tbl, "業外/稅前")[0] == pytest.approx(10.0)


# ── 成長率 ─────────────────────────────────────────────────────────────────

def test_revenue_qoq():
    tbl = build_ratio_table(_q_table(Revenue=[100.0, 110.0]))
    row = _find(tbl, "營收 QoQ")
    assert row[0] is None                    # 第一季沒有前一季
    assert row[1] == pytest.approx(10.0)

def test_revenue_yoy_needs_four_quarters_back():
    tbl = build_ratio_table(_q_table(Revenue=[100.0, 100.0, 100.0, 100.0, 125.0]))
    row = _find(tbl, "營收 YoY")
    assert row[3] is None                    # 第 4 季往前數不到第 0 季的前一年
    assert row[4] == pytest.approx(25.0)

def test_net_income_yoy():
    tbl = build_ratio_table(_q_table(**{
        "Net Income": [10.0, 10.0, 10.0, 10.0, 13.0],
    }))
    assert _find(tbl, "稅後淨利 YoY")[4] == pytest.approx(30.0)

def test_eps_yoy():
    tbl = build_ratio_table(_q_table(**{
        "Diluted EPS": [1.0, 1.0, 1.0, 1.0, 1.2],
    }))
    assert _find(tbl, "EPS YoY")[4] == pytest.approx(20.0)

def test_yoy_from_negative_base_is_none():
    """基期是負數時成長率沒有意義，不可算出誤導性的數字。"""
    tbl = build_ratio_table(_q_table(**{
        "Net Income": [-10.0, 1.0, 1.0, 1.0, 5.0],
    }))
    assert _find(tbl, "稅後淨利 YoY")[4] is None


# ── ROE（TTM 淨利 ÷ 期初期末平均權益）──────────────────────────────────────

def test_roe_uses_ttm_income_and_average_equity():
    tbl = build_ratio_table(_q_table(**{
        "Net Income":           [25.0, 25.0, 25.0, 25.0, 25.0],
        "Total Equity — Parent": [900.0, 950.0, 1000.0, 1050.0, 1100.0],
    }))
    # TTM 淨利 = 100；平均權益 = (950 + 1100) / 2 = 1025 → 9.756%
    assert _find(tbl, "ROE")[4] == pytest.approx(100.0 / 1025.0 * 100, rel=1e-4)

def test_roe_none_without_four_quarters():
    tbl = build_ratio_table(_q_table(**{
        "Net Income": [25.0, 25.0],
        "Total Equity — Parent": [900.0, 950.0],
    }))
    assert _find(tbl, "ROE")[1] is None


# ── 現金流與每股 ───────────────────────────────────────────────────────────

def test_fcf_margin():
    tbl = build_ratio_table(_q_table(**{
        "Revenue": [1000.0], "Free Cash Flow": [180.0],
    }))
    assert _find(tbl, "FCF Margin")[0] == pytest.approx(18.0)

def test_fcf_over_net_income_is_a_multiple_not_percent():
    """現金轉換率習慣看倍數，單位是 (x) 不是 (%)。"""
    tbl = build_ratio_table(_q_table(**{
        "Free Cash Flow": [120.0], "Net Income": [100.0],
    }))
    hits = [c for c in tbl.concepts if c.startswith("FCF/淨利")]
    assert hits and hits[0].endswith("(x)")
    assert _row(tbl, hits[0])[0] == pytest.approx(1.2)

def test_bvps_is_dollars():
    tbl = build_ratio_table(_q_table(**{
        "Total Equity — Parent": [1000.0], "Shares Outstanding": [250.0],
    }))
    hits = [c for c in tbl.concepts if c.startswith("BVPS")]
    assert hits and hits[0].endswith("($)")
    assert _row(tbl, hits[0])[0] == pytest.approx(4.0)

def test_dso_is_days():
    tbl = build_ratio_table(_q_table(**{
        "Revenue": [1000.0], "Accounts Receivable": [500.0],
    }))
    hits = [c for c in tbl.concepts if c.startswith("DSO")]
    assert hits and hits[0].endswith("(days)")
    # 單季營收年化：500 / (1000 * 4) * 365
    assert _row(tbl, hits[0])[0] == pytest.approx(500.0 / 4000.0 * 365.0)


# ── 表格結構 ───────────────────────────────────────────────────────────────

def test_sheet_name():
    tbl = build_ratio_table(_q_table(Revenue=[100.0]))
    assert tbl.sheet_name == "Data_Ratios"

def test_quarter_labels_copied_from_source():
    src = _q_table(Revenue=[100.0, 110.0])
    tbl = build_ratio_table(src)
    assert tbl.quarter_labels == src.quarter_labels

def test_every_ratio_row_present_even_when_uncomputable():
    """固定模板：算不出來也要有那一列，全 None。不可因為缺資料就少一列。"""
    tbl = build_ratio_table(_q_table(Revenue=[100.0]))
    assert len(tbl.concepts) == len(RATIO_DEFS)

def test_b_column_holds_the_formula_text():
    """B 欄寫算法——skill 讀值，人看得懂這格怎麼來的。"""
    tbl = build_ratio_table(_q_table(**{"Revenue": [1000.0], "Gross Profit": [380.0]}))
    idx = next(i for i, c in enumerate(tbl.concepts) if c.startswith("毛利率"))
    assert "Gross Profit" in tbl.labels[idx] and "Revenue" in tbl.labels[idx]

def test_every_row_name_carries_a_unit_suffix():
    """單位後綴是 excel_formatter 判斷格式的依據，不可漏。"""
    tbl = build_ratio_table(_q_table(Revenue=[100.0]))
    for name in tbl.concepts:
        assert name.endswith(("(%)", "(x)", "(days)", "($)")), name

def test_every_row_has_formula_text():
    tbl = build_ratio_table(_q_table(Revenue=[100.0]))
    assert all(lbl.strip() for lbl in tbl.labels)

def test_returns_none_for_empty_source():
    assert build_ratio_table(None) is None

def test_missing_concepts_do_not_raise():
    """來源表缺很多列是常態（金融股、小公司），不可炸。"""
    tbl = build_ratio_table(_q_table(**{"Cash": [10.0]}))
    assert tbl is not None
    assert all(v is None for row in tbl.values for v in row)


# ═════════════════════════════════════════════════════════════════════════════
# 季度不連續時的 YoY / QoQ（2026-08-02，ARLO 實跑抓到）
#
# ARLO 抓到的季度是 [FY2024Q1, Q2, Q3, FY2025Q1, Q2, Q3, FY2026Q1]——缺 Q4。
# 用「往前數 4 格」算 YoY 會拿到 5 季前的數字，而且看起來完全正常。
# 必須依季度標籤對齊，不是依欄位位置。
# ═════════════════════════════════════════════════════════════════════════════

def _q_table_labelled(labels, **rows):
    concepts = list(rows)
    return StatementTable(
        sheet_name="Data_Financials(Q)",
        quarter_labels=list(labels),
        filing_dates=[""] * len(labels),
        concepts=concepts,
        values=[list(rows[c]) for c in concepts],
        ticker="TEST",
        labels=[""] * len(concepts),
    )


def test_quarter_ordinal_is_monotonic_across_year_boundary():
    assert _quarter_ordinal("FY2025Q1") - _quarter_ordinal("FY2024Q4") == 1
    assert _quarter_ordinal("FY2025Q1") - _quarter_ordinal("FY2024Q1") == 4


def test_lag_index_finds_exact_quarter():
    labels = ["FY2024Q1", "FY2024Q2", "FY2024Q3", "FY2024Q4", "FY2025Q1"]
    assert _lag_index(labels, 4, 4) == 0          # FY2025Q1 往前 4 季 = FY2024Q1
    assert _lag_index(labels, 4, 1) == 3          # 往前 1 季 = FY2024Q4


def test_lag_index_returns_none_when_quarter_absent():
    """缺 Q4 時，FY2025Q1 的前一季不存在——不可退而用旁邊那欄。"""
    labels = ["FY2024Q1", "FY2024Q2", "FY2024Q3", "FY2025Q1"]
    assert _lag_index(labels, 3, 1) is None


def test_yoy_uses_label_not_position_when_a_quarter_is_missing():
    """ARLO 實際的季度序列：缺 FY2024Q4。

    FY2025Q2 的正確基期是 FY2024Q2（=200）→ +30%。
    若用「往前數 4 格」會抓到 index 0 的 FY2024Q1（=100）→ +160%，
    數字看起來完全正常，但錯得離譜。這條測試就是釘住這件事。
    """
    tbl = build_ratio_table(_q_table_labelled(
        ["FY2024Q1", "FY2024Q2", "FY2024Q3", "FY2025Q1", "FY2025Q2"],
        Revenue=[100.0, 200.0, 300.0, 400.0, 260.0],
    ))
    assert _find(tbl, "營收 YoY")[4] == pytest.approx(30.0)


def test_yoy_none_when_prior_year_quarter_truly_absent():
    """基期那一季根本沒抓到時要留空，不可拿旁邊的季度頂替。"""
    tbl = build_ratio_table(_q_table_labelled(
        ["FY2024Q2", "FY2024Q3", "FY2025Q1", "FY2025Q2", "FY2025Q3"],
        Revenue=[100.0, 100.0, 100.0, 100.0, 130.0],
    ))
    # FY2025Q3 的基期 FY2024Q3 存在 → 有值；FY2025Q1 的基期 FY2024Q1 不存在 → None
    assert _find(tbl, "營收 YoY")[2] is None
    assert _find(tbl, "營收 YoY")[4] == pytest.approx(30.0)


def test_yoy_correct_when_quarters_contiguous():
    tbl = build_ratio_table(_q_table_labelled(
        ["FY2024Q1", "FY2024Q2", "FY2024Q3", "FY2024Q4", "FY2025Q1"],
        Revenue=[100.0, 0.0, 0.0, 0.0, 130.0],
    ))
    assert _find(tbl, "營收 YoY")[4] == pytest.approx(30.0)


def test_qoq_skips_across_a_gap():
    tbl = build_ratio_table(_q_table_labelled(
        ["FY2024Q2", "FY2024Q3", "FY2025Q1"],
        Revenue=[100.0, 110.0, 120.0],
    ))
    row = _find(tbl, "營收 QoQ")
    assert row[1] == pytest.approx(10.0)     # Q3 vs Q2，連續
    assert row[2] is None                    # FY2025Q1 前一季 FY2024Q4 不存在


def test_ttm_requires_four_contiguous_quarters():
    """TTM 少一季就不可加總——會低估，而且錯得很像對的。"""
    tbl = build_ratio_table(_q_table_labelled(
        ["FY2024Q1", "FY2024Q2", "FY2024Q3", "FY2025Q1"],
        **{"Net Income": [25.0, 25.0, 25.0, 25.0],
           "Total Equity — Parent": [900.0, 950.0, 1000.0, 1050.0]},
    ))
    assert _find(tbl, "ROE")[3] is None


def test_ttm_works_when_contiguous():
    tbl = build_ratio_table(_q_table_labelled(
        ["FY2024Q1", "FY2024Q2", "FY2024Q3", "FY2024Q4"],
        **{"Net Income": [25.0, 25.0, 25.0, 25.0],
           "Total Equity — Parent": [1000.0, 1000.0, 1000.0, 1000.0]},
    ))
    assert _find(tbl, "ROE")[3] == pytest.approx(10.0)
