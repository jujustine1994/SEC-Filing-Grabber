"""Tests for nongaap_layout.py — Data_NonGAAP 的固定模板版面。"""
import pytest
from nongaap_layout import (
    build_nongaap_table, CORE_ROWS, ADDBACK_ROWS,
    SECTION_CORE, SECTION_RECON, SECTION_OTHER, SECTION_ANNUAL,
    RESIDUAL_ROW, GAAP_PREFIX,
)


def _built(per_quarter, labels=None, dates=None):
    labels = labels or sorted(per_quarter)
    return build_nongaap_table(
        "TEST", per_quarter, labels, dates or [""] * len(labels),
    )


def _row(tbl, name):
    assert name in tbl.concepts, f"缺列 {name}；現有：{tbl.concepts}"
    return tbl.values[tbl.concepts.index(name)]


# ── 固定模板 ───────────────────────────────────────────────────────────────

def test_core_rows_always_present_even_with_no_data():
    """完全沒有 Non-GAAP 資料的公司（AAPL/AMZN/COST）也要產出完整骨架。
    讀不到 sheet 和讀到空 sheet 是兩種訊號，前者無法區分「沒報」與「抓取失敗」。"""
    tbl = _built({"FY2025Q1": {}})
    for display, _key, _gaap in CORE_ROWS:
        assert display in tbl.concepts


def test_core_row_order_is_stable_across_companies():
    a = _built({"FY2025Q1": {"Non-GAAP Gross Margin": 40.0}})
    b = _built({"FY2025Q1": {"Adjusted EBITDA": 100.0}})
    core_a = [c for c in a.concepts if c in {r[0] for r in CORE_ROWS}]
    core_b = [c for c in b.concepts if c in {r[0] for r in CORE_ROWS}]
    assert core_a == core_b


def test_sheet_name():
    assert _built({"FY2025Q1": {}}).sheet_name == "Data_NonGAAP"


def test_section_headers_present_in_order():
    tbl = _built({"FY2025Q1": {"某個公司自訂指標": 1.0}})
    idx = {s: tbl.concepts.index(s) for s in
           (SECTION_CORE, SECTION_RECON, SECTION_OTHER)}
    assert idx[SECTION_CORE] < idx[SECTION_RECON] < idx[SECTION_OTHER]


def test_section_header_rows_have_no_values():
    tbl = _built({"FY2025Q1": {"Non-GAAP Gross Margin": 40.0}})
    assert all(v is None for v in _row(tbl, SECTION_CORE))


# ── Core 取值 ──────────────────────────────────────────────────────────────

def test_core_row_picks_up_its_metric():
    tbl = _built({"FY2025Q1": {"Non-GAAP Gross Margin": 40.0}})
    assert _row(tbl, "Non-GAAP Gross Margin")[0] == 40.0


def test_core_row_blank_when_company_does_not_report_it():
    tbl = _built({"FY2025Q1": {"Non-GAAP Gross Margin": 40.0}})
    assert _row(tbl, "Adjusted EBITDA")[0] is None


def test_core_row_matches_across_quarters():
    tbl = _built({
        "FY2025Q1": {"Non-GAAP Gross Margin": 40.0},
        "FY2025Q2": {"Non-GAAP Gross Margin": 42.0},
    })
    assert _row(tbl, "Non-GAAP Gross Margin") == [40.0, 42.0]


# ── GAAP 對照行 ────────────────────────────────────────────────────────────

def test_gaap_companion_row_exists_for_margins():
    tbl = _built({"FY2025Q1": {}})
    assert f"{GAAP_PREFIX}Gross Margin" in tbl.concepts
    assert f"{GAAP_PREFIX}Net Margin" in tbl.concepts


def test_gaap_companion_takes_value_from_same_press_release():
    """對照值必須來自同一份新聞稿，不是從 Data_Financials 拉——
    Non-GAAP 的季度標籤晚一季，跨表拉會變成錯開一季的無聲比較。"""
    tbl = _built({"FY2025Q1": {
        "Non-GAAP Gross Margin": 50.1,
        "GAAP Gross Margin": 48.3,
    }})
    assert _row(tbl, "Non-GAAP Gross Margin")[0] == 50.1
    assert _row(tbl, f"{GAAP_PREFIX}Gross Margin")[0] == 48.3


def test_gaap_companion_sits_right_after_its_nongaap_row():
    tbl = _built({"FY2025Q1": {}})
    i = tbl.concepts.index("Non-GAAP Gross Margin")
    assert tbl.concepts[i + 1] == f"{GAAP_PREFIX}Gross Margin"


def test_source_written_in_column_b():
    tbl = _built({"FY2025Q1": {"Non-GAAP Gross Margin": 50.1}})
    i = tbl.concepts.index("Non-GAAP Gross Margin")
    assert "8-K" in tbl.labels[i]


# ── 淨利率是推導的 ──────────────────────────────────────────────────────────

def test_net_margin_derived_from_net_income_and_revenue():
    """32 家沒有任何一家在新聞稿寫 non-GAAP net margin，只能推導。"""
    tbl = _built({"FY2025Q1": {
        "Non-GAAP Net Income": 120.0,
        "Non-GAAP Revenue": 1000.0,
    }})
    assert _row(tbl, "Non-GAAP Net Margin")[0] == pytest.approx(12.0)


def test_net_margin_falls_back_to_gaap_revenue():
    """多數公司不報 Non-GAAP Revenue（只有 38%），分母改用 GAAP Revenue。"""
    tbl = _built({"FY2025Q1": {
        "Non-GAAP Net Income": 120.0,
        "GAAP Revenue": 1000.0,
    }})
    assert _row(tbl, "Non-GAAP Net Margin")[0] == pytest.approx(12.0)


def test_net_margin_blank_without_a_denominator():
    tbl = _built({"FY2025Q1": {"Non-GAAP Net Income": 120.0}})
    assert _row(tbl, "Non-GAAP Net Margin")[0] is None


def test_derived_row_says_so_in_column_b():
    tbl = _built({"FY2025Q1": {}})
    i = tbl.concepts.index("Non-GAAP Net Margin")
    assert "DERIVED" in tbl.labels[i]


# ── 調節表與殘差 ───────────────────────────────────────────────────────────

def test_addback_rows_always_present():
    tbl = _built({"FY2025Q1": {}})
    for display, _key in ADDBACK_ROWS:
        assert display in tbl.concepts


def test_addback_picks_up_value():
    tbl = _built({"FY2025Q1": {"Stock-Based Compensation": 20.0}})
    assert _row(tbl, "  + 股權獎酬 SBC")[0] == 20.0


def test_residual_ties_the_bridge():
    """殘差 = Non-GAAP 淨利 − GAAP 淨利 − 具名項目合計。"""
    tbl = _built({"FY2025Q1": {
        "GAAP Net Income": 100.0,
        "Non-GAAP Net Income": 130.0,
        "Stock-Based Compensation": 20.0,
        "Amortization of Intangibles": 8.0,
        "Tax Effect of Adjustments": -6.0,
    }})
    # 100 + 20 + 8 - 6 = 122，距離 130 還差 8
    assert _row(tbl, RESIDUAL_ROW)[0] == pytest.approx(8.0)


def test_residual_is_zero_when_named_items_explain_everything():
    tbl = _built({"FY2025Q1": {
        "GAAP Net Income": 100.0,
        "Non-GAAP Net Income": 120.0,
        "Stock-Based Compensation": 20.0,
    }})
    assert _row(tbl, RESIDUAL_ROW)[0] == pytest.approx(0.0)


def test_residual_blank_without_both_ends_of_the_bridge():
    """缺 GAAP 或 Non-GAAP 淨利時算不出殘差，留空而不是當 0。"""
    tbl = _built({"FY2025Q1": {"Stock-Based Compensation": 20.0}})
    assert _row(tbl, RESIDUAL_ROW)[0] is None


def test_residual_ignores_unnamed_metrics_that_are_not_addbacks():
    """毛利率之類的指標不可被算進調節橋。"""
    tbl = _built({"FY2025Q1": {
        "GAAP Net Income": 100.0,
        "Non-GAAP Net Income": 120.0,
        "Stock-Based Compensation": 20.0,
        "Non-GAAP Gross Margin": 45.0,
    }})
    assert _row(tbl, RESIDUAL_ROW)[0] == pytest.approx(0.0)


# ── Overflow 與年度區 ──────────────────────────────────────────────────────

def test_unknown_metric_goes_to_overflow_not_dropped():
    tbl = _built({"FY2025Q1": {"Non-GAAP Service Gross Margin": 81.7}})
    assert "Non-GAAP Service Gross Margin" in tbl.concepts
    i = tbl.concepts.index("Non-GAAP Service Gross Margin")
    assert i > tbl.concepts.index(SECTION_OTHER)


def test_core_metric_does_not_also_appear_in_overflow():
    tbl = _built({"FY2025Q1": {"Non-GAAP Gross Margin": 45.0}})
    assert tbl.concepts.count("Non-GAAP Gross Margin") == 1


def test_annual_rows_go_to_their_own_section():
    tbl = _built({"FY2025Q1": {"Non-GAAP Gross Margin (FY)": 44.0}})
    assert SECTION_ANNUAL in tbl.concepts
    i = tbl.concepts.index("Non-GAAP Gross Margin (FY)")
    assert i > tbl.concepts.index(SECTION_ANNUAL)


def test_annual_section_omitted_when_no_annual_rows():
    tbl = _built({"FY2025Q1": {"Non-GAAP Gross Margin": 45.0}})
    assert SECTION_ANNUAL not in tbl.concepts


def test_overflow_section_omitted_when_everything_is_core():
    tbl = _built({"FY2025Q1": {"Non-GAAP Gross Margin": 45.0}})
    assert SECTION_OTHER not in tbl.concepts


# ── 結構完整性 ─────────────────────────────────────────────────────────────

def test_all_rows_have_same_width_as_quarters():
    tbl = _built({"FY2025Q1": {"a": 1.0}, "FY2025Q2": {"b": 2.0}})
    assert all(len(r) == 2 for r in tbl.values)


def test_labels_align_with_concepts():
    tbl = _built({"FY2025Q1": {"Non-GAAP Gross Margin": 45.0}})
    assert len(tbl.labels) == len(tbl.concepts) == len(tbl.values)


def test_no_data_at_all_still_returns_a_table():
    assert _built({}) is not None
