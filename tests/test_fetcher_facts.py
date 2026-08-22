"""fetcher_facts.py — SEC companyfacts API 取數路徑（TODO G11 spike）。

這條路徑是現行「逐份解 filing」的平行替代品，**還沒接上主流程**，目的是先
產出逐格比對報告給 CTH 決定要不要切換（見 `docs/TODO.md` G11）。

這裡只測純邏輯（期間分類、重編取版、concept 解析、組表），不打網路。
"""
import pytest

import fetcher_facts as ff


def _fact(start, end, val, *, form="10-Q", filed="2025-01-01", accn="a"):
    f = {"end": end, "val": val, "form": form, "filed": filed, "accn": accn}
    if start:
        f["start"] = start
    return f


# ── 期間分類 ────────────────────────────────────────────────────────────────

def test_duration_days_none_for_instant():
    """資產負債表是時點值，沒有 start。"""
    assert ff.duration_days(_fact(None, "2025-06-28", 1.0)) is None


def test_duration_days_counts_inclusive_span():
    assert ff.duration_days(_fact("2025-03-30", "2025-06-28", 1.0)) == 90


@pytest.mark.parametrize("start, end, expected", [
    # 13 週季：52/53 週制會落在 84~98 天，**不可以只認剛好 91 天**
    ("2025-03-30", "2025-06-28", "quarter"),   # 90
    ("2025-04-28", "2025-07-27", "quarter"),   # 90，NVDA
    ("2024-12-29", "2025-03-29", "quarter"),   # 90
    # 年度
    ("2024-12-29", "2025-12-27", "annual"),    # 363
    ("2025-01-27", "2026-01-25", "annual"),    # 363，NVDA
    # 半年 / YTD 九個月 —— 兩邊都不算，要丟掉
    ("2025-01-01", "2025-06-30", None),        # 180
    ("2025-01-01", "2025-09-30", None),        # 272
])
def test_classify_period_duration(start, end, expected):
    assert ff.classify_period(_fact(start, end, 1.0)) == expected


def test_classify_period_instant():
    assert ff.classify_period(_fact(None, "2025-06-28", 1.0)) == "instant"


# ── 重編取版 ────────────────────────────────────────────────────────────────

def test_pick_fact_as_reported_takes_the_earliest_filed():
    """預設取當初申報值——符合「回看當時看得到什麼」的分析直覺。"""
    facts = [
        _fact("2025-03-30", "2025-06-28", 111.0, filed="2026-02-01", accn="restated"),
        _fact("2025-03-30", "2025-06-28", 100.0, filed="2025-07-30", accn="original"),
    ]
    assert ff.pick_fact(facts, prefer="as_reported")["val"] == 100.0


def test_pick_fact_latest_takes_the_newest_filed():
    facts = [
        _fact("2025-03-30", "2025-06-28", 111.0, filed="2026-02-01", accn="restated"),
        _fact("2025-03-30", "2025-06-28", 100.0, filed="2025-07-30", accn="original"),
    ]
    assert ff.pick_fact(facts, prefer="latest")["val"] == 111.0


def test_pick_fact_is_deterministic_when_filed_dates_tie():
    """同一天申報時要有穩定的第二排序鍵，不可以看 dict 順序。"""
    a = _fact("2025-03-30", "2025-06-28", 1.0, filed="2025-07-30", accn="0001-25-000002")
    b = _fact("2025-03-30", "2025-06-28", 2.0, filed="2025-07-30", accn="0001-25-000001")
    assert ff.pick_fact([a, b], prefer="as_reported")["accn"] == "0001-25-000001"
    assert ff.pick_fact([b, a], prefer="as_reported")["accn"] == "0001-25-000001"


def test_pick_fact_empty_returns_none():
    assert ff.pick_fact([], prefer="as_reported") is None


def test_pick_fact_rejects_unknown_preference():
    with pytest.raises(ValueError):
        ff.pick_fact([_fact("2025-03-30", "2025-06-28", 1.0)], prefer="whatever")


# ── concept → {期末日: 值} ──────────────────────────────────────────────────

_RAW = {
    "facts": {
        "us-gaap": {
            "Revenues": {"label": "Revenues", "units": {"USD": [
                _fact("2025-03-30", "2025-06-28", 100.0, filed="2025-07-30"),
                _fact("2025-06-29", "2025-09-27", 120.0, filed="2025-10-30"),
                _fact("2024-12-29", "2025-12-27", 500.0, form="10-K", filed="2026-02-01"),
                _fact("2025-01-01", "2025-09-30", 999.0, filed="2025-10-30"),   # YTD，要丟掉
            ]}},
            "Assets": {"label": "Assets", "units": {"USD": [
                _fact(None, "2025-06-28", 900.0, filed="2025-07-30"),
            ]}},
        }
    }
}


def test_series_for_concept_returns_quarter_values_by_period_end():
    out = ff.series_for_concept(_RAW, "Revenues", kind="quarter", prefer="as_reported")
    assert out == {"2025-06-28": 100.0, "2025-09-27": 120.0}


def test_series_for_concept_drops_ytd_and_other_odd_durations():
    """九個月的 YTD 欄不可以混進單季序列——這是舊路徑最容易出錯的地方。"""
    assert "2025-09-30" not in ff.series_for_concept(
        _RAW, "Revenues", kind="quarter", prefer="as_reported")


def test_series_for_concept_annual():
    out = ff.series_for_concept(_RAW, "Revenues", kind="annual", prefer="as_reported")
    assert out == {"2025-12-27": 500.0}


def test_series_for_concept_instant_for_balance_sheet():
    out = ff.series_for_concept(_RAW, "Assets", kind="instant", prefer="as_reported")
    assert out == {"2025-06-28": 900.0}


def test_series_for_concept_missing_concept_returns_empty():
    assert ff.series_for_concept(_RAW, "NoSuchConcept", kind="quarter",
                                 prefer="as_reported") == {}


def test_series_for_concept_tries_fallback_when_primary_absent():
    """同一個經濟意義會跨 concept（NVDA 早年 Revenues、後來 RevenueFrom...）。"""
    out = ff.series_for_concept(_RAW, "RevenueFromContractWithCustomerExcludingAssessedTax",
                                kind="quarter", prefer="as_reported",
                                fallbacks=["Revenues"])
    assert out == {"2025-06-28": 100.0, "2025-09-27": 120.0}


def test_series_for_concept_primary_wins_over_fallback():
    raw = {"facts": {"us-gaap": {
        "Primary": {"units": {"USD": [_fact("2025-03-30", "2025-06-28", 1.0)]}},
        "Backup":  {"units": {"USD": [_fact("2025-03-30", "2025-06-28", 2.0)]}},
    }}}
    out = ff.series_for_concept(raw, "Primary", kind="quarter",
                                prefer="as_reported", fallbacks=["Backup"])
    assert out == {"2025-06-28": 1.0}


# ── 套用 mapping 組表 ───────────────────────────────────────────────────────

_SPEC = {
    "Revenue": {"concepts": ["RevenueFromContractWithCustomerExcludingAssessedTax",
                             "Revenues"], "kind": "quarter"},
    "Capex":   {"concepts": ["PaymentsToAcquirePropertyPlantAndEquipment"],
                "kind": "quarter", "negate": True},
    "Total Assets": {"concepts": ["Assets"], "kind": "instant"},
}

_RAW2 = {"facts": {"us-gaap": {
    "Revenues": {"units": {"USD": [
        _fact("2025-03-30", "2025-06-28", 100.0),
        _fact("2025-06-29", "2025-09-27", 120.0),
    ]}},
    "PaymentsToAcquirePropertyPlantAndEquipment": {"units": {"USD": [
        _fact("2025-03-30", "2025-06-28", 7.0),
    ]}},
    "Assets": {"units": {"USD": [
        _fact(None, "2025-06-28", 900.0),
        _fact(None, "2025-09-27", 950.0),
    ]}},
}}}


def test_resolve_row_applies_negate():
    """現行路徑的 Capex 是負數（現金流出），companyfacts 報正數。"""
    out = ff.resolve_row(_RAW2, _SPEC["Capex"], prefer="as_reported")
    assert out == {"2025-06-28": -7.0}


def test_resolve_row_without_negate_keeps_sign():
    out = ff.resolve_row(_RAW2, _SPEC["Revenue"], prefer="as_reported")
    assert out == {"2025-06-28": 100.0, "2025-09-27": 120.0}


def test_build_table_columns_are_the_union_of_all_rows_sorted_by_period_end():
    tbl = ff.build_table(_RAW2, _SPEC, sheet_name="Data_X",
                         fy_end_month=12, ticker="T", prefer="as_reported")
    assert tbl.period_ends == ["2025-06-28", "2025-09-27"]
    assert tbl.quarter_labels == ["FY2025Q2", "FY2025Q3"]
    assert tbl.sheet_name == "Data_X"
    assert tbl.ticker == "T"


def test_build_table_row_order_follows_the_spec_not_the_data():
    """列的順序是模板決定的，不能跟著 dict 或資料順序跑。"""
    tbl = ff.build_table(_RAW2, _SPEC, sheet_name="Data_X",
                         fy_end_month=12, ticker="T", prefer="as_reported")
    assert tbl.concepts == ["Revenue", "Capex", "Total Assets"]


def test_build_table_missing_period_is_none_not_dropped():
    """Capex 只有一期有值，另一期要是 None，不可以整列縮短。"""
    tbl = ff.build_table(_RAW2, _SPEC, sheet_name="Data_X",
                         fy_end_month=12, ticker="T", prefer="as_reported")
    assert tbl.values[tbl.concepts.index("Capex")] == [-7.0, None]
    assert tbl.values[tbl.concepts.index("Total Assets")] == [900.0, 950.0]


def test_build_table_uses_fiscal_year_end_month_for_labels():
    """NVDA 一月結算：2025-06-28 那季是 FY2026Q2，不是 FY2025Q2。"""
    tbl = ff.build_table(_RAW2, _SPEC, sheet_name="Data_X",
                         fy_end_month=1, ticker="T", prefer="as_reported")
    assert tbl.quarter_labels == ["FY2026Q2", "FY2026Q3"]


def test_build_table_empty_facts_gives_empty_table():
    tbl = ff.build_table({"facts": {"us-gaap": {}}}, _SPEC, sheet_name="Data_X",
                         fy_end_month=12, ticker="T", prefer="as_reported")
    assert tbl.quarter_labels == [] and tbl.concepts == list(_SPEC)


# ── 三表分開建、列序照模板（CTH 2026-08-22：原本模板架構要維持）─────────────

_SPEC_IS = {"Revenue": {"concepts": ["Revenues"], "kind": "quarter"},
            "Gross Profit": {"concepts": ["GrossProfit"], "kind": "quarter"}}
_SPEC_BS = {"Total Assets": {"concepts": ["Assets"], "kind": "instant"}}
_SPEC_CF = {"Capex": {"concepts": ["PaymentsToAcquirePropertyPlantAndEquipment"],
                      "kind": "quarter", "negate": True}}


def test_build_statement_tables_returns_three_separate_tables():
    """IS / BS / CF 三張表分開產出，不可以合成一張——下游 _merge_financials()
    與 Q4 合成都是吃三張表的形狀。"""
    tables = ff.build_statement_tables(
        _RAW2, {"IS": _SPEC_IS, "BS": _SPEC_BS, "CF": _SPEC_CF},
        fy_end_month=12, ticker="T", prefer="as_reported")
    assert [t.sheet_name for t in tables] == ["Data_IS", "Data_BS", "Data_CF"]


def test_build_statement_tables_keeps_template_row_order_per_statement():
    tables = ff.build_statement_tables(
        _RAW2, {"IS": _SPEC_IS, "BS": _SPEC_BS, "CF": _SPEC_CF},
        fy_end_month=12, ticker="T", prefer="as_reported")
    by_sheet = {t.sheet_name: t for t in tables}
    assert by_sheet["Data_IS"].concepts == ["Revenue", "Gross Profit"]
    assert by_sheet["Data_BS"].concepts == ["Total Assets"]
    assert by_sheet["Data_CF"].concepts == ["Capex"]


def test_build_statement_tables_share_one_period_axis():
    """三張表的期間欄要一致，否則 _merge_financials() 合出來會錯位。"""
    tables = ff.build_statement_tables(
        _RAW2, {"IS": _SPEC_IS, "BS": _SPEC_BS, "CF": _SPEC_CF},
        fy_end_month=12, ticker="T", prefer="as_reported")
    axes = {tuple(t.period_ends) for t in tables}
    assert len(axes) == 1
    assert axes.pop() == ("2025-06-28", "2025-09-27")


# ── 單位：EPS 是 USD/shares、股數是 shares（2026-08-22 50 家推導時發現）─────
#
# 只讀 USD 的話，Basic/Diluted EPS 與各種股數列永遠找不到 concept——那不是
# 模板列有問題，是取數漏了單位。

_RAW_UNITS = {"facts": {"us-gaap": {
    "EarningsPerShareDiluted": {"units": {"USD/shares": [
        _fact("2025-03-30", "2025-06-28", 1.23),
    ]}},
    "WeightedAverageNumberOfDilutedSharesOutstanding": {"units": {"shares": [
        _fact("2025-03-30", "2025-06-28", 1000.0),
    ]}},
    "Revenues": {"units": {"USD": [_fact("2025-03-30", "2025-06-28", 100.0)]}},
}}}


def test_series_for_concept_reads_usd_per_share_unit():
    out = ff.series_for_concept(_RAW_UNITS, "EarningsPerShareDiluted",
                                kind="quarter", prefer="as_reported", unit="USD/shares")
    assert out == {"2025-06-28": 1.23}


def test_series_for_concept_reads_shares_unit():
    out = ff.series_for_concept(_RAW_UNITS, "WeightedAverageNumberOfDilutedSharesOutstanding",
                                kind="quarter", prefer="as_reported", unit="shares")
    assert out == {"2025-06-28": 1000.0}


def test_series_for_concept_defaults_to_usd():
    """沒指定單位就是 USD——絕大多數列都是金額，不要每列都得寫。"""
    assert ff.series_for_concept(_RAW_UNITS, "Revenues", kind="quarter",
                                 prefer="as_reported") == {"2025-06-28": 100.0}


def test_series_for_concept_wrong_unit_returns_empty():
    """單位指錯要回空，不可以偷偷退回 USD——那會讓 EPS 抓到金額。"""
    assert ff.series_for_concept(_RAW_UNITS, "EarningsPerShareDiluted",
                                 kind="quarter", prefer="as_reported") == {}


def test_resolve_row_honours_unit_in_spec():
    spec = {"concepts": ["EarningsPerShareDiluted"], "kind": "quarter",
            "unit": "USD/shares"}
    assert ff.resolve_row(_RAW_UNITS, spec, prefer="as_reported") == {"2025-06-28": 1.23}


# ── taxonomy：流通股數在 dei 不在 us-gaap（2026-08-22 50 家推導時發現）──────

_RAW_DEI = {"facts": {
    "us-gaap": {"Assets": {"units": {"USD": [_fact(None, "2025-06-28", 900.0)]}}},
    "dei": {"EntityCommonStockSharesOutstanding": {"units": {"shares": [
        _fact(None, "2025-06-28", 2450.0),
    ]}}},
}}


def test_series_for_concept_reads_dei_taxonomy():
    out = ff.series_for_concept(_RAW_DEI, "EntityCommonStockSharesOutstanding",
                                kind="instant", prefer="as_reported",
                                unit="shares", taxonomy="dei")
    assert out == {"2025-06-28": 2450.0}


def test_series_for_concept_defaults_to_us_gaap_taxonomy():
    assert ff.series_for_concept(_RAW_DEI, "EntityCommonStockSharesOutstanding",
                                 kind="instant", prefer="as_reported",
                                 unit="shares") == {}


def test_resolve_row_honours_taxonomy_in_spec():
    spec = {"concepts": ["EntityCommonStockSharesOutstanding"], "kind": "instant",
            "unit": "shares", "taxonomy": "dei"}
    assert ff.resolve_row(_RAW_DEI, spec, prefer="as_reported") == {"2025-06-28": 2450.0}
