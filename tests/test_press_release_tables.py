"""press_release_tables 的確定性解析測試（TODO B3）。

全部離線：`tests/fixtures/arlo_recon_table.html` 是 ARLO 2026 Q2 新聞稿裡
那張 Non-GAAP 調節表的原始 HTML（只砍掉 style/class 屬性，欄位結構原封不動），
所以 Workiva 的空白間隔欄與重複欄位都還在，正是這個模組要清掉的東西。
"""
from pathlib import Path

import pytest

from press_release_tables import (
    PressTable,
    clean_grid,
    filter_nongaap,
    is_nongaap_table,
    parse_tables,
)

FIXTURE = Path(__file__).parent / "fixtures" / "arlo_recon_table.html"


# ── clean_grid：Workiva 版面雜訊 ─────────────────────────────────────────────

def test_drops_all_empty_rows_and_columns():
    grid = [
        ["", "", "", ""],
        ["Revenue", "", "100", ""],
        ["", "", "", ""],
    ]
    assert clean_grid(grid) == [["Revenue", "100"]]


def test_collapses_duplicated_value_cells():
    """Workiva 把同一個數字重複寫進相鄰欄，收斂成一格。"""
    grid = [
        ["SBC", "SBC", "SBC", "21710", "21710", ""],
    ]
    assert clean_grid(grid) == [["SBC", "21710"]]


def test_currency_symbol_column_merges_into_value():
    """`$` 自成一欄、數字在隔壁欄——輸出只留數字。"""
    grid = [
        ["GAAP net income", "GAAP net income", "$", "3028", ""],
    ]
    assert clean_grid(grid) == [["GAAP net income", "3028"]]


def test_percent_suffix_stays_with_the_number():
    grid = [
        ["Gross margin", "48.2", "48.2", "%"],
    ]
    assert clean_grid(grid) == [["Gross margin", "48.2%"]]


def test_parenthesised_negative_survives():
    grid = [
        ["Gain on sale", "Gain on sale", "(6,423)", "(6,423)", ""],
    ]
    assert clean_grid(grid) == [["Gain on sale", "(6,423)"]]


def test_all_empty_column_separates_two_value_groups():
    """全空欄是 Workiva 的間隔欄，用來切開兩個期間，不可以直接刪掉再合併。"""
    grid = [
        ["Revenue", "$", "155937", "", "$", "150382"],
        ["Cost",    "80721", "80721", "", "77714", "77714"],
    ]
    assert clean_grid(grid) == [
        ["Revenue", "155937", "150382"],
        ["Cost", "80721", "77714"],
    ]


def test_rows_are_rectangular():
    grid = [
        ["Revenue", "$", "155937", "", "$", "150382"],
        ["Note", "", "", "", "", ""],
    ]
    out = clean_grid(grid)
    assert len({len(r) for r in out}) == 1


def test_conflicting_values_in_one_group_are_kept_visible():
    """同一格擠出兩個不同數字時寧可原樣併排，也不要靜默丟掉一個。"""
    grid = [["Weird", "1", "2", ""]]
    assert clean_grid(grid) == [["Weird", "1 2"]]


def test_colspan_title_row_collapses_to_first_cell():
    """`NVIDIA CORPORATION` 被 colspan 展成 4 欄，只留一格。"""
    grid = [
        ["NVIDIA CORPORATION", "NVIDIA CORPORATION", "", "NVIDIA CORPORATION"],
        ["Revenue", "100", "", "200"],
    ]
    assert clean_grid(grid) == [
        ["NVIDIA CORPORATION", "", ""],
        ["Revenue", "100", "200"],
    ]


def test_identical_numbers_across_periods_are_not_collapsed():
    """各期數字剛好相同時不可以清空——那是真資料不是 colspan。"""
    grid = [["Shares", "108123", "", "108123"]]
    assert clean_grid(grid) == [["Shares", "108123", "108123"]]


def test_empty_grid():
    assert clean_grid([]) == []


# ── parse_tables：真實 ARLO 調節表 ──────────────────────────────────────────

@pytest.fixture(scope="module")
def arlo_tables() -> list[PressTable]:
    return parse_tables(FIXTURE.read_text(encoding="utf-8"))


def test_parses_single_table_from_fixture(arlo_tables):
    assert len(arlo_tables) == 1


def test_gaap_net_income_row_has_all_five_periods(arlo_tables):
    """調節表數字必須完整——這是 pandas.read_html 路線值不值得做的判準。"""
    row = _row_starting_with(arlo_tables[0], "GAAP net income")
    assert row[1:] == ["3028", "14877", "3124", "17905", "2289"]


def test_nongaap_net_income_row_complete(arlo_tables):
    row = _row_starting_with(arlo_tables[0], "Non-GAAP net income")
    assert row[1:] == ["31098", "30964", "18815", "62062", "35289"]


def test_negative_adjustment_preserved(arlo_tables):
    row = _row_starting_with(arlo_tables[0], "Gain on sale of long-term investment")
    assert "(6,423)" in row


def test_period_header_row_present(arlo_tables):
    """期間標籤要留著——沒有它，skill 沒辦法知道哪一欄是哪一季。"""
    header = arlo_tables[0].rows[1]
    assert header[1] == "June 28, 2026"
    assert "June 29, 2025" in header


def test_all_rows_same_width(arlo_tables):
    assert len({len(r) for r in arlo_tables[0].rows}) == 1


def test_cleaned_table_is_compact(arlo_tables):
    """(24, 30) 的原始網格清完應該剩個位數欄寬。"""
    tbl = arlo_tables[0]
    assert tbl.n_cols <= 8, tbl.rows[0]
    assert len(tbl.text()) < 2000


# ── 篩選 ────────────────────────────────────────────────────────────────────

def test_reconciliation_table_is_flagged(arlo_tables):
    assert is_nongaap_table(arlo_tables[0])


def test_plain_table_is_not_flagged():
    tbl = PressTable(index=0, rows=[["Revenue", "100"], ["Cost", "40"]])
    assert not is_nongaap_table(tbl)


def test_reconciliation_in_caption_is_enough():
    tbl = PressTable(index=0, rows=[["Net income", "100"]],
                     caption="RECONCILIATIONS OF GAAP MEASURES")
    assert is_nongaap_table(tbl)


def test_reconciliation_only_in_body_does_not_count():
    """現金流量表固定有「Reconciliation of cash...」一列，不能因此被當調節表。"""
    tbl = PressTable(
        index=0,
        rows=[["Reconciliation of cash, cash equivalents and restricted cash", ""],
              ["Cash and cash equivalents", "101382"]],
        caption="UNAUDITED CONDENSED CONSOLIDATED STATEMENTS OF CASH FLOWS",
    )
    assert not is_nongaap_table(tbl)


def test_filter_keeps_only_matching_tables(arlo_tables):
    plain = PressTable(index=1, rows=[["Revenue", "100"]])
    assert filter_nongaap(arlo_tables + [plain]) == arlo_tables


def _row_starting_with(tbl: PressTable, label: str) -> list[str]:
    for row in tbl.rows:
        if row and row[0].startswith(label):
            return row
    raise AssertionError(f"找不到以 {label!r} 開頭的列：{[r[0] for r in tbl.rows]}")
