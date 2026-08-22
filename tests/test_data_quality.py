"""data_quality.py — 抓取結果的缺漏判斷（TODO G5 改版）。

CTH 2026-08-22 定案的三個判斷，可信度由高到低：

    A 季度斷層          相鄰兩期差太多 → 中間漏了幾季      誤判率 0
    B 中間有洞          同一列有些期有、有些期沒有          誤判率 0
    C 整列全空且矛盾    空白但相關欄位顯示它應該要有        誤判率低

**三個都不需要「同業基準表」**。原本提案主打的「52 家普及率」被降級成參考
資訊——它做不到 C 做得到的事，而且正是「公司真的沒有這個科目就被永遠標紅」
這種誤判的來源。
"""
import pytest

import data_quality as dq
from fetcher_gaap import StatementTable


_KNOWN = frozenset({"Revenue", "ROU", "Old", "X", "A", "B", "Nothing",
                    "Long-term Debt", "Short-term Debt", "Interest Expense",
                    "Debt Proceeds", "Debt Repayments"}
                   | {f"R{i}" for i in range(12)})


def _many(pattern, n=10):
    """湊足 `_SPARSE_MIN_ROWS` 的列數——稀疏判斷在太少列的表上不啟動。"""
    return [f"R{i}" for i in range(n)], [list(pattern) for _ in range(n)]


def _tbl(concepts, values, ends):
    return StatementTable(
        sheet_name="Data_Financials(Q)",
        quarter_labels=[""] * len(ends),
        filing_dates=[""] * len(ends),
        concepts=concepts, values=values,
        ticker="T", labels=[""] * len(concepts), period_ends=ends,
    )


# ── A. 季度斷層 ─────────────────────────────────────────────────────────────

def test_missing_quarters_none_when_consecutive():
    ends = ["2025-03-29", "2025-06-28", "2025-09-27", "2025-12-27"]
    assert dq.missing_quarters(ends) == []


def test_missing_quarters_finds_one_gap():
    ends = ["2025-03-29", "2025-06-28", "2025-12-27"]      # 少了 9 月那季
    gaps = dq.missing_quarters(ends)
    assert len(gaps) == 1
    assert gaps[0].after == "2025-06-28" and gaps[0].before == "2025-12-27"
    assert gaps[0].count == 1


def test_missing_quarters_accepts_costco_16_week_quarter():
    """COSTCO 的第四季是 16 週（112~119 天），是正常的一季不是缺季。

    52 家 1,482 對相鄰期間裡，111~150 天的 16 筆全部是 COSTCO。
    **所以不能用固定門檻（例如「>120 天算缺」），要用 round 除法。**
    """
    assert dq.missing_quarters(["2025-05-11", "2025-08-31"]) == []


def test_missing_quarters_counts_two_missing():
    """相隔約三季 → 中間漏兩季。"""
    gaps = dq.missing_quarters(["2025-03-29", "2025-12-27"])
    assert gaps and gaps[0].count == 2


def test_missing_quarters_ignores_duplicate_period_end():
    """實測 SNOW 有兩欄期末日都是 2022-01-31，那是重複列不是缺季。"""
    assert dq.missing_quarters(["2025-03-29", "2025-03-29", "2025-06-28"]) == []


def test_missing_quarters_ignores_unparseable_dates():
    assert dq.missing_quarters(["", "2025-06-28", "2025-09-27"]) == []


# ── B. 中間有洞 ─────────────────────────────────────────────────────────────

_ENDS4 = ["2025-03-29", "2025-06-28", "2025-09-27", "2025-12-27"]


def test_holed_rows_flags_a_gap_in_the_middle():
    t = _tbl(["Revenue"], [[1.0, None, 3.0, 4.0]], _ENDS4)
    holes = dq.holed_rows(t, _KNOWN)
    assert [h.row for h in holes] == ["Revenue"]
    assert holes[0].have == 3 and holes[0].span == 4


def test_holed_rows_ignores_a_row_that_legitimately_starts_late():
    """`Operating Lease ROU Assets` 只有 28/67 期不是漏抓——租賃準則
    ASC 842 從 2019 才適用，之前本來就沒有這一列。

    **只看「第一個有值」到「最後一個有值」之間有沒有洞**，前後空白不算。
    這條沒處理好會製造一大堆假警報。
    """
    t = _tbl(["ROU"], [[None, None, 3.0, 4.0]], _ENDS4)
    assert dq.holed_rows(t, _KNOWN) == []


def test_holed_rows_ignores_a_row_that_stops_being_reported():
    t = _tbl(["Old"], [[1.0, 2.0, None, None]], _ENDS4)
    assert dq.holed_rows(t, _KNOWN) == []


def test_holed_rows_ignores_all_empty_row():
    """整列全空歸 C 判斷，不是洞。"""
    t = _tbl(["X"], [[None] * 4], _ENDS4)
    assert dq.holed_rows(t, _KNOWN) == []


def test_holed_rows_ignores_full_row():
    t = _tbl(["Revenue"], [[1.0, 2.0, 3.0, 4.0]], _ENDS4)
    assert dq.holed_rows(t, _KNOWN) == []


def test_holed_rows_sorted_worst_first():
    t = _tbl(["A", "B"],
             [[1.0, None, None, 4.0],      # 2/4
              [1.0, 2.0, None, 4.0]],      # 3/4
             _ENDS4)
    assert [h.row for h in dq.holed_rows(t, _KNOWN)] == ["A", "B"]


# ── C. 整列全空且與其他欄位矛盾 ────────────────────────────────────────────

def test_contradictions_flags_missing_debt_flows_when_debt_exists():
    """NVDA 實測：有 74 億長期負債、有利息費用，卻完全沒有借還款紀錄。"""
    t = _tbl(["Long-term Debt", "Interest Expense", "Debt Proceeds", "Debt Repayments"],
             [[100.0] * 4, [-5.0] * 4, [None] * 4, [None] * 4], _ENDS4)
    found = {c.row for c in dq.contradictions(t, _KNOWN)}
    assert found == {"Debt Proceeds", "Debt Repayments"}


def test_contradictions_stays_quiet_when_the_company_simply_has_no_debt():
    """**這條就是 CTH 擔心的誤判**：公司真的沒負債，就不該標紅。"""
    t = _tbl(["Long-term Debt", "Interest Expense", "Debt Proceeds", "Debt Repayments"],
             [[None] * 4, [None] * 4, [None] * 4, [None] * 4], _ENDS4)
    assert dq.contradictions(t, _KNOWN) == []


def test_contradictions_ignores_rows_that_have_values():
    """有值就不是「整列全空」，歸 B 判斷。"""
    t = _tbl(["Long-term Debt", "Debt Repayments"],
             [[100.0] * 4, [None, 2.0, None, None]], _ENDS4)
    assert [c.row for c in dq.contradictions(t, _KNOWN)] == []


def test_contradictions_reason_names_the_evidence():
    """理由要講出「憑什麼說它該有」，不然使用者無從判斷。"""
    t = _tbl(["Long-term Debt", "Debt Repayments"],
             [[100.0] * 4, [None] * 4], _ENDS4)
    c = dq.contradictions(t, _KNOWN)[0]
    assert "Long-term Debt" in c.evidence


def test_contradictions_missing_rows_do_not_crash():
    t = _tbl(["Revenue"], [[1.0] * 4], _ENDS4)
    assert dq.contradictions(t, _KNOWN) == []


# ── 整合 ────────────────────────────────────────────────────────────────────

def test_assess_returns_all_three_and_counts_rows():
    t = _tbl(["Long-term Debt", "Debt Repayments", "Revenue", "Nothing"],
             [[100.0, 100.0, 100.0],
              [None, None, None],
              [1.0, None, 3.0],
              [None, None, None]],
             ["2025-03-29", "2025-06-28", "2025-12-27"])
    r = dq.assess(t, _KNOWN)
    assert r.total_periods == 3
    assert len(r.missing_quarters) == 1          # 9 月那季不見了
    assert [h.row for h in r.holed] == ["Revenue"]
    assert [c.row for c in r.contradictions] == ["Debt Repayments"]
    # 「整列全空但沒有矛盾」的不列出來，只計數
    assert r.empty_but_plausible == 1            # Nothing


def test_overflow_rows_are_excluded_by_default():
    """`Other (as reported)` 的公司特有科目本來就會斷斷續續（見 TODO G4），
    算進來的話 NVDA 會報 85 列有洞、絕大多數是雜訊。預設只評估模板列。"""
    t = _tbl(["Revenue", "Premium amortization on investments, net"],
             [[1.0, None, 3.0, 4.0], [1.0, None, 3.0, 4.0]], _ENDS4)
    assert [h.row for h in dq.holed_rows(t)] == ["Revenue"]


# ── D. 整欄稀疏（欄位在、但整排幾乎都空）────────────────────────────────────
#
# 實測發現的：合成 Q4 失敗時，**每一個流量列都會在那一期出現一個洞**。
# 那不是 40 個獨立的列問題，是一個期間問題。不收攏的話 B 的清單會被同一期
# 洗版（KO/ARLO/JPM 都出現 Revenue(20/21)、Cost of Revenue(20/21)… 一整排）。

def test_sparse_periods_flags_a_mostly_empty_column():
    names, vals = _many([1.0, None, 3.0])
    vals[0] = [1.0, 2.0, 3.0]                      # 只有一列在中間那期有值
    t = _tbl(names, vals, ["2025-03-29", "2025-06-28", "2025-09-27"])
    sparse = dq.sparse_periods(t, _KNOWN)
    assert [s.period_end for s in sparse] == ["2025-06-28"]
    assert sparse[0].filled == 1 and sparse[0].total == 10


def test_sparse_periods_not_applied_to_a_tiny_table():
    """列數太少時不做稀疏判斷，否則一個 None 就把整欄判成稀疏。"""
    t = _tbl(["Revenue"], [[1.0, None, 3.0]],
             ["2025-03-29", "2025-06-28", "2025-09-27"])
    assert dq.sparse_periods(t, _KNOWN) == []


def test_sparse_periods_ignores_a_healthy_column():
    names, vals = _many([1.0, 2.0])
    assert dq.sparse_periods(_tbl(names, vals, ["2025-03-29", "2025-06-28"]), _KNOWN) == []


def test_holed_rows_skips_periods_that_are_sparse_for_everyone():
    """整欄稀疏的那一期不算進個別列的洞——否則同一件事會被報 40 次。"""
    names, vals = _many([1.0, None, 3.0])
    assert dq.holed_rows(_tbl(names, vals, ["2025-03-29", "2025-06-28", "2025-09-27"]),
                         _KNOWN) == []


def test_holed_rows_still_flags_a_row_specific_hole():
    """只有這一列缺、別人都有 → 還是要報。"""
    names, vals = _many([1.0, 2.0, 3.0])
    vals[3] = [1.0, None, 3.0]
    holes = dq.holed_rows(_tbl(names, vals, ["2025-03-29", "2025-06-28", "2025-09-27"]),
                          _KNOWN)
    assert [h.row for h in holes] == ["R3"]


# ── 模板不適用（2026-08-22，52 家實測發現）─────────────────────────────────
#
# BAC / GS / SCHW / PLD 四家的稀疏欄接近全部（21 期裡 19~21 欄）。那不是抓漏，
# 是 IS/BS/CF 模板為製造業設計、金融股與 REIT 的報表結構完全不同（TODO D8）。
# 對這種情況正確的訊息是「模板不適用」，不是列出 21 行稀疏欄。

def test_template_mismatch_when_most_periods_are_sparse():
    names, vals = _many([None, None, 3.0])      # 三期裡兩期整排空
    t = _tbl(names, vals, ["2025-03-29", "2025-06-28", "2025-09-27"])
    r = dq.assess(t, _KNOWN)
    assert r.template_mismatch is True


def test_template_mismatch_false_for_a_normal_company():
    names, vals = _many([1.0, 2.0, 3.0])
    r = dq.assess(_tbl(names, vals, ["2025-03-29", "2025-06-28", "2025-09-27"]), _KNOWN)
    assert r.template_mismatch is False
