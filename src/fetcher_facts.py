"""fetcher_facts.py — 走 SEC companyfacts API 的取數路徑（TODO G11 spike）。

**這條路徑還沒接上主流程。** 它是現行「逐份下載並解析 filing」的平行替代品，
目的是先產出逐格比對報告，讓 CTH 看數據決定要不要切換。

## 為什麼要有這條路

現行路徑對每一份 10-Q/10-K 各下載並解析一次 XBRL，而且**同一份 filing 會被
解析 4 次**（IS/BS/CF/segments 各一次，override 觸發時最多 7 次）。實測：

    下載 0.3 秒（之後被 edgartools 的 ~/.edgar/_tcache 快取成 0 秒）
    XBRL 解析 1.3~2.1 秒，**完全沒有快取**

六家公司 × 75 份 × 4 次 ≈ 1,800 次解析 ≈ 45 分鐘。

companyfacts 是 SEC 官方把「這家公司歷來 tag 過的所有 XBRL fact」整理成一包
JSON，**一家公司一個 request**。實測 NVDA：下載 0.32 秒、JSON 解析 0.02 秒、
4.0 MB、626 個 us-gaap concept、66 期單季資料（2008-07 起，比現行路徑還早一年）。

## 它順便解掉的問題

每筆 fact 自帶 `start`/`end`，所以：
  - 篩 3 個月的 duration 就是單季，**不用 YTD 拆算**（TODO G3 的根因消失）
  - 10-K 也會 tag Q4 的 3 個月 duration，**Q4 不用「年報−Q1−Q2−Q3」合成**
  - **完全不需要猜 `(Qn)`**（TODO D0-6 那類 label 碰撞不會發生）
每筆自帶 `accn`/`form`/`filed`，所以重編取哪一版變成明確選擇（TODO G8 消失）。

## 它拿不到的東西（**這是限制，不是待辦**）

fact 的欄位只有 `start`/`end`/`val`/`accn`/`fy`/`fp`/`form`/`filed`/`frame`
——**沒有任何維度欄位**，companyfacts 只收合併總數。所以：
  1. `Data_Segments`（帶維度的分類細項）**這條路拿不到**，非走解 filing 不可
  2. 沒有 presentation linkbase → 沒有報表結構、沒有公司自報的原文標籤。
     每個 concept 只有 US-GAAP 分類的**官方標準標籤**（如
     `Share-based Payment Arrangement, Noncash Expense`），不是公司自己寫的
"""

from __future__ import annotations

from datetime import date
from typing import Any, Iterable

COMPANYFACTS_URL = "https://data.sec.gov/api/xbrl/companyfacts/CIK{cik:010d}.json"

# 單季／年度的天數區間。**不可以只認「剛好 91 天」或「剛好 365 天」**——
# 美股多用 52/53 週制，13 週季實測落在 84~98 天，會計年度落在 357~371 天，
# 再加上公司偶爾有過渡期，區間要留寬一點，但不能寬到把半年（~180）或
# 九個月 YTD（~270）吃進來。
QUARTER_DAYS = (80, 100)
ANNUAL_DAYS = (330, 400)

_PREFERENCES = ("as_reported", "latest")


def duration_days(fact: dict) -> int | None:
    """fact 涵蓋幾天。時點值（沒有 `start`）回 None。"""
    start = fact.get("start")
    if not start:
        return None
    try:
        return (date.fromisoformat(fact["end"]) - date.fromisoformat(start)).days
    except (KeyError, TypeError, ValueError):
        return None


def classify_period(fact: dict) -> str | None:
    """fact → `"quarter"` / `"annual"` / `"instant"`。都不是回 None。

    回 None 的典型是半年報與九個月 YTD——**那些一定要丟掉**，混進單季序列
    就是舊路徑最容易出錯的地方。
    """
    days = duration_days(fact)
    if days is None:
        return "instant" if fact.get("end") else None
    if QUARTER_DAYS[0] <= days <= QUARTER_DAYS[1]:
        return "quarter"
    if ANNUAL_DAYS[0] <= days <= ANNUAL_DAYS[1]:
        return "annual"
    return None


def pick_fact(facts: Iterable[dict], *, prefer: str) -> dict | None:
    """同一期間有多筆（每份提到它的 filing 各一筆）時挑一筆。

    `prefer="as_reported"`  取 `filed` 最早的＝當初申報值。符合分析師
                            「回看那個時點看得到什麼」的直覺，也是預設。
    `prefer="latest"`       取 `filed` 最新的＝含後續重編。

    `filed` 同一天時用 `accn` 當第二排序鍵——**不可以看 list 順序**，
    那會讓同一份輸入在不同執行產出不同結果。
    """
    if prefer not in _PREFERENCES:
        raise ValueError(f"prefer must be one of {list(_PREFERENCES)}, got {prefer!r}")
    items = list(facts)
    if not items:
        return None
    key = lambda f: (f.get("filed", ""), f.get("accn", ""))
    return min(items, key=key) if prefer == "as_reported" else max(items, key=key)


def _unit_facts(raw: dict, concept: str, unit: str = "USD",
                taxonomy: str = "us-gaap") -> list[dict]:
    """取某個 us-gaap concept 在指定單位下的事實。沒有就回空 list。

    **指錯單位一定要回空，不可以偷偷退回 USD**——EPS 是 `USD/shares`、股數是
    `shares`，退回 USD 會讓每股盈餘抓到金額，而且數字看起來很合理不會有人發現。
    """
    node = raw.get("facts", {}).get(taxonomy, {}).get(concept)
    if not node:
        return []
    return list(node.get("units", {}).get(unit, []))


def series_for_concept(
    raw: dict,
    concept: str,
    *,
    kind: str,
    prefer: str,
    fallbacks: list[str] | None = None,
    unit: str = "USD",
    taxonomy: str = "us-gaap",
) -> dict[str, Any]:
    """concept → `{期末日: 值}`，只收 `kind` 那一種期間。

    `fallbacks` 依序試，**primary 有資料就不看 fallback**——同一個經濟意義
    會跨 concept（NVDA 早年是 `Revenues`、後來換成
    `RevenueFromContractWithCustomerExcludingAssessedTax`），但兩個都有值時
    以模板指定的 primary 為準，不要混著用。
    """
    for name in [concept, *(fallbacks or [])]:
        buckets: dict[str, list[dict]] = {}
        for fact in _unit_facts(raw, name, unit, taxonomy):
            if classify_period(fact) != kind:
                continue
            buckets.setdefault(fact["end"], []).append(fact)
        if buckets:
            return {
                end: pick_fact(items, prefer=prefer)["val"]
                for end, items in buckets.items()
            }
    return {}


# ── 套用 mapping 組表 ───────────────────────────────────────────────────────
#
# mapping 的形狀（見 `facts_mapping.py`）：
#
#     {"列名": {"concepts": [依序試的 us-gaap element name],
#               "kind": "quarter" | "annual" | "instant",
#               "unit": "USD" | "USD/shares" | "shares",   # 可省略，預設 USD
#               "taxonomy": "us-gaap" | "dei",             # 可省略，預設 us-gaap
#               "negate": True}}          # 可省略，預設 False
#
# **`unit` 為什麼需要**：EPS 是 `USD/shares`、股數是 `shares`。2026-08-22 跑
# 50 家推導時，Basic/Diluted EPS 與三個股數列「完全找不到 concept」，一度以為
# 是模板列有問題，實際上是取數只讀了 USD。
#
# **`negate` 為什麼需要**：現行路徑對某些列做過符號正規化（Capex 記成現金流出
# 的負數、Interest Expense 記成負數），companyfacts 給的是公司原始 tag 的正號。
# 要對齊既有輸出就得逐列標。哪些列要 negate 是 `scripts/spike_derive_mapping.py`
# 用現行路徑的數字反推出來的，不是憑印象填。


def resolve_row(raw: dict, spec: dict, *, prefer: str) -> dict[str, Any]:
    """一列的 mapping → `{期末日: 值}`。concepts 依序試，第一個有資料的就用。"""
    concepts = spec.get("concepts") or []
    if not concepts:
        return {}
    series = series_for_concept(raw, concepts[0], kind=spec["kind"], prefer=prefer,
                                fallbacks=list(concepts[1:]),
                                unit=spec.get("unit", "USD"),
                                taxonomy=spec.get("taxonomy", "us-gaap"))
    if spec.get("negate"):
        return {k: -v for k, v in series.items()}
    return series


def build_table(
    raw: dict,
    mapping: dict[str, dict],
    *,
    sheet_name: str,
    fy_end_month: int,
    ticker: str,
    prefer: str,
):
    """mapping + companyfacts → `StatementTable`，形狀跟現行路徑產出的一致。

    欄位是所有列期末日的**聯集**，依日期排序（不是依標籤字串——標籤排序在
    跨財年時不保證等於時間順序）。某一列在某一期沒有值就是 `None`，
    **不可以把那一欄整個丟掉**，否則使用者看不出有漏（TODO G6 的精神）。
    """
    from fetcher_gaap import StatementTable          # 延後 import，避免循環匯入
    from fiscal_input import fiscal_quarter_of, fy_start_month

    rows = {name: resolve_row(raw, spec, prefer=prefer) for name, spec in mapping.items()}
    ends = sorted({e for series in rows.values() for e in series})
    start_month = fy_start_month(fy_end_month)

    return StatementTable(
        sheet_name=sheet_name,
        quarter_labels=[fiscal_quarter_of(e, start_month) for e in ends],
        filing_dates=[""] * len(ends),
        concepts=list(mapping),
        values=[[rows[name].get(e) for e in ends] for name in mapping],
        ticker=ticker,
        labels=[""] * len(mapping),
        period_ends=list(ends),
    )


# 三張表的 sheet 名稱與順序。**跟現行路徑一致**，因為下游的
# `_merge_financials()` 與 `_synthesize_q4()` 都是吃這個形狀。
_SHEET_BY_STATEMENT = {"IS": "Data_IS", "BS": "Data_BS", "CF": "Data_CF"}
_STATEMENT_ORDER = ("IS", "BS", "CF")


def build_statement_tables(
    raw: dict,
    mappings: dict[str, dict],
    *,
    fy_end_month: int,
    ticker: str,
    prefer: str,
) -> list:
    """三張表分開產出，形狀跟現行路徑一致。

    CTH 2026-08-22 明確要求「原本模板的格式架構要維持，包含排序方式、三表分別」：
      - IS / BS / CF **分開三張表**，不合成一張（下游 `_merge_financials()`
        與 `_synthesize_q4()` 都是吃三張表的形狀）
      - 每張表的**列序完全照模板**，不跟著資料或 concept 名稱跑
      - 列名（機器鍵）維持英文原樣，下游腳本與 Excel 公式不受影響

    **三張表共用同一條期間軸**：期間欄取三張表的聯集後統一，否則
    `_merge_financials()` 合起來會錯位（IS 有 20 欄、BS 有 22 欄的話，
    合併時對不上同一期）。
    """
    ends = sorted({
        e
        for mapping in mappings.values()
        for spec in mapping.values()
        for e in resolve_row(raw, spec, prefer=prefer)
    })

    from fetcher_gaap import StatementTable
    from fiscal_input import fiscal_quarter_of, fy_start_month

    start_month = fy_start_month(fy_end_month)
    labels = [fiscal_quarter_of(e, start_month) for e in ends]

    tables = []
    for tag in _STATEMENT_ORDER:
        mapping = mappings.get(tag) or {}
        rows = {name: resolve_row(raw, spec, prefer=prefer)
                for name, spec in mapping.items()}
        tables.append(StatementTable(
            sheet_name=_SHEET_BY_STATEMENT[tag],
            quarter_labels=list(labels),
            filing_dates=[""] * len(ends),
            concepts=list(mapping),
            values=[[rows[name].get(e) for e in ends] for name in mapping],
            ticker=ticker,
            labels=[""] * len(mapping),
            period_ends=list(ends),
        ))
    return tables
