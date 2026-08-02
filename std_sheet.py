"""
std_sheet.py — `Data_Std`：跨公司版面完全固定的標準表。

═══════════════════════════════════════════════════════════════════════════════
要解決的問題
═══════════════════════════════════════════════════════════════════════════════
現行輸出有**三個軸同時在變**，所以沒辦法寫一份通用模板去對照不同公司。
實測 `output/` 既有 11 個檔案：

    檔案        總列  季欄數  Revenue  Cash  Total Assets  FCF   sheet 數
    AAPL         96    50        4     28        41        96     12
    AMD         217    45        4     47        60       137     29
    ARLO        124     4        4     35        48       106     10
    COHR        219    47        4     56        69       147     26
    TSLA        238    45        4     55        68       158     30

  1. sheet 名稱與數量：10～30 張，而且是 `Data_Seg_usgaapRevenueFromContr`
     這種截斷過的 XBRL 概念名
  2. 列位置：`Cash` 落在第 28～56 列之間。因為 overflow 行插在每個 section 的
     模板行之後，IS 多幾行 overflow，BS 整段就往下移幾行
  3. 欄位置：季度數 4～50

唯一能跨公司直接參照的格子只有 `C4`（最舊一季的營收）。`VLOOKUP` 也不安全——
`Net Income` 在 IS 與 CF 各出現一次，只會抓到第一個。

═══════════════════════════════════════════════════════════════════════════════
Data_Std 的三個保證
═══════════════════════════════════════════════════════════════════════════════
  1. **固定 sheet 名稱**，永遠產生（沒資料就整列空白）
  2. **固定列位**：只放固定模板行，overflow 一律不進來（那正是害列位浮動的
     原因）。列號由 `FROZEN_ROW_NUMBERS` 的測試釘住，任何人插入一列都會紅
  3. **固定機器鍵**（B 欄）：`IS.REVENUE` / `BS.CASH` / `CF.NET_INCOME` /
     `RATIO.毛利率`。用 `MATCH` 對機器鍵比用列號更保險，即使日後調整列序也不會壞

欄位維持**舊→新**（與其他 sheet 一致）。

═══════════════════════════════════════════════════════════════════════════════
兩種期間標籤：財季 vs 日曆季
═══════════════════════════════════════════════════════════════════════════════
現行標籤 `FY2026Q1` 是**公司財季**，不是日曆季。`_col_to_quarter_label()` 對非
12 月結算的公司會把年份往後推，所以：

    ARLO（12 月結算）FY2026Q1 → 2026 年 3 月底 → 日曆 2026Q1
    AAPL（ 9 月結算）FY2026Q1 → 2025 年 12 月底 → 日曆 2025Q4
    NVDA（ 1 月結算）FY2026Q1 → 2025 年  4 月底 → 日曆 2025Q2

三家的 `FY2026Q1` 差了將近一年。模板若照財季標籤對齊，會把 NVDA 的 2025 年 4 月
拿去比 ARLO 的 2026 年 3 月，而且完全看不出來。

所以 `Data_Std` 同時給兩種標籤，由模板自己選要用哪個對齊：

    第 1 列  財季標籤    FY2025Q3   FY2025Q4   FY2026Q1     ← 比較同一家的營運週期
    第 2 列  申報日      2025-11-06 2026-02-26 2026-05-07
    第 3 列  日曆季      2025Q3     2025Q4     2026Q1       ← 跨公司對齊用這列
    第 4 列  期末年月    2025-09    2025-12    2026-03
    第 5 列  資料版本    STD_V1     ...

跨公司比較用第 3 列（`HLOOKUP("2026Q1", ...)`）；同一家看營運週期用第 1 列。
"""

from __future__ import annotations

import re
from typing import Any

from fetcher_gaap import BS_TEMPLATE, CF_TEMPLATE, IS_TEMPLATE, StatementTable
from ratios import RATIO_DEFS

SCHEMA_VERSION = "STD_V1"
SHEET_NAME     = "Data_Std"

# 表頭列的機器鍵
ROW_CALENDAR   = "META.CALENDAR_QUARTER"
ROW_PERIOD_END = "META.PERIOD_END"
ROW_SCHEMA     = "META.SCHEMA"

# 來源表的 section 標題（`_merge_financials` 產生的）
SEC_IS = "Income Statement"
SEC_BS = "Balance Sheet"
SEC_CF = "Cash Flow"
SEC_RATIO = "Ratios"

_QUARTER_RE  = re.compile(r"FY(\d{4})Q([1-4])$")
_UNIT_SUFFIX = re.compile(r"\s*\((?:%|x|days|\$)\)\s*$")


# ── 期間換算 ────────────────────────────────────────────────────────────────

def _fiscal_period_end(label: str, fy_end_month: int) -> tuple[int, int] | None:
    """財季標籤 → (西元年, 月)，指該財季**結束**的年月。無法解析回 None。

    財年 Y 結束於西元 Y 年的 fy_end_month 月（SEC 慣例）。第 q 財季結束於
    財年結束前 (4-q) 季，也就是 fy_end_month - 3*(4-q) 個月。
    """
    m = _QUARTER_RE.match((label or "").strip())
    if m is None:
        return None
    year, q = int(m.group(1)), int(m.group(2))
    month = fy_end_month - 3 * (4 - q)
    while month <= 0:
        month += 12
        year -= 1
    return year, month


def _calendar_quarter(label: str, fy_end_month: int) -> str:
    """財季標籤 → 日曆季標籤（`2026Q1`）。年度標籤或無法解析回空字串。"""
    parsed = _fiscal_period_end(label, fy_end_month)
    if parsed is None:
        return ""
    year, month = parsed
    return f"{year}Q{(month - 1) // 3 + 1}"


def _period_end(label: str, fy_end_month: int) -> str:
    """財季標籤 → 期末年月（`2026-03`）。無法解析回空字串。"""
    parsed = _fiscal_period_end(label, fy_end_month)
    if parsed is None:
        return ""
    year, month = parsed
    return f"{year}-{month:02d}"


# ── 機器鍵 ──────────────────────────────────────────────────────────────────

def _slug(name: str) -> str:
    """顯示名 → 機器鍵尾段。非英數一律換底線，避免公式裡出現空白與符號。"""
    out = re.sub(r"[^0-9A-Za-z一-鿿]+", "_", name).strip("_")
    return out.upper() if out.isascii() else out


def _ratio_slug(name: str) -> str:
    """比率列名去掉單位後綴（`毛利率 (%)` → `毛利率`）。"""
    return _UNIT_SUFFIX.sub("", name).strip()


# ── 固定列定義 ──────────────────────────────────────────────────────────────
#
# (section, 顯示名, 機器鍵)。順序即 sheet 上的列序。
# 全部由既有模板衍生——新增 IS/BS/CF 模板行或比率時，這裡自動跟上，
# 但 FROZEN_ROW_NUMBERS 的測試會紅，提醒你列號變了、使用者的公式要跟著改。

def _build_std_rows() -> list[tuple[str, str, str]]:
    rows: list[tuple[str, str, str]] = []
    for section, template, prefix in (
        (SEC_IS, IS_TEMPLATE, "IS"),
        (SEC_BS, BS_TEMPLATE, "BS"),
        (SEC_CF, CF_TEMPLATE, "CF"),
    ):
        rows.append((section, section, ""))          # section 標題列
        for entry in template:
            display = entry[0]
            rows.append((section, display, f"{prefix}.{_slug(display)}"))

    rows.append((SEC_RATIO, SEC_RATIO, ""))
    for display, _formula, _fn in RATIO_DEFS:
        rows.append((SEC_RATIO, display, f"RATIO.{_ratio_slug(display)}"))
    return rows


STD_ROWS = _build_std_rows()

# ── 列號凍結表 ──────────────────────────────────────────────────────────────
#
# 目的是**擋住日後的漂移**，不是驗證初始版面正確（初始值本來就是照實作抄的）。
# 任何人插入或搬動一列，`test_row_numbers_are_frozen` 會立刻紅，提醒使用者的
# 跨檔公式需要同步更新。挑代表性的鍵即可，不必全列。
FROZEN_ROW_NUMBERS = {
    ROW_CALENDAR:            3,
    ROW_PERIOD_END:          4,
    ROW_SCHEMA:              5,
    "IS.REVENUE":            7,
    "IS.GROSS_PROFIT":       9,
    "IS.OPERATING_INCOME":  15,
    "IS.NET_INCOME":        22,
    "IS.DILUTED_EPS":       26,
    "BS.CASH":              30,
    "BS.TOTAL_ASSETS":      43,
    "BS.TOTAL_LIABILITIES": 60,
    "BS.SHARES_OUTSTANDING": 71,
    "CF.NET_INCOME":        73,
    "CF.OPERATING_CASH_FLOW": 82,
    "CF.CAPEX":             83,
    "CF.FREE_CASH_FLOW":    98,
    "RATIO.營收 YoY":       100,
    "RATIO.毛利率":         108,
    "RATIO.ROE":            133,
}


# ── 取值 ────────────────────────────────────────────────────────────────────

def _index_by_section(q_table: StatementTable) -> dict[str, dict[str, list[Any]]]:
    """把合併表拆成 {section: {concept: values}}。

    分區取值是必要的——`Net Income` 在 IS 與 CF 各出現一次，
    整表查名字只會拿到第一個。每個 section 內取首見的那筆。
    """
    sections: dict[str, dict[str, list[Any]]] = {}
    current = ""
    for concept, row in zip(q_table.concepts, q_table.values):
        name = str(concept or "").strip()
        if name in (SEC_IS, SEC_BS, SEC_CF):
            current = name
            sections.setdefault(current, {})
            continue
        if not name or not current:
            continue
        sections.setdefault(current, {}).setdefault(name, row)
    return sections


# ── Public API ──────────────────────────────────────────────────────────────

def build_std_table(
    q_table: StatementTable | None,
    ratio_table: StatementTable | None = None,
    fy_end_month: int = 12,
) -> StatementTable | None:
    """由 `Data_Financials(Q)` 與 `Data_Ratios` 組出版面固定的 `Data_Std`。

    來源為 None 時回 None。來源缺科目時該列全部留空——列一定在，值可以沒有，
    這樣跨公司的列號才會一致。
    """
    if q_table is None:
        return None

    labels_q = list(q_table.quarter_labels)
    n = len(labels_q)
    by_section = _index_by_section(q_table)

    ratio_values: dict[str, list[Any]] = {}
    if ratio_table is not None:
        ratio_index = {q: i for i, q in enumerate(ratio_table.quarter_labels)}
        for concept, row in zip(ratio_table.concepts, ratio_table.values):
            # 比率表的季度理論上與來源相同，仍依標籤對位以防萬一
            ratio_values[concept] = [
                row[ratio_index[q]] if q in ratio_index and ratio_index[q] < len(row) else None
                for q in labels_q
            ]

    concepts: list[str] = []
    keys: list[str] = []
    values: list[list[Any]] = []

    def add(display: str, key: str, row: list[Any]) -> None:
        concepts.append(display)
        keys.append(key)
        values.append(row)

    # 表頭三列
    add("日曆季", ROW_CALENDAR, [_calendar_quarter(q, fy_end_month) for q in labels_q])
    add("期末年月", ROW_PERIOD_END, [_period_end(q, fy_end_month) for q in labels_q])
    add("資料版本", ROW_SCHEMA, [SCHEMA_VERSION] * n)

    for section, display, key in STD_ROWS:
        if not key:                       # section 標題列
            add(display, "", [None] * n)
            continue
        if section == SEC_RATIO:
            row = ratio_values.get(display)
        else:
            row = by_section.get(section, {}).get(display)
        add(display, key, list(row) if row is not None else [None] * n)

    return StatementTable(
        sheet_name=SHEET_NAME,
        quarter_labels=labels_q,
        filing_dates=list(q_table.filing_dates),
        concepts=concepts,
        values=values,
        ticker=q_table.ticker,
        labels=keys,
    )
