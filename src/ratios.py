"""
ratios.py — 從 Data_Financials(Q) 算出常見財務比率，產生 Data_Ratios sheet。

設計決定（2026-08-02）：

1. **寫算好的「值」，不寫 Excel 公式。** Excel 公式的計算結果要等 Excel 開過並
   存檔才會寫進檔案，用 openpyxl 直接讀會拿到 None——下游 skill 就是這樣讀的。
   代價是你在 Excel 改了原始數字比率不會自動更新；要做 sensitivity 請在
   `My_*` sheet 自己拉公式，那本來就是保留給你的區域。

2. **B 欄寫算法文字**（如 `Gross Profit / Revenue`）。skill 讀 C 欄起的值，
   人看 B 欄知道這格怎麼來的。

3. **列名帶單位後綴** `(%)` / `(x)` / `(days)` / `($)`。這不是裝飾——
   `excel_formatter` 靠它決定數字格式，沒有後綴的話 `DSO` 會被當成金額除以
   1,000,000。後綴同時讓使用者一眼知道 1.2 是「1.2 倍」還是「1.2%」。

4. **算不出來就留 None，不猜。** 缺一個科目就整格空白，不用 0 代替、不用其他
   科目推估。金融股與小公司缺列是常態。

5. **成長率基期為負或零時回 None。** −10 成長到 +5 算出來的百分比沒有意義，
   放進去只會讓人誤讀。
"""

from __future__ import annotations

from typing import Any, Callable

from fetcher_gaap import StatementTable

# ── 基礎工具 ────────────────────────────────────────────────────────────────

Ctx = dict[str, list[Any]]

# 季度標籤放在 ctx 的保留鍵下，讓每個比率函式都能依標籤對齊期間
_LABELS_KEY = "__quarter_labels__"


def _labels(ctx: Ctx) -> list[str]:
    return ctx.get(_LABELS_KEY) or []


def _safe_div(numerator: Any, denominator: Any) -> float | None:
    """相除。任一邊是 None 或分母為 0 時回 None（不回 0，那會被誤讀成真的是 0）。"""
    if numerator is None or denominator is None:
        return None
    try:
        if float(denominator) == 0.0:
            return None
        return float(numerator) / float(denominator)
    except (TypeError, ValueError):
        return None


def _at(ctx: Ctx, concept: str, i: int) -> Any:
    """取某科目第 i 欄的值。科目不存在或索引越界回 None。"""
    row = ctx.get(concept)
    if row is None or i < 0 or i >= len(row):
        return None
    return row[i]


_QUARTER_RE = __import__("re").compile(r"FY(\d{4})Q([1-4])")


def _quarter_ordinal(label: str) -> int | None:
    """把 'FY2025Q1' 換成可相減的序號。非季度標籤回 None。"""
    m = _QUARTER_RE.fullmatch((label or "").strip())
    if m is None:
        return None
    return int(m.group(1)) * 4 + int(m.group(2)) - 1


def _lag_index(labels: list[str], i: int, lag: int) -> int | None:
    """找出「labels[i] 往前 lag 季」那一季的欄索引，不存在回 None。

    **不可用 i - lag 取代**。抓到的季度常有缺口（ARLO 實際抓到
    FY2024Q1/Q2/Q3 → FY2025Q1，中間缺 Q4），往前數格數會拿到 5 季前的數字，
    而且算出來的成長率看起來完全正常——這種錯最難發現。
    """
    if i < 0 or i >= len(labels):
        return None
    here = _quarter_ordinal(labels[i])
    if here is None:
        return None
    want = here - lag
    for j, label in enumerate(labels):
        if _quarter_ordinal(label) == want:
            return j
    return None


def _at_lag(ctx: Ctx, concept: str, i: int, lag: int) -> Any:
    """取某科目「往前 lag 季」的值，該季不存在回 None。"""
    j = _lag_index(_labels(ctx), i, lag)
    return None if j is None else _at(ctx, concept, j)


def _ttm(row: list[Any] | None, i: int, labels: list[str] | None = None) -> float | None:
    """近四季加總。不足四季、或四季中有任何缺值，回 None。

    缺值不可當 0 加總——那會低估 TTM，而且錯得很像對的。
    """
    if row is None or i < 0 or i >= len(row):
        return None

    if labels is None:
        # 沒給標籤時退回位置法（只有單元測試會這樣用）
        if i < 3:
            return None
        idxs = [i - 3, i - 2, i - 1, i]
    else:
        # 依標籤取近四季。缺任何一季就回 None——少加一季會低估 TTM
        idxs = []
        for lag in (3, 2, 1, 0):
            j = _lag_index(labels, i, lag)
            if j is None:
                return None
            idxs.append(j)

    window = [row[j] if 0 <= j < len(row) else None for j in idxs]
    if any(v is None for v in window):
        return None
    try:
        return float(sum(float(v) for v in window))
    except (TypeError, ValueError):
        return None


def _pct(value: float | None) -> float | None:
    """轉成百分比數字（0.382 → 38.2）。

    存百分比數字而非比例，是為了對齊 Data_NonGAAP 的慣例；
    excel_formatter 會再 ÷100 並套 0.0% 格式。
    """
    return None if value is None else value * 100.0


def _growth(ctx: Ctx, concept: str, i: int, lag: int) -> float | None:
    """成長率（%）。基期依**季度標籤**回推，不是往前數欄位。

    基期 <= 0 時回 None——負基期算成長率沒有意義（−10 成長到 +5 是幾 %？）。
    """
    row = ctx.get(concept)
    if row is None:
        return None
    j = _lag_index(_labels(ctx), i, lag)
    if j is None or j >= len(row) or i >= len(row):
        return None
    now, base = row[i], row[j]
    if now is None or base is None:
        return None
    try:
        if float(base) <= 0:
            return None
        return (float(now) / float(base) - 1.0) * 100.0
    except (TypeError, ValueError):
        return None


def _avg(a: Any, b: Any) -> float | None:
    if a is None or b is None:
        return None
    try:
        return (float(a) + float(b)) / 2.0
    except (TypeError, ValueError):
        return None


# ── 比率定義 ────────────────────────────────────────────────────────────────
#
# 每筆 = (顯示名稱含單位後綴, B 欄算法文字, 計算函式)
# 計算函式簽名 fn(ctx, i) -> float | None，i 是欄索引（0 = 最舊一季）。
#
# 要新增比率就加一行；順序即 sheet 上的列序。

RatioDef = tuple[str, str, Callable[[Ctx, int], Any]]

_DAYS_PER_QUARTER = 365.0 / 4.0

RATIO_DEFS: list[RatioDef] = [
    # ── 成長 ────────────────────────────────────────────────────────────
    ("Revenue YoY (%)", "Revenue[t] / Revenue[t-4] - 1",
     lambda c, i: _growth(c, "Revenue", i, 4)),
    ("Revenue QoQ (%)", "Revenue[t] / Revenue[t-1] - 1",
     lambda c, i: _growth(c, "Revenue", i, 1)),
    ("Gross Profit YoY (%)", "Gross Profit[t] / Gross Profit[t-4] - 1",
     lambda c, i: _growth(c, "Gross Profit", i, 4)),
    ("Operating Income YoY (%)", "Operating Income[t] / Operating Income[t-4] - 1",
     lambda c, i: _growth(c, "Operating Income", i, 4)),
    ("Net Income YoY (%)", "Net Income[t] / Net Income[t-4] - 1",
     lambda c, i: _growth(c, "Net Income", i, 4)),
    ("Net Income QoQ (%)", "Net Income[t] / Net Income[t-1] - 1",
     lambda c, i: _growth(c, "Net Income", i, 1)),
    ("EPS YoY (%)", "Diluted EPS[t] / Diluted EPS[t-4] - 1",
     lambda c, i: _growth(c, "Diluted EPS", i, 4)),
    ("Shares Outstanding YoY (%)", "Shares Outstanding[t] / Shares Outstanding[t-4] - 1",
     lambda c, i: _growth(c, "Shares Outstanding", i, 4)),

    # ── 利潤率 ──────────────────────────────────────────────────────────
    ("Gross Margin (%)", "Gross Profit / Revenue",
     lambda c, i: _pct(_safe_div(_at(c, "Gross Profit", i), _at(c, "Revenue", i)))),
    ("Opex Ratio (%)", "Total Operating Expense / Revenue",
     lambda c, i: _pct(_safe_div(_at(c, "Total Operating Expense", i), _at(c, "Revenue", i)))),
    ("R&D Ratio (%)", "R&D Expense / Revenue",
     lambda c, i: _pct(_safe_div(_at(c, "R&D Expense", i), _at(c, "Revenue", i)))),
    ("SG&A Ratio (%)", "SG&A Expense / Revenue",
     lambda c, i: _pct(_safe_div(_at(c, "SG&A Expense", i), _at(c, "Revenue", i)))),
    ("Operating Margin (%)", "Operating Income / Revenue",
     lambda c, i: _pct(_safe_div(_at(c, "Operating Income", i), _at(c, "Revenue", i)))),
    ("EBITDA Margin (%)", "(Operating Income + D&A) / Revenue",
     lambda c, i: _pct(_safe_div(
         None if _at(c, "Operating Income", i) is None or _at(c, "D&A", i) is None
         else float(_at(c, "Operating Income", i)) + float(_at(c, "D&A", i)),
         _at(c, "Revenue", i)))),
    ("Pre-tax Margin (%)", "Pre-tax Income / Revenue",
     lambda c, i: _pct(_safe_div(_at(c, "Pre-tax Income", i), _at(c, "Revenue", i)))),
    ("Net Margin (%)", "Net Income / Revenue",
     lambda c, i: _pct(_safe_div(_at(c, "Net Income", i), _at(c, "Revenue", i)))),

    # ── 結構 ────────────────────────────────────────────────────────────
    ("D&A / (COGS + Opex) (%)", "D&A / (Cost of Revenue + Total Operating Expense)",
     lambda c, i: _pct(_safe_div(
         _at(c, "D&A", i),
         None if _at(c, "Cost of Revenue", i) is None or _at(c, "Total Operating Expense", i) is None
         else float(_at(c, "Cost of Revenue", i)) + float(_at(c, "Total Operating Expense", i))))),
    ("Non-op / Pre-tax (%)", "Total Non-op Income/(Loss) / Pre-tax Income",
     lambda c, i: _pct(_safe_div(_at(c, "Total Non-op Income/(Loss)", i),
                                 _at(c, "Pre-tax Income", i)))),
    ("Effective Tax Rate (%)", "Income Tax / Pre-tax Income",
     lambda c, i: _pct(_safe_div(_at(c, "Income Tax", i), _at(c, "Pre-tax Income", i)))),
    ("SBC / Revenue (%)", "SBC / Revenue",
     lambda c, i: _pct(_safe_div(_at(c, "SBC", i), _at(c, "Revenue", i)))),
    ("SBC / Operating Cash Flow (%)", "SBC / Operating Cash Flow",
     lambda c, i: _pct(_safe_div(_at(c, "SBC", i), _at(c, "Operating Cash Flow", i)))),

    # ── 現金流 ──────────────────────────────────────────────────────────
    ("FCF Margin (%)", "Free Cash Flow / Revenue",
     lambda c, i: _pct(_safe_div(_at(c, "Free Cash Flow", i), _at(c, "Revenue", i)))),
    ("FCF / Net Income (x)", "Free Cash Flow / Net Income (cash conversion)",
     lambda c, i: _safe_div(_at(c, "Free Cash Flow", i), _at(c, "Net Income", i))),
    ("Capex / Revenue (%)", "abs(Capex) / Revenue",
     lambda c, i: _pct(_safe_div(
         None if _at(c, "Capex", i) is None else abs(float(_at(c, "Capex", i))),
         _at(c, "Revenue", i)))),
    ("Capex / D&A (x)", "abs(Capex) / D&A (reinvestment intensity)",
     lambda c, i: _safe_div(
         None if _at(c, "Capex", i) is None else abs(float(_at(c, "Capex", i))),
         _at(c, "D&A", i))),

    # ── 營運效率（單季數字年化後換算天數）─────────────────────────────────
    ("DSO (days)", "Accounts Receivable / (Revenue x 4) x 365",
     lambda c, i: _scale(_safe_div(_at(c, "Accounts Receivable", i),
                                   _mul(_at(c, "Revenue", i), 4)), 365.0)),
    ("DIO (days)", "Inventories / (Cost of Revenue x 4) x 365",
     lambda c, i: _scale(_safe_div(_at(c, "Inventories", i),
                                   _mul(_at(c, "Cost of Revenue", i), 4)), 365.0)),
    ("DPO (days)", "Accounts Payable / (Cost of Revenue x 4) x 365",
     lambda c, i: _scale(_safe_div(_at(c, "Accounts Payable", i),
                                   _mul(_at(c, "Cost of Revenue", i), 4)), 365.0)),
    ("Cash Conversion Cycle (days)", "DSO + DIO - DPO",
     lambda c, i: _ccc(c, i)),

    # ── 資產負債 ────────────────────────────────────────────────────────
    ("Current Ratio (x)", "Total Current Assets / Total Current Liabilities",
     lambda c, i: _safe_div(_at(c, "Total Current Assets", i),
                            _at(c, "Total Current Liabilities", i))),
    ("Quick Ratio (x)", "(Total Current Assets - Inventories) / Total Current Liabilities",
     lambda c, i: _safe_div(
         None if _at(c, "Total Current Assets", i) is None or _at(c, "Inventories", i) is None
         else float(_at(c, "Total Current Assets", i)) - float(_at(c, "Inventories", i)),
         _at(c, "Total Current Liabilities", i))),
    ("Net Debt / EBITDA (x)", "(Debt - Cash) / TTM EBITDA",
     lambda c, i: _net_debt_to_ebitda(c, i)),
    ("Interest Coverage (x)", "Operating Income / abs(Interest Expense)",
     lambda c, i: _safe_div(
         _at(c, "Operating Income", i),
         None if _at(c, "Interest Expense", i) is None
         else abs(float(_at(c, "Interest Expense", i))))),

    # ── 報酬率（TTM 淨利 ÷ 期初期末平均）──────────────────────────────────
    ("ROE (%)", "TTM Net Income / avg Total Equity — Parent",
     lambda c, i: _pct(_safe_div(_ttm(c.get("Net Income"), i, _labels(c)),
                                 _avg(_at_lag(c, "Total Equity — Parent", i, 3),
                                      _at(c, "Total Equity — Parent", i))))),
    ("ROA (%)", "TTM Net Income / avg Total Assets",
     lambda c, i: _pct(_safe_div(_ttm(c.get("Net Income"), i, _labels(c)),
                                 _avg(_at_lag(c, "Total Assets", i, 3),
                                      _at(c, "Total Assets", i))))),

    # ── 每股 ────────────────────────────────────────────────────────────
    ("BVPS ($)", "Total Equity — Parent / Shares Outstanding",
     lambda c, i: _safe_div(_at(c, "Total Equity — Parent", i),
                            _at(c, "Shares Outstanding", i))),
    ("FCF per Share ($)", "TTM Free Cash Flow / Shares Outstanding",
     lambda c, i: _safe_div(_ttm(c.get("Free Cash Flow"), i, _labels(c)),
                            _at(c, "Shares Outstanding", i))),
]


# ── 需要多步驟的計算 ────────────────────────────────────────────────────────

def _mul(value: Any, factor: float) -> float | None:
    if value is None:
        return None
    try:
        return float(value) * factor
    except (TypeError, ValueError):
        return None


def _scale(value: float | None, factor: float) -> float | None:
    return None if value is None else value * factor


def _ccc(ctx: Ctx, i: int) -> float | None:
    """現金循環天數 = DSO + DIO − DPO。任一段算不出來就整格 None。"""
    dso = _scale(_safe_div(_at(ctx, "Accounts Receivable", i),
                           _mul(_at(ctx, "Revenue", i), 4)), 365.0)
    dio = _scale(_safe_div(_at(ctx, "Inventories", i),
                           _mul(_at(ctx, "Cost of Revenue", i), 4)), 365.0)
    dpo = _scale(_safe_div(_at(ctx, "Accounts Payable", i),
                           _mul(_at(ctx, "Cost of Revenue", i), 4)), 365.0)
    if dso is None or dio is None or dpo is None:
        return None
    return dso + dio - dpo


def _net_debt_to_ebitda(ctx: Ctx, i: int) -> float | None:
    """淨負債 / TTM EBITDA。負債三段任一缺值就 None（少算一段會低估槓桿）。"""
    parts = [_at(ctx, "Short-term Debt", i),
             _at(ctx, "Current Portion of LT Debt", i),
             _at(ctx, "Long-term Debt", i)]
    if all(p is None for p in parts):
        return None
    debt = sum(float(p) for p in parts if p is not None)

    cash = _at(ctx, "Cash", i)
    if cash is None:
        return None

    op_row = ctx.get("Operating Income")
    da_row = ctx.get("D&A")
    ttm_op = _ttm(op_row, i, _labels(ctx))
    ttm_da = _ttm(da_row, i, _labels(ctx))
    if ttm_op is None or ttm_da is None:
        return None
    return _safe_div(debt - float(cash), ttm_op + ttm_da)


# ── Public API ──────────────────────────────────────────────────────────────

def build_ratio_table(q_table: StatementTable | None) -> StatementTable | None:
    """從 Data_Financials(Q) 算出 Data_Ratios。來源為 None 時回 None。

    固定模板：`RATIO_DEFS` 有幾筆就有幾列，算不出來的整列 None，
    不會因為某家公司缺科目就少幾列——跨公司列序一致，skill 才能用固定位置讀。
    """
    if q_table is None:
        return None

    ctx: Ctx = {c: v for c, v in zip(q_table.concepts, q_table.values)}
    ctx[_LABELS_KEY] = list(q_table.quarter_labels)
    n = len(q_table.quarter_labels)

    concepts: list[str] = []
    labels: list[str] = []
    values: list[list[Any]] = []

    for name, formula, fn in RATIO_DEFS:
        row: list[Any] = []
        for i in range(n):
            try:
                row.append(fn(ctx, i))
            except Exception:
                # 單一格算爆不該讓整張表消失。缺科目、型別怪異都落在這裡。
                row.append(None)
        concepts.append(name)
        labels.append(formula)
        values.append(row)

    return StatementTable(
        sheet_name="Data_Ratios",
        quarter_labels=list(q_table.quarter_labels),
        filing_dates=list(q_table.filing_dates),
        concepts=concepts,
        values=values,
        ticker=q_table.ticker,
        labels=labels,
    )
