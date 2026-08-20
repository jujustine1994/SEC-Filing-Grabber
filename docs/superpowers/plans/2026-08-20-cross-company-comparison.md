# 跨公司財務比較功能 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 讓使用者在 GUI 新增的 Tab4「跨公司比較」選多家公司、多個財務指標，輸出一份獨立的新格式 Excel（原始資料表 + 可互動快照 + 每指標一張歷史趨勢圖）。

**Architecture:** 新增 `src/comparison.py`（呼叫既有 `fetch_gaap_statements()` 取多家公司資料，用 `ratios.py::build_ratio_table()` 算比率，重組成 `{指標:{公司:{期間:值}}}`）+ `src/comparison_writer.py`（用 `openpyxl.chart` 寫出新格式 Excel）+ `main.py` 新增 Tab4 GUI。`ratios.py` 的 `RATIO_DEFS` 加 category 欄位並擴充約 20 個新比率，`excel_formatter.py` 新增可共用的格式判斷函式與 `($mm)` 單位後綴。

**Tech Stack:** Python 3、tkinter/ttk（既有 GUI 框架）、openpyxl（含這次新用的 `openpyxl.chart`）、pytest。

**Spec:** `docs/superpowers/specs/2026-08-20-cross-company-comparison-design.md`

## Global Constraints

- **不改動現有單一公司抓取流程**：`fetch_gaap_statements()`、Tab1/Tab2、既有 Excel 格式一律不動，新功能純外掛呼叫既有函式。
- **不做估值倍數**（P/E、EV/EBITDA、P/B）——需要股價資料，本次範圍外（TODO F2）。
- **不做獨立的 Q4 推算邏輯**——完全依賴 `fetcher_gaap.py::_synthesize_q4()`（D0-1 已驗證），本功能只呼叫 `fetch_gaap_statements()` 取現成結果。
- **`RATIO_DEFS` 新增的每一列名稱都要帶單位後綴**（`(%)`／`(x)`／`(days)`／`($)`／`($mm)`），這是 `excel_formatter.py` 判斷數字格式的依據，漏了後綴會格式跑掉。
- **算不出來就回 `None`，不猜、不用 0 代替**——沿用 `ratios.py` 檔頭原則。
- **成長率基期 ≤ 0 時回 `None`**——沿用 `_growth()` 既有原則，新寫的成長類公式要一致。
- **四個 locale 都要加**（`zh_tw`／`zh_cn`／`en`／`ja`），不可只改繁中。
- **`git commit` 前一律先跑相關測試**，不跳過 hook。

---

## 檔案結構總覽

| 檔案 | 動作 | 職責 |
|---|---|---|
| `src/excel_formatter.py` | 修改 | 新增 `($mm)` 單位後綴、抽出可共用的 `unit_format_for()` |
| `src/ratios.py` | 修改 | `RATIO_DEFS` 加 category 欄位、新增約 20 個比率與對應輔助函式 |
| `src/comparison.py` | 新增 | 多公司資料抓取與重組（`{指標:{公司:{期間:值}}}`） |
| `src/comparison_writer.py` | 新增 | 寫出 `Compare_Data`／`Snapshot`／`Snapshot_Manual`／`Chart_*` |
| `src/main.py` | 修改 | 新增 Tab4「跨公司比較」GUI（選擇視窗 + 主畫面 + 背景執行緒） |
| `src/locales/zh_tw.py`／`zh_cn.py`／`en.py`／`ja.py` | 修改 | Tab4 新增字串 |
| `tests/test_excel_formatter.py` | 修改 | `unit_format_for()` 與 `($mm)` 後綴測試 |
| `tests/test_ratios.py` | 修改 | 新比率的正確性測試 |
| `tests/test_comparison.py` | 新增 | `comparison.py` 單元測試 |
| `tests/test_comparison_writer.py` | 新增 | `comparison_writer.py` 單元測試 |

---

## Task 1: `excel_formatter.py` — 新增 `($mm)` 後綴、抽出可共用的格式判斷函式

**背景**：現有 `($)` 後綴（`FMT_EPS`，不除以 1,000,000）是給「每股」這種本來就是小數字的欄位用的（如 `BVPS ($)`）。這次要新增的 `EBITDA($mm)`／`Total Debt($mm)`／`Net Debt($mm)`／`Working Capital($mm)` 是原始金額量級（十億美元起跳），如果沿用 `($)` 後綴會被當成每股數字、少除 1,000,000，數字會大到看不懂。需要新的 `($mm)` 後綴對應 `FMT_FINANCIAL` + 除以 1,000,000。

同時，`comparison_writer.py`（Task 5-7）需要幫 `Compare_Data`／`Snapshot`／`Snapshot_Manual` 的每一格決定數字格式，邏輯要跟 `format_workbook()` 現有的判斷順序完全一致（後綴 → EPS → 百分比 → 股數 → 預設金額）。這裡先把該判斷邏輯抽成一個公開函式，兩邊共用，不要複製一份判斷邏輯出來維護兩次。

**Files:**
- Modify: `src/excel_formatter.py:183-188`（`_UNIT_SUFFIX_FORMATS`）、`src/excel_formatter.py:240-263`（`format_workbook()` 內的格式判斷迴圈）
- Test: `tests/test_excel_formatter.py`

**Interfaces:**
- Produces: `unit_format_for(concept: str) -> tuple[str, int]`（公開函式，回傳 `(Excel 數字格式字串, 除數)`），供 `comparison_writer.py` 使用

- [ ] **Step 1: 寫失敗測試**

在 `tests/test_excel_formatter.py` 尾端加入：

```python
from excel_formatter import unit_format_for, FMT_FINANCIAL, FMT_EPS, FMT_PERCENT, FMT_MULTIPLE, FMT_DAYS, FMT_SHARES


def test_unit_format_for_mm_suffix_uses_financial_format_and_million_divisor():
    fmt, divisor = unit_format_for("EBITDA ($mm)")
    assert fmt == FMT_FINANCIAL
    assert divisor == 1_000_000


def test_unit_format_for_dollar_suffix_uses_eps_format_no_divisor():
    fmt, divisor = unit_format_for("BVPS ($)")
    assert fmt == FMT_EPS
    assert divisor == 1


def test_unit_format_for_percent_suffix():
    fmt, divisor = unit_format_for("Revenue YoY (%)")
    assert fmt == FMT_PERCENT


def test_unit_format_for_multiple_suffix():
    fmt, _ = unit_format_for("Current Ratio (x)")
    assert fmt == FMT_MULTIPLE


def test_unit_format_for_days_suffix():
    fmt, _ = unit_format_for("DSO (days)")
    assert fmt == FMT_DAYS


def test_unit_format_for_no_suffix_defaults_to_financial_million():
    fmt, divisor = unit_format_for("Revenue")
    assert fmt == FMT_FINANCIAL
    assert divisor == 1_000_000


def test_unit_format_for_shares_concept():
    fmt, divisor = unit_format_for("Shares Outstanding")
    assert fmt == FMT_SHARES
    assert divisor == 1_000_000
```

- [ ] **Step 2: 執行測試確認失敗**

Run: `./venv/Scripts/python.exe -m pytest tests/test_excel_formatter.py -k unit_format_for -v`
Expected: FAIL，`ImportError: cannot import name 'unit_format_for'`

- [ ] **Step 3: 實作**

在 `src/excel_formatter.py` 找到 `_UNIT_SUFFIX_FORMATS`（約第 183 行），新增一行：

```python
_UNIT_SUFFIX_FORMATS = {
    "(%)":     (FMT_PERCENT,  100 if PERCENT_AS_EXCEL_RATIO else 1),
    "(x)":     (FMT_MULTIPLE, 1),
    "(days)":  (FMT_DAYS,     1),
    "($)":     (FMT_EPS,      1),
    "($mm)":   (FMT_FINANCIAL, 1_000_000),
}
```

緊接在 `_unit_suffix_rule()` 之後（原第 196 行之後）新增公開函式：

```python
def unit_format_for(concept: str) -> tuple[str, int]:
    """回傳 (Excel 數字格式, 除數)。判斷順序：單位後綴 → 每股 → 百分比 →
    股數 → 預設金額（÷1,000,000）。跟 format_workbook() 內的判斷順序完全一致，
    給 Data_Ratios 以外的表（如跨公司比較）共用同一套規則，不要各寫一份。
    """
    rule = _unit_suffix_rule(concept)
    if rule is not None:
        return rule
    if _is_eps_concept(concept):
        return FMT_EPS, 1
    if _is_percent_concept(concept):
        return FMT_PERCENT, 100 if PERCENT_AS_EXCEL_RATIO else 1
    if _is_shares_concept(concept):
        return FMT_SHARES, 1_000_000
    return FMT_FINANCIAL, 1_000_000
```

然後把 `format_workbook()` 裡原本第 247-263 行那段判斷邏輯改成呼叫這個新函式，避免同一套邏輯維護兩份：

```python
        # 單位後綴優先於一切關鍵字判斷（Data_Ratios 用）；判斷順序見 unit_format_for()
        fmt, divisor = unit_format_for(concept)
```

刪掉原本 `if suffix_rule is not None: ... elif _is_eps_concept ... else: fmt = FMT_FINANCIAL; divisor = 1_000_000` 那整段 if/elif 鏈（原 248-263 行），改成上面這一行呼叫。

- [ ] **Step 4: 執行測試確認通過**

Run: `./venv/Scripts/python.exe -m pytest tests/test_excel_formatter.py -v`
Expected: 全部 PASS（新測試 + 既有測試都要過，確認重構沒改變既有行為）

- [ ] **Step 5: Commit**

```bash
git add src/excel_formatter.py tests/test_excel_formatter.py
git commit -m "feat: 新增 (\$mm) 單位後綴，抽出可共用的 unit_format_for()"
```

---

## Task 2: `ratios.py` — `RATIO_DEFS` 加 category 欄位

**背景**：現有 `RATIO_DEFS` 是 `(名稱, 算法文字, 函式)` 三元組，只用註解分區塊。跨公司比較的選擇視窗要照分類（成長性／獲利能力／槓桿償債...）分組顯示指標，需要程式碼層面的分類欄位，不能只靠註解。

**Files:**
- Modify: `src/ratios.py:170-298`（`RatioDef` 型別定義、`RATIO_DEFS` 全部項目、`build_ratio_table()`）
- Test: `tests/test_ratios.py`

**Interfaces:**
- Produces: `RatioDef = tuple[str, str, str, Callable[[Ctx, int], Any]]`（名稱, 算法文字, category, 函式），`RATIO_CATEGORIES: list[str]`（分類清單，依 sheet 上出現順序）

- [ ] **Step 1: 寫失敗測試**

在 `tests/test_ratios.py` 尾端加入：

```python
def test_every_ratio_def_has_a_category():
    for name, formula, category, fn in RATIO_DEFS:
        assert category, f"{name} 缺 category"


def test_ratio_categories_lists_every_distinct_category_used():
    used = {category for _, _, category, _ in RATIO_DEFS}
    assert used == set(RATIO_CATEGORIES)
```

在檔案頂端 import 那行加入 `RATIO_CATEGORIES`：

```python
from ratios import build_ratio_table, RATIO_DEFS, RATIO_CATEGORIES, _safe_div, _ttm, _quarter_ordinal, _lag_index
```

- [ ] **Step 2: 執行測試確認失敗**

Run: `./venv/Scripts/python.exe -m pytest tests/test_ratios.py -k "category" -v`
Expected: FAIL，`ImportError: cannot import name 'RATIO_CATEGORIES'`

- [ ] **Step 3: 實作**

`src/ratios.py` 第 177 行的 `RatioDef` 型別改成四元組：

```python
RatioDef = tuple[str, str, str, Callable[[Ctx, int], Any]]

RATIO_CATEGORIES: list[str] = [
    "成長性", "獲利能力", "結構", "現金流", "營運效率", "槓桿償債", "報酬率", "每股",
]
```

`RATIO_DEFS` 現有 28 個項目全部從三元組改四元組，在算法文字後面插入分類字串。對照既有的區塊註解：

```python
RATIO_DEFS: list[RatioDef] = [
    # ── 成長 ────────────────────────────────────────────────────────────
    ("Revenue YoY (%)", "Revenue[t] / Revenue[t-4] - 1", "成長性",
     lambda c, i: _growth(c, "Revenue", i, 4)),
    ("Revenue QoQ (%)", "Revenue[t] / Revenue[t-1] - 1", "成長性",
     lambda c, i: _growth(c, "Revenue", i, 1)),
    ("Gross Profit YoY (%)", "Gross Profit[t] / Gross Profit[t-4] - 1", "成長性",
     lambda c, i: _growth(c, "Gross Profit", i, 4)),
    ("Operating Income YoY (%)", "Operating Income[t] / Operating Income[t-4] - 1", "成長性",
     lambda c, i: _growth(c, "Operating Income", i, 4)),
    ("Net Income YoY (%)", "Net Income[t] / Net Income[t-4] - 1", "成長性",
     lambda c, i: _growth(c, "Net Income", i, 4)),
    ("Net Income QoQ (%)", "Net Income[t] / Net Income[t-1] - 1", "成長性",
     lambda c, i: _growth(c, "Net Income", i, 1)),
    ("EPS YoY (%)", "Diluted EPS[t] / Diluted EPS[t-4] - 1", "成長性",
     lambda c, i: _growth(c, "Diluted EPS", i, 4)),
    ("Shares Outstanding YoY (%)", "Shares Outstanding[t] / Shares Outstanding[t-4] - 1", "成長性",
     lambda c, i: _growth(c, "Shares Outstanding", i, 4)),

    # ── 利潤率 ──────────────────────────────────────────────────────────
    ("Gross Margin (%)", "Gross Profit / Revenue", "獲利能力",
     lambda c, i: _pct(_safe_div(_at(c, "Gross Profit", i), _at(c, "Revenue", i)))),
    ("Opex Ratio (%)", "Total Operating Expense / Revenue", "獲利能力",
     lambda c, i: _pct(_safe_div(_at(c, "Total Operating Expense", i), _at(c, "Revenue", i)))),
    ("R&D Ratio (%)", "R&D Expense / Revenue", "獲利能力",
     lambda c, i: _pct(_safe_div(_at(c, "R&D Expense", i), _at(c, "Revenue", i)))),
    ("SG&A Ratio (%)", "SG&A Expense / Revenue", "獲利能力",
     lambda c, i: _pct(_safe_div(_at(c, "SG&A Expense", i), _at(c, "Revenue", i)))),
    ("Operating Margin (%)", "Operating Income / Revenue", "獲利能力",
     lambda c, i: _pct(_safe_div(_at(c, "Operating Income", i), _at(c, "Revenue", i)))),
    ("EBITDA Margin (%)", "(Operating Income + D&A) / Revenue", "獲利能力",
     lambda c, i: _pct(_safe_div(
         None if _at(c, "Operating Income", i) is None or _at(c, "D&A", i) is None
         else float(_at(c, "Operating Income", i)) + float(_at(c, "D&A", i)),
         _at(c, "Revenue", i)))),
    ("Pre-tax Margin (%)", "Pre-tax Income / Revenue", "獲利能力",
     lambda c, i: _pct(_safe_div(_at(c, "Pre-tax Income", i), _at(c, "Revenue", i)))),
    ("Net Margin (%)", "Net Income / Revenue", "獲利能力",
     lambda c, i: _pct(_safe_div(_at(c, "Net Income", i), _at(c, "Revenue", i)))),

    # ── 結構 ────────────────────────────────────────────────────────────
    ("D&A / (COGS + Opex) (%)", "D&A / (Cost of Revenue + Total Operating Expense)", "結構",
     lambda c, i: _pct(_safe_div(
         _at(c, "D&A", i),
         None if _at(c, "Cost of Revenue", i) is None or _at(c, "Total Operating Expense", i) is None
         else float(_at(c, "Cost of Revenue", i)) + float(_at(c, "Total Operating Expense", i))))),
    ("Non-op / Pre-tax (%)", "Total Non-op Income/(Loss) / Pre-tax Income", "結構",
     lambda c, i: _pct(_safe_div(_at(c, "Total Non-op Income/(Loss)", i),
                                 _at(c, "Pre-tax Income", i)))),
    ("Effective Tax Rate (%)", "Income Tax / Pre-tax Income", "結構",
     lambda c, i: _pct(_safe_div(_at(c, "Income Tax", i), _at(c, "Pre-tax Income", i)))),
    ("SBC / Revenue (%)", "SBC / Revenue", "結構",
     lambda c, i: _pct(_safe_div(_at(c, "SBC", i), _at(c, "Revenue", i)))),
    ("SBC / Operating Cash Flow (%)", "SBC / Operating Cash Flow", "結構",
     lambda c, i: _pct(_safe_div(_at(c, "SBC", i), _at(c, "Operating Cash Flow", i)))),

    # ── 現金流 ──────────────────────────────────────────────────────────
    ("FCF Margin (%)", "Free Cash Flow / Revenue", "現金流",
     lambda c, i: _pct(_safe_div(_at(c, "Free Cash Flow", i), _at(c, "Revenue", i)))),
    ("FCF / Net Income (x)", "Free Cash Flow / Net Income (cash conversion)", "現金流",
     lambda c, i: _safe_div(_at(c, "Free Cash Flow", i), _at(c, "Net Income", i))),
    ("Capex / Revenue (%)", "abs(Capex) / Revenue", "現金流",
     lambda c, i: _pct(_safe_div(
         None if _at(c, "Capex", i) is None else abs(float(_at(c, "Capex", i))),
         _at(c, "Revenue", i)))),
    ("Capex / D&A (x)", "abs(Capex) / D&A (reinvestment intensity)", "現金流",
     lambda c, i: _safe_div(
         None if _at(c, "Capex", i) is None else abs(float(_at(c, "Capex", i))),
         _at(c, "D&A", i))),

    # ── 營運效率（單季數字年化後換算天數）─────────────────────────────────
    ("DSO (days)", "Accounts Receivable / (Revenue x 4) x 365", "營運效率",
     lambda c, i: _scale(_safe_div(_at(c, "Accounts Receivable", i),
                                   _mul(_at(c, "Revenue", i), 4)), 365.0)),
    ("DIO (days)", "Inventories / (Cost of Revenue x 4) x 365", "營運效率",
     lambda c, i: _scale(_safe_div(_at(c, "Inventories", i),
                                   _mul(_at(c, "Cost of Revenue", i), 4)), 365.0)),
    ("DPO (days)", "Accounts Payable / (Cost of Revenue x 4) x 365", "營運效率",
     lambda c, i: _scale(_safe_div(_at(c, "Accounts Payable", i),
                                   _mul(_at(c, "Cost of Revenue", i), 4)), 365.0)),
    ("Cash Conversion Cycle (days)", "DSO + DIO - DPO", "營運效率",
     lambda c, i: _ccc(c, i)),

    # ── 資產負債 ────────────────────────────────────────────────────────
    ("Current Ratio (x)", "Total Current Assets / Total Current Liabilities", "槓桿償債",
     lambda c, i: _safe_div(_at(c, "Total Current Assets", i),
                            _at(c, "Total Current Liabilities", i))),
    ("Quick Ratio (x)", "(Total Current Assets - Inventories) / Total Current Liabilities", "槓桿償債",
     lambda c, i: _safe_div(
         None if _at(c, "Total Current Assets", i) is None or _at(c, "Inventories", i) is None
         else float(_at(c, "Total Current Assets", i)) - float(_at(c, "Inventories", i)),
         _at(c, "Total Current Liabilities", i))),
    ("Net Debt / EBITDA (x)", "(Debt - Cash) / TTM EBITDA", "槓桿償債",
     lambda c, i: _net_debt_to_ebitda(c, i)),
    ("Interest Coverage (x)", "Operating Income / abs(Interest Expense)", "槓桿償債",
     lambda c, i: _safe_div(
         _at(c, "Operating Income", i),
         None if _at(c, "Interest Expense", i) is None
         else abs(float(_at(c, "Interest Expense", i))))),

    # ── 報酬率（TTM 淨利 ÷ 期初期末平均）──────────────────────────────────
    ("ROE (%)", "TTM Net Income / avg Total Equity — Parent", "報酬率",
     lambda c, i: _pct(_safe_div(_ttm(c.get("Net Income"), i, _labels(c)),
                                 _avg(_at_lag(c, "Total Equity — Parent", i, 3),
                                      _at(c, "Total Equity — Parent", i))))),
    ("ROA (%)", "TTM Net Income / avg Total Assets", "報酬率",
     lambda c, i: _pct(_safe_div(_ttm(c.get("Net Income"), i, _labels(c)),
                                 _avg(_at_lag(c, "Total Assets", i, 3),
                                      _at(c, "Total Assets", i))))),

    # ── 每股 ────────────────────────────────────────────────────────────
    ("BVPS ($)", "Total Equity — Parent / Shares Outstanding", "每股",
     lambda c, i: _safe_div(_at(c, "Total Equity — Parent", i),
                            _at(c, "Shares Outstanding", i))),
    ("FCF per Share ($)", "TTM Free Cash Flow / Shares Outstanding", "每股",
     lambda c, i: _safe_div(_ttm(c.get("Free Cash Flow"), i, _labels(c)),
                            _at(c, "Shares Outstanding", i))),
]
```

`build_ratio_table()`（原第 370 行）的迴圈解構要從三元組改四元組：

```python
    for name, formula, category, fn in RATIO_DEFS:
```

（`category` 在這個函式裡不需要用到，只是解構時要跟著改，不然會 `ValueError: not enough values to unpack`）

- [ ] **Step 4: 執行測試確認通過**

Run: `./venv/Scripts/python.exe -m pytest tests/test_ratios.py -v`
Expected: 全部 PASS（既有 28 個比率的所有測試都要維持通過，只是多了 category 欄位）

- [ ] **Step 5: Commit**

```bash
git add src/ratios.py tests/test_ratios.py
git commit -m "refactor: RATIO_DEFS 加 category 欄位，供跨公司比較的指標分類用"
```

---

## Task 3: `ratios.py` — 新增約 20 個比率

**Files:**
- Modify: `src/ratios.py`（`RATIO_DEFS` 尾端新增項目、新增輔助函式）
- Test: `tests/test_ratios.py`

**Interfaces:**
- Consumes: Task 2 的 `RatioDef` 四元組格式、`RATIO_CATEGORIES`
- Produces: `_ebitda(ctx, i)`、`_total_debt(ctx, i)`、`_net_debt(ctx, i)`、`_ebitda_growth(ctx, i, lag)`、`_nopat(ctx, i)`、`_invested_capital(ctx, i)`、`_roic(ctx, i)` 輔助函式，供本檔內 `RATIO_DEFS` 使用

- [ ] **Step 1: 寫失敗測試**

在 `tests/test_ratios.py` 尾端加入（挑幾個公式邏輯較複雜的做正確性測試，其餘由既有的
`test_every_ratio_row_present_even_when_uncomputable`／`test_every_row_name_carries_a_unit_suffix`
等結構性測試自動涵蓋）：

```python
def test_debt_ratio():
    tbl = _q_table(**{
        "Total Liabilities": [600.0],
        "Total Assets": [1000.0],
    })
    rt = build_ratio_table(tbl)
    assert _find(rt, "Debt Ratio")[0] == pytest.approx(60.0)


def test_debt_to_equity():
    tbl = _q_table(**{
        "Total Liabilities": [600.0],
        "Total Equity — Parent": [400.0],
    })
    rt = build_ratio_table(tbl)
    assert _find(rt, "Debt-to-Equity")[0] == pytest.approx(1.5)


def test_da_over_revenue():
    tbl = _q_table(**{"D&A": [50.0], "Revenue": [1000.0]})
    rt = build_ratio_table(tbl)
    assert _find(rt, "D&A / Revenue")[0] == pytest.approx(5.0)


def test_ebitda_dollar_amount():
    tbl = _q_table(**{"Operating Income": [100.0], "D&A": [20.0]})
    rt = build_ratio_table(tbl)
    assert _find(rt, "EBITDA ($mm)")[0] == pytest.approx(120.0)


def test_ebitda_missing_da_returns_none():
    tbl = _q_table(**{"Operating Income": [100.0]})
    rt = build_ratio_table(tbl)
    assert _find(rt, "EBITDA ($mm)")[0] is None


def test_total_debt_sums_available_parts_only():
    tbl = _q_table(**{
        "Short-term Debt": [10.0],
        "Long-term Debt": [90.0],
        # Current Portion of LT Debt 缺，應視為 0 不影響其他兩段
    })
    rt = build_ratio_table(tbl)
    assert _find(rt, "Total Debt")[0] == pytest.approx(100.0)


def test_net_debt():
    tbl = _q_table(**{
        "Short-term Debt": [10.0],
        "Long-term Debt": [90.0],
        "Cash": [30.0],
    })
    rt = build_ratio_table(tbl)
    assert _find(rt, "Net Debt")[0] == pytest.approx(70.0)


def test_working_capital():
    tbl = _q_table(**{
        "Total Current Assets": [500.0],
        "Total Current Liabilities": [300.0],
    })
    rt = build_ratio_table(tbl)
    assert _find(rt, "Working Capital")[0] == pytest.approx(200.0)


def test_equity_multiplier():
    tbl = _q_table(**{"Total Assets": [1000.0], "Total Equity — Parent": [250.0]})
    rt = build_ratio_table(tbl)
    assert _find(rt, "Equity Multiplier")[0] == pytest.approx(4.0)


def test_cash_ratio():
    tbl = _q_table(**{"Cash": [50.0], "Total Current Liabilities": [200.0]})
    rt = build_ratio_table(tbl)
    assert _find(rt, "Cash Ratio")[0] == pytest.approx(0.25)


def test_cogs_ratio():
    tbl = _q_table(**{"Cost of Revenue": [600.0], "Revenue": [1000.0]})
    rt = build_ratio_table(tbl)
    assert _find(rt, "COGS Ratio")[0] == pytest.approx(60.0)


def test_operating_cf_margin():
    tbl = _q_table(**{"Operating Cash Flow": [200.0], "Revenue": [1000.0]})
    rt = build_ratio_table(tbl)
    assert _find(rt, "Operating CF Margin")[0] == pytest.approx(20.0)


def test_roic_approx():
    tbl = _q_table(**{
        "Operating Income": [200.0],
        "Income Tax": [40.0],
        "Pre-tax Income": [200.0],
        "Short-term Debt": [100.0],
        "Long-term Debt": [400.0],
        "Total Equity — Parent": [500.0],
        "Cash": [100.0],
    })
    rt = build_ratio_table(tbl)
    # NOPAT = 200 * (1 - 40/200) = 160；Invested Capital = 100+400+500-100 = 900
    assert _find(rt, "ROIC")[0] == pytest.approx(160.0 / 900.0 * 100.0)


def test_roic_none_when_pretax_zero():
    tbl = _q_table(**{
        "Operating Income": [200.0], "Income Tax": [0.0], "Pre-tax Income": [0.0],
        "Long-term Debt": [400.0], "Total Equity — Parent": [500.0], "Cash": [100.0],
    })
    rt = build_ratio_table(tbl)
    assert _find(rt, "ROIC")[0] is None


def test_ebitda_yoy_growth():
    labels = _consecutive_labels(5)
    tbl = StatementTable(
        sheet_name="Data_Financials(Q)", quarter_labels=labels, filing_dates=[""] * 5,
        concepts=["Operating Income", "D&A"],
        values=[[100.0, 100.0, 100.0, 100.0, 150.0], [10.0, 10.0, 10.0, 10.0, 20.0]],
        ticker="TEST", labels=["", ""],
    )
    rt = build_ratio_table(tbl)
    # base(第0欄) EBITDA=110，第4欄 EBITDA=170 → (170/110-1)*100
    assert _find(rt, "EBITDA YoY")[4] == pytest.approx((170.0 / 110.0 - 1.0) * 100.0)
```

- [ ] **Step 2: 執行測試確認失敗**

Run: `./venv/Scripts/python.exe -m pytest tests/test_ratios.py -k "debt_ratio or ebitda or total_debt or net_debt or working_capital or equity_multiplier or cash_ratio or cogs_ratio or operating_cf_margin or roic or da_over_revenue" -v`
Expected: FAIL（`AssertionError: 找不到以 ... 開頭的列` — 這些指標還不存在）

- [ ] **Step 3: 實作**

在 `src/ratios.py` 的 `_net_debt_to_ebitda()`（原第 329-348 行）**之後**新增輔助函式：

```python
def _total_debt(ctx: Ctx, i: int) -> float | None:
    """短期借款＋一年內到期長期負債＋長期借款。三段全缺才回 None，缺一段當 0 補。"""
    parts = [_at(ctx, "Short-term Debt", i),
             _at(ctx, "Current Portion of LT Debt", i),
             _at(ctx, "Long-term Debt", i)]
    if all(p is None for p in parts):
        return None
    return sum(float(p) for p in parts if p is not None)


def _net_debt(ctx: Ctx, i: int) -> float | None:
    debt = _total_debt(ctx, i)
    cash = _at(ctx, "Cash", i)
    if debt is None or cash is None:
        return None
    return debt - float(cash)


def _ebitda(ctx: Ctx, i: int) -> float | None:
    """單季 EBITDA＝Operating Income + D&A。跟 _net_debt_to_ebitda() 用的 TTM 版本不同，
    這是給規模型指標（EBITDA($mm)）與 EBITDA YoY 用的單季值。"""
    op = _at(ctx, "Operating Income", i)
    da = _at(ctx, "D&A", i)
    if op is None or da is None:
        return None
    return float(op) + float(da)


def _ebitda_growth(ctx: Ctx, i: int, lag: int) -> float | None:
    """比照 _growth() 的邏輯（依季度標籤回推基期、基期 <= 0 回 None），
    只是分子分母換成算出來的 EBITDA，不是 ctx 裡的原始科目。"""
    j = _lag_index(_labels(ctx), i, lag)
    if j is None:
        return None
    now = _ebitda(ctx, i)
    base = _ebitda(ctx, j)
    if now is None or base is None or base <= 0:
        return None
    return (now / base - 1.0) * 100.0


def _nopat(ctx: Ctx, i: int) -> float | None:
    """稅後淨營業利潤（近似值）＝Operating Income × (1 − 有效稅率)。
    有效稅率缺任一項或 Pre-tax Income 為 0 時回 None。"""
    op = _at(ctx, "Operating Income", i)
    tax = _at(ctx, "Income Tax", i)
    pretax = _at(ctx, "Pre-tax Income", i)
    if op is None or tax is None or pretax is None:
        return None
    try:
        if float(pretax) == 0:
            return None
        tax_rate = float(tax) / float(pretax)
        return float(op) * (1.0 - tax_rate)
    except (TypeError, ValueError):
        return None


def _invested_capital(ctx: Ctx, i: int) -> float | None:
    """投入資本（近似值）＝Total Debt + Total Equity — Parent − Cash。"""
    debt = _total_debt(ctx, i)
    equity = _at(ctx, "Total Equity — Parent", i)
    cash = _at(ctx, "Cash", i)
    if debt is None or equity is None or cash is None:
        return None
    return debt + float(equity) - float(cash)


def _roic(ctx: Ctx, i: int) -> float | None:
    """ROIC 近似值（業界慣用簡化版，未拆一次性項目）＝NOPAT / Invested Capital。"""
    return _pct(_safe_div(_nopat(ctx, i), _invested_capital(ctx, i)))
```

然後在 `RATIO_DEFS` 列表結尾（原本 `FCF per Share ($)` 那一行之後、`]` 之前）新增：

```python
    # ── 成長（QoQ／額外 YoY）────────────────────────────────────────────
    ("Gross Profit QoQ (%)", "Gross Profit[t] / Gross Profit[t-1] - 1", "成長性",
     lambda c, i: _growth(c, "Gross Profit", i, 1)),
    ("Operating Income QoQ (%)", "Operating Income[t] / Operating Income[t-1] - 1", "成長性",
     lambda c, i: _growth(c, "Operating Income", i, 1)),
    ("EPS QoQ (%)", "Diluted EPS[t] / Diluted EPS[t-1] - 1", "成長性",
     lambda c, i: _growth(c, "Diluted EPS", i, 1)),
    ("FCF YoY (%)", "Free Cash Flow[t] / Free Cash Flow[t-4] - 1", "成長性",
     lambda c, i: _growth(c, "Free Cash Flow", i, 4)),
    ("EBITDA YoY (%)", "EBITDA[t] / EBITDA[t-4] - 1", "成長性",
     lambda c, i: _ebitda_growth(c, i, 4)),

    # ── 槓桿償債 ────────────────────────────────────────────────────────
    ("Debt Ratio (%)", "Total Liabilities / Total Assets", "槓桿償債",
     lambda c, i: _pct(_safe_div(_at(c, "Total Liabilities", i), _at(c, "Total Assets", i)))),
    ("Debt-to-Equity (x)", "Total Liabilities / Total Equity — Parent", "槓桿償債",
     lambda c, i: _safe_div(_at(c, "Total Liabilities", i), _at(c, "Total Equity — Parent", i))),
    ("Equity Multiplier (x)", "Total Assets / Total Equity — Parent", "槓桿償債",
     lambda c, i: _safe_div(_at(c, "Total Assets", i), _at(c, "Total Equity — Parent", i))),
    ("LT Debt to Capital (%)", "Long-term Debt / (Long-term Debt + Total Equity — Parent)", "槓桿償債",
     lambda c, i: _pct(_safe_div(
         _at(c, "Long-term Debt", i),
         None if _at(c, "Long-term Debt", i) is None or _at(c, "Total Equity — Parent", i) is None
         else float(_at(c, "Long-term Debt", i)) + float(_at(c, "Total Equity — Parent", i))))),

    # ── 現金流（額外）───────────────────────────────────────────────────
    ("Operating CF Margin (%)", "Operating Cash Flow / Revenue", "現金流",
     lambda c, i: _pct(_safe_div(_at(c, "Operating Cash Flow", i), _at(c, "Revenue", i)))),
    ("OCF / Net Income (x)", "Operating Cash Flow / Net Income", "現金流",
     lambda c, i: _safe_div(_at(c, "Operating Cash Flow", i), _at(c, "Net Income", i))),

    # ── 營運效率（額外，週轉次數版）─────────────────────────────────────
    ("Asset Turnover (x)", "TTM Revenue / avg Total Assets", "營運效率",
     lambda c, i: _safe_div(_ttm(c.get("Revenue"), i, _labels(c)),
                            _avg(_at_lag(c, "Total Assets", i, 3), _at(c, "Total Assets", i)))),
    ("Inventory Turnover (x)", "TTM Cost of Revenue / avg Inventories", "營運效率",
     lambda c, i: _safe_div(_ttm(c.get("Cost of Revenue"), i, _labels(c)),
                            _avg(_at_lag(c, "Inventories", i, 3), _at(c, "Inventories", i)))),
    ("Receivables Turnover (x)", "TTM Revenue / avg Accounts Receivable", "營運效率",
     lambda c, i: _safe_div(_ttm(c.get("Revenue"), i, _labels(c)),
                            _avg(_at_lag(c, "Accounts Receivable", i, 3), _at(c, "Accounts Receivable", i)))),

    # ── 結構／規模 ──────────────────────────────────────────────────────
    ("D&A / Revenue (%)", "D&A / Revenue", "結構",
     lambda c, i: _pct(_safe_div(_at(c, "D&A", i), _at(c, "Revenue", i)))),
    ("EBITDA ($mm)", "Operating Income + D&A", "結構",
     lambda c, i: _ebitda(c, i)),
    ("Total Debt ($mm)", "Short-term Debt + Current Portion of LT Debt + Long-term Debt", "結構",
     lambda c, i: _total_debt(c, i)),
    ("Net Debt ($mm)", "Total Debt - Cash", "結構",
     lambda c, i: _net_debt(c, i)),
    ("Working Capital ($mm)", "Total Current Assets - Total Current Liabilities", "結構",
     lambda c, i: None if _at(c, "Total Current Assets", i) is None or _at(c, "Total Current Liabilities", i) is None
     else float(_at(c, "Total Current Assets", i)) - float(_at(c, "Total Current Liabilities", i))),
    ("Cash Ratio (x)", "Cash / Total Current Liabilities", "結構",
     lambda c, i: _safe_div(_at(c, "Cash", i), _at(c, "Total Current Liabilities", i))),
    ("COGS Ratio (%)", "Cost of Revenue / Revenue", "結構",
     lambda c, i: _pct(_safe_div(_at(c, "Cost of Revenue", i), _at(c, "Revenue", i)))),

    # ── 報酬率（近似值）─────────────────────────────────────────────────
    ("ROIC (%)", "Operating Income x (1 - Effective Tax Rate) / (Total Debt + Equity - Cash) [approx.]", "報酬率",
     lambda c, i: _roic(c, i)),
]
```

- [ ] **Step 4: 執行測試確認通過**

Run: `./venv/Scripts/python.exe -m pytest tests/test_ratios.py -v`
Expected: 全部 PASS（既有 28 個 + 新增 21 個，共 49 個比率的測試都要過）

- [ ] **Step 5: Commit**

```bash
git add src/ratios.py tests/test_ratios.py
git commit -m "feat: ratios.py 新增 21 個比率（跨公司比較用），含 ROIC/Debt Ratio/EBITDA 等"
```

---

## Task 4: `src/comparison.py` — 多公司資料抓取與重組

**Files:**
- Create: `src/comparison.py`
- Test: `tests/test_comparison.py`

**Interfaces:**
- Consumes: `fetcher_gaap.fetch_gaap_statements(ticker, identity, ...) -> list[StatementTable]`、`fetcher_gaap.StatementTable`（含 `.period_ends: list[str]`）、`ratios.build_ratio_table(q_table) -> StatementTable | None`
- Produces:
  - `CompanyFetchError`（dataclass：`ticker: str`、`error_type: str`）
  - `ComparisonResult`（dataclass：`metrics: dict[str, dict[str, dict[str, float | None]]]`
    ＝`{指標名: {ticker: {period_label: value}}}`、`period_ends: dict[str, dict[str, str]]`
    ＝`{ticker: {period_label: 期末結算日}}`、`failures: list[CompanyFetchError]`）
  - `build_comparison(tickers: list[str], identity: str, metric_names: list[str], *, frequency: str, start_year: int | None, end_year: int | None, max_filings: int = 80, max_annual_filings: int = 20) -> ComparisonResult`
    （`frequency` 是 `"quarterly"` 或 `"annual"`）

- [ ] **Step 1: 寫失敗測試**

Create `tests/test_comparison.py`:

```python
"""Tests for comparison.py — 跨公司比較的資料重組。"""
from unittest.mock import patch

import pytest

from fetcher_gaap import StatementTable
from comparison import build_comparison, ComparisonResult, CompanyFetchError


def _fake_q_table(ticker, revenue, gross_profit, period_ends):
    labels = [f"FY2024Q{i+1}" for i in range(len(revenue))]
    return StatementTable(
        sheet_name="Data_Financials(Q)",
        quarter_labels=labels,
        filing_dates=[""] * len(revenue),
        concepts=["Revenue", "Gross Profit"],
        values=[revenue, gross_profit],
        ticker=ticker,
        labels=["", ""],
        period_ends=period_ends,
    )


def test_build_comparison_extracts_raw_concept_across_companies():
    def fake_fetch(ticker, identity, **kwargs):
        data = {
            "NVDA": _fake_q_table("NVDA", [100.0, 110.0], [50.0, 60.0],
                                  ["2024-03-31", "2024-06-30"]),
            "AMD": _fake_q_table("AMD", [80.0, 90.0], [30.0, 35.0],
                                 ["2024-03-31", "2024-06-30"]),
        }
        return [data[ticker]]

    with patch("comparison.fetch_gaap_statements", side_effect=fake_fetch):
        result = build_comparison(
            ["NVDA", "AMD"], "test@example.com", ["Revenue"],
            frequency="quarterly", start_year=None, end_year=None,
        )

    assert isinstance(result, ComparisonResult)
    assert result.metrics["Revenue"]["NVDA"]["FY2024Q1"] == 100.0
    assert result.metrics["Revenue"]["AMD"]["FY2024Q2"] == 90.0
    assert result.period_ends["NVDA"]["FY2024Q1"] == "2024-03-31"
    assert result.failures == []


def test_build_comparison_extracts_ratio_metric():
    def fake_fetch(ticker, identity, **kwargs):
        return [_fake_q_table(ticker, [100.0, 200.0], [50.0, 80.0],
                              ["2024-03-31", "2024-06-30"])]

    with patch("comparison.fetch_gaap_statements", side_effect=fake_fetch):
        result = build_comparison(
            ["NVDA"], "test@example.com", ["Gross Margin (%)"],
            frequency="quarterly", start_year=None, end_year=None,
        )

    # Gross Margin = Gross Profit / Revenue: 50/100=50%, 80/200=40%
    assert result.metrics["Gross Margin (%)"]["NVDA"]["FY2024Q1"] == pytest.approx(50.0)
    assert result.metrics["Gross Margin (%)"]["NVDA"]["FY2024Q2"] == pytest.approx(40.0)


def test_build_comparison_skips_failed_company_and_continues():
    def fake_fetch(ticker, identity, **kwargs):
        if ticker == "BADTICKER":
            raise ValueError("No 10-Q filings found for ticker 'BADTICKER'.")
        return [_fake_q_table(ticker, [100.0], [50.0], ["2024-03-31"])]

    with patch("comparison.fetch_gaap_statements", side_effect=fake_fetch):
        result = build_comparison(
            ["NVDA", "BADTICKER"], "test@example.com", ["Revenue"],
            frequency="quarterly", start_year=None, end_year=None,
        )

    assert result.metrics["Revenue"]["NVDA"]["FY2024Q1"] == 100.0
    assert "BADTICKER" not in result.metrics["Revenue"]
    assert len(result.failures) == 1
    assert result.failures[0] == CompanyFetchError(ticker="BADTICKER", error_type="ValueError")


def test_build_comparison_annual_frequency_reads_data_financials_y():
    def fake_fetch(ticker, identity, **kwargs):
        q_tbl = _fake_q_table(ticker, [100.0], [50.0], ["2024-03-31"])
        y_tbl = StatementTable(
            sheet_name="Data_Financials(Y)", quarter_labels=["FY2024"],
            filing_dates=[""], concepts=["Revenue"], values=[[400.0]],
            ticker=ticker, labels=[""], period_ends=["2024-12-31"],
        )
        return [q_tbl, y_tbl]

    with patch("comparison.fetch_gaap_statements", side_effect=fake_fetch):
        result = build_comparison(
            ["NVDA"], "test@example.com", ["Revenue"],
            frequency="annual", start_year=None, end_year=None,
        )

    assert result.metrics["Revenue"]["NVDA"]["FY2024"] == 400.0


def test_build_comparison_filters_by_year_range():
    def fake_fetch(ticker, identity, **kwargs):
        labels = ["FY2022Q4", "FY2023Q1", "FY2024Q1"]
        return [StatementTable(
            sheet_name="Data_Financials(Q)", quarter_labels=labels,
            filing_dates=[""] * 3, concepts=["Revenue"], values=[[10.0, 20.0, 30.0]],
            ticker=ticker, labels=[""], period_ends=["2022-12-31", "2023-03-31", "2024-03-31"],
        )]

    with patch("comparison.fetch_gaap_statements", side_effect=fake_fetch):
        result = build_comparison(
            ["NVDA"], "test@example.com", ["Revenue"],
            frequency="quarterly", start_year=2023, end_year=2023,
        )

    assert list(result.metrics["Revenue"]["NVDA"].keys()) == ["FY2023Q1"]
```

- [ ] **Step 2: 執行測試確認失敗**

Run: `./venv/Scripts/python.exe -m pytest tests/test_comparison.py -v`
Expected: FAIL，`ModuleNotFoundError: No module named 'comparison'`

- [ ] **Step 3: 實作**

Create `src/comparison.py`:

```python
"""comparison.py — 跨公司財務比較的資料抓取與重組。

把多家公司各自的 fetch_gaap_statements() 結果，重組成
{指標名: {ticker: {period_label: value}}} 這種給 comparison_writer.py
直接寫表用的形狀。單一公司抓取失敗不中斷整體流程，記錄下來繼續下一家
（比照 fetcher_gaap.collect_gaps() 的「跳過不中斷」原則，但這裡是公司
層級的跳過，不是同一家公司內部的科目缺漏，所以不共用同一套機制）。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal

from fetcher_gaap import StatementTable, fetch_gaap_statements
from ratios import build_ratio_table, RATIO_DEFS


_RATIO_NAMES = {name for name, _, _, _ in RATIO_DEFS}


@dataclass(frozen=True)
class CompanyFetchError:
    ticker: str
    error_type: str


@dataclass
class ComparisonResult:
    metrics: dict[str, dict[str, dict[str, float | None]]] = field(default_factory=dict)
    period_ends: dict[str, dict[str, str]] = field(default_factory=dict)
    failures: list[CompanyFetchError] = field(default_factory=list)


def _sheet_name_for(frequency: Literal["quarterly", "annual"]) -> str:
    return "Data_Financials(Q)" if frequency == "quarterly" else "Data_Financials(Y)"


def _filter_by_year(table: StatementTable, start_year: int | None, end_year: int | None) -> StatementTable:
    """依 period_ends 的年份篩選欄位。沒有 period_ends 資料的欄一律保留
    （不因為篩不了就整欄丟掉，寧可多顯示也不要漏資料）。"""
    if start_year is None and end_year is None:
        return table

    keep = []
    for i, end in enumerate(table.period_ends or []):
        if not end:
            keep.append(i)
            continue
        try:
            year = int(end[:4])
        except (TypeError, ValueError):
            keep.append(i)
            continue
        if start_year is not None and year < start_year:
            continue
        if end_year is not None and year > end_year:
            continue
        keep.append(i)

    return StatementTable(
        sheet_name=table.sheet_name,
        quarter_labels=[table.quarter_labels[i] for i in keep],
        filing_dates=[table.filing_dates[i] for i in keep] if table.filing_dates else [],
        concepts=table.concepts,
        values=[[row[i] for i in keep] for row in table.values],
        ticker=table.ticker,
        labels=table.labels,
        period_ends=[table.period_ends[i] for i in keep] if table.period_ends else [],
    )


def build_comparison(
    tickers: list[str],
    identity: str,
    metric_names: list[str],
    *,
    frequency: Literal["quarterly", "annual"],
    start_year: int | None,
    end_year: int | None,
    max_filings: int = 80,
    max_annual_filings: int = 20,
) -> ComparisonResult:
    """對每個 ticker 抓資料、抽出選定指標，重組成跨公司比較用的資料結構。"""
    result = ComparisonResult(metrics={name: {} for name in metric_names})
    sheet_name = _sheet_name_for(frequency)

    for ticker in tickers:
        ticker = ticker.strip().upper()
        if not ticker:
            continue
        try:
            tables = fetch_gaap_statements(
                ticker, identity, max_filings=max_filings,
                max_annual_filings=max_annual_filings,
                fetch_quarterly=(frequency == "quarterly"),
                fetch_annual=(frequency == "annual"),
            )
        except Exception as e:
            result.failures.append(CompanyFetchError(ticker=ticker, error_type=type(e).__name__))
            continue

        raw_table = next((t for t in tables if t.sheet_name == sheet_name), None)
        if raw_table is None:
            result.failures.append(CompanyFetchError(ticker=ticker, error_type="NoDataForFrequency"))
            continue

        raw_table = _filter_by_year(raw_table, start_year, end_year)
        ratio_table = build_ratio_table(raw_table)

        period_map: dict[str, str] = {}
        for i, label in enumerate(raw_table.quarter_labels):
            end = raw_table.period_ends[i] if i < len(raw_table.period_ends or []) else ""
            period_map[label] = end
        result.period_ends[ticker] = period_map

        for metric_name in metric_names:
            source_table = ratio_table if metric_name in _RATIO_NAMES else raw_table
            if source_table is None or metric_name not in source_table.concepts:
                continue
            row = source_table.values[source_table.concepts.index(metric_name)]
            result.metrics.setdefault(metric_name, {})[ticker] = dict(
                zip(source_table.quarter_labels, row)
            )

    return result
```

- [ ] **Step 4: 執行測試確認通過**

Run: `./venv/Scripts/python.exe -m pytest tests/test_comparison.py -v`
Expected: 全部 PASS

- [ ] **Step 5: Commit**

```bash
git add src/comparison.py tests/test_comparison.py
git commit -m "feat: 新增 comparison.py，跨公司資料抓取與重組"
```

---

## Task 5: `src/comparison_writer.py` — `Compare_Data` sheet

**Files:**
- Create: `src/comparison_writer.py`
- Test: `tests/test_comparison_writer.py`

**Interfaces:**
- Consumes: Task 4 的 `ComparisonResult`、Task 1 的 `unit_format_for()`
- Produces: `write_compare_data_sheet(wb: Workbook, result: ComparisonResult, metric_names: list[str]) -> dict[str, tuple[int, int]]`
  （回傳 `{指標名: (資料起始列, 資料結束列)}`，供 Task 6/7 的 Snapshot 公式與 Task 8 的圖表定位用）

- [ ] **Step 1: 寫失敗測試**

Create `tests/test_comparison_writer.py`:

```python
"""Tests for comparison_writer.py — 跨公司比較 Excel 輸出。"""
from openpyxl import Workbook

from comparison import ComparisonResult
from comparison_writer import write_compare_data_sheet


def _sample_result():
    return ComparisonResult(
        metrics={
            "Revenue": {
                "NVDA": {"FY2024Q1": 100.0, "FY2024Q2": 110.0},
                "AMD": {"FY2024Q1": 80.0, "FY2024Q2": 90.0},
            },
        },
        period_ends={
            "NVDA": {"FY2024Q1": "2024-03-31", "FY2024Q2": "2024-06-30"},
            "AMD": {"FY2024Q1": "2024-03-31", "FY2024Q2": "2024-06-30"},
        },
        failures=[],
    )


def test_compare_data_sheet_has_metric_header_and_period_columns():
    wb = Workbook()
    result = _sample_result()
    write_compare_data_sheet(wb, result, ["Revenue"])
    ws = wb["Compare_Data"]

    # A1 是指標名稱區塊標題
    assert ws["A1"].value == "Revenue"
    # 第一列（標題列）是期間標籤，從 B 欄開始
    header_row = [c.value for c in ws[2]]
    assert "FY2024Q1" in header_row
    assert "FY2024Q2" in header_row


def test_compare_data_sheet_has_static_period_end_row():
    wb = Workbook()
    result = _sample_result()
    write_compare_data_sheet(wb, result, ["Revenue"])
    ws = wb["Compare_Data"]

    # 第三列是期末結算日，靜態文字，不是公式
    row3 = [c.value for c in ws[3]]
    assert "2024-03-31" in row3
    for cell in ws[3]:
        if cell.value:
            assert not str(cell.value).startswith("=")


def test_compare_data_sheet_lists_company_rows_with_values():
    wb = Workbook()
    result = _sample_result()
    write_compare_data_sheet(wb, result, ["Revenue"])
    ws = wb["Compare_Data"]

    company_col_a = [c.value for c in ws["A"]]
    assert "NVDA" in company_col_a
    assert "AMD" in company_col_a


def test_compare_data_sheet_returns_block_ranges():
    wb = Workbook()
    result = _sample_result()
    ranges = write_compare_data_sheet(wb, result, ["Revenue"])
    assert "Revenue" in ranges
    start, end = ranges["Revenue"]
    assert end > start


def test_compare_data_sheet_stacks_multiple_metric_blocks():
    wb = Workbook()
    result = _sample_result()
    result.metrics["Gross Margin (%)"] = {"NVDA": {"FY2024Q1": 50.0}}
    ranges = write_compare_data_sheet(wb, result, ["Revenue", "Gross Margin (%)"])
    rev_start, rev_end = ranges["Revenue"]
    gm_start, gm_end = ranges["Gross Margin (%)"]
    assert gm_start > rev_end
```

- [ ] **Step 2: 執行測試確認失敗**

Run: `./venv/Scripts/python.exe -m pytest tests/test_comparison_writer.py -v`
Expected: FAIL，`ModuleNotFoundError: No module named 'comparison_writer'`

- [ ] **Step 3: 實作**

Create `src/comparison_writer.py`:

```python
"""comparison_writer.py — 把 comparison.py 的資料結構寫成跨公司比較 Excel。

Sheet 結構（見 docs/superpowers/specs/2026-08-20-cross-company-comparison-design.md）：
  Compare_Data    — 唯一一張原始資料表，每個指標一個區塊往下疊
  Snapshot        — 活的，公式驅動的單一時間點快照
  Snapshot_Manual — 空白，供人工貼值凍結存檔
  Chart_<指標>     — 每個指標各一張，只放圖表
"""

from __future__ import annotations

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

from comparison import ComparisonResult
from excel_formatter import unit_format_for

_HEADER_FONT = Font(bold=True)
_BLOCK_GAP = 1  # 區塊之間空幾列


def write_compare_data_sheet(
    wb: Workbook, result: ComparisonResult, metric_names: list[str]
) -> dict[str, tuple[int, int]]:
    """寫 Compare_Data。回傳 {指標名: (資料列起, 資料列迄)}（不含標題/期末結算日列），
    給 Snapshot 的 MATCH 公式與 Chart 的資料來源 range 用。"""
    ws = wb.active
    ws.title = "Compare_Data"

    all_companies = sorted({
        company
        for metric_data in result.metrics.values()
        for company in metric_data
    })

    block_ranges: dict[str, tuple[int, int]] = {}
    row = 1
    for metric_name in metric_names:
        metric_data = result.metrics.get(metric_name, {})
        fmt, divisor = unit_format_for(metric_name)

        # 收集這個指標出現過的所有期間標籤，依標籤字串排序（FYyyyyQq 天然可字串排序）
        periods: list[str] = sorted({
            label for company_data in metric_data.values() for label in company_data
        })

        # 標題列
        title_cell = ws.cell(row=row, column=1, value=metric_name)
        title_cell.font = _HEADER_FONT
        header_row = row + 1
        ws.cell(row=header_row, column=1, value="公司")
        for col, period in enumerate(periods, start=2):
            ws.cell(row=header_row, column=col, value=period)

        # 期末結算日列（靜態文字，供 Snapshot 用）
        end_date_row = header_row + 1
        ws.cell(row=end_date_row, column=1, value="期末結算日")
        for col, period in enumerate(periods, start=2):
            end_date = ""
            for company in all_companies:
                end_date = result.period_ends.get(company, {}).get(period, "")
                if end_date:
                    break
            ws.cell(row=end_date_row, column=col, value=end_date)

        # 公司資料列
        data_start = end_date_row + 1
        for offset, company in enumerate(all_companies):
            r = data_start + offset
            ws.cell(row=r, column=1, value=company)
            company_data = metric_data.get(company, {})
            for col, period in enumerate(periods, start=2):
                value = company_data.get(period)
                cell = ws.cell(row=r, column=col, value=value)
                if isinstance(value, (int, float)):
                    cell.value = value / divisor
                    cell.number_format = fmt
        data_end = data_start + len(all_companies) - 1

        block_ranges[metric_name] = (data_start, data_end)
        row = data_end + 1 + _BLOCK_GAP

    return block_ranges
```

- [ ] **Step 4: 執行測試確認通過**

Run: `./venv/Scripts/python.exe -m pytest tests/test_comparison_writer.py -v`
Expected: 全部 PASS

- [ ] **Step 5: Commit**

```bash
git add src/comparison_writer.py tests/test_comparison_writer.py
git commit -m "feat: comparison_writer.py 新增 Compare_Data sheet 寫入"
```

---

## Task 6: `src/comparison_writer.py` — `Snapshot`（活公式）＋ `Snapshot_Manual`

**Files:**
- Modify: `src/comparison_writer.py`
- Test: `tests/test_comparison_writer.py`

**Interfaces:**
- Consumes: Task 5 的 `write_compare_data_sheet()` 回傳的 `block_ranges`
- Produces: `write_snapshot_sheets(wb: Workbook, result: ComparisonResult, metric_names: list[str], block_ranges: dict[str, tuple[int, int]], default_date: str = "") -> None`

- [ ] **Step 1: 寫失敗測試**

在 `tests/test_comparison_writer.py` 尾端加入：

```python
from comparison_writer import write_snapshot_sheets


def test_snapshot_sheet_has_yellow_input_cell_and_formulas():
    wb = Workbook()
    result = _sample_result()
    block_ranges = write_compare_data_sheet(wb, result, ["Revenue"])
    write_snapshot_sheets(wb, result, ["Revenue"], block_ranges, default_date="2024-03-31")

    ws = wb["Snapshot"]
    assert ws["B1"].value == "2024-03-31"
    assert ws["B1"].fill.fgColor.rgb in ("00FFFF00", "FFFFFF00")

    # 公司資料格要是 INDEX/MATCH 公式，不是寫死的值
    body = [[c.value for c in row] for row in ws.iter_rows(min_row=3)]
    formula_cells = [v for row in body for v in row if isinstance(v, str) and v.startswith("=")]
    assert formula_cells, "Snapshot 應該用公式，不是寫死的值"
    assert any("INDEX" in f and "MATCH" in f for f in formula_cells)


def test_snapshot_manual_sheet_is_blank_with_same_headers():
    wb = Workbook()
    result = _sample_result()
    block_ranges = write_compare_data_sheet(wb, result, ["Revenue"])
    write_snapshot_sheets(wb, result, ["Revenue"], block_ranges, default_date="2024-03-31")

    ws = wb["Snapshot_Manual"]
    header_row = [c.value for c in ws[1]]
    assert "Revenue" in header_row
    company_col = [c.value for c in ws["A"]]
    assert "NVDA" in company_col
    # 資料格是空的
    for row in ws.iter_rows(min_row=2, min_col=2):
        for cell in row:
            assert cell.value is None
```

- [ ] **Step 2: 執行測試確認失敗**

Run: `./venv/Scripts/python.exe -m pytest tests/test_comparison_writer.py -k snapshot -v`
Expected: FAIL，`ImportError: cannot import name 'write_snapshot_sheets'`

- [ ] **Step 3: 實作**

在 `src/comparison_writer.py` 尾端新增：

```python
from openpyxl.utils import get_column_letter

_YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")


def write_snapshot_sheets(
    wb: Workbook,
    result: ComparisonResult,
    metric_names: list[str],
    block_ranges: dict[str, tuple[int, int]],
    default_date: str = "",
) -> None:
    """寫 Snapshot（活公式）與 Snapshot_Manual（空白，供人工貼值）。

    Snapshot 用 INDEX/MATCH 對 Compare_Data 每個指標區塊的「期末結算日」列
    （每個區塊的資料起始列往上一列，見 write_compare_data_sheet 的排版）取值，
    改 B1 的日期 Excel 會自動重算——這是刻意選擇用真公式，不是寫死算好的值，
    因為這份只給人在 Excel 裡看，沒有下游腳本要讀它（讀 Snapshot_Manual）。
    """
    all_companies = sorted({
        company
        for metric_data in result.metrics.values()
        for company in metric_data
    })

    snap = wb.create_sheet("Snapshot")
    snap["A1"] = "時間點"
    snap["B1"] = default_date
    snap["B1"].fill = _YELLOW_FILL

    header_row = 2
    snap.cell(row=header_row, column=1, value="公司")
    for col, metric_name in enumerate(metric_names, start=2):
        snap.cell(row=header_row, column=col, value=metric_name)

    for r_offset, company in enumerate(all_companies):
        r = header_row + 1 + r_offset
        snap.cell(row=r, column=1, value=company)
        for col, metric_name in enumerate(metric_names, start=2):
            if metric_name not in block_ranges:
                continue
            data_start, data_end = block_ranges[metric_name]
            end_date_row = data_start - 1   # 期末結算日列緊接在資料列上方
            header_period_row = data_start - 2  # 期間標籤列
            company_row = None
            for rr in range(data_start, data_end + 1):
                if ws_company_at("Compare_Data", wb, rr) == company:
                    company_row = rr
                    break
            if company_row is None:
                continue

            last_col = wb["Compare_Data"].max_column
            end_date_range = (
                f"Compare_Data!$B${end_date_row}:${get_column_letter(last_col)}${end_date_row}"
            )
            data_range = (
                f"Compare_Data!$B${company_row}:${get_column_letter(last_col)}${company_row}"
            )
            formula = f'=INDEX({data_range},MATCH($B$1,{end_date_range},0))'
            snap.cell(row=r, column=col, value=formula)

    # Snapshot_Manual：同樣的表頭，資料格留空供人工貼值
    manual = wb.create_sheet("Snapshot_Manual")
    manual.cell(row=1, column=1, value="公司")
    for col, metric_name in enumerate(metric_names, start=2):
        manual.cell(row=1, column=col, value=metric_name)
    for r_offset, company in enumerate(all_companies):
        manual.cell(row=2 + r_offset, column=1, value=company)


def ws_company_at(sheet_name: str, wb: Workbook, row: int) -> str | None:
    """讀 Compare_Data 某一列 A 欄的公司代號。"""
    return wb[sheet_name].cell(row=row, column=1).value
```

- [ ] **Step 4: 執行測試確認通過**

Run: `./venv/Scripts/python.exe -m pytest tests/test_comparison_writer.py -v`
Expected: 全部 PASS

- [ ] **Step 5: Commit**

```bash
git add src/comparison_writer.py tests/test_comparison_writer.py
git commit -m "feat: comparison_writer.py 新增 Snapshot（活公式）與 Snapshot_Manual"
```

---

## Task 7: `src/comparison_writer.py` — `Chart_<指標>` 歷史趨勢圖

**Files:**
- Modify: `src/comparison_writer.py`
- Test: `tests/test_comparison_writer.py`

**Interfaces:**
- Consumes: Task 5 的 `block_ranges`
- Produces: `write_chart_sheets(wb: Workbook, metric_names: list[str], block_ranges: dict[str, tuple[int, int]]) -> None`

- [ ] **Step 1: 寫失敗測試**

在 `tests/test_comparison_writer.py` 尾端加入：

```python
from comparison_writer import write_chart_sheets


def test_chart_sheet_created_per_metric_with_line_chart():
    wb = Workbook()
    result = _sample_result()
    block_ranges = write_compare_data_sheet(wb, result, ["Revenue"])
    write_chart_sheets(wb, ["Revenue"], block_ranges)

    assert "Chart_Revenue" in wb.sheetnames
    ws = wb["Chart_Revenue"]
    assert len(ws._charts) == 1


def test_chart_sheet_name_truncates_long_metric_names():
    wb = Workbook()
    result = _sample_result()
    long_name = "A Very Long Metric Name That Exceeds Excel Sheet Name Limit (%)"
    result.metrics[long_name] = {"NVDA": {"FY2024Q1": 1.0}}
    block_ranges = write_compare_data_sheet(wb, result, ["Revenue", long_name])
    write_chart_sheets(wb, ["Revenue", long_name], block_ranges)

    # Excel sheet 名稱上限 31 字元
    assert all(len(name) <= 31 for name in wb.sheetnames)
```

- [ ] **Step 2: 執行測試確認失敗**

Run: `./venv/Scripts/python.exe -m pytest tests/test_comparison_writer.py -k chart -v`
Expected: FAIL，`ImportError: cannot import name 'write_chart_sheets'`

- [ ] **Step 3: 實作**

在 `src/comparison_writer.py` 尾端新增：

```python
from openpyxl.chart import LineChart, Reference


def _chart_sheet_name(metric_name: str) -> str:
    """Chart_<指標> 但要塞進 Excel 的 31 字元 sheet 名稱上限。"""
    prefix = "Chart_"
    max_metric_len = 31 - len(prefix)
    safe_name = "".join(ch for ch in metric_name if ch not in '[]:*?/\\')
    return prefix + safe_name[:max_metric_len]


def write_chart_sheets(
    wb: Workbook, metric_names: list[str], block_ranges: dict[str, tuple[int, int]]
) -> None:
    """每個指標各一張 sheet，只放一張折線圖（歷史趨勢，一條線一家公司）。
    使用者要看長條圖版本，在 Excel 裡對圖表右鍵「變更圖表類型」自己切，
    這裡不用同一指標產兩份圖表物件。"""
    data_ws = wb["Compare_Data"]

    for metric_name in metric_names:
        if metric_name not in block_ranges:
            continue
        data_start, data_end = block_ranges[metric_name]
        header_row = data_start - 2
        last_col = data_ws.max_column

        chart = LineChart()
        chart.title = metric_name
        chart.style = 2
        chart.y_axis.title = metric_name
        chart.x_axis.title = "期間"

        data_ref = Reference(
            data_ws, min_col=1, max_col=last_col, min_row=data_start, max_row=data_end
        )
        chart.add_data(data_ref, titles_from_data=True, from_rows=True)

        categories_ref = Reference(
            data_ws, min_col=2, max_col=last_col, min_row=header_row, max_row=header_row
        )
        chart.set_categories(categories_ref)

        sheet_name = _chart_sheet_name(metric_name)
        chart_ws = wb.create_sheet(sheet_name)
        chart_ws.add_chart(chart, "B2")
```

- [ ] **Step 4: 執行測試確認通過**

Run: `./venv/Scripts/python.exe -m pytest tests/test_comparison_writer.py -v`
Expected: 全部 PASS

- [ ] **Step 5: Commit**

```bash
git add src/comparison_writer.py tests/test_comparison_writer.py
git commit -m "feat: comparison_writer.py 新增每指標一張的歷史趨勢圖表 sheet"
```

---

## Task 8: `src/comparison_writer.py` — 整合入口 `write_comparison_workbook()`

**Files:**
- Modify: `src/comparison_writer.py`
- Test: `tests/test_comparison_writer.py`

**Interfaces:**
- Consumes: Task 5-7 的三個寫入函式
- Produces: `write_comparison_workbook(result: ComparisonResult, metric_names: list[str], output_path: Path, snapshot_date: str = "") -> None`

- [ ] **Step 1: 寫失敗測試**

在 `tests/test_comparison_writer.py` 尾端加入：

```python
import tempfile
from pathlib import Path

from openpyxl import load_workbook

from comparison_writer import write_comparison_workbook


def test_write_comparison_workbook_produces_all_expected_sheets():
    result = _sample_result()
    with tempfile.TemporaryDirectory() as tmp:
        out_path = Path(tmp) / "compare_test.xlsx"
        write_comparison_workbook(result, ["Revenue"], out_path, snapshot_date="2024-03-31")

        assert out_path.exists()
        wb = load_workbook(out_path)
        assert wb.sheetnames == ["Compare_Data", "Snapshot", "Snapshot_Manual", "Chart_Revenue"]
```

- [ ] **Step 2: 執行測試確認失敗**

Run: `./venv/Scripts/python.exe -m pytest tests/test_comparison_writer.py -k write_comparison_workbook -v`
Expected: FAIL，`ImportError`

- [ ] **Step 3: 實作**

在 `src/comparison_writer.py` 尾端新增：

```python
from pathlib import Path


def write_comparison_workbook(
    result: ComparisonResult,
    metric_names: list[str],
    output_path: Path,
    snapshot_date: str = "",
) -> None:
    """組出完整跨公司比較 Excel 並存檔。"""
    wb = Workbook()
    block_ranges = write_compare_data_sheet(wb, result, metric_names)
    write_snapshot_sheets(wb, result, metric_names, block_ranges, default_date=snapshot_date)
    write_chart_sheets(wb, metric_names, block_ranges)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
```

- [ ] **Step 4: 執行測試確認通過**

Run: `./venv/Scripts/python.exe -m pytest tests/test_comparison_writer.py -v`
Expected: 全部 PASS

- [ ] **Step 5: Commit**

```bash
git add src/comparison_writer.py tests/test_comparison_writer.py
git commit -m "feat: comparison_writer.py 新增 write_comparison_workbook() 整合入口"
```

---

## Task 9: `main.py` — Tab4「跨公司比較」選擇視窗

**背景**：這是 GUI 任務，`tkinter`/`ttk` 沒有官方的自動化測試框架，這個專案現有的 GUI 程式碼
（`main.py`）也沒有針對 widget 互動寫單元測試（測試都集中在 `fetcher_gaap`／`ratios`／
`excel_formatter` 這些邏輯層）。這個任務改成「實作 + 手動驗證」，不用 TDD 的 fail-first 流程。

**Files:**
- Modify: `src/main.py`（新增 `_build_tab4()`、選擇視窗方法、公司自動完成邏輯）
- Modify: `src/locales/zh_tw.py`／`zh_cn.py`／`en.py`／`ja.py`（新增 Tab4 字串，見 Task 11）

**Interfaces:**
- Consumes: `company_cache.json`（既有，`CACHE_PATH` 已定義於 `main.py:106`）、Task 2/3 的
  `ratios.RATIO_CATEGORIES`、`ratios.RATIO_DEFS`、`fetcher_gaap.py` 的
  `IS_TEMPLATE`／`BS_TEMPLATE`／`CF_TEMPLATE`（`src/fetcher_gaap.py:253/282/334`，
  已依報表類型分開的科目定義清單，每筆 tuple 第 0 欄是顯示名稱）
- Produces: `self.compare_selected_tickers: list[tuple[str, str]]`（ticker, 公司名）、
  `self.compare_selected_metrics: list[str]`、`self.compare_start_year`／`end_year`／
  `frequency`／`snapshot_date` 這幾個 `tk.Variable`

- [ ] **Step 1: 找到 Tab 註冊點與既有 Tab1 pattern**

Read `src/main.py:434-438`（`self.notebook = ttk.Notebook(...)` 之後的 `_build_tab1()`／
`_build_tab2()`／`_build_tab3()` 呼叫）與 `src/main.py:468-636`（`_build_tab1()` 完整內容，
GUI 元件排版 pattern 的範本）。

- [ ] **Step 2: 新增公司快取查詢輔助方法**

在 `main.py` 找一個既有的小工具方法附近（例如 `_load_company_cache` 尚不存在，新增在
`_confirm_company` 方法之前），新增：

```python
    def _load_company_cache(self) -> dict[str, str]:
        """讀 company_cache.json，回傳 {ticker: 公司名}。讀不到就回空字典，
        不要讓自動完成功能因為快取檔案缺失而整個掛掉。"""
        if not CACHE_PATH.exists():
            return {}
        try:
            with open(CACHE_PATH, encoding="utf-8") as f:
                data = json.load(f)
            return data.get("companies", {})
        except (json.JSONDecodeError, OSError):
            return {}
```

- [ ] **Step 3: 新增 Tab4 主畫面 `_build_tab4()`**

在 `main.py` 的 `_build_tab3()` 方法定義之前（維持既有 `_build_tab1/2/3` 的順序慣例）新增：

```python
    def _build_tab4(self):
        """Build Tab 4 (跨公司比較): 選擇視窗按鈕、輸出設定、執行按鈕、進度條與 log。"""
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text=t("gui.tab.compare"))
        tab.columnconfigure(0, weight=1)

        self.compare_selected_tickers: list[tuple[str, str]] = []
        self.compare_selected_metrics: list[str] = []
        self.compare_start_year = tk.StringVar(value="")
        self.compare_end_year = tk.StringVar(value="")
        self.compare_frequency = tk.StringVar(value="quarterly")
        self.compare_snapshot_date = tk.StringVar(value="")

        summary_frame = ttk.Frame(tab, relief="groove", borderwidth=1, padding=8)
        summary_frame.grid(row=0, column=0, sticky="ew", pady=4)
        summary_frame.columnconfigure(0, weight=1)
        self.compare_summary_label = ttk.Label(summary_frame, text=t("gui.compare.no_selection"),
                                                justify="left")
        self.compare_summary_label.grid(row=0, column=0, sticky="w")

        ttk.Button(tab, text=t("gui.btn.compare_select"),
                   command=self._open_compare_selection_window).grid(row=1, column=0, pady=8)

        out_row = ttk.Frame(tab)
        out_row.grid(row=2, column=0, sticky="ew", pady=4)
        ttk.Label(out_row, text=t("gui.lbl.save_location")).pack(side="left")
        self.compare_outdir_var = tk.StringVar(value=str(PROJECT_ROOT / "output" / "compare"))
        ttk.Entry(out_row, textvariable=self.compare_outdir_var).pack(
            side="left", fill="x", expand=True, padx=(6, 6))
        ttk.Button(out_row, text=t("gui.btn.browse"), width=5,
                   command=self._browse_compare_output_dir).pack(side="left")

        self.compare_run_btn = ttk.Button(tab, text=t("gui.btn.compare_run"),
                                           command=self._run_comparison)
        self.compare_run_btn.grid(row=3, column=0, pady=8)

        self.compare_progress = ttk.Progressbar(tab, mode="determinate")
        self.compare_progress.grid(row=4, column=0, sticky="ew", pady=4)
        self.compare_status_label = ttk.Label(tab, text="")
        self.compare_status_label.grid(row=5, column=0, sticky="w")

        self.compare_log = scrolledtext.ScrolledText(tab, height=10, state="disabled")
        self.compare_log.grid(row=6, column=0, sticky="nsew", pady=(4, 0))
        tab.rowconfigure(6, weight=1)

    def _browse_compare_output_dir(self):
        from tkinter import filedialog
        current = self.compare_outdir_var.get().strip() or str(PROJECT_ROOT / "output" / "compare")
        folder = filedialog.askdirectory(title=t("gui.dlg.choose_output_dir"), initialdir=current)
        if folder:
            self.compare_outdir_var.set(folder)

    def _update_compare_summary(self):
        if not self.compare_selected_tickers or not self.compare_selected_metrics:
            self.compare_summary_label.config(text=t("gui.compare.no_selection"))
            return
        tickers_str = "、".join(tk_ for tk_, _ in self.compare_selected_tickers)
        metrics_str = "、".join(self.compare_selected_metrics[:5])
        if len(self.compare_selected_metrics) > 5:
            metrics_str += f" ...({len(self.compare_selected_metrics)})"
        freq_label = t("gui.compare.freq_quarterly") if self.compare_frequency.get() == "quarterly" \
            else t("gui.compare.freq_annual")
        text = (f"{t('gui.compare.companies')}: {tickers_str}\n"
                f"{t('gui.compare.period')}: {self.compare_start_year.get()}"
                f"~{self.compare_end_year.get()} ({freq_label})\n"
                f"{t('gui.compare.metrics')}: {metrics_str}")
        self.compare_summary_label.config(text=text)
```

`_build_tab4()` 呼叫要加進 `__init__` 既有的 `self._build_tab1(); self._build_tab2();
self._build_tab3()` 那三行之後（`src/main.py:434-438`）：

```python
        self._build_tab1()
        self._build_tab2()
        self._build_tab3()
        self._build_tab4()
```

- [ ] **Step 4: 新增選擇視窗 `_open_compare_selection_window()`**

同一個 class 裡新增：

```python
    def _open_compare_selection_window(self):
        win = tk.Toplevel(self.root)
        win.title(t("gui.compare.select_title"))
        win.geometry("560x640")

        # ── ① 選公司 ──────────────────────────────────────────────
        ttk.Label(win, text=t("gui.compare.step1_company"), font=("", 11, "bold")).pack(
            anchor="w", padx=10, pady=(10, 2))

        ticker_row = ttk.Frame(win)
        ticker_row.pack(fill="x", padx=10)
        ttk.Label(ticker_row, text=t("gui.compare.ticker_input")).pack(side="left")
        ticker_var = tk.StringVar()
        ticker_entry = ttk.Entry(ticker_row, textvariable=ticker_var, width=30)
        ticker_entry.pack(side="left", padx=(6, 0), fill="x", expand=True)

        suggest_listbox = tk.Listbox(win, height=4)
        cache = self._load_company_cache()

        def _on_ticker_type(*_):
            text = ticker_var.get().strip().upper()
            suggest_listbox.delete(0, "end")
            if not text or "," in text:
                return
            matches = [(tk_, name) for tk_, name in cache.items() if tk_.startswith(text)][:8]
            for tk_, name in matches:
                suggest_listbox.insert("end", f"{tk_}  {name}")

        ticker_var.trace_add("write", _on_ticker_type)

        chips_frame = ttk.Frame(win)
        chips_frame.pack(fill="x", padx=10, pady=(4, 0))

        def _refresh_company_chips():
            for child in chips_frame.winfo_children():
                child.destroy()
            for tk_, name in self.compare_selected_tickers:
                chip = ttk.Frame(chips_frame, relief="raised", borderwidth=1)
                chip.pack(side="left", padx=2, pady=2)
                ttk.Label(chip, text=f"{tk_} {name}").pack(side="left", padx=(4, 0))
                ttk.Button(chip, text="✕", width=2,
                           command=lambda t_=tk_: _remove_company(t_)).pack(side="left")

        def _add_company(ticker: str):
            ticker = ticker.strip().upper()
            if not ticker or any(t_ == ticker for t_, _ in self.compare_selected_tickers):
                return
            name = cache.get(ticker, "")
            if not name:
                messagebox.showwarning(
                    t("gui.compare.unknown_ticker_title"),
                    t("gui.compare.unknown_ticker_msg").format(ticker=ticker))
                return
            self.compare_selected_tickers.append((ticker, name))
            _refresh_company_chips()

        def _remove_company(ticker: str):
            self.compare_selected_tickers = [
                (t_, n) for t_, n in self.compare_selected_tickers if t_ != ticker
            ]
            _refresh_company_chips()

        def _on_ticker_submit(event=None):
            text = ticker_var.get().strip()
            if "," in text:
                for part in text.split(","):
                    _add_company(part)
            else:
                _add_company(text)
            ticker_var.set("")
            suggest_listbox.delete(0, "end")

        ticker_entry.bind("<Return>", _on_ticker_submit)

        def _on_suggest_pick(event):
            selection = suggest_listbox.curselection()
            if not selection:
                return
            picked = suggest_listbox.get(selection[0]).split()[0]
            _add_company(picked)
            ticker_var.set("")
            suggest_listbox.delete(0, "end")

        suggest_listbox.bind("<<ListboxSelect>>", _on_suggest_pick)
        suggest_listbox.pack(fill="x", padx=10)

        _refresh_company_chips()

        ttk.Separator(win, orient="horizontal").pack(fill="x", padx=10, pady=8)

        # ── ② 選指標 ──────────────────────────────────────────────
        ttk.Label(win, text=t("gui.compare.step2_metrics"), font=("", 11, "bold")).pack(
            anchor="w", padx=10)

        period_row = ttk.Frame(win)
        period_row.pack(fill="x", padx=10, pady=4)
        ttk.Label(period_row, text=t("gui.compare.start_year")).pack(side="left")
        start_entry = ttk.Entry(period_row, textvariable=self.compare_start_year, width=6)
        start_entry.pack(side="left", padx=(2, 8))
        ttk.Label(period_row, text=t("gui.compare.end_year")).pack(side="left")
        end_entry = ttk.Entry(period_row, textvariable=self.compare_end_year, width=6)
        end_entry.pack(side="left", padx=(2, 8))
        ttk.Label(period_row, text=t("gui.compare.frequency")).pack(side="left")
        freq_combo = ttk.Combobox(period_row, textvariable=self.compare_frequency,
                                   values=["quarterly", "annual"], state="readonly", width=10)
        freq_combo.pack(side="left", padx=(2, 0))

        category_row = ttk.Frame(win)
        category_row.pack(fill="x", padx=10, pady=4)
        ttk.Label(category_row, text=t("gui.compare.metric_category")).pack(side="left")

        from ratios import RATIO_CATEGORIES, RATIO_DEFS
        all_categories = ["損益表", "資產負債表", "現金流"] + RATIO_CATEGORIES
        category_var = tk.StringVar(value=all_categories[0])
        category_combo = ttk.Combobox(category_row, textvariable=category_var,
                                       values=all_categories, state="readonly", width=14)
        category_combo.pack(side="left", padx=(4, 0))

        metric_check_frame = ttk.Frame(win)
        metric_check_frame.pack(fill="x", padx=10)
        metric_vars: dict[str, tk.BooleanVar] = {}

        def _metrics_for_category(category: str) -> list[str]:
            if category in ("損益表", "資產負債表", "現金流"):
                tag = {"損益表": "IS", "資產負債表": "BS", "現金流": "CF"}[category]
                return self._raw_concepts_for_tag(tag)
            return [name for name, _, cat, _ in RATIO_DEFS if cat == category]

        def _refresh_metric_checkboxes(*_):
            for child in metric_check_frame.winfo_children():
                child.destroy()
            names = _metrics_for_category(category_var.get())
            for idx, name in enumerate(names):
                var = metric_vars.setdefault(name, tk.BooleanVar(
                    value=name in self.compare_selected_metrics))

                def _on_toggle(name_=name, var_=var):
                    if var_.get() and name_ not in self.compare_selected_metrics:
                        self.compare_selected_metrics.append(name_)
                    elif not var_.get() and name_ in self.compare_selected_metrics:
                        self.compare_selected_metrics.remove(name_)
                    _refresh_metric_chips()

                ttk.Checkbutton(metric_check_frame, text=name, variable=var,
                                command=_on_toggle).grid(
                    row=idx // 2, column=idx % 2, sticky="w", padx=4)

        category_combo.bind("<<ComboboxSelected>>", _refresh_metric_checkboxes)

        metric_chips_frame = ttk.Frame(win)
        metric_chips_frame.pack(fill="x", padx=10, pady=(4, 0))

        def _refresh_metric_chips():
            for child in metric_chips_frame.winfo_children():
                child.destroy()
            for name in self.compare_selected_metrics:
                chip = ttk.Frame(metric_chips_frame, relief="raised", borderwidth=1)
                chip.pack(side="left", padx=2, pady=2)
                ttk.Label(chip, text=name).pack(side="left", padx=(4, 0))

                def _remove(name_=name):
                    self.compare_selected_metrics.remove(name_)
                    if name_ in metric_vars:
                        metric_vars[name_].set(False)
                    _refresh_metric_chips()

                ttk.Button(chip, text="✕", width=2, command=_remove).pack(side="left")

        _refresh_metric_checkboxes()
        _refresh_metric_chips()

        snapshot_row = ttk.Frame(win)
        snapshot_row.pack(fill="x", padx=10, pady=6)
        ttk.Label(snapshot_row, text=t("gui.compare.snapshot_date")).pack(side="left")
        ttk.Entry(snapshot_row, textvariable=self.compare_snapshot_date, width=14).pack(
            side="left", padx=(4, 0))

        btn_row = ttk.Frame(win)
        btn_row.pack(fill="x", padx=10, pady=10)
        ttk.Button(btn_row, text=t("gui.btn.cancel"), command=win.destroy).pack(
            side="right", padx=4)

        def _confirm():
            self._update_compare_summary()
            win.destroy()

        ttk.Button(btn_row, text=t("gui.btn.confirm"), command=_confirm).pack(
            side="right", padx=4)

    def _raw_concepts_for_tag(self, tag: str) -> list[str]:
        """從 fetcher_gaap 的科目定義表取某一類（IS/BS/CF）的欄位名稱清單。

        `IS_TEMPLATE`／`BS_TEMPLATE`／`CF_TEMPLATE` 是 fetcher_gaap.py 裡已經
        依報表類型分開的 module-level 清單（見 src/fetcher_gaap.py:253/282/334），
        每筆 tuple 的第 0 欄就是顯示名稱，不用再照第 4 欄的 "IS"/"BS"/"CF"
        標籤篩一次——那個標籤是給 merge 邏輯用的，這三份清單本身已經是分好的。
        """
        from fetcher_gaap import IS_TEMPLATE, BS_TEMPLATE, CF_TEMPLATE
        source = {"IS": IS_TEMPLATE, "BS": BS_TEMPLATE, "CF": CF_TEMPLATE}[tag]
        return [row[0] for row in source]
```

- [ ] **Step 5: 手動驗證**

啟動 GUI（`./venv/Scripts/python.exe src/main.py`），切到 Tab4：
1. 確認「跨公司比較」分頁出現，位置在進階設定之後
2. 點「選擇比較內容」，輸入 `nvda` 應該跳出自動完成建議，選擇後 chip 正確出現
3. 貼上 `amd, avgo` 逗號分隔，兩家公司都要正確加入 chip 列表
4. 輸入一個不存在的 ticker（如 `ZZZZZ`），應該跳出警告視窗，不會加進清單
5. 切換指標分類下拉，勾選框內容要正確換成該分類的指標，勾選狀態要跨分類保留
6. 按「確定」後，Tab4 主畫面摘要要正確顯示選擇的公司/期間/指標

若第 4 步驟卡住（`_raw_concepts_for_tag` 的 import 路徑跟實際 `fetcher_gaap.py` 不符），
先跑 `grep -n "^_IS_ROWS\|^_BS_ROWS\|^_CF_ROWS\|IS_TEMPLATE\|BS_TEMPLATE\|CF_TEMPLATE" src/fetcher_gaap.py`
找到目前實際的變數名稱再修正。

- [ ] **Step 6: Commit**

```bash
git add src/main.py
git commit -m "feat: main.py 新增 Tab4 跨公司比較選擇視窗"
```

---

## Task 10: `main.py` — Tab4 執行流程（背景執行緒、進度、錯誤處理）

**Files:**
- Modify: `src/main.py`

**Interfaces:**
- Consumes: Task 4 的 `comparison.build_comparison()`、Task 8 的
  `comparison_writer.write_comparison_workbook()`、Task 9 的 Tab4 GUI 變數

- [ ] **Step 1: 新增 `_run_comparison()` 與背景執行緒**

在 `main.py`，比照既有 `_worker_batch`（`src/main.py:2170` 附近）的 pattern，新增：

```python
    def _run_comparison(self):
        if not self.compare_selected_tickers:
            messagebox.showwarning(t("gui.compare.select_title"), t("gui.compare.no_company_warn"))
            return
        if not self.compare_selected_metrics:
            messagebox.showwarning(t("gui.compare.select_title"), t("gui.compare.no_metric_warn"))
            return
        identity = self.cfg.get("identity", "")
        if not identity:
            messagebox.showwarning(t("gui.compare.select_title"), t("gui.lbl.identity_missing"))
            return

        self.compare_run_btn.config(state="disabled")
        threading.Thread(target=self._compare_worker, daemon=True).start()

    def _compare_worker(self):
        from comparison import build_comparison
        from comparison_writer import write_comparison_workbook

        identity = self.cfg.get("identity", "")
        tickers = [t_ for t_, _ in self.compare_selected_tickers]
        metrics = list(self.compare_selected_metrics)
        start_year = int(self.compare_start_year.get()) if self.compare_start_year.get().strip() else None
        end_year = int(self.compare_end_year.get()) if self.compare_end_year.get().strip() else None
        frequency = self.compare_frequency.get()

        self.msg_queue.put(("compare_log", t("gui.compare.log_start", n=len(tickers))))
        try:
            result = build_comparison(
                tickers, identity, metrics, frequency=frequency,
                start_year=start_year, end_year=end_year,
            )
        except Exception as e:
            self.msg_queue.put(("compare_error", f"{type(e).__name__}{_exc_status(e)}"))
            return

        for failure in result.failures:
            self.msg_queue.put(("compare_log",
                                 t("gui.compare.log_company_failed",
                                   ticker=failure.ticker, error_type=failure.error_type)))

        if not any(result.metrics.get(m) for m in metrics):
            self.msg_queue.put(("compare_error", t("gui.compare.nothing_fetched")))
            return

        out_dir = Path(self.compare_outdir_var.get().strip() or str(PROJECT_ROOT / "output" / "compare"))
        names = "_".join(tickers[:3])
        filename = f"比較_{names}_{date.today().strftime('%Y%m%d')}.xlsx"
        out_path = out_dir / filename

        try:
            write_comparison_workbook(
                result, metrics, out_path, snapshot_date=self.compare_snapshot_date.get().strip()
            )
        except Exception as e:
            self.msg_queue.put(("compare_error", f"{type(e).__name__}{_exc_status(e)}"))
            return

        self.msg_queue.put(("compare_done", str(out_path)))
```

- [ ] **Step 2: 在既有的 `_poll_queue()` 訊息處理裡加上這幾個新 message type**

Read `src/main.py` 裡 `_poll_queue`（處理 `self.msg_queue` 各種 tuple 類型的地方，搜尋
`msg_queue.put(("` 找到既有 pattern，例如 `("last_output_folder", ...)`）的完整實作，
依既有 `if kind == "..."` 分支寫法加入：

```python
                elif kind == "compare_log":
                    self._compare_log(data)
                elif kind == "compare_error":
                    self._compare_log(f"錯誤：{data}", level="ERROR")
                    self.compare_run_btn.config(state="normal")
                elif kind == "compare_done":
                    self._compare_log(t("gui.compare.log_done", path=data))
                    self.compare_run_btn.config(state="normal")
```

（實際加入位置要對照 `_poll_queue` 現有的 if/elif 鏈縮排與變數名稱，這裡只示意分支內容。）

新增對應的 log 輔助方法（緊接在 `_log()` 方法之後）：

```python
    def _compare_log(self, msg: str, level: str = "INFO"):
        self.compare_log.config(state="normal")
        self.compare_log.insert("end", f"{msg}\n")
        self.compare_log.see("end")
        self.compare_log.config(state="disabled")
```

- [ ] **Step 3: 手動驗證**

1. 在 Tab4 選 2-3 家真實公司（如 `NVDA, AMD`）、勾選 3-4 個指標（含至少一個原始科目與一個
   比率）、頻率選季度、期間 2023~2024
2. 按「產生比較 Excel」，觀察 log 是否出現抓取進度，最後出現「完成」訊息
3. 打開輸出的 Excel，確認 `Compare_Data`／`Snapshot`／`Snapshot_Manual`／`Chart_*` 都存在
4. 在 `Snapshot` 的 B1 改一個日期（用 `Compare_Data` 裡實際出現過的期末結算日），確認
   下方數值正確變化
5. 故意在選擇視窗塞一個不存在或無法抓 10-Q/10-K 的 ticker 混在正常公司裡一起送出，
   確認該公司被跳過、其餘公司正常產出、log 有標記失敗原因

- [ ] **Step 4: Commit**

```bash
git add src/main.py
git commit -m "feat: main.py 新增 Tab4 跨公司比較執行流程（背景執行緒/進度/錯誤處理）"
```

---

## Task 11: 四語系 i18n 字串

**Files:**
- Modify: `src/locales/zh_tw.py`、`src/locales/zh_cn.py`、`src/locales/en.py`、`src/locales/ja.py`

**Interfaces:**
- Consumes: Task 9/10 裡所有 `t("gui.compare.*")`／`t("gui.tab.compare")`／`t("gui.btn.compare_*")` 呼叫

- [ ] **Step 1: 列出 Task 9/10 用到的所有翻譯 key**

```
gui.tab.compare
gui.btn.compare_select
gui.btn.compare_run
gui.compare.no_selection
gui.compare.companies
gui.compare.period
gui.compare.metrics
gui.compare.freq_quarterly
gui.compare.freq_annual
gui.compare.select_title
gui.compare.step1_company
gui.compare.ticker_input
gui.compare.unknown_ticker_title
gui.compare.unknown_ticker_msg
gui.compare.step2_metrics
gui.compare.start_year
gui.compare.end_year
gui.compare.frequency
gui.compare.metric_category
gui.compare.snapshot_date
gui.compare.no_company_warn
gui.compare.no_metric_warn
gui.compare.log_start
gui.compare.log_company_failed
gui.compare.nothing_fetched
gui.compare.log_done
```

- [ ] **Step 2: 加入 `src/locales/zh_tw.py`**（在既有 `"gui.tab.*"`／`"gui.btn.*"`／`"gui.log.*"`
  區塊對應位置插入，比照該檔既有的 key 排序慣例）：

```python
    "gui.tab.compare": '跨公司比較',
    "gui.btn.compare_select": '🔧 選擇比較內容...',
    "gui.btn.compare_run": '▶ 產生比較 Excel',
    "gui.compare.no_selection": '尚未選擇公司與指標',
    "gui.compare.companies": '比較公司',
    "gui.compare.period": '期間',
    "gui.compare.metrics": '已選指標',
    "gui.compare.freq_quarterly": '季度',
    "gui.compare.freq_annual": '年度',
    "gui.compare.select_title": '選擇比較內容',
    "gui.compare.step1_company": '① 選擇公司',
    "gui.compare.ticker_input": '輸入 ticker：',
    "gui.compare.unknown_ticker_title": '找不到公司',
    "gui.compare.unknown_ticker_msg": '{ticker} 不在快取清單中，請確認代號是否正確',
    "gui.compare.step2_metrics": '② 選擇比較指標',
    "gui.compare.start_year": '起始年',
    "gui.compare.end_year": '結束年',
    "gui.compare.frequency": '頻率',
    "gui.compare.metric_category": '指標分類',
    "gui.compare.snapshot_date": '快照時間點（如 2025/12/31）：',
    "gui.compare.no_company_warn": '請先選擇至少一家公司',
    "gui.compare.no_metric_warn": '請先選擇至少一個指標',
    "gui.compare.log_start": '開始抓取 {n} 家公司資料...',
    "gui.compare.log_company_failed": '[{ticker}] 抓取失敗，跳過 -> {error_type}',
    "gui.compare.nothing_fetched": '所有公司都抓取失敗，沒有可輸出的資料',
    "gui.compare.log_done": '比較 Excel 已產出：{path}',
```

- [ ] **Step 3: 加入 `src/locales/zh_cn.py`**（簡體轉譯，key 相同）：

```python
    'gui.tab.compare': '跨公司比较',
    'gui.btn.compare_select': '🔧 选择比较内容...',
    'gui.btn.compare_run': '▶ 生成比较 Excel',
    'gui.compare.no_selection': '尚未选择公司与指标',
    'gui.compare.companies': '比较公司',
    'gui.compare.period': '期间',
    'gui.compare.metrics': '已选指标',
    'gui.compare.freq_quarterly': '季度',
    'gui.compare.freq_annual': '年度',
    'gui.compare.select_title': '选择比较内容',
    'gui.compare.step1_company': '① 选择公司',
    'gui.compare.ticker_input': '输入 ticker：',
    'gui.compare.unknown_ticker_title': '找不到公司',
    'gui.compare.unknown_ticker_msg': '{ticker} 不在缓存清单中，请确认代号是否正确',
    'gui.compare.step2_metrics': '② 选择比较指标',
    'gui.compare.start_year': '起始年',
    'gui.compare.end_year': '结束年',
    'gui.compare.frequency': '频率',
    'gui.compare.metric_category': '指标分类',
    'gui.compare.snapshot_date': '快照时间点（如 2025/12/31）：',
    'gui.compare.no_company_warn': '请先选择至少一家公司',
    'gui.compare.no_metric_warn': '请先选择至少一个指标',
    'gui.compare.log_start': '开始抓取 {n} 家公司资料...',
    'gui.compare.log_company_failed': '[{ticker}] 抓取失败，跳过 -> {error_type}',
    'gui.compare.nothing_fetched': '所有公司都抓取失败，没有可输出的数据',
    'gui.compare.log_done': '比较 Excel 已产出：{path}',
```

- [ ] **Step 4: 加入 `src/locales/en.py`**：

```python
    'gui.tab.compare': 'Cross-Company Compare',
    'gui.btn.compare_select': '🔧 Select comparison...',
    'gui.btn.compare_run': '▶ Generate comparison Excel',
    'gui.compare.no_selection': 'No companies or metrics selected yet',
    'gui.compare.companies': 'Companies',
    'gui.compare.period': 'Period',
    'gui.compare.metrics': 'Selected metrics',
    'gui.compare.freq_quarterly': 'Quarterly',
    'gui.compare.freq_annual': 'Annual',
    'gui.compare.select_title': 'Select comparison content',
    'gui.compare.step1_company': '① Select companies',
    'gui.compare.ticker_input': 'Enter ticker:',
    'gui.compare.unknown_ticker_title': 'Company not found',
    'gui.compare.unknown_ticker_msg': '{ticker} is not in the cached list, please check the symbol',
    'gui.compare.step2_metrics': '② Select metrics',
    'gui.compare.start_year': 'Start year',
    'gui.compare.end_year': 'End year',
    'gui.compare.frequency': 'Frequency',
    'gui.compare.metric_category': 'Metric category',
    'gui.compare.snapshot_date': 'Snapshot date (e.g. 2025/12/31):',
    'gui.compare.no_company_warn': 'Please select at least one company',
    'gui.compare.no_metric_warn': 'Please select at least one metric',
    'gui.compare.log_start': 'Fetching data for {n} companies...',
    'gui.compare.log_company_failed': '[{ticker}] fetch failed, skipped -> {error_type}',
    'gui.compare.nothing_fetched': 'All companies failed to fetch, nothing to output',
    'gui.compare.log_done': 'Comparison Excel produced: {path}',
```

- [ ] **Step 5: 加入 `src/locales/ja.py`**：

```python
    "gui.tab.compare": '企業間比較',
    "gui.btn.compare_select": '🔧 比較内容を選択...',
    "gui.btn.compare_run": '▶ 比較 Excel を生成',
    "gui.compare.no_selection": '会社と指標が未選択です',
    "gui.compare.companies": '比較対象企業',
    "gui.compare.period": '期間',
    "gui.compare.metrics": '選択済み指標',
    "gui.compare.freq_quarterly": '四半期',
    "gui.compare.freq_annual": '年次',
    "gui.compare.select_title": '比較内容を選択',
    "gui.compare.step1_company": '① 企業を選択',
    "gui.compare.ticker_input": 'ティッカー入力：',
    "gui.compare.unknown_ticker_title": '企業が見つかりません',
    "gui.compare.unknown_ticker_msg": '{ticker} はキャッシュに存在しません。銘柄コードを確認してください',
    "gui.compare.step2_metrics": '② 比較指標を選択',
    "gui.compare.start_year": '開始年',
    "gui.compare.end_year": '終了年',
    "gui.compare.frequency": '頻度',
    "gui.compare.metric_category": '指標カテゴリ',
    "gui.compare.snapshot_date": 'スナップショット日付（例：2025/12/31）：',
    "gui.compare.no_company_warn": '少なくとも1社選択してください',
    "gui.compare.no_metric_warn": '少なくとも1つの指標を選択してください',
    "gui.compare.log_start": '{n} 社のデータを取得中...',
    "gui.compare.log_company_failed": '[{ticker}] 取得失敗、スキップ -> {error_type}',
    "gui.compare.nothing_fetched": '全社の取得に失敗しました。出力するデータがありません',
    "gui.compare.log_done": '比較 Excel を生成しました：{path}',
```

- [ ] **Step 6: 執行既有 i18n 完整性測試**

這個專案的 `docs/ARCHITECTURE.md` 提到多語言鐵則有對應的自動化檢查（四個 locale 的 key
集合要一致）。執行：

Run: `./venv/Scripts/python.exe -m pytest tests/ -k i18n -v`
Expected: 全部 PASS（四個 locale 檔的 key 集合完全一致，沒有漏翻）

如果沒有找到專門的 i18n key 一致性測試，改跑全部測試確認沒有因為字串缺失而報錯：
Run: `./venv/Scripts/python.exe -m pytest tests/ -v`

- [ ] **Step 7: Commit**

```bash
git add src/locales/zh_tw.py src/locales/zh_cn.py src/locales/en.py src/locales/ja.py
git commit -m "i18n: 新增 Tab4 跨公司比較的四語系字串"
```

---

## Task 12: 端對端驗證

**Files:** 無程式碼異動，純驗證

- [ ] **Step 1: 執行全部自動化測試**

Run: `./venv/Scripts/python.exe -m pytest tests/ -v`
Expected: 全部 PASS，數量應該是既有測試數 + 本次新增的測試數（`test_excel_formatter.py`
+7、`test_ratios.py` +2（category）+16（新比率）、`test_comparison.py` +5、
`test_comparison_writer.py` +9 左右）

- [ ] **Step 2: 拿真實股票跑一次完整流程**

啟動 GUI，Tab4 選 `NVDA, AMD, AVGO` 三家、季度、2023~2024、勾 `Revenue`、
`Gross Margin (%)`、`Debt Ratio (%)`、`ROIC (%)` 四個指標，快照日期填一個確定在範圍內
的期末結算日，執行。

檢查輸出 Excel：
1. `Compare_Data` 四個區塊都有資料，數字量級合理（Revenue 是十億美元級、Debt Ratio 是
   0~100 之間的百分比）
2. `Snapshot` 的 `B1` 改成另一個合法的期末結算日，數字要正確跟著變
3. `Chart_Revenue` 等四張圖表 sheet 都有一張折線圖，三條線分別對應三家公司，趨勢看起來
   合理（不是全部疊在 0）
4. 在 Excel 裡對任一張圖表右鍵「變更圖表類型」→ 長條圖，確認可以正常切換不出錯

- [ ] **Step 3: 確認不影響既有功能**

跑一次 Tab1 單一公司抓取（任一支股票），確認輸出 Excel 跟這次改動前一致——`ratios.py`
的 `RATIO_DEFS` 從三元組改四元組、`excel_formatter.py` 抽出 `unit_format_for()` 都是這次
異動範圍，必須确认 `Data_Ratios` sheet 的既有 28 個比率數值與格式完全沒變。

- [ ] **Step 4: 更新 TODO.md**

把 `docs/TODO.md` 的 F1 項目狀態從「暫緩開發」改成「已完成」，內容搬去
`docs/CHANGELOG.md`（比照專案既有慣例：「已完成項目一律搬去 CHANGELOG，TODO 不留殘骸」，
見 `docs/TODO.md` E 段開頭的說明）。

```bash
git add docs/TODO.md docs/CHANGELOG.md
git commit -m "docs: 跨公司比較功能完成，TODO F1 搬入 CHANGELOG"
```
