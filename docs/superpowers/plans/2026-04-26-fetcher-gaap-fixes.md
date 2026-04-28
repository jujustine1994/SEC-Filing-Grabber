# Fetcher GAAP Three-Statement Fixes Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 修正 `fetcher_gaap.py` 中八個已知的資料正確性、覆蓋率、效能與標籤問題。

**Architecture:** 所有改動集中在 `fetcher_gaap.py`（模板定義、match 邏輯、post-processing）與對應的 `tests/test_fetcher_gaap.py`。每個 Task 獨立可測試，不跨 Task 相依。Task 8（FY label）幅度最大，最後執行。

**Tech Stack:** Python 3.12+、pytest、pandas、edgartools

---

## 檔案地圖

| 檔案 | 本次異動 |
|------|---------|
| `fetcher_gaap.py` | 全部 8 個 Task 的核心修改 |
| `tests/test_fetcher_gaap.py` | 每個 Task 對應新增單元測試 |

不需新增檔案，不需動 `excel_writer.py`、`override_engine.py`、`main.py`。

---

## Task 1：pre-XBRL 早期終止（效能）

**Problem:** AMD 等老公司有數十年 pre-XBRL 申報，loop 從 1994 年逐一嘗試，大幅拖慢 fetch 時間。EDGAR 回傳的 filings 是**最新在前**，所以遇到 2008 年以前的 filing 可以直接 `break`。

**Files:**
- Modify: `fetcher_gaap.py` — `_build_is_table`、`_build_bs_table`、`_build_cf_table`、`_build_segment_tables`（四個 filing loop 開頭各加一條 break）

- [ ] **Step 1: 寫失敗測試**

在 `tests/test_fetcher_gaap.py` 最後加：

```python
# ── Task 1: pre-XBRL early exit ───────────────────────────────────────────────

from datetime import date as _date
from unittest.mock import PropertyMock

def _make_old_filing(filing_date_str: str):
    """Mock filing with given date but no XBRL (pre-2008)."""
    f = MagicMock()
    f.filing_date = _date.fromisoformat(filing_date_str)
    f.obj.side_effect = Exception("No XBRL")
    return f


def test_build_is_table_stops_before_pre_xbrl():
    """Filings dated before 2008-01-01 should cause early loop termination, not exception."""
    modern = _make_filing(period_col="2024-03-31 (Q1)", val=100.0, filing_date="2024-04-30")
    modern.filing_date = _date(2024, 4, 30)
    old = _make_old_filing("2007-04-30")

    # Should not raise even though old.obj() raises Exception
    gaap_tbl, _ = _build_is_table([modern, old], max_filings=80)
    assert len(gaap_tbl.quarter_labels) == 1   # only the modern one


def test_build_bs_table_stops_before_pre_xbrl():
    modern = _make_filing(period_col="2024-03-31 (Q1)", val=100.0, filing_date="2024-04-30")
    modern.filing_date = _date(2024, 4, 30)
    old = _make_old_filing("2007-04-30")
    gaap_tbl, _ = _build_bs_table([modern, old], max_filings=80)
    assert len(gaap_tbl.quarter_labels) == 1


def test_build_cf_table_stops_before_pre_xbrl():
    modern = _make_cf_filing("2024-03-31 (Q1)", "2024-03-31 (Q1)", 100.0, 150.0, "2024-04-30")
    modern.filing_date = _date(2024, 4, 30)
    old = _make_old_filing("2007-04-30")
    gaap_tbl, _ = _build_cf_table([modern, old], max_filings=80)
    assert len(gaap_tbl.quarter_labels) == 1
```

- [ ] **Step 2: 跑測試確認失敗**

```
pytest tests/test_fetcher_gaap.py::test_build_is_table_stops_before_pre_xbrl -v
```

Expected: `AttributeError` 或 `FAILED`（`filing_date` 未被判斷）

- [ ] **Step 3: 實作**

在 `fetcher_gaap.py` 頂部加常數（在 `META_COLS` 定義後面）：

```python
from datetime import date as _date

_XBRL_CUTOFF: _date = _date(2008, 1, 1)
```

在 `_build_is_table`、`_build_bs_table`、`_build_cf_table`、`_build_segment_tables` 的 filing loop 開頭，緊接在 `if len(periods) >= max_filings: break` 之後加：

```python
        if getattr(filing, "filing_date", None) and filing.filing_date < _XBRL_CUTOFF:
            break   # filings are newest-first; everything older is also pre-XBRL
```

（四個函數各加一次，位置完全相同）

- [ ] **Step 4: 跑測試確認通過**

```
pytest tests/test_fetcher_gaap.py::test_build_is_table_stops_before_pre_xbrl tests/test_fetcher_gaap.py::test_build_bs_table_stops_before_pre_xbrl tests/test_fetcher_gaap.py::test_build_cf_table_stops_before_pre_xbrl -v
```

Expected: 3 PASSED

- [ ] **Step 5: 跑全套 unit tests 確認沒有 regression**

```
pytest tests/test_fetcher_gaap.py -v
```

Expected: all PASSED

- [ ] **Step 6: Commit**

```
git add fetcher_gaap.py tests/test_fetcher_gaap.py
git commit -m "perf: stop IS/BS/CF/seg loops at pre-XBRL cutoff (2008)"
```

---

## Task 2：Dividends std_concept Bug Fix

**Problem:** `CF_TEMPLATE` 裡 Dividends Paid 的 `std_concept` 設為 `"DistributionsToMinorityInterests"`（NCI 分配），優先級最高，若公司有此 XBRL tag 就會抓到錯誤行。正確的 fallback concept 是 `PaymentsOfDividends`。

**Files:**
- Modify: `fetcher_gaap.py` — `CF_TEMPLATE` 第 19 行（Dividends Paid）

- [ ] **Step 1: 寫失敗測試**

```python
# ── Task 2: Dividends std_concept bug ─────────────────────────────────────────

def _make_cf_dividends_df():
    """CF df with NCI distribution row AND a real dividends row."""
    return pd.DataFrame({
        "concept":               [
            "us-gaap_NetCashProvidedByUsedInOperatingActivities",
            "us-gaap_DistributionsToMinorityInterests",   # NCI — must NOT be picked
            "us-gaap_PaymentsOfDividendsCommonStock",     # real dividends — must be picked
        ],
        "label":                 ["Net cash from ops", "Distributions to NCI", "Dividends paid"],
        "standard_concept":      ["NetCashFromOperatingActivities", "DistributionsToMinorityInterests", None],
        "abstract":              [False, False, False],
        "is_breakdown":          [False, False, False],
        "level":                 [3, 4, 4],
        "dimension_member_label":[None, None, None],
        "2024-03-31 (Q1)":       [500.0, 30.0, 80.0],
    })


def test_dividends_paid_does_not_pick_nci_distribution():
    """Dividends Paid must pick PaymentsOfDividends, not DistributionsToMinorityInterests."""
    df = _make_cf_dividends_df()
    mock_is = MagicMock(); mock_is.to_dataframe.return_value = _make_is_df_minimal("2024-03-31 (Q1)")
    mock_cf = MagicMock(); mock_cf.to_dataframe.return_value = df
    mock_fin = MagicMock()
    mock_fin.income_statement.return_value = mock_is
    mock_fin.cashflow_statement.return_value = mock_cf
    mock_tenq = MagicMock(); mock_tenq.financials = mock_fin
    filing = MagicMock(); filing.obj.return_value = mock_tenq; filing.filing_date = "2024-04-30"

    gaap_tbl, _ = _build_cf_table([filing], max_filings=1)
    div_idx = gaap_tbl.concepts.index("Dividends Paid")
    assert gaap_tbl.values[div_idx][0] == pytest.approx(80.0), (
        f"Expected 80.0 (real dividends), got {gaap_tbl.values[div_idx][0]}"
    )
```

- [ ] **Step 2: 跑測試確認失敗**

```
pytest tests/test_fetcher_gaap.py::test_dividends_paid_does_not_pick_nci_distribution -v
```

Expected: FAILED（抓到 30.0 而非 80.0）

- [ ] **Step 3: 實作**

在 `fetcher_gaap.py` 的 `CF_TEMPLATE` 找到：

```python
    ("Dividends Paid",             "DistributionsToMinorityInterests",   "PaymentsOfDividends",                                   "CF", "first", "dividend"),
```

改為：

```python
    ("Dividends Paid",             None,                                  "PaymentsOfDividends|PaymentsOfDividendsCommonStock|PaymentsOfOrdinaryDividends", "CF", "first", "dividend"),
```

- [ ] **Step 4: 跑測試確認通過**

```
pytest tests/test_fetcher_gaap.py::test_dividends_paid_does_not_pick_nci_distribution -v
```

Expected: PASSED

- [ ] **Step 5: 跑全套 unit tests**

```
pytest tests/test_fetcher_gaap.py -v
```

Expected: all PASSED

- [ ] **Step 6: Commit**

```
git add fetcher_gaap.py tests/test_fetcher_gaap.py
git commit -m "fix: remove wrong DistributionsToMinorityInterests from Dividends Paid template"
```

---

## Task 3：Net Income fallback chain（加入 NetIncomeLossAttributableToParent）

**Problem:** 當 `std_concept == "NetIncome"` 不命中時，post-processing 直接 fallback 到 `ProfitLoss`（含 NCI 的合併淨利），語義與「歸母淨利」不同。應先試 `NetIncomeLossAttributableToParent`（parent-only），再 fallback 到 `ProfitLoss`。

**Files:**
- Modify: `fetcher_gaap.py` — `_build_is_table` post-processing（Net Income fallback 段）、`IS_TEMPLATE` 第 16 行

- [ ] **Step 1: 寫失敗測試**

```python
# ── Task 3: Net Income fallback chain ─────────────────────────────────────────

def test_build_is_table_prefers_attributable_to_parent_over_profitloss():
    """NetIncomeLossAttributableToParent should be picked before ProfitLoss."""
    df = pd.DataFrame({
        "concept":               ["us-gaap_ProfitLoss",                       "us-gaap_NetIncomeLossAttributableToParent"],
        "label":                 ["Net income incl. NCI",                      "Net income attributable to common"],
        "standard_concept":      ["ProfitLoss",                                "NetIncomeLossAttributableToParent"],
        "abstract":              [False,                                        False],
        "is_breakdown":          [False,                                        False],
        "level":                 [3,                                            3],
        "dimension_member_label":[None,                                         None],
        "2024-03-31 (Q1)":       [300.0,                                        280.0],   # 300 = 280 parent + 20 NCI
        "2023-03-31 (Q1)":       [270.0,                                        255.0],
    })
    mock_stmt = MagicMock(); mock_stmt.to_dataframe.return_value = df
    mock_fin = MagicMock()
    mock_fin.income_statement.return_value = mock_stmt
    mock_fin.cashflow_statement.return_value = mock_stmt
    mock_tenq = MagicMock(); mock_tenq.financials = mock_fin
    filing = MagicMock(); filing.obj.return_value = mock_tenq; filing.filing_date = "2024-04-30"

    gaap_tbl, _ = _build_is_table([filing], max_filings=1)
    ni_idx = gaap_tbl.concepts.index("Net Income")
    assert gaap_tbl.values[ni_idx][0] == pytest.approx(280.0), (
        f"Expected 280.0 (parent-only), got {gaap_tbl.values[ni_idx][0]}"
    )
```

- [ ] **Step 2: 跑測試確認失敗**

```
pytest tests/test_fetcher_gaap.py::test_build_is_table_prefers_attributable_to_parent_over_profitloss -v
```

Expected: FAILED（目前抓到 300.0）

- [ ] **Step 3: 實作**

在 `fetcher_gaap.py` 的 `IS_TEMPLATE` 找到：

```python
    ("Net Income",                 "NetIncome",                      "NetIncomeLoss",                                          "IS", "first", None),
```

改為（fallback_suffix 加上 parent-only 的 XBRL concept）：

```python
    ("Net Income",                 "NetIncome",                      "NetIncomeLoss|NetIncomeLossAttributableToParent",         "IS", "first", None),
```

在 `_build_is_table` 找到 ProfitLoss fallback 段：

```python
        # 2. Net Income: ProfitLoss fallback (BA, TSLA, XOM, WMT)
        if row_vals.get(_NET_INCOME_IDX) is None:
            idx = _match_is_row(df, "ProfitLoss", "ProfitLoss")
```

在 ProfitLoss 嘗試之前插入：

```python
        # 2. Net Income fallback chain
        if row_vals.get(_NET_INCOME_IDX) is None:
            # 2a. Parent-only net income (more precise than ProfitLoss)
            idx = _match_is_row(df, "NetIncomeLossAttributableToParent",
                                 "NetIncomeLossAttributableToParent")
            if idx is not None:
                consumed.add(idx)
                row_vals[_NET_INCOME_IDX] = _to_python_val(df.loc[idx, q_col])
                if _NET_INCOME_IDX not in row_labels:
                    row_labels[_NET_INCOME_IDX] = unicodedata.normalize(
                        "NFKC", str(df.loc[idx, "label"] or ""))

        if row_vals.get(_NET_INCOME_IDX) is None:
            # 2b. ProfitLoss last resort (includes NCI — use only when parent-only unavailable)
            idx = _match_is_row(df, "ProfitLoss", "ProfitLoss")
            if idx is not None:
                consumed.add(idx)
                row_vals[_NET_INCOME_IDX] = _to_python_val(df.loc[idx, q_col])
                if _NET_INCOME_IDX not in row_labels:
                    row_labels[_NET_INCOME_IDX] = unicodedata.normalize(
                        "NFKC", str(df.loc[idx, "label"] or ""))
```

（同時刪除原本合在一起的 ProfitLoss 段，改為上面的兩段）

- [ ] **Step 4: 跑測試確認通過**

```
pytest tests/test_fetcher_gaap.py::test_build_is_table_prefers_attributable_to_parent_over_profitloss tests/test_fetcher_gaap.py::test_build_is_table_net_income_profitloss_fallback -v
```

Expected: 2 PASSED（原 ProfitLoss 測試也要繼續通過）

- [ ] **Step 5: 跑全套**

```
pytest tests/test_fetcher_gaap.py -v
```

- [ ] **Step 6: Commit**

```
git add fetcher_gaap.py tests/test_fetcher_gaap.py
git commit -m "fix: prefer NetIncomeLossAttributableToParent over ProfitLoss for Net Income fallback"
```

---

## Task 4：Revenue fallback 擴展

**Problem:** `IS_TEMPLATE` Revenue 的 `fallback_suffix` 只涵蓋 `RevenueFromContractWithCustomer` 系列，不涵蓋 `us-gaap_Revenues`（許多舊公司）和 `SalesRevenueNet`（零售類）。

**Files:**
- Modify: `fetcher_gaap.py` — `IS_TEMPLATE` 第 1 行（Revenue）

- [ ] **Step 1: 寫失敗測試**

```python
# ── Task 4: Revenue fallback expansion ────────────────────────────────────────

def _make_is_df_revenues_only(period_col="2024-03-31 (Q1)", val=1000.0):
    """IS df where Revenue uses us-gaap_Revenues (not RevenueFromContractWithCustomer)."""
    return pd.DataFrame({
        "concept":               ["us-gaap_Revenues", "us-gaap_NetIncomeLoss"],
        "label":                 ["Revenues",         "Net income"],
        "standard_concept":      ["TotalRevenues",    "NetIncome"],   # not "Revenue"
        "abstract":              [False, False],
        "is_breakdown":          [False, False],
        "level":                 [3, 3],
        "dimension_member_label":[None, None],
        period_col:              [val, val * 0.1],
        "2023-03-31 (Q1)":       [val * 0.9, val * 0.08],
    })


def _make_filing_revenues_only(**kwargs):
    df = _make_is_df_revenues_only(**kwargs)
    mock_stmt = MagicMock(); mock_stmt.to_dataframe.return_value = df
    mock_fin = MagicMock()
    mock_fin.income_statement.return_value = mock_stmt
    mock_fin.cashflow_statement.return_value = mock_stmt
    mock_tenq = MagicMock(); mock_tenq.financials = mock_fin
    filing = MagicMock(); filing.obj.return_value = mock_tenq; filing.filing_date = "2024-04-30"
    return filing


def test_revenue_fallback_picks_us_gaap_revenues():
    """Revenue must resolve when XBRL uses us-gaap_Revenues (concept ends with _Revenues)."""
    filing = _make_filing_revenues_only(val=2000.0)
    gaap_tbl, _ = _build_is_table([filing], max_filings=1)
    rev_idx = gaap_tbl.concepts.index("Revenue")
    assert gaap_tbl.values[rev_idx][0] == pytest.approx(2000.0), (
        f"Expected 2000.0, got {gaap_tbl.values[rev_idx][0]}"
    )


def test_revenue_fallback_picks_sales_revenue_net():
    """Revenue must resolve when XBRL uses us-gaap_SalesRevenueNet."""
    df = pd.DataFrame({
        "concept":               ["us-gaap_SalesRevenueNet", "us-gaap_NetIncomeLoss"],
        "label":                 ["Net sales",               "Net income"],
        "standard_concept":      ["SalesRevenueNet",          "NetIncome"],
        "abstract":              [False, False],
        "is_breakdown":          [False, False],
        "level":                 [3, 3],
        "dimension_member_label":[None, None],
        "2024-03-31 (Q1)":       [3000.0, 300.0],
        "2023-03-31 (Q1)":       [2700.0, 270.0],
    })
    mock_stmt = MagicMock(); mock_stmt.to_dataframe.return_value = df
    mock_fin = MagicMock()
    mock_fin.income_statement.return_value = mock_stmt
    mock_fin.cashflow_statement.return_value = mock_stmt
    mock_tenq = MagicMock(); mock_tenq.financials = mock_fin
    filing = MagicMock(); filing.obj.return_value = mock_tenq; filing.filing_date = "2024-04-30"

    gaap_tbl, _ = _build_is_table([filing], max_filings=1)
    rev_idx = gaap_tbl.concepts.index("Revenue")
    assert gaap_tbl.values[rev_idx][0] == pytest.approx(3000.0)
```

- [ ] **Step 2: 跑測試確認失敗**

```
pytest tests/test_fetcher_gaap.py::test_revenue_fallback_picks_us_gaap_revenues tests/test_fetcher_gaap.py::test_revenue_fallback_picks_sales_revenue_net -v
```

Expected: 2 FAILED

- [ ] **Step 3: 實作**

在 `fetcher_gaap.py` 的 `IS_TEMPLATE` 找到：

```python
    ("Revenue",                    "Revenue",                        "RevenueFromContractWithCustomer",                        "IS", "first", None),
```

改為：

```python
    ("Revenue",                    "Revenue",                        r"RevenueFromContractWithCustomer|SalesRevenueNet|SalesRevenueGoodsNet|_Revenues$|^Revenues$", "IS", "first", None),
```

- [ ] **Step 4: 跑測試確認通過**

```
pytest tests/test_fetcher_gaap.py::test_revenue_fallback_picks_us_gaap_revenues tests/test_fetcher_gaap.py::test_revenue_fallback_picks_sales_revenue_net -v
```

Expected: 2 PASSED

- [ ] **Step 5: 跑全套**

```
pytest tests/test_fetcher_gaap.py -v
```

- [ ] **Step 6: Commit**

```
git add fetcher_gaap.py tests/test_fetcher_gaap.py
git commit -m "fix: expand Revenue fallback to cover Revenues and SalesRevenueNet concepts"
```

---

## Task 5：Total Non-op Derived Guard（排除 Discontinued Operations 干擾）

**Problem:** `Total Non-op = Pre-tax − Operating Income` 是 derived 公式。若公司有 discontinued operations income，這個差值會包含 discont. ops，讓 Non-op 虛增。應在偵測到 discont. ops 行時，放棄 derived，回傳 None。

**Files:**
- Modify: `fetcher_gaap.py` — `_build_is_table` post-processing（Total Non-op derived 段）

- [ ] **Step 1: 寫失敗測試**

```python
# ── Task 5: Total Non-op derived guard ────────────────────────────────────────

def test_total_nonop_not_derived_when_discontinued_ops_present():
    """When discontinued operations exist, Total Non-op should NOT be derived (return None)."""
    df = pd.DataFrame({
        "concept":               [
            "us-gaap_OperatingIncomeLoss",
            "us-gaap_IncomeLossFromDiscontinuedOperationsNetOfTax",
            "us-gaap_IncomeLossFromContinuingOperationsBeforeIncomeTax",
        ],
        "label":                 ["Operating income", "Discont. ops income", "Income before taxes"],
        "standard_concept":      ["OperatingIncomeLoss", None, "PretaxIncomeLoss"],
        "abstract":              [False, False, False],
        "is_breakdown":          [False, False, False],
        "level":                 [3, 3, 3],
        "dimension_member_label":[None, None, None],
        "2024-03-31 (Q1)":       [100.0, 20.0, 120.0],   # Non-op = 0, Discont = 20
        "2023-03-31 (Q1)":       [90.0, 15.0, 105.0],
    })
    mock_stmt = MagicMock(); mock_stmt.to_dataframe.return_value = df
    mock_fin = MagicMock()
    mock_fin.income_statement.return_value = mock_stmt
    mock_fin.cashflow_statement.return_value = mock_stmt
    mock_tenq = MagicMock(); mock_tenq.financials = mock_fin
    filing = MagicMock(); filing.obj.return_value = mock_tenq; filing.filing_date = "2024-04-30"

    gaap_tbl, _ = _build_is_table([filing], max_filings=1)
    nonop_idx = gaap_tbl.concepts.index("Total Non-op Income/(Loss)")
    assert gaap_tbl.values[nonop_idx][0] is None, (
        f"Expected None (discontinued ops present), got {gaap_tbl.values[nonop_idx][0]}"
    )


def test_total_nonop_derived_normally_without_discontinued_ops():
    """When no discontinued ops, Total Non-op should still be derived as Pretax - Operating."""
    df = pd.DataFrame({
        "concept":               ["us-gaap_OperatingIncomeLoss", "us-gaap_IncomeLossFromContinuingOperationsBeforeIncomeTax"],
        "label":                 ["Operating income",            "Income before taxes"],
        "standard_concept":      ["OperatingIncomeLoss",         "PretaxIncomeLoss"],
        "abstract":              [False, False],
        "is_breakdown":          [False, False],
        "level":                 [3, 3],
        "dimension_member_label":[None, None],
        "2024-03-31 (Q1)":       [100.0, 115.0],
        "2023-03-31 (Q1)":       [90.0, 103.0],
    })
    mock_stmt = MagicMock(); mock_stmt.to_dataframe.return_value = df
    mock_fin = MagicMock()
    mock_fin.income_statement.return_value = mock_stmt
    mock_fin.cashflow_statement.return_value = mock_stmt
    mock_tenq = MagicMock(); mock_tenq.financials = mock_fin
    filing = MagicMock(); filing.obj.return_value = mock_tenq; filing.filing_date = "2024-04-30"

    gaap_tbl, _ = _build_is_table([filing], max_filings=1)
    nonop_idx = gaap_tbl.concepts.index("Total Non-op Income/(Loss)")
    assert gaap_tbl.values[nonop_idx][0] == pytest.approx(15.0)  # 115 - 100
```

- [ ] **Step 2: 跑測試確認失敗**

```
pytest tests/test_fetcher_gaap.py::test_total_nonop_not_derived_when_discontinued_ops_present -v
```

Expected: FAILED（目前還是 derived 出 20.0）

- [ ] **Step 3: 實作**

在 `_build_is_table` 找到 Total Non-op derived 段（post-processing 1.）：

```python
        # 1. Total Non-op: DERIVED = Pre-tax − Operating Income
        if row_vals.get(_NONOP_TOTAL_IDX) is None:
            op_val     = row_vals.get(_OP_INCOME_IDX)
            pretax_val = row_vals.get(_PRETAX_IDX)
            if op_val is not None and pretax_val is not None:
                row_vals[_NONOP_TOTAL_IDX] = pretax_val - op_val
```

改為：

```python
        # 1. Total Non-op: DERIVED = Pre-tax − Operating Income
        #    Guard: skip if discontinued operations present (would distort the difference)
        if row_vals.get(_NONOP_TOTAL_IDX) is None:
            op_val     = row_vals.get(_OP_INCOME_IDX)
            pretax_val = row_vals.get(_PRETAX_IDX)
            has_discontinued = _match_is_row(df, None, "DiscontinuedOperations") is not None
            if op_val is not None and pretax_val is not None and not has_discontinued:
                row_vals[_NONOP_TOTAL_IDX] = pretax_val - op_val
```

- [ ] **Step 4: 跑測試確認通過**

```
pytest tests/test_fetcher_gaap.py::test_total_nonop_not_derived_when_discontinued_ops_present tests/test_fetcher_gaap.py::test_total_nonop_derived_normally_without_discontinued_ops tests/test_fetcher_gaap.py::test_build_is_table_total_nonop_derived_from_pretax_minus_operating -v
```

Expected: 3 PASSED

- [ ] **Step 5: 跑全套**

```
pytest tests/test_fetcher_gaap.py -v
```

- [ ] **Step 6: Commit**

```
git add fetcher_gaap.py tests/test_fetcher_gaap.py
git commit -m "fix: skip Total Non-op derivation when discontinued operations present"
```

---

## Task 6：Investment Proceeds 多概念加總

**Problem:** `Investment Proceeds` 沒有單一 XBRL 加總行，取 `first` match 不可靠。應改為對所有相關 proceeds 概念加總。

**Files:**
- Modify: `fetcher_gaap.py` — 新增 `_sum_matching_rows` helper、新增 CF index constants、`_build_cf_table` post-processing

- [ ] **Step 1: 寫失敗測試**

```python
# ── Task 6: Investment Proceeds multi-sum ─────────────────────────────────────

def _make_cf_df_investment_proceeds():
    """CF df with multiple investment proceeds lines (no single sum row)."""
    return pd.DataFrame({
        "concept":               [
            "us-gaap_NetCashProvidedByUsedInOperatingActivities",
            "us-gaap_ProceedsFromSaleOfAvailableForSaleSecurities",
            "us-gaap_ProceedsFromMaturitiesPrepaymentsAndCallsOfAvailableForSaleSecurities",
            "us-gaap_ProceedsFromSaleOfShortTermInvestments",
        ],
        "label":                 [
            "Net cash from ops",
            "Proceeds from sale of AFS securities",
            "Proceeds from maturities of AFS securities",
            "Proceeds from sale of ST investments",
        ],
        "standard_concept":      ["NetCashFromOperatingActivities", None, None, None],
        "abstract":              [False, False, False, False],
        "is_breakdown":          [False, False, False, False],
        "level":                 [3, 4, 4, 4],
        "dimension_member_label":[None, None, None, None],
        "2024-03-31 (Q1)":       [500.0, 200.0, 150.0, 100.0],
    })


def test_investment_proceeds_sums_multiple_concepts():
    """Investment Proceeds must sum all ProceedsFrom*Investment*|AFS*|ShortTerm* rows."""
    is_df = _make_is_df_minimal("2024-03-31 (Q1)")
    cf_df = _make_cf_df_investment_proceeds()
    mock_is = MagicMock(); mock_is.to_dataframe.return_value = is_df
    mock_cf = MagicMock(); mock_cf.to_dataframe.return_value = cf_df
    mock_fin = MagicMock()
    mock_fin.income_statement.return_value = mock_is
    mock_fin.cashflow_statement.return_value = mock_cf
    mock_tenq = MagicMock(); mock_tenq.financials = mock_fin
    filing = MagicMock(); filing.obj.return_value = mock_tenq; filing.filing_date = "2024-04-30"

    gaap_tbl, _ = _build_cf_table([filing], max_filings=1)
    proc_idx = gaap_tbl.concepts.index("Investment Proceeds")
    # 200 + 150 + 100 = 450
    assert gaap_tbl.values[proc_idx][0] == pytest.approx(450.0), (
        f"Expected 450.0 (sum), got {gaap_tbl.values[proc_idx][0]}"
    )
```

- [ ] **Step 2: 跑測試確認失敗**

```
pytest tests/test_fetcher_gaap.py::test_investment_proceeds_sums_multiple_concepts -v
```

Expected: FAILED（目前只取第一個 match，值為 200.0）

- [ ] **Step 3: 實作**

**3a.** 在 `fetcher_gaap.py` 加 helper（放在 `_collect_overflow` 函數之後）：

```python
def _sum_matching_rows(
    df: pd.DataFrame,
    col: str,
    patterns: list[str],
    consumed: set[int],
) -> tuple[Any, list[int]]:
    """Sum values from consolidated rows whose concept matches any pattern in patterns.

    Skips rows already in consumed. Returns (total_or_None, list_of_matched_indices).
    """
    mask = _consolidated_mask(df)
    df_c = df[mask]
    total: float | None = None
    indices: list[int] = []
    seen_concepts: set[str] = set()
    for pattern in patterns:
        matches = df_c[df_c["concept"].astype(str).str.contains(pattern, case=False, na=False, regex=True)]
        for idx, row in matches.iterrows():
            if idx in consumed:
                continue
            concept = str(row.get("concept", "") or "")
            if concept in seen_concepts:
                continue
            seen_concepts.add(concept)
            val = _to_python_val(row.get(col))
            if val is not None:
                total = (total or 0.0) + val
                indices.append(idx)
    return total, indices
```

**3b.** 在 `_CF_IDX` 常數段（`_CF_FCF_IDX` 下面）加，並同時在 module level 加 pattern 常數：

```python
_CF_INV_PURCHASES_IDX  = _CF_IDX["Investment Purchases"]
_CF_INV_PROCEEDS_IDX   = _CF_IDX["Investment Proceeds"]
_CF_DEBT_PROCEEDS_IDX   = _CF_IDX["Debt Proceeds"]
_CF_DEBT_REPAYMENTS_IDX = _CF_IDX["Debt Repayments"]

# Module-level pattern lists for multi-concept CF sums
_INV_PROCEEDS_PATTERNS: list[str] = [
    r"ProceedsFromSaleOfInvestments",
    r"ProceedsFromSaleOfAvailableForSaleSecurities",
    r"ProceedsFromMaturitiesPrepaymentsAndCallsOfAvailableForSaleSecurities",
    r"ProceedsFromSaleAndMaturityOfMarketableSecurities",
    r"ProceedsFromSaleOfShortTermInvestments",
    r"ProceedsFromSaleMaturityAndCollectionOfShorttermInvestments",
]
```

**3c.** ⚠️ 正確位置：在 `_build_cf_table` 的 filing loop 內，主 template for-loop 結束之後、overflow collection 開始之前（即 `df_c = df[_consolidated_mask(df)]` 這行之前）。這樣 `consumed` 才能在 overflow 收集前被更新，避免被 sum 的行重複進入 overflow：

```python
        # Post-processing (BEFORE overflow collection): Investment Proceeds — sum all relevant rows
        inv_proc_val, inv_proc_indices = _sum_matching_rows(df, data_col, _INV_PROCEEDS_PATTERNS, consumed)
        if inv_proc_val is not None:
            row_vals[_CF_INV_PROCEEDS_IDX] = inv_proc_val
            consumed.update(inv_proc_indices)

        # Collect raw overflow for all filings ...  ← overflow block starts here
        df_c = df[_consolidated_mask(df)]
```

- [ ] **Step 4: 跑測試確認通過**

```
pytest tests/test_fetcher_gaap.py::test_investment_proceeds_sums_multiple_concepts -v
```

Expected: PASSED

- [ ] **Step 5: 跑全套**

```
pytest tests/test_fetcher_gaap.py -v
```

- [ ] **Step 6: Commit**

```
git add fetcher_gaap.py tests/test_fetcher_gaap.py
git commit -m "fix: sum multiple Investment Proceeds XBRL concepts instead of first-match"
```

---

## Task 7：Debt Proceeds / Repayments 多概念加總

**Problem:** 公司常按 LT/ST/信用額度分開申報 debt proceeds 和 repayments，沒有單一合計行。當前只取 `first` match，漏掉其他分項。

**Files:**
- Modify: `fetcher_gaap.py` — 新增 CF index constants、`_build_cf_table` post-processing（延伸 Task 6 的 `_sum_matching_rows`）

- [ ] **Step 1: 寫失敗測試**

```python
# ── Task 7: Debt Proceeds/Repayments multi-sum ────────────────────────────────

def _make_cf_df_debt_lines():
    """CF df with separate LT and ST debt proceeds/repayments, no summary row."""
    return pd.DataFrame({
        "concept":               [
            "us-gaap_NetCashProvidedByUsedInOperatingActivities",
            "us-gaap_ProceedsFromIssuanceOfLongTermDebt",
            "us-gaap_ProceedsFromShortTermBorrowings",
            "us-gaap_RepaymentsOfLongTermDebt",
            "us-gaap_RepaymentsOfShortTermDebt",
        ],
        "label":                 [
            "Net cash from ops",
            "Proceeds from LT debt",
            "Proceeds from ST borrowings",
            "Repayments of LT debt",
            "Repayments of ST debt",
        ],
        "standard_concept":      ["NetCashFromOperatingActivities", None, None, None, None],
        "abstract":              [False, False, False, False, False],
        "is_breakdown":          [False, False, False, False, False],
        "level":                 [3, 4, 4, 4, 4],
        "dimension_member_label":[None, None, None, None, None],
        "2024-03-31 (Q1)":       [500.0, 1000.0, 200.0, 800.0, 100.0],
    })


def test_debt_proceeds_sums_lt_and_st():
    """Debt Proceeds must sum LT + ST debt issuance proceeds."""
    is_df = _make_is_df_minimal("2024-03-31 (Q1)")
    cf_df = _make_cf_df_debt_lines()
    mock_is = MagicMock(); mock_is.to_dataframe.return_value = is_df
    mock_cf = MagicMock(); mock_cf.to_dataframe.return_value = cf_df
    mock_fin = MagicMock()
    mock_fin.income_statement.return_value = mock_is
    mock_fin.cashflow_statement.return_value = mock_cf
    mock_tenq = MagicMock(); mock_tenq.financials = mock_fin
    filing = MagicMock(); filing.obj.return_value = mock_tenq; filing.filing_date = "2024-04-30"

    gaap_tbl, _ = _build_cf_table([filing], max_filings=1)
    proc_idx = gaap_tbl.concepts.index("Debt Proceeds")
    assert gaap_tbl.values[proc_idx][0] == pytest.approx(1200.0), (  # 1000 + 200
        f"Expected 1200.0, got {gaap_tbl.values[proc_idx][0]}"
    )


def test_debt_repayments_sums_lt_and_st():
    """Debt Repayments must sum LT + ST repayments."""
    is_df = _make_is_df_minimal("2024-03-31 (Q1)")
    cf_df = _make_cf_df_debt_lines()
    mock_is = MagicMock(); mock_is.to_dataframe.return_value = is_df
    mock_cf = MagicMock(); mock_cf.to_dataframe.return_value = cf_df
    mock_fin = MagicMock()
    mock_fin.income_statement.return_value = mock_is
    mock_fin.cashflow_statement.return_value = mock_cf
    mock_tenq = MagicMock(); mock_tenq.financials = mock_fin
    filing = MagicMock(); filing.obj.return_value = mock_tenq; filing.filing_date = "2024-04-30"

    gaap_tbl, _ = _build_cf_table([filing], max_filings=1)
    rep_idx = gaap_tbl.concepts.index("Debt Repayments")
    assert gaap_tbl.values[rep_idx][0] == pytest.approx(900.0), (  # 800 + 100
        f"Expected 900.0, got {gaap_tbl.values[rep_idx][0]}"
    )
```

- [ ] **Step 2: 跑測試確認失敗**

```
pytest tests/test_fetcher_gaap.py::test_debt_proceeds_sums_lt_and_st tests/test_fetcher_gaap.py::test_debt_repayments_sums_lt_and_st -v
```

Expected: 2 FAILED

- [ ] **Step 3: 實作**

**3a.** `_CF_DEBT_PROCEEDS_IDX`、`_CF_DEBT_REPAYMENTS_IDX` 已在 Task 6 Step 3b 一併定義，不需重複加。在同一段 module-level 常數後面補：

```python
_DEBT_PROCEEDS_PATTERNS: list[str] = [
    r"ProceedsFromIssuanceOfDebt$",
    r"ProceedsFromIssuanceOfLongTermDebt",
    r"ProceedsFromShortTermBorrowings",
    r"ProceedsFromLinesOfCredit",
    r"ProceedsFromIssuanceOfMediumTermNotes",
    r"ProceedsFromIssuanceOfSeniorLongTermDebt",
]
_DEBT_REPAYMENTS_PATTERNS: list[str] = [
    r"RepaymentsOfDebt$",
    r"RepaymentsOfLongTermDebt",
    r"RepaymentsOfShortTermDebt",
    r"RepaymentsOfLinesOfCredit",
    r"RepaymentsOfMediumTermNotes",
    r"RepaymentsOfSeniorDebt",
]
```

**3b.** ⚠️ 正確位置：與 Task 6 的 Investment Proceeds 緊接在一起（同樣在 overflow collection 之前）：

```python
        # Post-processing (BEFORE overflow collection): Investment Proceeds
        inv_proc_val, inv_proc_indices = _sum_matching_rows(df, data_col, _INV_PROCEEDS_PATTERNS, consumed)
        if inv_proc_val is not None:
            row_vals[_CF_INV_PROCEEDS_IDX] = inv_proc_val
            consumed.update(inv_proc_indices)

        # Post-processing (BEFORE overflow collection): Debt Proceeds
        debt_proc_val, debt_proc_indices = _sum_matching_rows(df, data_col, _DEBT_PROCEEDS_PATTERNS, consumed)
        if debt_proc_val is not None:
            row_vals[_CF_DEBT_PROCEEDS_IDX] = debt_proc_val
            consumed.update(debt_proc_indices)

        # Post-processing (BEFORE overflow collection): Debt Repayments
        debt_rep_val, debt_rep_indices = _sum_matching_rows(df, data_col, _DEBT_REPAYMENTS_PATTERNS, consumed)
        if debt_rep_val is not None:
            row_vals[_CF_DEBT_REPAYMENTS_IDX] = debt_rep_val
            consumed.update(debt_rep_indices)

        # Collect raw overflow for all filings ...  ← overflow block starts here
        df_c = df[_consolidated_mask(df)]
```

- [ ] **Step 4: 跑測試確認通過**

```
pytest tests/test_fetcher_gaap.py::test_debt_proceeds_sums_lt_and_st tests/test_fetcher_gaap.py::test_debt_repayments_sums_lt_and_st -v
```

Expected: 2 PASSED

- [ ] **Step 5: 跑全套**

```
pytest tests/test_fetcher_gaap.py -v
```

- [ ] **Step 6: Commit**

```
git add fetcher_gaap.py tests/test_fetcher_gaap.py
git commit -m "fix: sum LT+ST debt proceeds and repayments instead of first-match"
```

---

## Task 8：FY Label 對齊公司財年

**Problem:** `_col_to_quarter_label` 用 period-end 曆年標 FY，導致 AAPL（Sep FY end）的 Q1 FY2024（Oct-Dec 2023）被標成 `FY2023Q1`，與官方說法差一個財年。Microsoft（Jun）也有相同問題。

**修法：** 新增 `_detect_fy_end_month` 從 10-K filing 偵測財年結束月份；`_col_to_quarter_label` 加 `fy_end_month` 參數（預設 12，向後相容）；所有 build 函數透傳此參數。

**Files:**
- Modify: `fetcher_gaap.py` — `_col_to_quarter_label`、新增 `_detect_fy_end_month`、`_build_is_table`/`_build_bs_table`/`_build_cf_table`/`_build_segment_tables`/`_build_dynamic_table`（全加 `fy_end_month` 參數）、`fetch_gaap_statements`（偵測並透傳）
- Modify: `tests/test_fetcher_gaap.py` — 新增非十二月 FY 測試；已有的函數呼叫不須改動（預設 12 = 現行行為）

- [ ] **Step 1: 寫失敗測試**

```python
# ── Task 8: FY label fiscal year alignment ────────────────────────────────────

def test_col_to_quarter_label_default_december_fy_unchanged():
    """Default fy_end_month=12 must produce identical output to current behaviour."""
    assert _col_to_quarter_label("2023-12-30 (Q1)") == "FY2023Q1"
    assert _col_to_quarter_label("2023-09-30 (FY)") == "FY2023"


def test_col_to_quarter_label_sep_fy_q1_increments_year():
    """AAPL-style Sep FY: Q1 ends Dec 2023 → label FY2024Q1 (company's FY2024)."""
    assert _col_to_quarter_label("2023-12-30 (Q1)", fy_end_month=9) == "FY2024Q1"


def test_col_to_quarter_label_sep_fy_q2_same_year():
    """AAPL Q2 ends Mar 2024 → label FY2024Q2 (same FY2024)."""
    assert _col_to_quarter_label("2024-03-30 (Q2)", fy_end_month=9) == "FY2024Q2"


def test_col_to_quarter_label_sep_fy_q3_same_year():
    """AAPL Q3 ends Jun 2024 → label FY2024Q3 (same FY2024)."""
    assert _col_to_quarter_label("2024-06-29 (Q3)", fy_end_month=9) == "FY2024Q3"


def test_col_to_quarter_label_sep_fy_annual_unchanged():
    """AAPL annual FY ends Sep 2024 → label FY2024 (unchanged, period is FY not Q)."""
    assert _col_to_quarter_label("2024-09-28 (FY)", fy_end_month=9) == "FY2024"


def test_col_to_quarter_label_jun_fy_q1_increments_year():
    """MSFT-style Jun FY: Q1 ends Sep 2024 → label FY2025Q1."""
    assert _col_to_quarter_label("2024-09-30 (Q1)", fy_end_month=6) == "FY2025Q1"


def test_col_to_quarter_label_jun_fy_q2_increments_year():
    """MSFT Q2 ends Dec 2024 → label FY2025Q2."""
    assert _col_to_quarter_label("2024-12-31 (Q2)", fy_end_month=6) == "FY2025Q2"


def test_col_to_quarter_label_jun_fy_q3_same_year():
    """MSFT Q3 ends Mar 2025 → label FY2025Q3."""
    assert _col_to_quarter_label("2025-03-31 (Q3)", fy_end_month=6) == "FY2025Q3"


def test_detect_fy_end_month_returns_9_for_sep_end():
    """_detect_fy_end_month should return 9 when the 10-K FY column ends in September."""
    df = pd.DataFrame({
        "concept": ["us-gaap_RevenueFromContractWithCustomer"],
        "label":   ["Revenue"],
        "standard_concept": ["Revenue"],
        "abstract": [False], "is_breakdown": [False], "level": [3],
        "dimension_member_label": [None],
        "2024-09-28 (FY)": [1000.0],
    })
    from fetcher_gaap import _detect_fy_end_month
    mock_stmt = MagicMock(); mock_stmt.to_dataframe.return_value = df
    mock_fin = MagicMock(); mock_fin.income_statement.return_value = mock_stmt
    mock_tenq = MagicMock(); mock_tenq.financials = mock_fin
    filing = MagicMock(); filing.obj.return_value = mock_tenq
    assert _detect_fy_end_month([filing]) == 9


def test_detect_fy_end_month_defaults_to_12_on_failure():
    """_detect_fy_end_month should return 12 when no 10-K filings given."""
    from fetcher_gaap import _detect_fy_end_month
    assert _detect_fy_end_month([]) == 12
```

- [ ] **Step 2: 跑測試確認失敗**

```
pytest tests/test_fetcher_gaap.py::test_col_to_quarter_label_sep_fy_q1_increments_year tests/test_fetcher_gaap.py::test_detect_fy_end_month_returns_9_for_sep_end -v
```

Expected: FAILED（`_col_to_quarter_label` 沒有 `fy_end_month` 參數，`_detect_fy_end_month` 不存在）

- [ ] **Step 3: 實作 `_col_to_quarter_label` 與 `_detect_fy_end_month`**

在 `fetcher_gaap.py` 找到 `_col_to_quarter_label`，改為：

```python
def _col_to_quarter_label(col_name: str, fy_end_month: int = 12) -> str:
    """Convert edgartools period column name to FY label.

    fy_end_month: company's fiscal year end month (1-12). Default 12 = calendar year.
    For non-December FY companies, Q labels in months after the FY end belong to the
    next fiscal year (e.g. AAPL Sep FY: Dec 2023 Q1 → FY2024Q1).

    Examples (default fy_end_month=12):
        "2023-03-31 (Q1)"  -> "FY2023Q1"
        "2024-12-31 (FY)"  -> "FY2024"
    Examples (fy_end_month=9, AAPL):
        "2023-12-30 (Q1)"  -> "FY2024Q1"
        "2024-09-28 (FY)"  -> "FY2024"   (annual label unchanged)
    """
    m = re.match(r"(\d{4})-(\d{2})-\d{2}\s+\((\w+)\)", col_name.strip())
    if m:
        year, month, period = int(m.group(1)), int(m.group(2)), m.group(3)
        if period.upper() == "FY":
            return f"FY{year}"
        if fy_end_month < 12 and month > fy_end_month:
            year += 1
        return f"FY{year}{period}"
    return col_name
```

在 `_col_to_quarter_label` 下面加新函數：

```python
def _detect_fy_end_month(filings_k: list) -> int:
    """Detect company's fiscal year end month from 10-K filings.

    Looks for a column labeled '(FY)' in the IS statement of the first 3 10-K filings.
    Returns the month number (1-12), defaulting to 12 (December) if not detected.
    """
    for filing in filings_k[:3]:
        try:
            tenq = filing.obj()
            is_stmt = tenq.financials.income_statement()
            if is_stmt is None:
                continue
            df = is_stmt.to_dataframe()
            for col in df.columns:
                if col in META_COLS:
                    continue
                mm = re.search(r"\d{4}-(\d{2})-\d{2}\s+\(FY\)", col)
                if mm:
                    return int(mm.group(1))
        except Exception:
            continue
    return 12
```

- [ ] **Step 4: 透傳 `fy_end_month` 到全部 build 函數**

共六個函數（含 `_build_template_table` 和 `_build_dynamic_table`），各加 `fy_end_month: int = 12` 參數並更新所有 `_col_to_quarter_label(...)` 呼叫：

```python
# _build_is_table
def _build_is_table(filings, max_filings: int, is_overrides=None, fy_end_month: int = 12):
    ...
    label = _col_to_quarter_label(q_col, fy_end_month)

# _build_bs_table
def _build_bs_table(filings, max_filings: int, bs_overrides=None, fy_end_month: int = 12):
    ...
    label = _col_to_quarter_label(is_q_col, fy_end_month) if is_q_col else _col_to_quarter_label(bs_col, fy_end_month)

# _build_cf_table
def _build_cf_table(filings, max_filings: int, cf_overrides=None, fy_end_month: int = 12):
    ...
    label = _col_to_quarter_label(q_col, fy_end_month)   # standalone Q
    ...
    label = _col_to_quarter_label(is_q_col, fy_end_month)  # YTD path

# _build_segment_tables
def _build_segment_tables(filings, max_filings: int, fy_end_month: int = 12):
    ...
    period_label = _col_to_quarter_label(q_col, fy_end_month)

# _build_template_table（generic builder，目前未被 fetch 呼叫，但需保持一致）
def _build_template_table(filings, template, sheet_name, stmt_method, max_filings, fy_end_month: int = 12):
    ...
    label = _col_to_quarter_label(q_col, fy_end_month)

# _build_dynamic_table（fallback builder，需保持一致）
def _build_dynamic_table(filings, stmt_method, sheet_name, max_filings, fy_end_month: int = 12):
    ...
    label = _col_to_quarter_label(q_col, fy_end_month)
```

- [ ] **Step 5: 在 `fetch_gaap_statements` 偵測並透傳**

在 `fetch_gaap_statements` 的 `filings_k = list(company.get_filings(...))` 之後加：

```python
    fy_end_month = _detect_fy_end_month(filings_k) if filings_k else 12
```

並把所有 `_build_is_table`、`_build_bs_table`、`_build_cf_table`、`_build_segment_tables` 的呼叫加上 `fy_end_month=fy_end_month`（共 8 次呼叫）。

- [ ] **Step 6: 跑測試確認通過**

```
pytest tests/test_fetcher_gaap.py -v
```

Expected: all PASSED（現有測試預設 `fy_end_month=12`，行為不變；新測試通過）

- [ ] **Step 7: Commit**

```
git add fetcher_gaap.py tests/test_fetcher_gaap.py
git commit -m "feat: align quarterly FY labels with company fiscal year end month"
```

---

## 最終驗收

- [ ] **跑完整 unit test suite**

```
pytest -v
```

Expected: 全部 PASSED（原有 110 個 + 新增約 25 個）

- [ ] **Live smoke test：TSLA、AMD（重新 fetch）**

```python
# 在專案根目錄執行
python -c "
from config import load_config
from fetcher_gaap import fetch_gaap_statements
from excel_writer import write_statements
from pathlib import Path
cfg = load_config()
for ticker in ['TSLA', 'AMD']:
    tables = fetch_gaap_statements(ticker, cfg['identity'], max_filings=cfg['max_filings'], ai_config=cfg.get('ai',{}))
    write_statements(tables, Path(cfg['output_dir']) / f'{ticker}_v2.xlsx')
    fin = next(t for t in tables if t.sheet_name == 'Data_Financials(Q)')
    print(ticker, 'quarters:', fin.quarter_labels[-3:])
"
```

人工確認：
- TSLA 最新季度 label 是否為 `FY2026Q1`（不變，Dec FY end）
- AMD 最新季度 label 是否正確
- Dividends Paid 數字是否合理（非 NCI 金額）
- OCF / Revenue 是否有值

- [ ] **Final commit（如有任何調整）**

```
git add -A
git commit -m "chore: post-verification minor tweaks"
```
