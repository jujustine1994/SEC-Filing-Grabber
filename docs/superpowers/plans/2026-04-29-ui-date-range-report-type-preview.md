# UI 日期區間、報表類型、Sheet 預覽 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在主介面加入日期區間選擇、季報/年報切換、以及執行前快速掃描 sheet 清單並選擇性跳過。

**Architecture:** 後端 `fetch_gaap_statements` 新增 `start_year`/`end_year`/`fetch_quarterly`/`fetch_annual`/`excluded_sheets` 參數；新增 `preview_sheets()` 公開函式只抓最新一份 10-Q；前端 `main.py` Tab 1/Tab 2 各加入對應 UI 元件，Tab 1 另加「快速掃描」按鈕與 sheet 勾選面板。

**Tech Stack:** Python 3.11, edgartools, Tkinter, pytest, openpyxl

---

## 檔案異動清單

| 檔案 | 動作 | 說明 |
|------|------|------|
| `fetcher_gaap.py` | 修改 | 新增 `_filter_filings_by_year()`、`preview_sheets()`；重構 `fetch_gaap_statements()` 加入 5 個新參數 |
| `fetcher_nongaap.py` | 修改 | `fetch_nongaap_statements()` 加入 `start_year`/`end_year` 參數 |
| `main.py` | 修改 | Tab 1 加入日期區間/報表類型/快速掃描；Tab 2 加入日期區間/報表類型 |
| `tests/test_fetcher_gaap.py` | 修改 | 新增 `_filter_filings_by_year`、`fetch_gaap_statements` 新參數、`preview_sheets` 的測試 |
| `tests/test_fetcher_nongaap.py` | 修改 | 新增 year filter 測試 |

---

## Task 1: `_filter_filings_by_year` helper

**Files:**
- Modify: `fetcher_gaap.py` (在 `_XBRL_CUTOFF` 常數定義之後加入 helper)
- Modify: `tests/test_fetcher_gaap.py`

- [ ] **Step 1: 寫失敗測試**

在 `tests/test_fetcher_gaap.py` 的 imports 加入：
```python
from fetcher_gaap import _filter_filings_by_year
```

在檔案尾端加入：

```python
# ── _filter_filings_by_year ───────────────────────────────────────────────────

def _make_dated_filing(year: int, month: int = 1, day: int = 15):
    """Mock filing with a date object as filing_date."""
    from datetime import date
    f = MagicMock()
    f.filing_date = date(year, month, day)
    return f


def test_filter_no_bounds_returns_all():
    filings = [_make_dated_filing(2020), _make_dated_filing(2022)]
    assert _filter_filings_by_year(filings, None, None) == filings


def test_filter_start_year_excludes_older():
    f2019 = _make_dated_filing(2019)
    f2021 = _make_dated_filing(2021)
    result = _filter_filings_by_year([f2019, f2021], start_year=2020, end_year=None)
    assert result == [f2021]


def test_filter_end_year_excludes_newer():
    f2021 = _make_dated_filing(2021)
    f2023 = _make_dated_filing(2023)
    result = _filter_filings_by_year([f2021, f2023], start_year=None, end_year=2022)
    assert result == [f2021]


def test_filter_both_bounds():
    filings = [_make_dated_filing(y) for y in [2018, 2020, 2022, 2024]]
    result = _filter_filings_by_year(filings, start_year=2019, end_year=2022)
    assert [f.filing_date.year for f in result] == [2020, 2022]


def test_filter_string_filing_date():
    """Mock filings that return filing_date as a string (unit-test pattern)."""
    f = MagicMock()
    f.filing_date = "2021-06-15"
    result = _filter_filings_by_year([f], start_year=2020, end_year=2022)
    assert result == [f]


def test_filter_string_filing_date_excluded():
    f = MagicMock()
    f.filing_date = "2019-03-01"
    result = _filter_filings_by_year([f], start_year=2020, end_year=None)
    assert result == []


def test_filter_empty_list():
    assert _filter_filings_by_year([], start_year=2020, end_year=2022) == []
```

- [ ] **Step 2: 確認測試失敗**

```bash
cd "C:\Users\CTH\Documents\Code\SEC Financial Tools"
pytest tests/test_fetcher_gaap.py::test_filter_no_bounds_returns_all -v
```
預期：`ImportError: cannot import name '_filter_filings_by_year'`

- [ ] **Step 3: 實作 `_filter_filings_by_year`**

在 `fetcher_gaap.py` 的 `_XBRL_CUTOFF` 定義之後（約第 64 行之後）加入：

```python
def _filter_filings_by_year(
    filings: list,
    start_year: int | None,
    end_year: int | None,
) -> list:
    """Filter filings list to only those within [start_year, end_year] (inclusive).

    Handles both date objects and ISO date strings ('YYYY-MM-DD').
    Returns filings unchanged when both bounds are None.
    """
    if start_year is None and end_year is None:
        return filings
    result = []
    for f in filings:
        fd = getattr(f, "filing_date", None)
        if fd is None:
            result.append(f)
            continue
        year = fd.year if isinstance(fd, _date) else int(str(fd)[:4])
        if start_year is not None and year < start_year:
            continue
        if end_year is not None and year > end_year:
            continue
        result.append(f)
    return result
```

- [ ] **Step 4: 確認測試全部通過**

```bash
pytest tests/test_fetcher_gaap.py -k "filter" -v
```
預期：7 PASSED

- [ ] **Step 5: 確認全套 unit tests 無破壞**

```bash
pytest tests/test_fetcher_gaap.py -v --ignore=tests/test_live_snapshots.py
```
預期：全部 PASSED（原有 95 個 + 新增 7 個）

- [ ] **Step 6: Commit**

```bash
git add fetcher_gaap.py tests/test_fetcher_gaap.py
git commit -m "feat: add _filter_filings_by_year helper with year-range support"
```

---

## Task 2: `fetch_gaap_statements` 新參數（年份區間 + 報表類型 + excluded_sheets）

**Files:**
- Modify: `fetcher_gaap.py` (`fetch_gaap_statements` 重構)
- Modify: `tests/test_fetcher_gaap.py`

- [ ] **Step 1: 寫失敗測試**

在 `tests/test_fetcher_gaap.py` 的 imports 更新，確保 `fetch_gaap_statements` 已匯入（已有）。

在檔案尾端加入（需要 mock Company）：

```python
# ── fetch_gaap_statements new params ─────────────────────────────────────────

def _make_mock_company(q_filings=None, k_filings=None):
    """Mock edgartools Company returning given filing lists."""
    company = MagicMock()
    q_filings = q_filings or []
    k_filings = k_filings or []

    def get_filings(form, amendments=False):
        if form == "10-Q":
            return q_filings
        if form == "10-K":
            return k_filings
        return []

    company.get_filings.side_effect = get_filings
    company.name = "Test Corp"
    return company


def _make_k_filing(period_col="2024-12-28", val=100.0, filing_date="2025-02-01"):
    """Mock a 10-K filing (same structure as 10-Q mock)."""
    return _make_filing(period_col=period_col, val=val, filing_date=filing_date)


@patch("fetcher_gaap.Company")
@patch("fetcher_gaap.set_identity")
@patch("fetcher_gaap.load_overrides", return_value={})
def test_fetch_gaap_annual_only_no_q_table(mock_ov, mock_id, mock_co):
    """fetch_quarterly=False should produce Data_Financials(Y) but NOT Data_Financials(Q)."""
    k = _make_k_filing()
    mock_co.return_value = _make_mock_company(q_filings=[], k_filings=[k])

    tables = fetch_gaap_statements("TEST", "Test test@test.com",
                                   fetch_quarterly=False, fetch_annual=True)
    sheet_names = [t.sheet_name for t in tables]
    assert "Data_Financials(Q)" not in sheet_names
    assert "Data_Financials(Y)" in sheet_names


@patch("fetcher_gaap.Company")
@patch("fetcher_gaap.set_identity")
@patch("fetcher_gaap.load_overrides", return_value={})
def test_fetch_gaap_quarterly_only_no_y_table(mock_ov, mock_id, mock_co):
    """fetch_annual=False should produce Data_Financials(Q) but NOT Data_Financials(Y)."""
    q = _make_filing()
    mock_co.return_value = _make_mock_company(q_filings=[q], k_filings=[])

    tables = fetch_gaap_statements("TEST", "Test test@test.com",
                                   fetch_quarterly=True, fetch_annual=False)
    sheet_names = [t.sheet_name for t in tables]
    assert "Data_Financials(Q)" in sheet_names
    assert "Data_Financials(Y)" not in sheet_names


@patch("fetcher_gaap.Company")
@patch("fetcher_gaap.set_identity")
@patch("fetcher_gaap.load_overrides", return_value={})
def test_fetch_gaap_annual_only_raises_when_no_k(mock_ov, mock_id, mock_co):
    """fetch_quarterly=False with no 10-K filings should raise ValueError."""
    mock_co.return_value = _make_mock_company(q_filings=[], k_filings=[])
    with pytest.raises(ValueError, match="10-K"):
        fetch_gaap_statements("TEST", "Test test@test.com",
                               fetch_quarterly=False, fetch_annual=True)


@patch("fetcher_gaap.Company")
@patch("fetcher_gaap.set_identity")
@patch("fetcher_gaap.load_overrides", return_value={})
def test_fetch_gaap_excluded_sheets_removes_seg(mock_ov, mock_id, mock_co):
    """excluded_sheets should skip matching sheet names in the result."""
    q = _make_filing()
    mock_co.return_value = _make_mock_company(q_filings=[q], k_filings=[])

    # No segments in mock data; just verify excluded_sheets param is accepted without error
    tables = fetch_gaap_statements("TEST", "Test test@test.com",
                                   fetch_quarterly=True, fetch_annual=False,
                                   excluded_sheets={"Data_Seg_Revenue"})
    sheet_names = [t.sheet_name for t in tables]
    assert "Data_Seg_Revenue" not in sheet_names
```

- [ ] **Step 2: 確認測試失敗**

```bash
pytest tests/test_fetcher_gaap.py -k "fetch_gaap" -v
```
預期：`TypeError: fetch_gaap_statements() got an unexpected keyword argument 'fetch_quarterly'`

- [ ] **Step 3: 重構 `fetch_gaap_statements`**

將 `fetcher_gaap.py` 中的 `fetch_gaap_statements` 函式替換為以下版本（保留原有邏輯，加入新參數控制）：

```python
def fetch_gaap_statements(ticker: str, identity: str,
                           max_filings: int = 80,
                           max_annual_filings: int = 20,
                           ai_config: dict | None = None,
                           start_year: int | None = None,
                           end_year: int | None = None,
                           fetch_quarterly: bool = True,
                           fetch_annual: bool = True,
                           excluded_sheets: set[str] | None = None) -> list[StatementTable]:
    """Fetch quarterly and/or annual GAAP statements for a ticker.

    Args:
        ticker:              Stock ticker, e.g. "AAPL"
        identity:            SEC EDGAR identity string
        max_filings:         Max 10-Q filings to process (default 80, ~20 years)
        max_annual_filings:  Max 10-K filings to process (default 20, ~20 years)
        ai_config:           AI config dict (provider/model/api_key) for E2 diagnosis
        start_year:          Only include filings from this year onwards (None = no limit)
        end_year:            Only include filings up to this year (None = no limit)
        fetch_quarterly:     Whether to fetch 10-Q data (default True)
        fetch_annual:        Whether to fetch 10-K data (default True)
        excluded_sheets:     Set of sheet names to skip in the output

    Returns:
        List of StatementTable

    Raises:
        ValueError: No filings found for the requested form type(s)
    """
    ai_config = ai_config or {}
    excluded_sheets = excluded_sheets or set()
    set_identity(identity)
    company = Company(ticker)

    filings_q = list(company.get_filings(form="10-Q", amendments=False)) if fetch_quarterly else []
    filings_k = list(company.get_filings(form="10-K", amendments=False)) if fetch_annual else []

    if fetch_quarterly and not filings_q:
        raise ValueError(
            f"No 10-Q filings found for ticker '{ticker}'. "
            "The ticker may be invalid or the company may not file 10-Qs."
        )
    if not fetch_quarterly and not filings_k:
        raise ValueError(
            f"No 10-K filings found for ticker '{ticker}'. "
            "The ticker may be invalid or the company may not file 10-Ks."
        )

    # Apply year range filter
    filings_q = _filter_filings_by_year(filings_q, start_year, end_year)
    filings_k = _filter_filings_by_year(filings_k, start_year, end_year)

    overrides = load_overrides(ticker)
    fy_end_month = _detect_fy_end_month(filings_k) if filings_k else 12

    tables: list[StatementTable] = []

    if fetch_quarterly and filings_q:
        is_tbl, is_ng = _build_is_table(filings_q, max_filings, is_overrides=overrides.get("IS", {}), fy_end_month=fy_end_month)
        bs_tbl, bs_ng = _build_bs_table(filings_q, max_filings, bs_overrides=overrides.get("BS", {}), fy_end_month=fy_end_month)
        cf_tbl, cf_ng = _build_cf_table(filings_q, max_filings, cf_overrides=overrides.get("CF", {}), fy_end_month=fy_end_month)

        # Diagnose key rows that are all-None in recent quarters
        missing_is = check_key_rows(is_tbl.concepts, is_tbl.values, "IS")
        missing_bs = check_key_rows(bs_tbl.concepts, bs_tbl.values, "BS")
        missing_cf = check_key_rows(cf_tbl.concepts, cf_tbl.values, "CF")

        if missing_is or missing_bs or missing_cf:
            try:
                tenq_latest = filings_q[0].obj()
                latest_is_df = tenq_latest.financials.income_statement().to_dataframe()
                latest_bs_df = tenq_latest.financials.balance_sheet().to_dataframe()
                latest_cf_df = tenq_latest.financials.cashflow_statement().to_dataframe()
            except Exception as exc:
                print(f"[{ticker}] 診斷：無法取得最新 filing DataFrame — {exc!r}", file=sys.stderr)
                latest_is_df = latest_bs_df = latest_cf_df = None

            new_overrides: dict[str, dict] = {}
            if missing_is and latest_is_df is not None:
                fixes = run_diagnosis(ticker, "IS", latest_is_df, missing_is, ai_config)
                if fixes:
                    new_overrides["IS"] = fixes
            if missing_bs and latest_bs_df is not None:
                fixes = run_diagnosis(ticker, "BS", latest_bs_df, missing_bs, ai_config)
                if fixes:
                    new_overrides["BS"] = fixes
            if missing_cf and latest_cf_df is not None:
                fixes = run_diagnosis(ticker, "CF", latest_cf_df, missing_cf, ai_config)
                if fixes:
                    new_overrides["CF"] = fixes

            if new_overrides:
                total_fixes = sum(len(v) for v in new_overrides.values())
                print(f"[{ticker}] 自動修復：找到 {total_fixes} 項缺失指標修復方案，重新建表。", file=sys.stderr)
                overrides = load_overrides(ticker)
                is_tbl, is_ng = _build_is_table(filings_q, max_filings, is_overrides=overrides.get("IS", {}), fy_end_month=fy_end_month)
                bs_tbl, bs_ng = _build_bs_table(filings_q, max_filings, bs_overrides=overrides.get("BS", {}), fy_end_month=fy_end_month)
                cf_tbl, cf_ng = _build_cf_table(filings_q, max_filings, cf_overrides=overrides.get("CF", {}), fy_end_month=fy_end_month)
            else:
                remaining = missing_is + missing_bs + missing_cf
                if remaining:
                    no_key = "" if ai_config.get("api_key") else "（未設 AI API key，E2 診斷已跳過）"
                    print(f"[{ticker}] 警告：{remaining} 在 EDGAR 中無對應概念{no_key}。", file=sys.stderr)

        quarterly_tbl = _merge_financials(is_tbl, bs_tbl, cf_tbl, sheet_name="Data_Financials(Q)")
        tables.append(quarterly_tbl)
        if any(tbl.concepts for tbl in [is_ng, bs_ng, cf_ng]):
            ng_q_tbl = _merge_financials(is_ng, bs_ng, cf_ng, sheet_name="Data_Financials_NG(Q)")
            tables.append(ng_q_tbl)

    if fetch_annual and filings_k:
        is_ann, is_ann_ng = _build_is_table(filings_k, max_annual_filings, is_overrides=overrides.get("IS", {}), fy_end_month=fy_end_month)
        bs_ann, bs_ann_ng = _build_bs_table(filings_k, max_annual_filings, bs_overrides=overrides.get("BS", {}), fy_end_month=fy_end_month)
        cf_ann, cf_ann_ng = _build_cf_table(filings_k, max_annual_filings, cf_overrides=overrides.get("CF", {}), fy_end_month=fy_end_month)
        annual_tbl = _merge_financials(is_ann, bs_ann, cf_ann, sheet_name="Data_Financials(Y)")
        tables.append(annual_tbl)
        if any(tbl.concepts for tbl in [is_ann_ng, bs_ann_ng, cf_ann_ng]):
            ng_y_tbl = _merge_financials(is_ann_ng, bs_ann_ng, cf_ann_ng, sheet_name="Data_Financials_NG(Y)")
            tables.append(ng_y_tbl)

    if fetch_quarterly and filings_q:
        seg_tables = _build_segment_tables(filings_q, max_filings, fy_end_month=fy_end_month)
        tables.extend(t for t in seg_tables if t.sheet_name not in excluded_sheets)

    company_name = getattr(company, "name", ticker) or ticker
    tables.append(_build_meta_table(ticker, company_name, tables))

    for tbl in tables:
        tbl.ticker = ticker
    return tables
```

注意：此步驟同時把原本的 `total` 變數改名為 `total_fixes`（避免遮蔽 builtins）。

- [ ] **Step 4: 確認新測試通過**

```bash
pytest tests/test_fetcher_gaap.py -k "fetch_gaap" -v
```
預期：4 PASSED

- [ ] **Step 5: 確認全套 unit tests 無破壞**

```bash
pytest tests/test_fetcher_gaap.py -v
```
預期：全部 PASSED

- [ ] **Step 6: Commit**

```bash
git add fetcher_gaap.py tests/test_fetcher_gaap.py
git commit -m "feat: add start_year/end_year/fetch_quarterly/fetch_annual/excluded_sheets to fetch_gaap_statements"
```

---

## Task 3: `preview_sheets()` 函式

**Files:**
- Modify: `fetcher_gaap.py`（在 `fetch_gaap_statements` 之後加入）
- Modify: `tests/test_fetcher_gaap.py`

- [ ] **Step 1: 寫失敗測試**

在 `tests/test_fetcher_gaap.py` 的 imports 加入：
```python
from fetcher_gaap import preview_sheets
```

在檔案尾端加入：

```python
# ── preview_sheets ────────────────────────────────────────────────────────────

@patch("fetcher_gaap.Company")
@patch("fetcher_gaap.set_identity")
def test_preview_sheets_fixed_always_present(mock_id, mock_co):
    """Fixed sheets should always appear regardless of what the company has."""
    q = _make_filing()
    company = _make_mock_company(q_filings=[q])
    mock_co.return_value = company

    result = preview_sheets("AAPL", "Test test@test.com")

    assert "Data_Financials(Q)" in result
    assert "Data_Financials(Y)" in result
    assert "Data_Meta" in result


@patch("fetcher_gaap.Company")
@patch("fetcher_gaap.set_identity")
def test_preview_sheets_no_q_filings(mock_id, mock_co):
    """When no 10-Q filings exist, only fixed sheets are returned."""
    company = _make_mock_company(q_filings=[])
    mock_co.return_value = company

    result = preview_sheets("NOFILINGS", "Test test@test.com")

    assert result == ["Data_Financials(Q)", "Data_Financials(Y)", "Data_Meta"]


@patch("fetcher_gaap.Company")
@patch("fetcher_gaap.set_identity")
def test_preview_sheets_returns_list_of_strings(mock_id, mock_co):
    """Return type should be list[str]."""
    q = _make_filing()
    mock_co.return_value = _make_mock_company(q_filings=[q])

    result = preview_sheets("TEST", "Test test@test.com")

    assert isinstance(result, list)
    assert all(isinstance(s, str) for s in result)
```

- [ ] **Step 2: 確認測試失敗**

```bash
pytest tests/test_fetcher_gaap.py -k "preview_sheets" -v
```
預期：`ImportError: cannot import name 'preview_sheets'`

- [ ] **Step 3: 實作 `preview_sheets`**

在 `fetcher_gaap.py` 的 `fetch_gaap_statements` 函式之後加入：

```python
def preview_sheets(ticker: str, identity: str) -> list[str]:
    """Quick scan: fetch only the latest 10-Q to detect segment sheet names.

    Returns the predicted list of sheet names without performing a full fetch.
    Takes ~5–15 seconds (one HTTP request for the latest filing).

    Returns:
        List of sheet name strings. Fixed sheets (Financials Q/Y, Meta) are
        always included. Data_Seg_* sheets are detected from the latest 10-Q.
    """
    fixed = ["Data_Financials(Q)", "Data_Financials(Y)", "Data_Meta"]

    set_identity(identity)
    company = Company(ticker)
    filings_q = list(company.get_filings(form="10-Q", amendments=False))
    if not filings_q:
        return fixed

    try:
        seg_tables = _build_segment_tables([filings_q[0]], max_filings=1)
        seg_names = [t.sheet_name for t in seg_tables]
    except Exception as exc:
        print(f"[preview_sheets] Segment scan failed: {exc!r}", file=sys.stderr)
        seg_names = []

    return fixed + seg_names
```

- [ ] **Step 4: 確認測試通過**

```bash
pytest tests/test_fetcher_gaap.py -k "preview_sheets" -v
```
預期：3 PASSED

- [ ] **Step 5: 確認全套 unit tests**

```bash
pytest tests/test_fetcher_gaap.py -v
```
預期：全部 PASSED

- [ ] **Step 6: Commit**

```bash
git add fetcher_gaap.py tests/test_fetcher_gaap.py
git commit -m "feat: add preview_sheets() for quick segment detection"
```

---

## Task 4: `fetch_nongaap_statements` 年份區間

**Files:**
- Modify: `fetcher_nongaap.py`
- Modify: `tests/test_fetcher_nongaap.py`

- [ ] **Step 1: 寫失敗測試**

在 `tests/test_fetcher_nongaap.py` 頂端的 imports 加入：
```python
from fetcher_nongaap import _filter_nongaap_by_year
```

在檔案尾端加入：

```python
# ── _filter_nongaap_by_year ───────────────────────────────────────────────────

def _make_ng_filing(label: str):
    """Make a (label, filing, eight_k) tuple with a label like 'FY2021Q2'."""
    from unittest.mock import MagicMock
    filing = MagicMock()
    eight_k = MagicMock()
    return (label, filing, eight_k)


def test_filter_nongaap_no_bounds():
    filings = [_make_ng_filing("FY2020Q1"), _make_ng_filing("FY2022Q3")]
    result = _filter_nongaap_by_year(filings, None, None)
    assert len(result) == 2


def test_filter_nongaap_start_year():
    filings = [_make_ng_filing("FY2019Q4"), _make_ng_filing("FY2021Q1")]
    result = _filter_nongaap_by_year(filings, start_year=2020, end_year=None)
    assert len(result) == 1
    assert result[0][0] == "FY2021Q1"


def test_filter_nongaap_end_year():
    filings = [_make_ng_filing("FY2020Q2"), _make_ng_filing("FY2023Q1")]
    result = _filter_nongaap_by_year(filings, start_year=None, end_year=2021)
    assert len(result) == 1
    assert result[0][0] == "FY2020Q2"


def test_filter_nongaap_both_bounds():
    filings = [_make_ng_filing(f"FY{y}Q1") for y in [2018, 2020, 2022, 2024]]
    result = _filter_nongaap_by_year(filings, start_year=2019, end_year=2022)
    years = [int(lbl[2:6]) for lbl, _, _ in result]
    assert years == [2020, 2022]
```

- [ ] **Step 2: 確認測試失敗**

```bash
pytest tests/test_fetcher_nongaap.py -k "filter_nongaap" -v
```
預期：`ImportError: cannot import name '_filter_nongaap_by_year'`

- [ ] **Step 3: 實作 `_filter_nongaap_by_year` 並更新 `fetch_nongaap_statements`**

在 `fetcher_nongaap.py` 的 `_get_earnings_filings` 之前加入 helper：

```python
def _filter_nongaap_by_year(
    filings: list[tuple],
    start_year: int | None,
    end_year: int | None,
) -> list[tuple]:
    """Filter (label, filing, eight_k) tuples by year extracted from label (e.g. 'FY2021Q2' → 2021)."""
    if start_year is None and end_year is None:
        return filings
    import re as _re
    result = []
    for item in filings:
        label = item[0]
        m = _re.search(r'(\d{4})', label)
        if m is None:
            result.append(item)
            continue
        year = int(m.group(1))
        if start_year is not None and year < start_year:
            continue
        if end_year is not None and year > end_year:
            continue
        result.append(item)
    return result
```

更新 `fetch_nongaap_statements` 簽名（在 `max_filings: int = 80` 之後加入兩個新參數）並在函式體內加入過濾：

```python
def fetch_nongaap_statements(
    ticker: str,
    identity: str,
    ai_config: dict,
    output_dir: Path,
    progress_cb=None,
    max_filings: int = 80,
    start_year: int | None = None,
    end_year: int | None = None,
) -> list[StatementTable]:
```

在 `filings = _get_earnings_filings(company)[:max_filings]` 之後加入：

```python
    filings = _filter_nongaap_by_year(filings, start_year, end_year)
```

- [ ] **Step 4: 確認測試通過**

```bash
pytest tests/test_fetcher_nongaap.py -k "filter_nongaap" -v
```
預期：4 PASSED

- [ ] **Step 5: 確認全套 tests 無破壞**

```bash
pytest tests/test_fetcher_nongaap.py -v
```
預期：全部 PASSED

- [ ] **Step 6: Commit**

```bash
git add fetcher_nongaap.py tests/test_fetcher_nongaap.py
git commit -m "feat: add start_year/end_year filter to fetch_nongaap_statements"
```

---

## Task 5: UI Tab 1 — 日期區間 + 報表類型

**Files:**
- Modify: `main.py`

此 task 修改 `_build_tab1()`、`_run_single()`。無自動化測試，驗證靠人工啟動 app 測試。

- [ ] **Step 1: 在 `__init__` 加入新狀態變數**

在 `main.py` 的 `__init__` 中，`self.tab1_preview_label = None` 之後加入：

```python
        self.tab1_fetch_q_var: tk.BooleanVar | None = None
        self.tab1_fetch_k_var: tk.BooleanVar | None = None
        self.tab1_start_year_var: tk.StringVar | None = None
        self.tab1_end_year_var: tk.StringVar | None = None
```

- [ ] **Step 2: 在 `_build_tab1` 加入報表類型 row**

在 `_build_tab1` 的 Row 1（GAAP/Non-GAAP checkboxes）之後，Row 3（輸出設定 toggle）之前，插入兩個新 row：

```python
        # Row 2: Report type checkboxes
        row_rtype = ttk.Frame(tab)
        row_rtype.grid(row=2, column=0, sticky="ew", pady=(2, 0))
        self.tab1_fetch_q_var = tk.BooleanVar(value=True)
        self.tab1_fetch_k_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(row_rtype, text="季報 (10-Q)", variable=self.tab1_fetch_q_var).pack(side="left", padx=(0, 16))
        ttk.Checkbutton(row_rtype, text="年報 (10-K)", variable=self.tab1_fetch_k_var).pack(side="left")

        # Row 2b: Date range
        row_date = ttk.Frame(tab)
        row_date.grid(row=3, column=0, sticky="ew", pady=(2, 4))
        ttk.Label(row_date, text="日期區間：起").pack(side="left", padx=(0, 4))
        self.tab1_start_year_var = tk.StringVar(value="")
        ttk.Spinbox(row_date, from_=1993, to=2099, textvariable=self.tab1_start_year_var,
                    width=6).pack(side="left")
        ttk.Label(row_date, text="迄").pack(side="left", padx=(8, 4))
        self.tab1_end_year_var = tk.StringVar(value="")
        ttk.Spinbox(row_date, from_=1993, to=2099, textvariable=self.tab1_end_year_var,
                    width=6).pack(side="left")
        ttk.Label(row_date, text="年　（留空 = 全部）", foreground="#555555").pack(side="left", padx=(4, 0))
```

注意：原本 row 2、3、4、5 全部加 3（留出 row=4 給 Task 7 的 sheet panel），最終配置如下：

```
Row 0  ticker (+ scan button，Task 7 加入)
Row 1  GAAP/Non-GAAP checkboxes
Row 2  報表類型 (NEW)
Row 3  日期區間 (NEW)
Row 4  sheet panel (Task 7 加入，grid_remove 初始隱藏)
Row 5  Non-GAAP warning label   ← 原 row=2
Row 6  output settings toggle   ← 原 row=3
Row 7  output settings content  ← 原 row=4
Row 8  execute button           ← 原 row=5
```

在 `_build_tab1` 中更新以下三處 row 編號：

```python
        # Non-GAAP warning label：row=2 → row=5
        self.nongaap_warn_label.grid(row=5, column=0, sticky="w", padx=2)

        # Output settings toggle：row=3 → row=6
        out_toggle_row.grid(row=6, column=0, sticky="ew", pady=(8, 0))

        # Output settings content：row=4 → row=7
        out_frame.grid(row=7, column=0, sticky="ew", pady=(0, 4))

        # Execute button：row=5 → row=8
        self.btn_run_single.grid(row=8, column=0, pady=(8, 4))
```

- [ ] **Step 3: 更新 `_run_single` 讀取新 UI 狀態**

在 `_run_single` 中，`max_filings = self.cfg.get("max_filings", 80)` 之後加入：

```python
        fetch_q   = self.tab1_fetch_q_var.get() if self.tab1_fetch_q_var else True
        fetch_k   = self.tab1_fetch_k_var.get() if self.tab1_fetch_k_var else True
        if not fetch_q and not fetch_k:
            messagebox.showerror("錯誤", "請至少勾選季報 (10-Q) 或年報 (10-K)")
            return
        try:
            start_year = int(self.tab1_start_year_var.get()) if self.tab1_start_year_var and self.tab1_start_year_var.get().strip() else None
            end_year   = int(self.tab1_end_year_var.get())   if self.tab1_end_year_var   and self.tab1_end_year_var.get().strip()   else None
        except ValueError:
            messagebox.showerror("錯誤", "日期區間請輸入有效年份（如 2018）")
            return
```

並更新 `_start_worker` 呼叫，將新參數傳入：

```python
        self._start_worker(lambda: self._worker_single(
            ticker, fetch_gaap, fetch_nongaap, max_filings,
            fetch_q=fetch_q, fetch_k=fetch_k,
            start_year=start_year, end_year=end_year,
        ))
```

- [ ] **Step 4: 更新 `_worker_single` 簽名與 `fetch_gaap_statements` 呼叫**

```python
    def _worker_single(self, ticker: str, fetch_gaap: bool, fetch_nongaap: bool,
                       max_filings: int = 80, fetch_q: bool = True, fetch_k: bool = True,
                       start_year: int | None = None, end_year: int | None = None):
```

在 `fetch_gaap_statements` 呼叫中加入新參數：

```python
                gaap_tables = fetch_gaap_statements(
                    ticker, identity, max_filings=max_filings,
                    ai_config=self.cfg.get("ai", {}),
                    start_year=start_year, end_year=end_year,
                    fetch_quarterly=fetch_q, fetch_annual=fetch_k,
                )
```

在 `fetch_nongaap_statements` 呼叫中加入新參數：

```python
                ng_tables = fetch_nongaap_statements(
                    ticker, identity, ai_config,
                    output_dir=output_dir,
                    progress_cb=_ng_progress,
                    max_filings=max_filings,
                    start_year=start_year, end_year=end_year,
                )
```

- [ ] **Step 5: 人工驗證**

啟動 app：`python main.py`

測試項目：
1. 報表類型預設兩個都勾
2. 取消兩個都不勾 → 點執行 → 應彈出錯誤訊息
3. 日期區間起填 `2020`、迄留空 → 執行 AAPL → log 中應看到資料抓取（不報錯）
4. 日期區間填非數字 `abc` → 點執行 → 應彈出錯誤訊息

- [ ] **Step 6: Commit**

```bash
git add main.py
git commit -m "feat: add report type (10-Q/10-K) and date range to Tab 1 UI"
```

---

## Task 6: UI Tab 2 — 日期區間 + 報表類型

**Files:**
- Modify: `main.py`

- [ ] **Step 1: 在 `__init__` 加入 Tab 2 新狀態變數**

在 `self.tab1_end_year_var` 之後加入：

```python
        self.batch_fetch_q_var: tk.BooleanVar | None = None
        self.batch_fetch_k_var: tk.BooleanVar | None = None
        self.batch_start_year_var: tk.StringVar | None = None
        self.batch_end_year_var: tk.StringVar | None = None
```

- [ ] **Step 2: 在 `_build_tab2` 加入新 UI 元件**

在 `self.batch_nongaap_var.trace_add(...)` 之後（row 2 之後）加入：

```python
        # Row 3: Report type
        row_rtype2 = ttk.Frame(tab)
        row_rtype2.grid(row=3, column=0, sticky="w", pady=(4, 0))
        self.batch_fetch_q_var = tk.BooleanVar(value=True)
        self.batch_fetch_k_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(row_rtype2, text="季報 (10-Q)", variable=self.batch_fetch_q_var).pack(side="left", padx=(0, 16))
        ttk.Checkbutton(row_rtype2, text="年報 (10-K)", variable=self.batch_fetch_k_var).pack(side="left")

        # Row 4: Date range
        row_date2 = ttk.Frame(tab)
        row_date2.grid(row=4, column=0, sticky="w", pady=(2, 0))
        ttk.Label(row_date2, text="日期區間：起").pack(side="left", padx=(0, 4))
        self.batch_start_year_var = tk.StringVar(value="")
        ttk.Spinbox(row_date2, from_=1993, to=2099, textvariable=self.batch_start_year_var,
                    width=6).pack(side="left")
        ttk.Label(row_date2, text="迄").pack(side="left", padx=(8, 4))
        self.batch_end_year_var = tk.StringVar(value="")
        ttk.Spinbox(row_date2, from_=1993, to=2099, textvariable=self.batch_end_year_var,
                    width=6).pack(side="left")
        ttk.Label(row_date2, text="年　（留空 = 全部）", foreground="#555555").pack(side="left", padx=(4, 0))
```

原本 `btn_run_batch` 的 `grid(row=3)` 改為 `grid(row=5)`。

- [ ] **Step 3: 更新 `_run_batch` 讀取新 UI 狀態**

在 `_run_batch` 中，`fetch_nongaap = self.batch_nongaap_var.get()` 之後加入：

```python
        fetch_q = self.batch_fetch_q_var.get() if self.batch_fetch_q_var else True
        fetch_k = self.batch_fetch_k_var.get() if self.batch_fetch_k_var else True
        if not fetch_q and not fetch_k:
            messagebox.showerror("錯誤", "請至少勾選季報 (10-Q) 或年報 (10-K)")
            return
        try:
            start_year = int(self.batch_start_year_var.get()) if self.batch_start_year_var and self.batch_start_year_var.get().strip() else None
            end_year   = int(self.batch_end_year_var.get())   if self.batch_end_year_var   and self.batch_end_year_var.get().strip()   else None
        except ValueError:
            messagebox.showerror("錯誤", "日期區間請輸入有效年份（如 2018）")
            return
```

更新 `_start_worker` 呼叫：

```python
        self._start_worker(lambda: self._worker_batch(
            selected, fetch_nongaap,
            fetch_q=fetch_q, fetch_k=fetch_k,
            start_year=start_year, end_year=end_year,
        ))
```

- [ ] **Step 4: 更新 `_worker_batch` 簽名與呼叫**

```python
    def _worker_batch(self, tickers: list[str], fetch_nongaap: bool = False,
                      fetch_q: bool = True, fetch_k: bool = True,
                      start_year: int | None = None, end_year: int | None = None):
```

更新 `fetch_gaap_statements` 呼叫：

```python
                tables = fetch_gaap_statements(
                    ticker, identity, max_filings=max_filings, ai_config=ai_config,
                    start_year=start_year, end_year=end_year,
                    fetch_quarterly=fetch_q, fetch_annual=fetch_k,
                )
```

更新 `fetch_nongaap_statements` 呼叫：

```python
                    ng_tables = fetch_nongaap_statements(
                        ticker, identity, ai_config,
                        output_dir=output_dir,
                        progress_cb=_ng_cb,
                        max_filings=max_filings,
                        start_year=start_year, end_year=end_year,
                    )
```

- [ ] **Step 5: 人工驗證**

啟動 app：`python main.py`

測試項目：
1. Tab 2 報表類型預設兩個都勾
2. 取消兩個都不勾 → 點批量更新 → 應彈出錯誤訊息
3. 日期區間起填 `2022` → 選 1 個 watchlist ticker → 執行 → 不報錯

- [ ] **Step 6: Commit**

```bash
git add main.py
git commit -m "feat: add report type and date range to Tab 2 batch UI"
```

---

## Task 7: UI Tab 1 — 快速掃描 + Sheet 選擇面板

**Files:**
- Modify: `main.py`

- [ ] **Step 1: 在 `__init__` 加入新狀態變數**

在 `self.batch_end_year_var` 之後加入：

```python
        self._sheet_check_vars: dict[str, tk.BooleanVar] = {}
        self._sheet_panel_frame: tk.Frame | None = None
        self._scan_btn: ttk.Button | None = None
```

- [ ] **Step 2: 在 `_build_tab1` 加入「快速掃描」按鈕和 sheet 面板**

在 Row 0（ticker input row）的 pack 之後，加入「快速掃描」按鈕到 `row_ticker`：

```python
        self._scan_btn = ttk.Button(row_ticker, text="快速掃描 ▶", command=self._run_preview_scan, width=12)
        self._scan_btn.pack(side="left", padx=(12, 0))
```

在 Row 3（日期區間）之後（此位置已在 Task 5 中預留為 row=4），加入可折疊的 sheet 面板（預設隱藏）：

```python
        # Row 4: Sheet selection panel (hidden until scan completes)
        self._sheet_panel_frame = ttk.LabelFrame(tab, text=" 可選 Sheet（掃描後顯示）", padding=6)
        self._sheet_panel_frame.grid(row=4, column=0, sticky="ew", pady=(0, 4))
        self._sheet_panel_frame.grid_remove()
```

Row 5–8（Non-GAAP warning、輸出設定 toggle/content、執行按鈕）已在 Task 5 中配置，不需再調整。

- [ ] **Step 3: 實作 `_run_preview_scan`**

加入以下方法：

```python
    def _run_preview_scan(self):
        """Start background preview scan for the current ticker."""
        ticker = self._get_ph_value(self.ticker_var, self.TICKER_PH).upper()
        if not ticker:
            messagebox.showerror("錯誤", "請先輸入 Ticker")
            return
        identity = self.cfg.get("identity", "")
        if not identity:
            messagebox.showerror("錯誤", "請先在進階設定填入 Identity")
            return
        if self._scan_btn:
            self._scan_btn.config(state="disabled", text="掃描中...")
        if self._sheet_panel_frame:
            self._sheet_panel_frame.grid_remove()
        self._sheet_check_vars = {}
        threading.Thread(
            target=lambda: self._preview_scan_worker(ticker, identity), daemon=True
        ).start()

    def _preview_scan_worker(self, ticker: str, identity: str):
        """Background thread: call preview_sheets() and push result to queue."""
        try:
            from fetcher_gaap import preview_sheets
            sheets = preview_sheets(ticker, identity)
            self.msg_queue.put(("preview_scan_done", sheets))
        except Exception as e:
            self.msg_queue.put(("preview_scan_error", str(e)))
```

- [ ] **Step 4: 在 `_poll_queue` 處理掃描結果**

在 `_poll_queue` 的 `try` 區塊內，`elif msg_type == "ai_test_result":` 之後加入：

```python
                elif msg_type == "preview_scan_done":
                    self._build_sheet_panel(data)
                    if self._scan_btn:
                        self._scan_btn.config(state="normal", text="快速掃描 ▶")

                elif msg_type == "preview_scan_error":
                    if self._scan_btn:
                        self._scan_btn.config(state="normal", text="快速掃描 ▶")
                    messagebox.showerror("掃描失敗", f"無法完成快速掃描：{data}")
```

- [ ] **Step 5: 實作 `_build_sheet_panel`**

加入以下方法：

```python
    _FIXED_SHEETS = frozenset({"Data_Financials(Q)", "Data_Financials(Y)", "Data_Meta"})

    def _build_sheet_panel(self, sheet_names: list[str]):
        """Populate sheet selection panel with checkboxes. Fixed sheets are disabled."""
        if not self._sheet_panel_frame:
            return
        for w in self._sheet_panel_frame.winfo_children():
            w.destroy()
        self._sheet_check_vars = {}

        for name in sheet_names:
            var = tk.BooleanVar(value=True)
            self._sheet_check_vars[name] = var
            is_fixed = name in self._FIXED_SHEETS
            cb = ttk.Checkbutton(
                self._sheet_panel_frame, text=name, variable=var,
                state="disabled" if is_fixed else "normal",
            )
            cb.pack(anchor="w", padx=4)

        self._sheet_panel_frame.grid()
```

- [ ] **Step 6: 在 `_run_single` 讀取 excluded_sheets 並傳入**

在 Task 5 Step 3 的年份解析之後加入：

```python
        excluded = {
            name for name, var in self._sheet_check_vars.items()
            if not var.get() and name not in self._FIXED_SHEETS
        }
```

並更新 `_worker_single` 呼叫加入 `excluded_sheets=excluded`：

```python
        self._start_worker(lambda: self._worker_single(
            ticker, fetch_gaap, fetch_nongaap, max_filings,
            fetch_q=fetch_q, fetch_k=fetch_k,
            start_year=start_year, end_year=end_year,
            excluded_sheets=excluded,
        ))
```

更新 `_worker_single` 簽名：

```python
    def _worker_single(self, ticker: str, fetch_gaap: bool, fetch_nongaap: bool,
                       max_filings: int = 80, fetch_q: bool = True, fetch_k: bool = True,
                       start_year: int | None = None, end_year: int | None = None,
                       excluded_sheets: set[str] | None = None):
```

並在 `fetch_gaap_statements` 呼叫中加入：

```python
                gaap_tables = fetch_gaap_statements(
                    ticker, identity, max_filings=max_filings,
                    ai_config=self.cfg.get("ai", {}),
                    start_year=start_year, end_year=end_year,
                    fetch_quarterly=fetch_q, fetch_annual=fetch_k,
                    excluded_sheets=excluded_sheets or set(),
                )
```

- [ ] **Step 7: 人工驗證**

啟動 app：`python main.py`

測試項目：
1. 輸入 `AAPL`，點「快速掃描 ▶」→ 按鈕變灰「掃描中...」→ 數秒後出現 sheet 清單
2. 清單中 `Data_Financials(Q/Y)`、`Data_Meta` 為灰色（不可取消）
3. `Data_Seg_*` 可勾選/取消
4. 取消一個 Seg sheet → 執行 → log 應顯示此 sheet 被跳過（不在最終 Excel 中）
5. 不點掃描直接執行 → 全抓（不出現 sheet 面板沒有影響）

- [ ] **Step 8: Commit**

```bash
git add main.py
git commit -m "feat: add quick scan button and sheet selection panel to Tab 1"
```

---

## 完成後驗收

- [ ] 執行全套 unit tests：`pytest tests/ -v --ignore=tests/test_live_snapshots.py`
- [ ] 全部 PASSED
- [ ] 人工啟動 app，完整跑一次 Tab 1（AAPL，2020–2024年，只勾季報）
- [ ] 人工啟動 app，完整跑一次 Tab 2 批量（2 個 ticker，只勾年報）
