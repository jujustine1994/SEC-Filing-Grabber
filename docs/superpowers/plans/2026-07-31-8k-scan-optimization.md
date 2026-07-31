# 8-K 掃描效率優化 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 讓 `fetch_nongaap_statements()` 的下載量與「使用者要幾季」成正比，而非與「公司發過幾份 8-K」成正比。

**Architecture:** SEC 的申報清單本身就帶 `items`（8-K 項目代號）與 `period_of_report`（期間結束日）兩個欄位，取得清單即有，無需下載檔案。把「挑財報 → 算季度標籤 → 去重 → 套年份 → 切 max_filings」整段搬到清單階段完成，只對真正要用的申報呼叫 `filing.obj()`。清單篩選可能漏掉未標 `2.02` 的財報，因此在季度標籤序列上偵測缺口，僅針對缺口季度回退逐筆下載檢查。

**Tech Stack:** Python 3.12、edgartools（`EntityFilings` / `EntityFiling`）、pytest

## Global Constraints

- 對外簽章 `fetch_nongaap_statements(ticker, identity, ai_config, output_dir, progress_cb=None, max_filings=80, start_year=None, end_year=None)` 不得變動，GUI 呼叫端不修改
- `nongaap_cache.json` 格式不變（`{ticker: {quarter_label: {filing_date, eps_recon, metrics}}}`）
- 季度標籤格式為 `FY2024Q1`，一律用既有的 `_period_to_quarter_label()` 產生，不另寫轉換邏輯
- 例外處理只印 `type(exc).__name__` 與 `_exc_status(exc)`，**禁止** `f"{exc}"` 或 `{exc!r}`（LLM/HTTP SDK 的例外訊息挾帶 URL 與金鑰）
- 不修改 `fetcher_gaap.py`、`main.py`、`excel_writer.py`、`excel_formatter.py`
- 新測試放在既有的 `tests/test_fetcher_nongaap.py`，不新建測試檔
- 連網測試一律加 `@pytest.mark.slow`
- 測試指令為 `python -m pytest`（此機器 `pytest` 不在 PATH）

---

## 背景：edgartools 已驗證的物件行為

實作前先知道這兩件事，否則會誤以為需要下載：

```python
f = edgar.Company('ARLO').get_filings(form='8-K')   # EntityFilings
x = f[0]                                            # EntityFiling
x.items              # '5.07'          ← 逗號分隔字串，如 '2.02,9.01'
x.period_of_report   # '2026-06-18'    ← 字串
x.filing_date        # datetime.date(2026, 6, 23)
x.obj()              # ← 只有這個會下載
```

`get_filings()` 回傳**新→舊**排序。

## File Structure

| 檔案 | 動作 | 責任 |
|---|---|---|
| `fetcher_nongaap.py` | 修改 | 移除 `_get_earnings_filings()`；新增 `_list_earnings_filings()`（清單階段篩選）、`_find_missing_quarters()`（缺季偵測）、`_recover_missing_quarters()`（缺口補掃）；改寫 `fetch_nongaap_statements()` 的下載時機 |
| `tests/test_fetcher_nongaap.py` | 修改 | 新增假申報物件與 10 個測試 |

不新建檔案。`fetcher_nongaap.py` 目前 500 餘行、職責單一（8-K → Non-GAAP），不需要拆分。

---

### Task 1: 清單階段挑出財報 8-K

**Files:**
- Modify: `fetcher_nongaap.py`（在 `_filter_nongaap_by_year` 之後、`_get_earnings_filings` 之前插入新函式，約 445 行處）
- Test: `tests/test_fetcher_nongaap.py`

**Interfaces:**
- Consumes: 既有 `_period_to_quarter_label(period_of_report: str) -> str`、`_filter_nongaap_by_year(filings: list[tuple], start_year, end_year) -> list[tuple]`
- Produces: `_list_earnings_filings(company, start_year: int | None = None, end_year: int | None = None, max_filings: int = 80) -> list[tuple[str, Any]]` — 回傳 `(quarter_label, filing)`，**新→舊**排序，未下載任何檔案

- [ ] **Step 1: 寫失敗測試**

加到 `tests/test_fetcher_nongaap.py` 檔尾。先加 import 與假物件（後續 Task 共用）：

```python
import pytest
from fetcher_nongaap import _list_earnings_filings


class FakeFiling:
    """Stand-in for edgartools EntityFiling — only the listing-level attributes."""
    def __init__(self, items, period_of_report, accession="0000000000-00-000000"):
        self.items = items
        self.period_of_report = period_of_report
        self.accession_no = accession
        self.filing_date = period_of_report
        self.obj_called = False

    def obj(self):
        self.obj_called = True
        raise AssertionError("obj() must not be called during listing-stage filtering")


class FakeCompany:
    """Stand-in for edgartools Company. Filings are newest-first, as EDGAR returns them."""
    def __init__(self, filings):
        self._filings = filings

    def get_filings(self, **kwargs):
        return self._filings


def test_list_earnings_filings_keeps_only_item_202():
    company = FakeCompany([
        FakeFiling("2.02,9.01", "2024-09-30"),
        FakeFiling("5.07",      "2024-08-15"),
        FakeFiling("8.01,9.01", "2024-07-01"),
        FakeFiling("2.02",      "2024-06-30"),
    ])
    result = _list_earnings_filings(company)
    assert [label for label, _ in result] == ["FY2024Q3", "FY2024Q2"]


def test_list_earnings_filings_never_downloads():
    """Listing-stage filtering must not call obj() — that is the whole point."""
    filings = [
        FakeFiling("2.02,9.01", "2024-09-30"),
        FakeFiling("5.07",      "2024-08-15"),
    ]
    _list_earnings_filings(FakeCompany(filings))
    assert all(f.obj_called is False for f in filings)


def test_list_earnings_filings_tolerates_missing_items():
    """items may be None or empty on some filings — must not raise."""
    company = FakeCompany([
        FakeFiling(None, "2024-09-30"),
        FakeFiling("",   "2024-06-30"),
        FakeFiling("2.02", "2024-03-31"),
    ])
    result = _list_earnings_filings(company)
    assert [label for label, _ in result] == ["FY2024Q1"]


def test_list_earnings_filings_skips_malformed_period():
    company = FakeCompany([
        FakeFiling("2.02", None),
        FakeFiling("2.02", ""),
        FakeFiling("2.02", "2024-03-31"),
    ])
    result = _list_earnings_filings(company)
    assert [label for label, _ in result] == ["FY2024Q1"]


def test_list_earnings_filings_dedupes_keeping_oldest():
    """Same quarter filed twice (e.g. 8-K then corrected 8-K): keep the oldest filing."""
    oldest = FakeFiling("2.02", "2024-09-30", accession="OLDEST")
    newer  = FakeFiling("2.02", "2024-09-30", accession="NEWER")
    company = FakeCompany([newer, oldest])   # newest-first, as EDGAR returns
    result = _list_earnings_filings(company)
    assert len(result) == 1
    assert result[0][1].accession_no == "OLDEST"


def test_list_earnings_filings_applies_year_range():
    company = FakeCompany([
        FakeFiling("2.02", "2024-03-31"),
        FakeFiling("2.02", "2023-03-31"),
        FakeFiling("2.02", "2022-03-31"),
        FakeFiling("2.02", "2021-03-31"),
    ])
    result = _list_earnings_filings(company, start_year=2022, end_year=2023)
    assert [label for label, _ in result] == ["FY2023Q1", "FY2022Q1"]


def test_list_earnings_filings_max_filings_keeps_newest():
    company = FakeCompany([
        FakeFiling("2.02", "2024-12-31"),
        FakeFiling("2.02", "2024-09-30"),
        FakeFiling("2.02", "2024-06-30"),
        FakeFiling("2.02", "2024-03-31"),
    ])
    result = _list_earnings_filings(company, max_filings=2)
    assert [label for label, _ in result] == ["FY2024Q4", "FY2024Q3"]


def test_list_earnings_filings_year_filter_runs_before_max_filings():
    """Year range narrows the pool first; max_filings then trims the narrowed pool."""
    company = FakeCompany([
        FakeFiling("2.02", "2024-12-31"),
        FakeFiling("2.02", "2024-09-30"),
        FakeFiling("2.02", "2023-12-31"),
        FakeFiling("2.02", "2023-09-30"),
    ])
    result = _list_earnings_filings(company, end_year=2023, max_filings=1)
    assert [label for label, _ in result] == ["FY2023Q4"]
```

- [ ] **Step 2: 執行測試確認失敗**

Run: `python -m pytest tests/test_fetcher_nongaap.py -k list_earnings_filings -v`
Expected: FAIL — `ImportError: cannot import name '_list_earnings_filings' from 'fetcher_nongaap'`

- [ ] **Step 3: 寫最小實作**

在 `fetcher_nongaap.py` 的 `_filter_nongaap_by_year()` 之後插入：

```python
def _list_earnings_filings(
    company,
    start_year: int | None = None,
    end_year: int | None = None,
    max_filings: int = 80,
) -> list[tuple[str, Any]]:
    """Return [(quarter_label, filing)] for earnings 8-Ks, newest first.

    Filters entirely on listing metadata (``items`` and ``period_of_report``),
    which EDGAR supplies with the filing index — no document is downloaded here.
    Callers download only the filings they actually need.

    Item 2.02 is "Results of Operations and Financial Condition", i.e. the
    earnings release. SEC adopted that numbering on 2004-08-23; earlier filings
    used Item 12 or Item 5 and are not matched. See the design doc for why that
    is acceptable (max_filings defaults to 80 quarters ≈ 20 years).
    """
    candidates: list[tuple[str, Any]] = []
    for filing in company.get_filings(form="8-K", amendments=False):
        items = str(getattr(filing, "items", "") or "")
        if "2.02" not in items:
            continue
        period = str(getattr(filing, "period_of_report", "") or "").replace("-", "")
        if len(period) < 8:
            continue
        candidates.append((_period_to_quarter_label(period), filing))

    # Dedupe by quarter, keeping the oldest filing for each — matches prior behaviour
    # where a corrected re-filing does not displace the original release.
    seen: set[str] = set()
    deduped: list[tuple[str, Any]] = []
    for label, filing in reversed(candidates):      # oldest → newest
        if label not in seen:
            seen.add(label)
            deduped.append((label, filing))
    deduped.reverse()                               # back to newest → oldest

    deduped = _filter_nongaap_by_year(deduped, start_year, end_year)
    return deduped[:max_filings]
```

- [ ] **Step 4: 執行測試確認通過**

Run: `python -m pytest tests/test_fetcher_nongaap.py -k list_earnings_filings -v`
Expected: 8 passed

- [ ] **Step 5: Commit**

```bash
git add fetcher_nongaap.py tests/test_fetcher_nongaap.py
git commit -m "feat: 新增 _list_earnings_filings，於申報清單階段篩選財報 8-K 不下載檔案"
```

---

### Task 2: 缺季偵測

**Files:**
- Modify: `fetcher_nongaap.py`（緊接 `_list_earnings_filings` 之後）
- Test: `tests/test_fetcher_nongaap.py`

**Interfaces:**
- Consumes: 無（純字串運算）
- Produces: `_find_missing_quarters(labels: list[str]) -> list[str]` — 輸入季度標籤（順序不拘），回傳最舊與最新之間缺少的標籤，**由舊到新**排序

- [ ] **Step 1: 寫失敗測試**

```python
from fetcher_nongaap import _find_missing_quarters


def test_find_missing_quarters_none_missing():
    assert _find_missing_quarters(["FY2024Q3", "FY2024Q2", "FY2024Q1"]) == []


def test_find_missing_quarters_single_gap():
    assert _find_missing_quarters(["FY2024Q3", "FY2024Q1"]) == ["FY2024Q2"]


def test_find_missing_quarters_spans_year_boundary():
    assert _find_missing_quarters(["FY2025Q1", "FY2024Q3"]) == ["FY2024Q4"]


def test_find_missing_quarters_ignores_outside_range():
    """Gaps only exist between the oldest and newest label — never before or after."""
    assert _find_missing_quarters(["FY2024Q2", "FY2024Q3"]) == []


def test_find_missing_quarters_multiple_gaps():
    assert _find_missing_quarters(["FY2024Q4", "FY2024Q2", "FY2023Q4"]) == [
        "FY2024Q1", "FY2024Q3",
    ]


def test_find_missing_quarters_handles_empty_and_single():
    assert _find_missing_quarters([]) == []
    assert _find_missing_quarters(["FY2024Q1"]) == []


def test_find_missing_quarters_ignores_unparseable_labels():
    assert _find_missing_quarters(["FY2024Q3", "GARBAGE", "FY2024Q1"]) == ["FY2024Q2"]
```

- [ ] **Step 2: 執行測試確認失敗**

Run: `python -m pytest tests/test_fetcher_nongaap.py -k find_missing_quarters -v`
Expected: FAIL — `ImportError: cannot import name '_find_missing_quarters'`

- [ ] **Step 3: 寫最小實作**

```python
def _quarter_ordinal(label: str) -> int | None:
    """Convert 'FY2024Q3' to a sortable integer (2024*4 + 2). None if unparseable."""
    m = re.fullmatch(r"FY(\d{4})Q([1-4])", label.strip())
    if m is None:
        return None
    return int(m.group(1)) * 4 + (int(m.group(2)) - 1)


def _ordinal_to_quarter(ordinal: int) -> str:
    """Inverse of _quarter_ordinal."""
    return f"FY{ordinal // 4}Q{ordinal % 4 + 1}"


def _find_missing_quarters(labels: list[str]) -> list[str]:
    """Return quarter labels absent between the oldest and newest supplied label.

    A gap means the listing-stage filter missed an earnings release — usually an
    8-K that omitted Item 2.02. Nothing outside the supplied span counts as a gap:
    a company simply has no filings before its IPO or after its latest report.
    """
    ordinals = sorted(o for o in (_quarter_ordinal(x) for x in labels) if o is not None)
    if len(ordinals) < 2:
        return []
    present = set(ordinals)
    return [
        _ordinal_to_quarter(o)
        for o in range(ordinals[0], ordinals[-1] + 1)
        if o not in present
    ]
```

- [ ] **Step 4: 執行測試確認通過**

Run: `python -m pytest tests/test_fetcher_nongaap.py -k find_missing_quarters -v`
Expected: 7 passed

- [ ] **Step 5: Commit**

```bash
git add fetcher_nongaap.py tests/test_fetcher_nongaap.py
git commit -m "feat: 新增 _find_missing_quarters 偵測季度序列缺口"
```

---

### Task 3: 缺口補掃

**Files:**
- Modify: `fetcher_nongaap.py`（緊接 `_find_missing_quarters` 之後）
- Test: `tests/test_fetcher_nongaap.py`

**Interfaces:**
- Consumes: `_period_to_quarter_label()`、`_exc_status()`（已由 `from errsafe import _exc_status` 匯入）
- Produces: `_recover_missing_quarters(company, missing: list[str]) -> list[tuple[str, Any]]` — 只對期間落在 `missing` 且未標 `2.02` 的申報呼叫 `obj()`，用 `has_earnings` 判斷是否為財報，回傳 `(quarter_label, filing)`

- [ ] **Step 1: 寫失敗測試**

`FakeFiling.obj()` 在 Task 1 是設計成拋錯的，這裡需要一個會回傳物件的版本：

```python
from fetcher_nongaap import _recover_missing_quarters


class FakeEightK:
    def __init__(self, has_earnings):
        self.has_earnings = has_earnings


class RecoverableFiling:
    """FakeFiling variant whose obj() returns a parsed 8-K instead of raising."""
    def __init__(self, items, period_of_report, has_earnings, accession="ACC"):
        self.items = items
        self.period_of_report = period_of_report
        self.accession_no = accession
        self.filing_date = period_of_report
        self._has_earnings = has_earnings
        self.obj_calls = 0

    def obj(self):
        self.obj_calls += 1
        return FakeEightK(self._has_earnings)


def test_recover_missing_quarters_finds_mislabelled_earnings():
    target = RecoverableFiling("8.01,9.01", "2024-06-30", has_earnings=True)
    company = FakeCompany([
        RecoverableFiling("2.02", "2024-09-30", has_earnings=True),
        target,
        RecoverableFiling("2.02", "2024-03-31", has_earnings=True),
    ])
    result = _recover_missing_quarters(company, ["FY2024Q2"])
    assert [label for label, _ in result] == ["FY2024Q2"]
    assert result[0][1] is target


def test_recover_missing_quarters_only_downloads_gap_candidates():
    """Filings outside the gap, or already tagged 2.02, must not be downloaded."""
    in_gap     = RecoverableFiling("5.07",  "2024-06-30", has_earnings=True)
    tagged     = RecoverableFiling("2.02",  "2024-06-30", has_earnings=True)
    other_qtr  = RecoverableFiling("8.01",  "2024-09-30", has_earnings=True)
    _recover_missing_quarters(FakeCompany([other_qtr, tagged, in_gap]), ["FY2024Q2"])
    assert in_gap.obj_calls == 1
    assert tagged.obj_calls == 0
    assert other_qtr.obj_calls == 0


def test_recover_missing_quarters_ignores_non_earnings():
    company = FakeCompany([RecoverableFiling("5.07", "2024-06-30", has_earnings=False)])
    assert _recover_missing_quarters(company, ["FY2024Q2"]) == []


def test_recover_missing_quarters_empty_gap_list_downloads_nothing():
    f = RecoverableFiling("5.07", "2024-06-30", has_earnings=True)
    assert _recover_missing_quarters(FakeCompany([f]), []) == []
    assert f.obj_calls == 0


def test_recover_missing_quarters_survives_download_failure():
    """A filing that fails to parse must not abort recovery of the others."""
    class ExplodingFiling(RecoverableFiling):
        def obj(self):
            self.obj_calls += 1
            raise ValueError("parse failed")

    good = RecoverableFiling("5.07", "2024-06-30", has_earnings=True)
    bad  = ExplodingFiling("8.01", "2024-06-30", has_earnings=True, accession="BAD")
    result = _recover_missing_quarters(FakeCompany([bad, good]), ["FY2024Q2"])
    assert [label for label, _ in result] == ["FY2024Q2"]
    assert result[0][1] is good
```

- [ ] **Step 2: 執行測試確認失敗**

Run: `python -m pytest tests/test_fetcher_nongaap.py -k recover_missing_quarters -v`
Expected: FAIL — `ImportError: cannot import name '_recover_missing_quarters'`

- [ ] **Step 3: 寫最小實作**

```python
def _recover_missing_quarters(company, missing: list[str]) -> list[tuple[str, Any]]:
    """Deep-scan only the quarters the listing filter came up short on.

    Downloads a filing only when its period falls in a missing quarter and it was
    not already tagged Item 2.02 — typically a handful of filings, versus the
    hundreds a full historical scan would fetch.
    """
    if not missing:
        return []

    wanted = set(missing)
    found: dict[str, Any] = {}
    for filing in company.get_filings(form="8-K", amendments=False):
        items = str(getattr(filing, "items", "") or "")
        if "2.02" in items:
            continue
        period = str(getattr(filing, "period_of_report", "") or "").replace("-", "")
        if len(period) < 8:
            continue
        label = _period_to_quarter_label(period)
        if label not in wanted or label in found:
            continue
        try:
            eight_k = filing.obj()
        except Exception as exc:
            print(
                f"[fetcher_nongaap] gap scan {label} -> "
                f"{type(exc).__name__}{_exc_status(exc)}",
                file=sys.stderr,
            )
            continue
        if getattr(eight_k, "has_earnings", False):
            found[label] = filing

    return [(label, found[label]) for label in sorted(found, key=_quarter_ordinal)]
```

- [ ] **Step 4: 執行測試確認通過**

Run: `python -m pytest tests/test_fetcher_nongaap.py -k recover_missing_quarters -v`
Expected: 5 passed

- [ ] **Step 5: Commit**

```bash
git add fetcher_nongaap.py tests/test_fetcher_nongaap.py
git commit -m "feat: 新增 _recover_missing_quarters，僅對缺季區間回退深度掃描"
```

---

### Task 4: 接上 fetch_nongaap_statements 並移除舊掃描

**Files:**
- Modify: `fetcher_nongaap.py:447-484`（刪除 `_get_earnings_filings`）與 `fetcher_nongaap.py:508-533`（`fetch_nongaap_statements` 內的取件與迴圈）
- Test: `tests/test_fetcher_nongaap.py`

**Interfaces:**
- Consumes: `_list_earnings_filings()`、`_find_missing_quarters()`、`_recover_missing_quarters()`（Task 1–3）
- Produces: `fetch_nongaap_statements()` 對外簽章與回傳型別不變

- [ ] **Step 1: 確認沒有其他呼叫端**

Run: `grep -rn "_get_earnings_filings" --include="*.py" . | grep -v venv`
Expected: 只剩 `fetcher_nongaap.py` 內的定義與呼叫，加上 `tests/test_fetcher_nongaap.py:134` 與 `:142` 兩行**註解**（那個測試自己複製了去重邏輯，不匯入該函式，故不受影響）。若出現其他 `.py` 的實際呼叫，停下來回報。

- [ ] **Step 2: 寫失敗測試**

```python
def test_fetch_nongaap_downloads_only_uncached_quarters(tmp_path, monkeypatch):
    """The whole point: one obj() call per quarter actually needed, and none for
    quarters already in the cache."""
    import fetcher_nongaap as fn

    q3 = RecoverableFiling("2.02", "2024-09-30", has_earnings=True)
    q2 = RecoverableFiling("2.02", "2024-06-30", has_earnings=True)
    q1 = RecoverableFiling("2.02", "2024-03-31", has_earnings=True)
    company = FakeCompany([q3, q2, q1])

    monkeypatch.setattr(fn, "set_identity", lambda *a, **k: None)
    monkeypatch.setattr(fn, "Company", lambda ticker: company)
    monkeypatch.setattr(fn, "_extract_eps_recon", lambda ek: {})
    monkeypatch.setattr(fn, "_extract_nongaap_metrics", lambda ek, cfg: {"Non-GAAP EPS": 1.0})

    # FY2024Q1 is already cached — it must not be downloaded again
    fn._save_cache(tmp_path / fn.CACHE_FILENAME, "TEST",
                   {"FY2024Q1": {"filing_date": "2024-04-30", "eps_recon": {}, "metrics": {}}})

    fn.fetch_nongaap_statements("TEST", "CTH x@y.com", {"api_key": "k"}, tmp_path)

    assert q3.obj_calls == 1
    assert q2.obj_calls == 1
    assert q1.obj_calls == 0


def test_fetch_nongaap_writes_metrics_to_cache(tmp_path, monkeypatch):
    import fetcher_nongaap as fn

    company = FakeCompany([RecoverableFiling("2.02", "2024-09-30", has_earnings=True)])
    monkeypatch.setattr(fn, "set_identity", lambda *a, **k: None)
    monkeypatch.setattr(fn, "Company", lambda ticker: company)
    monkeypatch.setattr(fn, "_extract_eps_recon", lambda ek: {})
    monkeypatch.setattr(fn, "_extract_nongaap_metrics", lambda ek, cfg: {"Non-GAAP EPS": 2.5})

    fn.fetch_nongaap_statements("TEST", "CTH x@y.com", {"api_key": "k"}, tmp_path)

    cached = fn._load_cache(tmp_path / fn.CACHE_FILENAME, "TEST")
    assert cached["FY2024Q3"]["metrics"]["Non-GAAP EPS"] == 2.5
```

- [ ] **Step 3: 執行測試確認失敗**

Run: `python -m pytest tests/test_fetcher_nongaap.py -k "fetch_nongaap_downloads or fetch_nongaap_writes" -v`
Expected: FAIL — 目前 `fetch_nongaap_statements` 走 `_get_earnings_filings()`，該函式對每份申報都呼叫 `obj()`，因此 `q1.obj_calls` 會是 1 而非 0

- [ ] **Step 4: 刪除 `_get_earnings_filings`**

刪掉 `fetcher_nongaap.py:447-484` 整個函式（自 `def _get_earnings_filings(company)` 起，至 `return list(reversed(deduped))` 止）。它已被 Task 1–3 取代：清單篩選由 `_list_earnings_filings` 負責，逐筆深掃只在缺季時由 `_recover_missing_quarters` 執行。

- [ ] **Step 5: 改寫取件段落**

把 `fetch_nongaap_statements()` 內原本這兩行：

```python
    filings = _get_earnings_filings(company)[:max_filings]  # newest max_filings quarters only
    filings = _filter_nongaap_by_year(filings, start_year, end_year)
```

換成：

```python
    filings = _list_earnings_filings(company, start_year, end_year, max_filings)

    # A hole in the quarter sequence means the listing filter missed a release —
    # scan just that range rather than every 8-K the company ever filed.
    missing = _find_missing_quarters([label for label, _ in filings])
    if missing:
        recovered = _recover_missing_quarters(company, missing)
        if recovered:
            filings = sorted(
                filings + recovered,
                key=lambda item: _quarter_ordinal(item[0]) or 0,
                reverse=True,
            )
        still_missing = sorted(set(missing) - {label for label, _ in recovered})
        if still_missing:
            print(
                f"[fetcher_nongaap] {ticker} 無此季財報 8-K: {', '.join(still_missing)}",
                file=sys.stderr,
            )
```

- [ ] **Step 6: 改寫下載迴圈**

把原本的：

```python
    new_filings = [(lbl, f, ek) for lbl, f, ek in filings if lbl not in cache]
    total = len(new_filings)

    for i, (quarter_label, filing, eight_k) in enumerate(new_filings, 1):
        if progress_cb:
            progress_cb(i, total, f"Non-GAAP {ticker} {quarter_label} ({i}/{total})")

        try:
            eps_recon = _extract_eps_recon(eight_k)
```

換成（差異在於 tuple 只有兩元素，且 `obj()` 移進 try 內）：

```python
    new_filings = [(lbl, f) for lbl, f in filings if lbl not in cache]
    total = len(new_filings)

    for i, (quarter_label, filing) in enumerate(new_filings, 1):
        if progress_cb:
            progress_cb(i, total, f"Non-GAAP {ticker} {quarter_label} ({i}/{total})")

        try:
            eight_k = filing.obj()      # the only download in this loop
            eps_recon = _extract_eps_recon(eight_k)
```

該 `try` 區塊其餘部分（`metrics = ...` 起至 `_save_cache(...)` 與 `except` 分支）不動。

- [ ] **Step 7: 更新 docstring**

`fetch_nongaap_statements()` 的 docstring 中 `max_filings` 那行改為：

```python
        max_filings: Max number of earnings quarters to process (newest first, default 80).
                     Applied after the year range narrows the pool.
```

- [ ] **Step 8: 執行新測試確認通過**

Run: `python -m pytest tests/test_fetcher_nongaap.py -k "fetch_nongaap_downloads or fetch_nongaap_writes" -v`
Expected: 2 passed

- [ ] **Step 9: 執行全套測試確認無回歸**

Run: `python -m pytest tests/ --ignore=tests/test_live_snapshots.py -q`
Expected: 全數 PASSED（改動前基準為 250 passed，本次新增 22 個測試，應為 272 passed）。若有 FAILED，停下來回報，不要修改測試去迎合實作。

- [ ] **Step 10: Commit**

```bash
git add fetcher_nongaap.py tests/test_fetcher_nongaap.py
git commit -m "perf: 8-K 改為清單階段篩選，只下載需要的季度並移除全量掃描"
```

---

### Task 5: 連網驗收

**Files:**
- Test: `tests/test_fetcher_nongaap.py`

**Interfaces:**
- Consumes: Task 1–4 的全部函式
- Produces: 無新程式碼，僅測試

- [ ] **Step 1: 寫連網測試**

```python
@pytest.mark.slow
def test_live_listing_filter_matches_deep_scan_arlo():
    """The listing filter must not lose quarters a full scan would have found.

    ARLO is small (67 8-Ks total) so the deep comparison stays under a minute.
    """
    import edgar
    import config as cfgmod
    from fetcher_nongaap import _list_earnings_filings

    cfg = cfgmod.load_config()
    if not cfg.get("identity"):
        pytest.skip("no SEC identity configured")
    edgar.set_identity(cfg["identity"])

    company = edgar.Company("ARLO")
    fast_labels = {label for label, _ in _list_earnings_filings(company)}

    deep_labels = set()
    for filing in company.get_filings(form="8-K", amendments=False):
        period = str(filing.period_of_report or "").replace("-", "")
        if len(period) < 8:
            continue
        items = str(getattr(filing, "items", "") or "")
        if "2.02" in items:
            deep_labels.add(_period_to_quarter_label(period))
            continue
        try:
            if getattr(filing.obj(), "has_earnings", False):
                deep_labels.add(_period_to_quarter_label(period))
        except Exception:
            continue

    assert deep_labels - fast_labels == set(), (
        f"listing filter missed quarters the deep scan found: {deep_labels - fast_labels}"
    )


@pytest.mark.slow
def test_live_year_range_limits_download_count_crm():
    """Asking for two quarters must download two filings, not CRM's 290 8-Ks."""
    import edgar
    import config as cfgmod
    from fetcher_nongaap import _list_earnings_filings

    cfg = cfgmod.load_config()
    if not cfg.get("identity"):
        pytest.skip("no SEC identity configured")
    edgar.set_identity(cfg["identity"])

    result = _list_earnings_filings(edgar.Company("CRM"), max_filings=2)
    assert len(result) == 2
    # CRM's fiscal year ends in January, so Q labels must not all be Q4
    assert len({label for label, _ in result}) == 2
```

- [ ] **Step 2: 執行連網測試**

Run: `python -m pytest tests/test_fetcher_nongaap.py -m slow -k "live_listing or live_year_range" -v`
Expected: 2 passed。若 `test_live_listing_filter_matches_deep_scan_arlo` 失敗，代表 ARLO 有未標 2.02 的財報且缺季偵測沒補到——回報實際缺的季度，不要放寬斷言。

- [ ] **Step 3: 實跑三家計時**

Run:

```bash
python -c "
import time, config as cfgmod
from pathlib import Path
from fetcher_nongaap import fetch_nongaap_statements
cfg = cfgmod.load_config()
for t in ['CRM', 'PANW', 'ARLO']:
    t0 = time.time()
    tables = fetch_nongaap_statements(t, cfg['identity'], cfg['ai'], Path('output'),
                                      max_filings=4)
    print(f'{t}: {time.time()-t0:.1f} 秒, sheets={[x.sheet_name for x in tables]}')
"
```

Expected: 每家 60 秒內完成（含 4 次 AI 呼叫），並產生 `Data_NonGAAP`。實際秒數記錄下來，寫進 CHANGELOG。

- [ ] **Step 4: Commit**

```bash
git add tests/test_fetcher_nongaap.py
git commit -m "test: 新增 8-K 清單篩選與下載量的連網驗收測試"
```

---

### Task 6: 文件更新

**Files:**
- Modify: `CHANGELOG.md`（在「更新記錄」下方插入新段落）
- Modify: `TODO.md`（移除第 2 項）
- Modify: `ARCHITECTURE.md`（Non-GAAP 資料流段落）

**Interfaces:**
- Consumes: Task 5 Step 3 量到的實際秒數
- Produces: 無程式碼

- [ ] **Step 1: 寫 CHANGELOG**

在 `CHANGELOG.md` 的 `## 更新記錄` 底下、既有最新日期段落之前插入：

```markdown
### 2026-07-31

**8-K 掃描效率優化**

設計文件：`docs/superpowers/specs/2026-07-31-8k-scan-optimization-design.md`

- **`fetcher_nongaap.py`**：
  - 新增 `_list_earnings_filings()`：改在 SEC 申報清單階段以 `items`（Item 2.02）與 `period_of_report` 完成篩選、去重、年份過濾與 `max_filings` 切割，全程不下載檔案
  - 新增 `_quarter_ordinal()` / `_ordinal_to_quarter()` / `_find_missing_quarters()`：偵測季度序列缺口
  - 新增 `_recover_missing_quarters()`：僅對缺季區間回退逐筆 `obj()` 深掃，用 `has_earnings` 找回未標 2.02 的財報；補不到的季度寫 stderr
  - 移除 `_get_earnings_filings()`（對全部歷史 8-K 逐筆下載）
  - `fetch_nongaap_statements()`：`obj()` 移進迴圈，只對未快取的季度下載；年份過濾改在 `max_filings` 之前套用
- **實測**：AAPL 全部 8-K 235 份、含 2.02 者 94 份；抓 4 季的下載次數由 235 降至 4
- **已知邊界**：SEC 自 2004-08-23 才啟用 2.02 編號，更早的財報 8-K（Item 12/5）不會被抓到
- **`tests/test_fetcher_nongaap.py`**：新增 22 個測試（清單篩選 8、缺季偵測 7、缺口補掃 5、下載時機 2）+ 2 個 `slow` 連網驗收
```

- [ ] **Step 2: 更新 TODO**

刪除 `TODO.md` 第 2 項（8-K 掃描效率優化），其餘項目重新編號為 1–5。

- [ ] **Step 3: 更新 ARCHITECTURE**

Run: `grep -n "8-K\|nongaap\|Non-GAAP" ARCHITECTURE.md`

找到描述 Non-GAAP 資料流的段落，把「掃描全部 8-K」的敘述改為「清單階段以 Item 2.02 篩選，只下載需要的季度；偵測到缺季才對該區間深掃」。若原文沒有描述掃描方式，則在 Non-GAAP 段落補一句說明。

- [ ] **Step 4: Commit**

```bash
git add CHANGELOG.md TODO.md ARCHITECTURE.md
git commit -m "docs: 記錄 8-K 掃描效率優化並更新 TODO 與架構文件"
```
