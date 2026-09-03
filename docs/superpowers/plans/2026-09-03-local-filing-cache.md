# 本地 filing 快取 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 讓 `fetcher_gaap._filing_obj()` 先讀本機磁碟快取，只有 SEC 上新出現的 filing 才真的打網路解析，並在 Tab3 提供快取容量檢視與清除。

**Architecture:** 新增獨立模組 `src/filing_cache.py`，負責路徑、原子寫入、三張 DataFrame 的序列化／還原、替身物件（`_CachedFiling`／`_CachedFinancials`／`_CachedStatement`）、manifest 重建、GUI 用的容量統計。`fetcher_gaap` 只在 `_filing_obj()` 一個掛勾點呼叫它（IS/BS/CF/segment 四個 builder 共用這個函式，改一處全部受益），並用一個 `_disk_cache_scope()` context manager 把「這趟是哪家公司、cik 多少、命中幾份」綁在一次抓取內。`main.py` 在 Tab3 加一塊自帶固定高度捲動的「本地資料快取」區塊。

**Tech Stack:** Python 3.13、pandas、edgartools 5.29.0、tkinter/ttk、pytest、`json` + `os.replace()`（不引入資料庫、不新增第三方套件）

**Spec:** `docs/superpowers/specs/2026-09-03-local-filing-cache-design.md`

## Global Constraints

- **不新增任何第三方依賴**。只用 stdlib（`json`／`os`／`pathlib`／`importlib.metadata`／`shutil`）＋既有的 pandas。
- **不改變任何輸出內容**。快取是純加速層，`fetch_gaap_statements()` 的回傳值、Excel 逐格內容、呼叫端（Tab1／批次／`comparison.py`）一律不動。
- 快取根目錄：`%APPDATA%\SEC Financial Tools\filing_cache\`（沿用 `config.py` 的**有空格**版本，不是 `override_engine.py` 的 `SEC_Financial_Tools`）；沒有 `%APPDATA%` 時退回 `~/.sec_financial_tools/filing_cache`。
- 一個 ticker 一個資料夾（大寫），一份 filing 一個 `<accession>.json`；accession 格式必須符合 `^\d{10}-\d{2}-\d{6}$`，不符就不當快取鍵（防路徑注入）。
- `SCHEMA_VERSION = 1`（filing 檔與 manifest 各自帶）。
- edgartools 版本一律用 `importlib.metadata.version("edgartools")` 取得；**不可用 `edgar.__version__`（該屬性不存在）**。取不到版本（回 `None`）就整個停用快取（不讀也不寫），不得寫入 `"unknown"` 之類的預設值。
- 所有 JSON 寫入（filing 檔與 manifest）都走 tmp + `os.replace()`，tmp 檔名帶 PID。
- **`<accession>.json` 是否存在才是事實來源**；`_manifest.json` 只是衍生索引，壞掉／遺失／對不上一律從資料夾內容重建。
- 快取讀取前必須同時通過四道閘：JSON 可解析、`schema_version` 相符、`cik` 相符、`edgartools_version` 相符。任一不符 → 視同無快取，正常重抓，**不得拋例外**。
- 網路失敗（`NetworkDownError` 或任何例外）**絕對不可寫入任何快取檔**（正向或負向都不行）。只有「`filing.obj()` 成功但 `financials` 是 `None`」才寫 `has_financials: false` 的負向快取。
- 寫入是**逐份即時落檔**，不累積到整趟結束。
- 快取寫入失敗（磁碟滿／權限）只記 log 繼續跑，不跳提示、不中斷抓取。
- `src/` 內不得出現寫死的中日文字面（`tests/test_i18n.py` 第 3 條會擋）；新增的畫面字串一律進四個 locale（`zh_tw`／`zh_cn`／`en`／`ja`）的 `gui.*` key。
- `logs/app.log` 的訊息一律英文（2026-09-02 起的既有規則）。
- Tab3 是固定高度捲動容器（`main.py` `_TAB3_HEIGHT = 355`）；新區塊必須自帶第二層固定高度捲動（`_build_fixed_height_scrollable`，100~120px），加完要重新量測 `_TAB3_HEIGHT`。

## ⚠ 對 spec 的一處刻意偏離（實作前要讓 CTH 知道）

Spec §三 要求把 `cache 24/25` 放進 log 的**起始 `===` 行**。做不到：起始行在抓取**開始前**就寫出去，那時候還沒列 filing 清單，命中數不存在。本計畫改成在 GAAP 抓完後補一行獨立的 `[INFO ]` 快取行（Task 9），語意相同、位置不同。其他規格照 spec 不變。

---

## File Structure

| 檔案 | 責任 |
|------|------|
| `src/filing_cache.py`（新增，約 300 行） | 快取的全部知識：路徑、原子寫入、DataFrame 序列化／還原、替身物件、單份 filing 讀寫、manifest 重建、GUI 統計與清除。**不 import `fetcher_gaap`**（單向依賴）。 |
| `src/fetcher_gaap.py`（修改） | 只加：`_disk_cache_scope()` / `_bind_disk_cache()` / `last_cache_stats()`，以及 `_filing_obj()` 內的查快取與落檔。約 60 行。 |
| `src/main.py`（修改） | Tab3 「本地資料快取」區塊、容量格式化純函式、清除按鈕鎖定、刷新時機、log 快取行。 |
| `src/locales/{zh_tw,zh_cn,en,ja}.py`（修改） | 新增的 `gui.*` 字串。 |
| `tests/test_filing_cache.py`（新增） | 序列化、替身物件、四道閘、負向快取、manifest 重建、原子寫入、統計與清除。 |
| `tests/test_fetcher_gaap_cache.py`（新增） | `_filing_obj()` 的整合行為：命中不打網路、未命中才打、網路失敗不留快取、查詢邊界不漏抓、逐份即時落檔。 |
| `tests/test_gui_helpers.py`（修改） | 容量格式化、清除鈕狀態這兩個純函式。 |
| `docs/ARCHITECTURE.md`（修改） | 新增「本地 filing 快取」一節 + 更新 `_TAB3_HEIGHT` 的實測值。 |

執行環境：所有指令用專案的 venv。
```bash
./venv/Scripts/python.exe -m pytest tests/ -q
```

---

## Task 1: DataFrame 序列化與還原

三張 DataFrame 存成 `json.loads(df.to_json(orient="split"))` 的**物件**（不是字串）＋ `dtypes` map。要能分辨「這張表不存在（`None`）」與「存在但是空表」。

**Files:**
- Create: `src/filing_cache.py`
- Test: `tests/test_filing_cache.py`

**Interfaces:**
- Consumes: 無（第一個任務）
- Produces:
  - `filing_cache.df_to_payload(df: pd.DataFrame | None) -> dict | None`
  - `filing_cache.payload_to_df(payload: dict | None) -> pd.DataFrame | None`
  - payload 形狀：`{"data": {"columns": [...], "index": [...], "data": [[...]]}, "dtypes": {"欄名": "dtype 字串"}}`

- [ ] **Step 1: 寫失敗的測試**

建立 `tests/test_filing_cache.py`：

```python
"""Tests for filing_cache.py — 本地 filing 快取的儲存層。

快取的是「解析層」的輸出（edgartools 解出來的三張 DataFrame），比對層
永遠在快取之上即時重跑。所以這裡釘的是「存進去再讀回來，跟原本一模一樣」，
以及「任何一種對不上的情況都要安靜地退回無快取，不能拋例外、不能餵錯資料」。
"""
import json
import os
from pathlib import Path

import pandas as pd
import pytest

import filing_cache


def _sample_df() -> pd.DataFrame:
    """一張像 edgartools 真的會吐出來的表：str / float64 / int64 / bool 四種
    dtype，外加一整欄都是 None——那一欄是 read_json 會推錯成 float64 的地雷。"""
    return pd.DataFrame({
        "concept":  ["us-gaap_Revenue", "us-gaap_NetIncomeLoss"],
        "label":    ["Net sales", "Net income"],
        "level":    [4, 3],
        "abstract": [False, False],
        "2025-12-27 (Q1)": [1000.0, 200.0],
        "dimension_member_label": [None, None],
    })


# ── 序列化：存進去讀回來要一模一樣 ────────────────────────────────────────

def test_payload_roundtrip_keeps_values_and_dtypes():
    df = _sample_df()
    back = filing_cache.payload_to_df(filing_cache.df_to_payload(df))
    pd.testing.assert_frame_equal(back, df, check_like=False)


def test_payload_is_a_json_object_not_an_escaped_string():
    """存 `df.to_json()` 的字串會讓整份內容被逃逸一次，檔案膨脹 10~15%
    而且文字編輯器打開完全不能看。要存解析過的物件。"""
    payload = filing_cache.df_to_payload(_sample_df())
    assert isinstance(payload["data"], dict)
    assert set(payload["data"]) >= {"columns", "index", "data"}


def test_all_null_column_keeps_its_original_dtype():
    """`to_json(orient="split")` 不含 dtype，整欄皆 null 會被推成 float64。
    這是 spike 實測抓到的唯一一個還原不回去的細節。"""
    df = _sample_df()
    back = filing_cache.payload_to_df(filing_cache.df_to_payload(df))
    assert back["dimension_member_label"].dtype == df["dimension_member_label"].dtype


def test_none_and_empty_dataframe_are_not_the_same_thing():
    """`is_stmt is None` 與「空表」在下游（`_current_q_col`）行為不同，
    存檔再讀回來不可以混成同一種。"""
    assert filing_cache.df_to_payload(None) is None
    assert filing_cache.payload_to_df(None) is None

    empty = _sample_df().iloc[0:0]
    back = filing_cache.payload_to_df(filing_cache.df_to_payload(empty))
    assert back is not None
    assert len(back) == 0
    assert list(back.columns) == list(empty.columns)
```

- [ ] **Step 2: 跑測試確認失敗**

Run: `./venv/Scripts/python.exe -m pytest tests/test_filing_cache.py -q`
Expected: FAIL — `ModuleNotFoundError: No module named 'filing_cache'`

- [ ] **Step 3: 寫最小實作**

建立 `src/filing_cache.py`：

```python
"""
filing_cache.py — 本地 filing 解析快取（%APPDATA%\\SEC Financial Tools\\filing_cache）。

快取卡在**解析層與比對層之間**：存的是 edgartools 解出來的三張 DataFrame
（income statement / balance sheet / cashflow statement），比對層
（`IS/BS/CF_TEMPLATE` 那套科目對照）永遠在快取之上即時重跑。所以以後改
hint regex、加比率、調 Q4 合成邏輯都不會讓快取失效——但 **edgartools 升版
會**，那是另一條軸線，靠 `edgartools_version` 欄位擋（見 `load_filing`）。

事實來源是 `<accession>.json` 檔案本身；`_manifest.json` 只是給 GUI 看的
衍生索引，壞了直接從資料夾重建。
"""
from __future__ import annotations

import json
import os
import re
from pathlib import Path

import pandas as pd

SCHEMA_VERSION = 1

# SEC 的 accession number 格式固定，拿來當檔名前先驗——這同時是路徑注入的防線。
ACCESSION_RE = re.compile(r"^\d{10}-\d{2}-\d{6}$")

STATEMENT_KEYS = ("income_statement", "balance_sheet", "cashflow_statement")


# ── DataFrame 序列化 ──────────────────────────────────────────────────────

def df_to_payload(df: pd.DataFrame | None) -> dict | None:
    """DataFrame → 可放進 JSON 的物件。`None` 原樣傳遞（代表這張表不存在）。

    存 `json.loads(...)` 的**物件**不是 `to_json()` 的字串：字串塞進外層
    JSON 會被整份逃逸一次，檔案膨脹 10~15%，而且打開來完全不能看。
    """
    if df is None:
        return None
    return {
        "data": json.loads(df.to_json(orient="split")),
        "dtypes": {str(col): str(dt) for col, dt in df.dtypes.items()},
    }


def payload_to_df(payload: dict | None) -> pd.DataFrame | None:
    """payload → DataFrame。`None` 原樣傳遞。

    `orient="split"` 不帶 dtype，整欄皆 null 的欄位會被推成 float64——所以
    一定要照存檔時記下的 `dtypes` 明確 `astype()` 回去，不能靠自動推斷。
    """
    if payload is None:
        return None
    raw = payload["data"]
    df = pd.DataFrame(raw["data"], index=raw["index"], columns=raw["columns"])
    for col, dtype in (payload.get("dtypes") or {}).items():
        if col not in df.columns:
            continue
        try:
            df[col] = df[col].astype(dtype)
        except (TypeError, ValueError):
            # 型別還原失敗不該讓整份快取報廢——數值本身是對的，
            # 下游只有極少數地方在意 dtype。
            pass
    return df
```

- [ ] **Step 4: 跑測試確認通過**

Run: `./venv/Scripts/python.exe -m pytest tests/test_filing_cache.py -q`
Expected: PASS（4 passed）

- [ ] **Step 5: Commit**

```bash
git add src/filing_cache.py tests/test_filing_cache.py
git commit -m "feat(cache): DataFrame 序列化與 dtype 還原（本地 filing 快取第一層）

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01BADSq3w2c82y1Ky5yWXD8M"
```

---

## Task 2: 替身物件（`_CachedFiling` / `_CachedFinancials` / `_CachedStatement`）

快取命中時 `_filing_obj()` 沒辦法生出真的 edgartools 物件，要回傳一個長得像它、
只實作有人真的在用的那幾條路徑的替身。**這是全案最容易卡住的部分，所以排在前面。**

兩條隱性規則必須各有一條測試釘住：未知屬性要拋 `AttributeError`（不可回 `None`）、
`None` 與空 DataFrame 不可混淆。

**Files:**
- Modify: `src/filing_cache.py`
- Test: `tests/test_filing_cache.py`

**Interfaces:**
- Consumes: `df_to_payload()` / `payload_to_df()`（Task 1）
- Produces:
  - `filing_cache.cached_filing(entry: dict) -> _CachedFiling`
  - `_CachedFiling.financials` → `_CachedFinancials | None`（`entry["has_financials"]` 為 false 時是 `None`）
  - `_CachedFinancials.income_statement() / balance_sheet() / cashflow_statement()` → `_CachedStatement | None`
  - `_CachedStatement.to_dataframe()` → `pd.DataFrame`
  - `entry` 形狀：`{"has_financials": bool, "dataframes": {"income_statement": payload|None, ...}, ...}`

- [ ] **Step 1: 寫失敗的測試**

追加到 `tests/test_filing_cache.py`：

```python
# ── 替身物件 ──────────────────────────────────────────────────────────────
#
# 快取命中時 `_filing_obj()` 回傳的東西。四個 builder 對 filing 物件的用法
# 只有一種：`.financials` → `income_statement()`/`balance_sheet()`/
# `cashflow_statement()` → `.to_dataframe()`，全部無參數。

def _entry(has_financials=True, is_df="sample", bs_df="sample", cf_df=None):
    def _p(v):
        if v == "sample":
            return filing_cache.df_to_payload(_sample_df())
        return filing_cache.df_to_payload(v)
    return {
        "schema_version": filing_cache.SCHEMA_VERSION,
        "accession_no": "0001045810-25-000123",
        "form": "10-Q",
        "filing_date": "2025-08-27",
        "cik": 1045810,
        "has_financials": has_financials,
        "dataframes": None if not has_financials else {
            "income_statement": _p(is_df),
            "balance_sheet": _p(bs_df),
            "cashflow_statement": _p(cf_df),
        },
    }


def test_cached_filing_exposes_financials_attribute_directly():
    """有兩處直接寫 `tenq.financials.xxx()` 繞過 `_financials_of()`
    （fetcher_gaap.py:1123、2584-2586），這兩條路也要能吃替身。"""
    filing = filing_cache.cached_filing(_entry())
    df = filing.financials.income_statement().to_dataframe()
    pd.testing.assert_frame_equal(df, _sample_df())


def test_cached_filing_raises_attribute_error_for_anything_else():
    """`_financials_of()` 是 `getattr(tenq, "financials", None)`。替身若對未知
    屬性用 `__getattr__` 兜底回 None，以後有人新用到 filing 的其他屬性時，
    快取命中的路徑會**安靜地把整份 filing 當成沒資料**，清快取重跑卻是好的
    ——這種 bug 極難查。所以未知屬性一律照 Python 預設拋 AttributeError。"""
    filing = filing_cache.cached_filing(_entry())
    with pytest.raises(AttributeError):
        filing.has_earnings
    with pytest.raises(AttributeError):
        filing.obj


def test_missing_statement_comes_back_as_none_not_empty():
    """`is_stmt is None` 這種判斷在 fetcher_gaap.py:1330 / 1499 都有。"""
    filing = filing_cache.cached_filing(_entry(cf_df=None))
    assert filing.financials.cashflow_statement() is None


def test_empty_statement_comes_back_as_an_empty_dataframe_not_none():
    """跟上一條成對：空表不是「沒有這張表」，下游行為不一樣。"""
    empty = _sample_df().iloc[0:0]
    filing = filing_cache.cached_filing(_entry(cf_df=empty))
    stmt = filing.financials.cashflow_statement()
    assert stmt is not None
    assert len(stmt.to_dataframe()) == 0


def test_negative_cache_entry_has_financials_none():
    """pre-XBRL 舊申報：`financials` 本來就是 None，替身要照樣回 None，
    讓 `_financials_of()` 走既有的 `continue` 分支。"""
    filing = filing_cache.cached_filing(_entry(has_financials=False))
    assert filing.financials is None
```

- [ ] **Step 2: 跑測試確認失敗**

Run: `./venv/Scripts/python.exe -m pytest tests/test_filing_cache.py -q`
Expected: FAIL — `AttributeError: module 'filing_cache' has no attribute 'cached_filing'`

- [ ] **Step 3: 寫最小實作**

追加到 `src/filing_cache.py`：

```python
# ── 替身物件（快取命中時 `_filing_obj()` 的回傳值）──────────────────────────
#
# ⚠ 這三個類別**刻意不定義 `__getattr__`**。`_financials_of()` 是
# `getattr(tenq, "financials", None)`，替身若對未知屬性兜底回 None，以後有人
# 在某個 builder 新用到 filing 物件的其他屬性，快取命中的路徑會安靜地把整份
# filing 當成沒資料、清快取重跑卻是好的——這種 bug 極難查。只實作有人真的
# 在用的那幾條路徑，其餘一律照 Python 預設失敗。

class _CachedStatement:
    """替身的一張報表。只有 `to_dataframe()`，因為呼叫端只用這一個。"""

    def __init__(self, payload: dict):
        self._payload = payload

    def to_dataframe(self) -> pd.DataFrame:
        return payload_to_df(self._payload)


class _CachedFinancials:
    """替身的 `financials`。三個 getter 全部無參數，跟真物件的用法一致。"""

    def __init__(self, dataframes: dict):
        self._dfs = dataframes or {}

    def _stmt(self, key: str) -> _CachedStatement | None:
        payload = self._dfs.get(key)
        return None if payload is None else _CachedStatement(payload)

    def income_statement(self):
        return self._stmt("income_statement")

    def balance_sheet(self):
        return self._stmt("balance_sheet")

    def cashflow_statement(self):
        return self._stmt("cashflow_statement")


class _CachedFiling:
    """替身的 filing 物件。只有 `.financials` 一個屬性。"""

    def __init__(self, entry: dict):
        self.financials = (_CachedFinancials(entry.get("dataframes") or {})
                           if entry.get("has_financials") else None)


def cached_filing(entry: dict) -> _CachedFiling:
    """快取檔內容 → 可以餵給既有 builder 的替身 filing 物件。"""
    return _CachedFiling(entry)
```

- [ ] **Step 4: 跑測試確認通過**

Run: `./venv/Scripts/python.exe -m pytest tests/test_filing_cache.py -q`
Expected: PASS（9 passed）

- [ ] **Step 5: Commit**

```bash
git add src/filing_cache.py tests/test_filing_cache.py
git commit -m "feat(cache): 快取命中時的替身 filing 物件（未知屬性照樣 AttributeError）

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01BADSq3w2c82y1Ky5yWXD8M"
```

---

## Task 3: 路徑、原子寫入、edgartools 版本

**Files:**
- Modify: `src/filing_cache.py`
- Test: `tests/test_filing_cache.py`

**Interfaces:**
- Consumes: 無
- Produces:
  - `filing_cache.cache_root() -> Path`（每次呼叫重讀 `%APPDATA%`，測試可 monkeypatch）
  - `filing_cache.ticker_dir(ticker: str) -> Path`
  - `filing_cache.filing_path(ticker: str, accession: str) -> Path | None`（accession 格式不合法回 `None`）
  - `filing_cache.edgartools_version() -> str | None`
  - `filing_cache.atomic_write_json(path: Path, obj) -> bool`（成功 True，失敗 False 不拋）

- [ ] **Step 1: 寫失敗的測試**

追加到 `tests/test_filing_cache.py`：

```python
# ── 路徑與原子寫入 ────────────────────────────────────────────────────────

@pytest.fixture
def cache_dir(tmp_path, monkeypatch):
    """把快取根目錄導到 tmp_path。`cache_root()` 每次呼叫重讀環境變數，
    所以 monkeypatch 就夠了，不用改模組層常數。"""
    monkeypatch.setenv("APPDATA", str(tmp_path))
    return tmp_path / "SEC Financial Tools" / "filing_cache"


def test_cache_root_follows_the_config_py_appdata_convention(cache_dir):
    assert filing_cache.cache_root() == cache_dir


def test_ticker_dir_is_uppercase(cache_dir):
    assert filing_cache.ticker_dir("nvda").name == "NVDA"


def test_filing_path_rejects_anything_that_is_not_an_accession_number(cache_dir):
    assert filing_cache.filing_path("NVDA", "0001045810-25-000123") is not None
    for bad in ("../../etc/passwd", "", "abc", "0001045810-25-00012"):
        assert filing_cache.filing_path("NVDA", bad) is None


def test_atomic_write_leaves_no_tmp_file_behind(cache_dir):
    path = cache_dir / "NVDA" / "x.json"
    assert filing_cache.atomic_write_json(path, {"a": 1}) is True
    assert json.loads(path.read_text(encoding="utf-8")) == {"a": 1}
    assert list(path.parent.glob("*.tmp")) == []


def test_atomic_write_uses_a_pid_suffixed_tmp_then_replaces(cache_dir, monkeypatch):
    """兩個實例（批次抓取＋跨公司比較）有機會同時寫同一個檔名。
    tmp 檔名不帶 PID 的話兩邊會蓋到對方的暫存檔。"""
    seen = {}

    real_replace = os.replace

    def _spy(src, dst):
        seen["src"] = str(src)
        return real_replace(src, dst)

    monkeypatch.setattr(filing_cache.os, "replace", _spy)
    path = cache_dir / "NVDA" / "y.json"
    filing_cache.atomic_write_json(path, {"a": 1})
    assert str(os.getpid()) in seen["src"]
    assert seen["src"].endswith(".tmp")


def test_atomic_write_returns_false_instead_of_raising_when_disk_write_fails(
        cache_dir, monkeypatch):
    """磁碟滿／權限問題時只記 log 繼續跑，不能拖垮整趟抓取。"""
    def _boom(*a, **kw):
        raise OSError("disk full")

    monkeypatch.setattr(filing_cache, "open", _boom, raising=False)
    assert filing_cache.atomic_write_json(cache_dir / "NVDA" / "z.json", {"a": 1}) is False


def test_edgartools_version_is_read_from_package_metadata():
    """實測 `edgar.__version__` 不存在（AttributeError），只能走
    importlib.metadata。取不到就回 None，不可以填 "unknown" 之類的預設值
    ——那會讓版本比對永遠成功或永遠失敗，兩種都是錯的。"""
    v = filing_cache.edgartools_version()
    assert v is None or (isinstance(v, str) and v[0].isdigit())
```

- [ ] **Step 2: 跑測試確認失敗**

Run: `./venv/Scripts/python.exe -m pytest tests/test_filing_cache.py -q`
Expected: FAIL — `AttributeError: module 'filing_cache' has no attribute 'cache_root'`

- [ ] **Step 3: 寫最小實作**

追加到 `src/filing_cache.py`（放在 `SCHEMA_VERSION` 那組常數之後、序列化之前）：

```python
# ── 路徑 ──────────────────────────────────────────────────────────────────
#
# 沿用 `config.py` 的 `%APPDATA%\SEC Financial Tools\`（**有空格**那個；
# `override_engine.py` 用的是底線版 `SEC_Financial_Tools`，兩者歷史上就分岔了，
# 這裡跟 config.py 對齊）。每次呼叫重讀環境變數，測試才好導到 tmp。

def cache_root() -> Path:
    appdata = os.environ.get("APPDATA")
    if appdata:
        return Path(appdata) / "SEC Financial Tools" / "filing_cache"
    return Path.home() / ".sec_financial_tools" / "filing_cache"


def ticker_dir(ticker: str) -> Path:
    """一個 ticker 一個資料夾，名稱就是大寫 ticker——不用查表，在檔案總管
    肉眼就看得出哪些公司有快取、大概多大。"""
    return cache_root() / (ticker or "").strip().upper()


def filing_path(ticker: str, accession: str) -> Path | None:
    """`<accession>.json` 的完整路徑。accession 格式不合法回 None（不快取）
    ——這同時擋掉把奇怪字串當檔名寫出去的可能。"""
    if not ACCESSION_RE.match(str(accession or "")):
        return None
    return ticker_dir(ticker) / f"{accession}.json"


# ── 原子寫入 ──────────────────────────────────────────────────────────────

def atomic_write_json(path: Path, obj) -> bool:
    """tmp + `os.replace()`。tmp 檔名帶 PID，避免兩個實例互相蓋到暫存檔。

    寫不進去（磁碟滿、權限）只回 False，不拋——快取只是加速層，
    寫入失敗不該影響這次抓取的結果。
    """
    tmp = Path(str(path) + f".{os.getpid()}.tmp")
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(obj, f, ensure_ascii=False)
        os.replace(tmp, path)
        return True
    except OSError:
        try:
            tmp.unlink()
        except OSError:
            pass
        return False


# ── edgartools 版本 ───────────────────────────────────────────────────────

def edgartools_version() -> str | None:
    """實測 `edgar.__version__` **不存在**（AttributeError），只能走
    package metadata（實測回 "5.29.0"）。取不到回 None，呼叫端把 None 當成
    「這次不要用快取」——不可以填一個預設值混進檔案裡。"""
    try:
        from importlib.metadata import version
        return version("edgartools")
    except Exception:
        return None
```

- [ ] **Step 4: 跑測試確認通過**

Run: `./venv/Scripts/python.exe -m pytest tests/test_filing_cache.py -q`
Expected: PASS（16 passed）

- [ ] **Step 5: Commit**

```bash
git add src/filing_cache.py tests/test_filing_cache.py
git commit -m "feat(cache): 快取路徑、PID 帶號的原子寫入、edgartools 版本取得

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01BADSq3w2c82y1Ky5yWXD8M"
```

---

## Task 4: 單份 filing 的讀寫與四道閘

`load_filing()` 要通過四道閘（JSON 可解析、`schema_version`、`cik`、`edgartools_version`）
才回內容，任一不符回 `None`（視同無快取），**不得拋例外**。
`save_filing()` 逐份即時落檔，支援負向快取（`has_financials: false`）。

**Files:**
- Modify: `src/filing_cache.py`
- Test: `tests/test_filing_cache.py`

**Interfaces:**
- Consumes: `filing_path()`、`atomic_write_json()`、`edgartools_version()`、`df_to_payload()`
- Produces:
  - `filing_cache.load_filing(ticker: str, accession: str, cik: int) -> dict | None`
  - `filing_cache.save_filing(ticker: str, accession: str, *, form: str, filing_date: str, cik: int, dataframes: dict[str, pd.DataFrame | None] | None, has_financials: bool) -> bool`

- [ ] **Step 1: 寫失敗的測試**

追加到 `tests/test_filing_cache.py`：

```python
# ── 單份 filing 的讀寫與四道閘 ────────────────────────────────────────────

ACC = "0001045810-25-000123"


def _save_sample(ticker="NVDA", cik=1045810, cf=None):
    return filing_cache.save_filing(
        ticker, ACC, form="10-Q", filing_date="2025-08-27", cik=cik,
        dataframes={"income_statement": _sample_df(),
                    "balance_sheet": _sample_df(),
                    "cashflow_statement": cf},
        has_financials=True,
    )


def test_save_then_load_returns_the_same_three_dataframes(cache_dir):
    assert _save_sample() is True
    entry = filing_cache.load_filing("NVDA", ACC, 1045810)
    assert entry is not None
    filing = filing_cache.cached_filing(entry)
    pd.testing.assert_frame_equal(
        filing.financials.income_statement().to_dataframe(), _sample_df())
    assert filing.financials.cashflow_statement() is None


def test_load_returns_none_when_the_file_does_not_exist(cache_dir):
    assert filing_cache.load_filing("NVDA", ACC, 1045810) is None


def test_load_returns_none_when_the_cik_differs(cache_dir):
    """ticker 會換手（公司更名、代號被回收給別家）。這種錯不會報例外，
    只會安靜地把別家公司的數字餵給使用者，比任何其他失效情境都危險。"""
    _save_sample(cik=1045810)
    assert filing_cache.load_filing("NVDA", ACC, 99999) is None


def test_load_returns_none_when_the_edgartools_version_differs(cache_dir, monkeypatch):
    """套件升版可能讓同一份 filing 解出不一樣的結果（新的 standardization
    mapping、XBRL parser 修 bug）。這條軸線跟我們自己的比對規則是兩件事。"""
    _save_sample()
    monkeypatch.setattr(filing_cache, "edgartools_version", lambda: "99.0.0")
    assert filing_cache.load_filing("NVDA", ACC, 1045810) is None


def test_load_returns_none_when_the_schema_version_differs(cache_dir):
    _save_sample()
    path = filing_cache.filing_path("NVDA", ACC)
    raw = json.loads(path.read_text(encoding="utf-8"))
    raw["schema_version"] = filing_cache.SCHEMA_VERSION + 1
    path.write_text(json.dumps(raw), encoding="utf-8")
    assert filing_cache.load_filing("NVDA", ACC, 1045810) is None


def test_load_returns_none_for_corrupt_json_instead_of_raising(cache_dir):
    """損毀的快取不可以拖垮整趟抓取——當作沒這份，重抓後覆蓋掉。"""
    _save_sample()
    filing_cache.filing_path("NVDA", ACC).write_text("{ not json", encoding="utf-8")
    assert filing_cache.load_filing("NVDA", ACC, 1045810) is None


def test_cache_is_disabled_entirely_when_the_version_is_unavailable(
        cache_dir, monkeypatch):
    """取不到 edgartools 版本時不讀也不寫，不留下無法比對版本的檔案。"""
    monkeypatch.setattr(filing_cache, "edgartools_version", lambda: None)
    assert _save_sample() is False
    assert filing_cache.load_filing("NVDA", ACC, 1045810) is None


def test_negative_cache_records_a_filing_with_no_financials(cache_dir):
    """pre-XBRL 舊申報（2009 年前常見）：記著「這份試過了、沒有財務資料」，
    下次不用再打一次 SEC 重試。"""
    assert filing_cache.save_filing(
        "NVDA", ACC, form="10-Q", filing_date="2008-05-01", cik=1045810,
        dataframes=None, has_financials=False) is True
    entry = filing_cache.load_filing("NVDA", ACC, 1045810)
    assert entry is not None
    assert entry["has_financials"] is False
    assert filing_cache.cached_filing(entry).financials is None


def test_save_rejects_a_malformed_accession(cache_dir):
    assert filing_cache.save_filing(
        "NVDA", "not-an-accession", form="10-Q", filing_date="2025-08-27",
        cik=1045810, dataframes=None, has_financials=False) is False
```

- [ ] **Step 2: 跑測試確認失敗**

Run: `./venv/Scripts/python.exe -m pytest tests/test_filing_cache.py -q`
Expected: FAIL — `AttributeError: module 'filing_cache' has no attribute 'save_filing'`

- [ ] **Step 3: 寫最小實作**

追加到 `src/filing_cache.py`：

```python
# ── 單份 filing 的讀寫 ────────────────────────────────────────────────────
#
# `<accession>.json` **是否存在，才是「這份 filing 有沒有快取」的事實來源**。
# manifest 只是給 GUI 看的衍生索引，查快取不問它。

def load_filing(ticker: str, accession: str, cik: int) -> dict | None:
    """讀一份快取。四道閘任一沒過就回 None（視同無快取，照舊打 SEC 重抓）：
    JSON 可解析、`schema_version`、`cik`、`edgartools_version`。

    正確性優先於速度——寧可那次變慢，也不要餵錯公司的資料或吃到舊版
    parser 的 bug。任何情況都不拋例外。
    """
    version = edgartools_version()
    if version is None:
        return None
    path = filing_path(ticker, accession)
    if path is None or not path.exists():
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            entry = json.load(f)
    except (OSError, ValueError):
        return None
    if not isinstance(entry, dict):
        return None
    if entry.get("schema_version") != SCHEMA_VERSION:
        return None
    if entry.get("edgartools_version") != version:
        return None
    if cik is not None and entry.get("cik") != cik:
        return None
    return entry


def save_filing(ticker: str, accession: str, *, form: str, filing_date: str,
                cik: int, dataframes: dict | None, has_financials: bool) -> bool:
    """寫一份快取。**逐份即時落檔**——一趟抓取可能好幾分鐘，中途斷線或關視窗
    時，已經抓到的進度不該全部白費。

    `has_financials=False` 是負向快取（pre-XBRL 舊申報）。**網路失敗絕對不能
    走到這裡**——那是暫時性的，交給既有的 D11-B 缺漏帳本，每次都該重試。
    """
    version = edgartools_version()
    if version is None:
        return False
    path = filing_path(ticker, accession)
    if path is None:
        return False
    payloads = None
    if has_financials:
        payloads = {k: df_to_payload((dataframes or {}).get(k)) for k in STATEMENT_KEYS}
    entry = {
        "schema_version": SCHEMA_VERSION,
        "accession_no": accession,
        "form": form,
        "filing_date": filing_date,
        "cached_at": _now_iso(),
        "cik": cik,
        "edgartools_version": version,
        "has_financials": bool(has_financials),
        "dataframes": payloads,
    }
    return atomic_write_json(path, entry)
```

同時在 import 區加上 `from datetime import datetime` 並補這個小工具（放在常數之後）：

```python
def _now_iso() -> str:
    """本地時間帶時區偏移，例如 "2026-09-03T14:22:10+08:00"。純顯示用。"""
    return datetime.now().astimezone().isoformat(timespec="seconds")
```

- [ ] **Step 4: 跑測試確認通過**

Run: `./venv/Scripts/python.exe -m pytest tests/test_filing_cache.py -q`
Expected: PASS（25 passed）

- [ ] **Step 5: Commit**

```bash
git add src/filing_cache.py tests/test_filing_cache.py
git commit -m "feat(cache): 單份 filing 讀寫與四道閘（schema/cik/版本/損毀）＋負向快取

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01BADSq3w2c82y1Ky5yWXD8M"
```

---

## Task 5: `_manifest.json` 從磁碟重建

manifest 是衍生索引不是事實來源。遺失、損毀、或跟磁碟對不上都走**同一條路徑**：
掃資料夾裡有哪些 `<accession>.json`，逐份讀出 metadata 組回 `filings` 清單。

**Files:**
- Modify: `src/filing_cache.py`
- Test: `tests/test_filing_cache.py`

**Interfaces:**
- Consumes: `ticker_dir()`、`atomic_write_json()`
- Produces:
  - `filing_cache.rebuild_manifest(ticker: str, cik: int | None = None) -> dict`（同時寫回磁碟；回傳寫出去的 manifest）
  - `filing_cache.read_manifest(ticker: str) -> dict | None`

- [ ] **Step 1: 寫失敗的測試**

追加到 `tests/test_filing_cache.py`：

```python
# ── manifest（衍生索引，壞了直接重建）────────────────────────────────────

ACC2 = "0001045810-24-000456"


def test_manifest_is_rebuilt_from_whatever_is_actually_on_disk(cache_dir):
    _save_sample()
    filing_cache.save_filing("NVDA", ACC2, form="10-K", filing_date="2025-02-26",
                             cik=1045810, dataframes=None, has_financials=False)
    manifest = filing_cache.rebuild_manifest("NVDA", cik=1045810)
    accs = {f["accession_no"] for f in manifest["filings"]}
    assert accs == {ACC, ACC2}
    assert manifest["cik"] == 1045810
    assert manifest["schema_version"] == filing_cache.SCHEMA_VERSION
    assert manifest["last_checked_at"]


def test_manifest_rows_carry_the_fields_the_gui_needs(cache_dir):
    _save_sample()
    row = filing_cache.rebuild_manifest("NVDA", cik=1045810)["filings"][0]
    assert set(row) >= {"accession_no", "form", "filing_date", "cached_at",
                        "edgartools_version", "has_financials", "size_bytes"}
    assert row["size_bytes"] > 0


def test_a_corrupt_manifest_is_simply_replaced(cache_dir):
    """manifest 壞掉不需要特別的「修正」邏輯，跟「本來就不存在」同一條路徑。"""
    _save_sample()
    (filing_cache.ticker_dir("NVDA") / "_manifest.json").write_text(
        "{ garbage", encoding="utf-8")
    assert filing_cache.read_manifest("NVDA") is None
    manifest = filing_cache.rebuild_manifest("NVDA", cik=1045810)
    assert len(manifest["filings"]) == 1
    assert filing_cache.read_manifest("NVDA") is not None


def test_a_manifest_out_of_sync_with_disk_is_corrected_by_rebuild(cache_dir):
    """manifest 說有兩份、磁碟只有一份 → 以磁碟為準。"""
    _save_sample()
    filing_cache.rebuild_manifest("NVDA", cik=1045810)
    filing_cache.filing_path("NVDA", ACC).unlink()
    assert filing_cache.rebuild_manifest("NVDA", cik=1045810)["filings"] == []


def test_rebuilding_a_ticker_with_no_cache_directory_is_harmless(cache_dir):
    assert filing_cache.rebuild_manifest("ZZZZ", cik=1)["filings"] == []
```

- [ ] **Step 2: 跑測試確認失敗**

Run: `./venv/Scripts/python.exe -m pytest tests/test_filing_cache.py -q`
Expected: FAIL — `AttributeError: module 'filing_cache' has no attribute 'rebuild_manifest'`

- [ ] **Step 3: 寫最小實作**

追加到 `src/filing_cache.py`：

```python
# ── manifest ──────────────────────────────────────────────────────────────
#
# ⚠ manifest **不是事實來源**，是衍生索引，只給 GUI 顯示用。查快取一律直接問
# `<accession>.json` 存不存在（掃 100 個檔名是微秒級成本，比維護「manifest
# 跟磁碟是否同步」的心智負擔低得多）。所以這裡沒有「修正」邏輯，只有重建。

MANIFEST_NAME = "_manifest.json"


def read_manifest(ticker: str) -> dict | None:
    """讀 manifest。不存在、壞掉、版本不符一律回 None（呼叫端重建）。"""
    path = ticker_dir(ticker) / MANIFEST_NAME
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, ValueError):
        return None
    if not isinstance(data, dict) or data.get("schema_version") != SCHEMA_VERSION:
        return None
    return data


def rebuild_manifest(ticker: str, cik: int | None = None) -> dict:
    """從資料夾實際內容重建並寫回。manifest 遺失／損毀／跟磁碟對不上都走這條。"""
    directory = ticker_dir(ticker)
    rows: list[dict] = []
    try:
        paths = sorted(directory.glob("*.json"))
    except OSError:
        paths = []
    for path in paths:
        if path.name == MANIFEST_NAME or not ACCESSION_RE.match(path.stem):
            continue
        try:
            size = path.stat().st_size
            with open(path, "r", encoding="utf-8") as f:
                entry = json.load(f)
        except (OSError, ValueError):
            continue    # 壞掉的那一份不進索引，下次抓取會重抓並覆蓋
        if not isinstance(entry, dict):
            continue
        rows.append({
            "accession_no": entry.get("accession_no", path.stem),
            "form": entry.get("form", ""),
            "filing_date": entry.get("filing_date", ""),
            "cached_at": entry.get("cached_at", ""),
            "edgartools_version": entry.get("edgartools_version", ""),
            "has_financials": bool(entry.get("has_financials")),
            "size_bytes": size,
        })
    manifest = {
        "schema_version": SCHEMA_VERSION,
        "ticker": (ticker or "").strip().upper(),
        "cik": cik,
        # 上次去 SEC 查 filing 清單的時間。純顯示用，**不是有效期限**
        # ——每次抓取都會重查，不靠這個判斷要不要查。
        "last_checked_at": _now_iso(),
        "filings": rows,
    }
    if rows or directory.exists():
        atomic_write_json(directory / MANIFEST_NAME, manifest)
    return manifest
```

- [ ] **Step 4: 跑測試確認通過**

Run: `./venv/Scripts/python.exe -m pytest tests/test_filing_cache.py -q`
Expected: PASS（30 passed）

- [ ] **Step 5: Commit**

```bash
git add src/filing_cache.py tests/test_filing_cache.py
git commit -m "feat(cache): _manifest.json 從磁碟重建（衍生索引，不是事實來源）

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01BADSq3w2c82y1Ky5yWXD8M"
```

---

## Task 6: GUI 要用的統計與清除 API

**Files:**
- Modify: `src/filing_cache.py`
- Test: `tests/test_filing_cache.py`

**Interfaces:**
- Consumes: `cache_root()`、`ticker_dir()`、`read_manifest()`、`rebuild_manifest()`
- Produces:
  - `filing_cache.list_cached_tickers() -> list[dict]`，每筆 `{"ticker": str, "count": int, "size_bytes": int}`，依 `size_bytes` 由大到小
  - `filing_cache.total_size_bytes() -> int`
  - `filing_cache.clear_ticker(ticker: str) -> bool`
  - `filing_cache.clear_all() -> int`（回傳刪掉幾家）

- [ ] **Step 1: 寫失敗的測試**

追加到 `tests/test_filing_cache.py`：

```python
# ── GUI 用的統計與清除 ────────────────────────────────────────────────────

def test_listing_scans_the_folder_no_global_index_needed(cache_dir):
    """快取了哪些公司直接掃資料夾就知道——不維護容易跟磁碟脫鉤的全域清單。"""
    _save_sample(ticker="NVDA")
    _save_sample(ticker="AMD")
    rows = filing_cache.list_cached_tickers()
    assert {r["ticker"] for r in rows} == {"NVDA", "AMD"}
    assert all(r["count"] == 1 and r["size_bytes"] > 0 for r in rows)


def test_listing_is_sorted_by_size_descending(cache_dir):
    _save_sample(ticker="BIG")
    filing_cache.save_filing("SML", ACC, form="10-Q", filing_date="2025-08-27",
                             cik=1, dataframes=None, has_financials=False)
    rows = filing_cache.list_cached_tickers()
    assert [r["ticker"] for r in rows] == ["BIG", "SML"]


def test_listing_is_empty_when_nothing_is_cached(cache_dir):
    assert filing_cache.list_cached_tickers() == []
    assert filing_cache.total_size_bytes() == 0


def test_listing_survives_a_file_vanishing_mid_scan(cache_dir, monkeypatch):
    """正在被清除、或另一個實例正在寫入時檔案可能瞬間消失。"""
    _save_sample()

    def _boom(self):
        raise OSError("gone")

    monkeypatch.setattr(Path, "stat", _boom)
    rows = filing_cache.list_cached_tickers()
    assert rows == [] or rows[0]["size_bytes"] == 0


def test_total_size_is_the_sum_of_every_ticker(cache_dir):
    _save_sample(ticker="NVDA")
    _save_sample(ticker="AMD")
    rows = filing_cache.list_cached_tickers()
    assert filing_cache.total_size_bytes() == sum(r["size_bytes"] for r in rows)


def test_clear_ticker_removes_the_whole_folder(cache_dir):
    _save_sample(ticker="NVDA")
    _save_sample(ticker="AMD")
    assert filing_cache.clear_ticker("NVDA") is True
    assert not filing_cache.ticker_dir("NVDA").exists()
    assert [r["ticker"] for r in filing_cache.list_cached_tickers()] == ["AMD"]


def test_clear_ticker_on_something_that_is_not_cached_is_harmless(cache_dir):
    assert filing_cache.clear_ticker("ZZZZ") is False


def test_clear_all_removes_every_ticker(cache_dir):
    _save_sample(ticker="NVDA")
    _save_sample(ticker="AMD")
    assert filing_cache.clear_all() == 2
    assert filing_cache.list_cached_tickers() == []
```

- [ ] **Step 2: 跑測試確認失敗**

Run: `./venv/Scripts/python.exe -m pytest tests/test_filing_cache.py -q`
Expected: FAIL — `AttributeError: module 'filing_cache' has no attribute 'list_cached_tickers'`

- [ ] **Step 3: 寫最小實作**

在 `src/filing_cache.py` 的 import 區加 `import shutil`，並追加：

```python
# ── GUI：統計與清除 ───────────────────────────────────────────────────────

def _dir_stats(directory: Path) -> tuple[int, int]:
    """(filing 份數, 位元組數)。`except OSError` 是必要的——清除中或另一個
    實例正在寫入時，檔案可能在 glob 與 stat 之間消失。"""
    count = 0
    size = 0
    try:
        paths = list(directory.glob("*.json"))
    except OSError:
        return 0, 0
    for path in paths:
        try:
            size += path.stat().st_size
        except OSError:
            continue
        if path.name != MANIFEST_NAME and ACCESSION_RE.match(path.stem):
            count += 1
    return count, size


def list_cached_tickers() -> list[dict]:
    """快取了哪些公司：直接掃 `filing_cache/` 底下有哪些子資料夾。
    公司數量最多幾十家，掃資料夾的成本可以忽略。"""
    root = cache_root()
    rows: list[dict] = []
    try:
        entries = sorted(p for p in root.iterdir() if p.is_dir())
    except OSError:
        return []
    for directory in entries:
        count, size = _dir_stats(directory)
        if count == 0 and size == 0:
            continue
        rows.append({"ticker": directory.name, "count": count, "size_bytes": size})
    rows.sort(key=lambda r: (-r["size_bytes"], r["ticker"]))
    return rows


def total_size_bytes() -> int:
    return sum(r["size_bytes"] for r in list_cached_tickers())


def clear_ticker(ticker: str) -> bool:
    """整個刪掉那家公司的資料夾。下次抓這家會當作全新開始。"""
    directory = ticker_dir(ticker)
    if not directory.exists():
        return False
    try:
        shutil.rmtree(directory)
        return True
    except OSError:
        return False


def clear_all() -> int:
    """刪掉所有公司的快取，回傳刪掉幾家。"""
    return sum(1 for row in list_cached_tickers() if clear_ticker(row["ticker"]))
```

- [ ] **Step 4: 跑測試確認通過**

Run: `./venv/Scripts/python.exe -m pytest tests/test_filing_cache.py -q`
Expected: PASS（38 passed）

- [ ] **Step 5: Commit**

```bash
git add src/filing_cache.py tests/test_filing_cache.py
git commit -m "feat(cache): 快取容量統計與清除 API（掃資料夾，不維護全域索引）

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01BADSq3w2c82y1Ky5yWXD8M"
```

---

## Task 7: 掛進 `fetcher_gaap._filing_obj()`

唯一的掛勾點。四個 builder（IS/BS/CF/segment）都已經共用它，改這一個函式全部受益。

**Files:**
- Modify: `src/fetcher_gaap.py`（import 區、`_parse_cache_scope()` 附近、`_filing_obj()`、`fetch_gaap_statements()`、`_fetch_gaap_impl()`）
- Test: `tests/test_fetcher_gaap_cache.py`（新增）

**Interfaces:**
- Consumes: `filing_cache.load_filing()` / `save_filing()` / `cached_filing()` / `rebuild_manifest()`（Task 2、4、5）
- Produces:
  - `fetcher_gaap._disk_cache_scope()`：context manager，yield 一個 `{"ticker": None, "cik": None, "hits": 0, "misses": 0}` dict
  - `fetcher_gaap._bind_disk_cache(ticker: str, cik: int | None) -> None`
  - `fetcher_gaap.last_cache_stats() -> tuple[int, int]`（(命中數, 這趟處理的份數)；沒開過快取回 `(0, 0)`）

- [ ] **Step 1: 寫失敗的測試**

建立 `tests/test_fetcher_gaap_cache.py`：

```python
"""本地 filing 快取掛在 `_filing_obj()` 上的整合行為。

`filing_cache.py` 自己的儲存層測試在 tests/test_filing_cache.py。這裡釘的是
「什麼時候會打網路、什麼時候不會」，以及那幾條踩到會餵錯資料的邊界。
"""
from unittest.mock import MagicMock

import pandas as pd
import pytest

import filing_cache
from fetcher_gaap import (
    _filing_obj,
    _bind_disk_cache,
    _disk_cache_scope,
    _parse_cache_scope,
    last_cache_stats,
)
from net_retry import NetworkDownError

ACC = "0001045810-25-000123"
ACC_OLD = "0001045810-19-000001"


@pytest.fixture
def cache_dir(tmp_path, monkeypatch):
    monkeypatch.setenv("APPDATA", str(tmp_path))
    return tmp_path / "SEC Financial Tools" / "filing_cache"


def _df():
    return pd.DataFrame({
        "concept": ["us-gaap_Revenue"],
        "label": ["Net sales"],
        "2025-12-27 (Q1)": [1000.0],
    })


def _fake_filing(acc=ACC, filing_date="2025-08-27", form="10-Q", financials=True):
    """一份假的 edgartools filing。`obj()` 被呼叫幾次是這組測試的主要斷言。"""
    stmt = MagicMock()
    stmt.to_dataframe.return_value = _df()
    fin = MagicMock()
    fin.income_statement.return_value = stmt
    fin.balance_sheet.return_value = stmt
    fin.cashflow_statement.return_value = stmt

    obj = MagicMock()
    obj.financials = fin if financials else None

    filing = MagicMock()
    filing.accession_no = acc
    filing.filing_date = filing_date
    filing.form = form
    filing.obj.return_value = obj
    return filing


# ── 命中／未命中 ──────────────────────────────────────────────────────────

def test_first_fetch_hits_the_network_and_writes_the_cache_file(cache_dir):
    filing = _fake_filing()
    with _disk_cache_scope(), _parse_cache_scope():
        _bind_disk_cache("NVDA", 1045810)
        _filing_obj(filing)
    assert filing.obj.call_count == 1
    assert filing_cache.filing_path("NVDA", ACC).exists()
    assert any(f["accession_no"] == ACC
               for f in filing_cache.read_manifest("NVDA")["filings"])


def test_second_fetch_reads_the_cache_and_never_touches_the_network(cache_dir):
    first = _fake_filing()
    with _disk_cache_scope(), _parse_cache_scope():
        _bind_disk_cache("NVDA", 1045810)
        _filing_obj(first)

    second = _fake_filing()
    with _disk_cache_scope(), _parse_cache_scope():
        _bind_disk_cache("NVDA", 1045810)
        tenq = _filing_obj(second)
        df = tenq.financials.income_statement().to_dataframe()
    assert second.obj.call_count == 0
    pd.testing.assert_frame_equal(df, _df())
    assert last_cache_stats() == (1, 1)


def test_a_cache_hit_still_goes_through_the_in_memory_parse_cache(cache_dir):
    """G9 的記憶體快取存的就是 `_filing_obj()` 的回傳值——存真物件或存替身
    對它沒有差別，兩層快取不衝突。"""
    warm = _fake_filing()
    with _disk_cache_scope(), _parse_cache_scope():
        _bind_disk_cache("NVDA", 1045810)
        _filing_obj(warm)

    filing = _fake_filing()
    with _disk_cache_scope(), _parse_cache_scope():
        _bind_disk_cache("NVDA", 1045810)
        assert _filing_obj(filing) is _filing_obj(filing)


def test_cache_is_off_when_nothing_is_bound(cache_dir):
    """沒綁定 ticker/cik（例如拿不到 cik）時，行為跟改動前完全一樣。"""
    filing = _fake_filing()
    with _parse_cache_scope():
        _filing_obj(filing)
    assert filing.obj.call_count == 1
    assert not filing_cache.cache_root().exists()


# ── 負向快取 vs 網路失敗 ──────────────────────────────────────────────────

def test_a_filing_with_no_financials_is_cached_negatively(cache_dir):
    """pre-XBRL 舊申報：記著「試過了、沒有財務資料」，下次不用再打 SEC。"""
    first = _fake_filing(acc=ACC_OLD, filing_date="2008-05-01", financials=False)
    with _disk_cache_scope(), _parse_cache_scope():
        _bind_disk_cache("NVDA", 1045810)
        _filing_obj(first)

    second = _fake_filing(acc=ACC_OLD, filing_date="2008-05-01", financials=False)
    with _disk_cache_scope(), _parse_cache_scope():
        _bind_disk_cache("NVDA", 1045810)
        tenq = _filing_obj(second)
    assert second.obj.call_count == 0
    assert tenq.financials is None


def test_a_network_failure_is_never_recorded_as_a_negative_cache(cache_dir):
    """網路失敗是暫時性的，交給既有的 D11-B 缺漏帳本，每次都該重試。
    誤記成「沒有 financials」會讓那一期**永久**消失。"""
    boom = _fake_filing()
    boom.obj.side_effect = NetworkDownError("network down")
    with _disk_cache_scope(), _parse_cache_scope():
        _bind_disk_cache("NVDA", 1045810)
        with pytest.raises(NetworkDownError):
            _filing_obj(boom)
    assert not filing_cache.filing_path("NVDA", ACC).exists()

    retry = _fake_filing()
    with _disk_cache_scope(), _parse_cache_scope():
        _bind_disk_cache("NVDA", 1045810)
        _filing_obj(retry)
    assert retry.obj.call_count == 1


def test_a_parse_failure_after_download_is_not_cached_either(cache_dir):
    """`financials` 拿得到但 `to_dataframe()` 炸掉——一樣不留快取，下次重試。"""
    filing = _fake_filing()
    filing.obj.return_value.financials.income_statement.side_effect = RuntimeError("bad xbrl")
    with _disk_cache_scope(), _parse_cache_scope():
        _bind_disk_cache("NVDA", 1045810)
        _filing_obj(filing)
    assert not filing_cache.filing_path("NVDA", ACC).exists()


# ── 逐份即時落檔 ──────────────────────────────────────────────────────────

def test_filings_parsed_before_a_crash_stay_on_disk(cache_dir):
    """一趟抓取可能好幾分鐘。中途斷線或使用者關視窗時，已經抓到的進度
    不可以全部白費——所以是逐份落檔，不是整趟跑完才一次寫入。"""
    ok = _fake_filing(acc=ACC)
    boom = _fake_filing(acc=ACC_OLD)
    boom.obj.side_effect = NetworkDownError("network down")

    with pytest.raises(NetworkDownError):
        with _disk_cache_scope(), _parse_cache_scope():
            _bind_disk_cache("NVDA", 1045810)
            _filing_obj(ok)
            _filing_obj(boom)

    assert filing_cache.filing_path("NVDA", ACC).exists()
    assert not filing_cache.filing_path("NVDA", ACC_OLD).exists()


# ── 查詢邊界不能漏抓 ──────────────────────────────────────────────────────

def test_a_wider_query_still_fetches_filings_that_were_never_cached(cache_dir):
    """先用窄範圍抓（只碰到 2020 年以後的 filing），再用完整期間抓同一家，
    第二次一定要真的去補抓從沒進過快取的舊 filing——不能因為「這家公司
    已經有快取」就少抓。快取的粒度是**一份 filing**，不是一家公司。"""
    recent = _fake_filing(acc=ACC, filing_date="2025-08-27")
    with _disk_cache_scope(), _parse_cache_scope():
        _bind_disk_cache("NVDA", 1045810)
        _filing_obj(recent)

    recent2 = _fake_filing(acc=ACC, filing_date="2025-08-27")
    old = _fake_filing(acc=ACC_OLD, filing_date="2019-05-01")
    with _disk_cache_scope(), _parse_cache_scope():
        _bind_disk_cache("NVDA", 1045810)
        _filing_obj(recent2)
        _filing_obj(old)
    assert recent2.obj.call_count == 0      # 已經有的照樣讀快取
    assert old.obj.call_count == 1          # 沒有的一定要補抓
    assert filing_cache.filing_path("NVDA", ACC_OLD).exists()
    assert last_cache_stats() == (1, 2)


def test_the_cache_is_keyed_per_company_not_shared(cache_dir):
    """ticker 會換手。`cik` 不符時整包視同無快取——這種錯不會報例外，
    只會安靜地把別家公司的數字餵給使用者。"""
    filing = _fake_filing()
    with _disk_cache_scope(), _parse_cache_scope():
        _bind_disk_cache("NVDA", 1045810)
        _filing_obj(filing)

    other = _fake_filing()
    with _disk_cache_scope(), _parse_cache_scope():
        _bind_disk_cache("NVDA", 99999)     # 同 ticker、不同公司
        _filing_obj(other)
    assert other.obj.call_count == 1
```

- [ ] **Step 2: 跑測試確認失敗**

Run: `./venv/Scripts/python.exe -m pytest tests/test_fetcher_gaap_cache.py -q`
Expected: FAIL — `ImportError: cannot import name '_disk_cache_scope' from 'fetcher_gaap'`

- [ ] **Step 3: 寫最小實作**

3a. `src/fetcher_gaap.py` import 區加一行（在 `from fetch_ledger import FetchLedger` 附近）：

```python
import filing_cache
```

3b. 在 `_cache_key()` 之後、`_list_filings()` 之前插入磁碟快取的範圍管理：

```python
# ── 本地磁碟快取（跨執行有效）────────────────────────────────────────────
#
# 跟上面的 `_parse_cache`（G9，只活在一次執行的記憶體裡）是**兩層不同的快取**，
# 不衝突：磁碟快取讀回來的結果一樣會進記憶體快取，避免同一次執行內重複讀檔。
#
# 綁定分兩步是為了避免把 `_fetch_gaap_impl()` 整段重新縮排：範圍在
# `fetch_gaap_statements()` 開，`ticker`/`cik` 等 `Company(ticker)` 建好之後
# 才由 `_bind_disk_cache()` 填進去。沒綁定（拿不到 cik）就整個不用快取，
# 行為跟改動前一模一樣。
_disk_cache: dict | None = None
_last_cache_stats: tuple[int, int] = (0, 0)


@contextmanager
def _disk_cache_scope():
    """一次抓取的磁碟快取範圍。離開時更新 manifest 並記下命中統計。"""
    global _disk_cache, _last_cache_stats
    outer = _disk_cache
    ctx = {"ticker": None, "cik": None, "hits": 0, "misses": 0} if outer is None else outer
    _disk_cache = ctx
    try:
        yield ctx
    finally:
        if outer is None:
            if ctx["ticker"]:
                # 收尾重建索引：反映這趟跑完後磁碟上實際的內容。
                # manifest 只是給 GUI 看的，重建失敗不影響任何抓取結果。
                try:
                    filing_cache.rebuild_manifest(ctx["ticker"], ctx["cik"])
                except OSError:
                    pass
            _last_cache_stats = (ctx["hits"], ctx["hits"] + ctx["misses"])
            _disk_cache = None


def _bind_disk_cache(ticker: str, cik) -> None:
    """把這趟的公司身分填進磁碟快取範圍。cik 拿不到就不啟用快取——
    ticker 只是別名，會換手；cik 才是跟 SEC 打交道真正的鍵。"""
    if _disk_cache is None or not ticker or cik is None:
        return
    try:
        _disk_cache["cik"] = int(cik)
    except (TypeError, ValueError):
        return
    _disk_cache["ticker"] = str(ticker).strip().upper()


def last_cache_stats() -> tuple[int, int]:
    """上一趟抓取的 (命中份數, 處理份數)。給 log 用（見 main.py）。"""
    return _last_cache_stats


def _save_to_disk_cache(ctx: dict, filing, obj) -> None:
    """把剛解析出來的三張 DataFrame 逐份即時落檔。

    ⚠ 只有「`filing.obj()` 成功回來」才會走到這裡——網路失敗會在上面直接
    往外拋，不留任何快取（正向或負向都不留），下次照樣重試。
    解析成功但 `financials` 是 None（pre-XBRL）才寫負向快取。
    """
    acc = _cache_key(filing)
    fd = getattr(filing, "filing_date", None)
    meta = {
        "form": str(getattr(filing, "form", "") or ""),
        "filing_date": str(fd) if fd else "",
        "cik": ctx["cik"],
    }
    fin = _financials_of(obj)
    if fin is None:
        filing_cache.save_filing(ctx["ticker"], acc, dataframes=None,
                                 has_financials=False, **meta)
        return
    try:
        dfs = {}
        for key, getter in (("income_statement", fin.income_statement),
                            ("balance_sheet", fin.balance_sheet),
                            ("cashflow_statement", fin.cashflow_statement)):
            stmt = getter()
            dfs[key] = None if stmt is None else stmt.to_dataframe()
    except Exception:
        return   # 解析失敗不留快取，下次重試（跟網路失敗同一個處理）
    filing_cache.save_filing(ctx["ticker"], acc, dataframes=dfs,
                             has_financials=True, **meta)
```

3c. 改 `_filing_obj()`（`src/fetcher_gaap.py:317`）——在既有的記憶體快取判斷之後、`with_retry` 之前插入磁碟查詢，並在成功之後落檔：

```python
def _filing_obj(filing):
    """下載並解析一份 filing，網路問題會退避重試（CTH 2026-08-17）。

    重試是為了救閃斷——一次瞬斷不該讓那一季永久消失。救不回來就拋
    `NetworkDownError`，呼叫端記進帳本後**繼續抓下一期**，不中止。

    帳本判定網路已經斷掉之後（連續多期失敗）就不再退避重試：整個網路
    斷掉時 40 份財報各等 2+4 秒等於乾等 4 分鐘，剩下的快速失敗完，
    照樣寫檔、照樣提示。

    快取有兩層：記憶體（G9，一次執行內）與磁碟（跨執行）。磁碟命中時回傳的
    是 `filing_cache` 的替身物件，只實作 `.financials` 那一條鏈——四個 builder
    對 filing 物件的用法就只有這一種。
    """
    key = _cache_key(filing)
    if _parse_cache is not None and key is not None and key in _parse_cache:
        return _parse_cache[key]

    ctx = _disk_cache
    if ctx is not None and ctx["ticker"] and key is not None:
        entry = filing_cache.load_filing(ctx["ticker"], key, ctx["cik"])
        if entry is not None:
            ctx["hits"] += 1
            obj = filing_cache.cached_filing(entry)
            if _parse_cache is not None:
                _parse_cache[key] = obj
            return obj

    led = _ledger()
    attempts = 1 if (led is not None and led.give_up_retrying) else 3
    obj = with_retry(lambda: filing.obj(), attempts=attempts)

    if ctx is not None and ctx["ticker"] and key is not None:
        ctx["misses"] += 1
        _save_to_disk_cache(ctx, filing, obj)

    if _parse_cache is not None and key is not None:
        _parse_cache[key] = obj
    return obj
```

3d. 在 `fetch_gaap_statements()` 把磁碟快取範圍跟記憶體快取範圍開在一起（`src/fetcher_gaap.py:2496` 與 `2503` 兩處 `with` 都要加）：

```python
    with _disk_cache_scope(), _parse_cache_scope():
        tables = _fetch_gaap_impl(
            ticker, identity, max_filings, max_annual_filings, ai_config,
            start_year, end_year, fetch_quarterly, fetch_annual, excluded_sheets,
        )

    def _retry_once() -> tuple[list[StatementTable], FetchLedger]:
        with collect_gaps() as retry_led, _disk_cache_scope(), _parse_cache_scope():
            retry_tables = _fetch_gaap_impl(
                ticker, identity, max_filings, max_annual_filings, ai_config,
                start_year, end_year, fetch_quarterly, fetch_annual, excluded_sheets,
            )
        return retry_tables, retry_led
```

3e. 在 `_fetch_gaap_impl()` 建好 `Company` 之後綁定（`src/fetcher_gaap.py` 的 `company = Company(ticker)` 那一行後面）：

```python
    company = Company(ticker)
    # cik 才是跟 SEC 打交道真正的鍵；ticker 只是會換手的別名。拿不到就不用快取。
    _bind_disk_cache(ticker, getattr(company, "cik", None))
```

- [ ] **Step 4: 跑測試確認通過**

Run: `./venv/Scripts/python.exe -m pytest tests/test_fetcher_gaap_cache.py tests/test_fetcher_gaap.py -q`
Expected: PASS（新檔 11 passed，`test_fetcher_gaap.py` 既有測試全數不變）

- [ ] **Step 5: 跑全套確認沒有迴歸**

Run: `./venv/Scripts/python.exe -m pytest tests/ -q`
Expected: PASS（原本 701 passed，加上本任務前的新測試）

- [ ] **Step 6: Commit**

```bash
git add src/fetcher_gaap.py tests/test_fetcher_gaap_cache.py
git commit -m "feat(cache): _filing_obj() 先查本機快取，四個 builder 一起受益

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01BADSq3w2c82y1Ky5yWXD8M"
```

---

## Task 8: log 記快取命中數

耗時變快也可能只是那天 SEC 比較順。沒有命中數字，使用者跟維護者都無法判斷
「這次到底有沒有吃到快取」。

⚠ **這裡對 spec 有一處偏離**：spec 寫的是塞進起始 `===` 行，但那一行在抓取開始
前就寫出去了，那時命中數還不存在。改成 GAAP 抓完後補一行獨立的 `[INFO ]`。

**Files:**
- Modify: `src/main.py`（`_worker_single()` 裡 GAAP 抓取結束處）
- Test: `tests/test_gui_helpers.py`

**Interfaces:**
- Consumes: `fetcher_gaap.last_cache_stats()`（Task 7）
- Produces: `main.cache_log_line(ticker: str, hits: int, total: int) -> str | None`（`total == 0` 回 `None`）

- [ ] **Step 1: 寫失敗的測試**

追加到 `tests/test_gui_helpers.py`：

```python
# ── 快取命中數的 log 行（2026-09-03）────────────────────────────────────────
#
# 耗時變快也可能只是那天 SEC 比較順。沒有命中數字就無法判斷這次有沒有吃到快取。
# `logs/app.log` 一律英文（2026-09-02 起的既有規則）。

def test_cache_log_line_reports_hits_over_total():
    assert main.cache_log_line("NVDA", 24, 25) == "NVDA cache 24/25"


def test_cache_log_line_is_skipped_when_nothing_was_processed():
    """沒抓任何 filing 時不要在 log 留一行 `cache 0/0` 的雜訊。"""
    assert main.cache_log_line("NVDA", 0, 0) is None


def test_cache_log_line_is_english_only():
    """log 的讀者是維護者與 AI，而且這個檔同時被 PowerShell 寫，全英文
    可以整類避開 cp950 的編碼地雷。"""
    line = main.cache_log_line("NVDA", 0, 25)
    assert all(ord(ch) < 128 for ch in line)
```

- [ ] **Step 2: 跑測試確認失敗**

Run: `./venv/Scripts/python.exe -m pytest tests/test_gui_helpers.py -q`
Expected: FAIL — `AttributeError: module 'main' has no attribute 'cache_log_line'`

- [ ] **Step 3: 寫最小實作**

3a. 在 `src/main.py` 的 `format_elapsed()` 附近加：

```python
def cache_log_line(ticker: str, hits: int, total: int) -> str | None:
    """本地 filing 快取的命中率，寫進 `logs/app.log`（一律英文）。

    ⚠ 設計文件原本要求塞進起始 `===` 行，做不到——那一行在抓取**開始前**
    就寫出去了，那時候還沒列 filing 清單、命中數不存在。改成 GAAP 抓完後
    補一行獨立的 INFO，語意相同。
    """
    if not total:
        return None
    return f"{ticker} cache {hits}/{total}"
```

3b. 在 `_worker_single()` 的 GAAP 抓取結束處（`fetch_gaap_statements(...)` 呼叫回來之後）加：

```python
                _hits, _total = fetcher_gaap.last_cache_stats()
                _cache_line = cache_log_line(ticker, _hits, _total)
                if _cache_line:
                    _write_log(_cache_line)
```

（若該處是以 `from fetcher_gaap import fetch_gaap_statements` 的形式引用，就改成
在函式內 `from fetcher_gaap import last_cache_stats` 後直接呼叫——沿用該檔既有的
import 風格，不要為了這一行改動全檔的 import 方式。）

- [ ] **Step 4: 跑測試確認通過**

Run: `./venv/Scripts/python.exe -m pytest tests/test_gui_helpers.py -q`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/main.py tests/test_gui_helpers.py
git commit -m "feat(cache): log 記錄快取命中數（cache 24/25）

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01BADSq3w2c82y1Ky5yWXD8M"
```

---

## Task 9: GUI 字串與兩個純函式

Tab3 區塊本身很難不開視窗測，所以把「會出錯的邏輯」抽成純函式先測掉：
容量格式化、清除鈕的 enable/disable 狀態。

**Files:**
- Modify: `src/locales/zh_tw.py`、`src/locales/zh_cn.py`、`src/locales/en.py`、`src/locales/ja.py`
- Modify: `src/main.py`
- Test: `tests/test_gui_helpers.py`

**Interfaces:**
- Consumes: 無
- Produces:
  - `main.format_size(num_bytes: int) -> str`（`39_200_000` → `"39.2 MB"`）
  - `main.cache_buttons_state(is_running: bool) -> str`（`"disabled"` / `"normal"`）
  - 新的 i18n key：`gui.frame.filing_cache`、`gui.lbl.cache_total`、`gui.lbl.cache_filings`、`gui.lbl.cache_empty`、`gui.btn.cache_open_folder`、`gui.btn.cache_clear`、`gui.btn.cache_clear_all`、`gui.dlg.cache_clear_all_title`、`gui.msg.cache_clear_all_body`

- [ ] **Step 1: 寫失敗的測試**

追加到 `tests/test_gui_helpers.py`：

```python
# ── 快取容量顯示 ────────────────────────────────────────────────────────────
#
# CTH 選的是手動清、不做自動上限，所以「現在到底佔多少」是他做決定的依據，
# 這個數字必須一眼看得懂——不是 41104179 這種原始位元組數。

@pytest.mark.parametrize("num_bytes, expected", [
    (0, "0 KB"),
    (512, "0.5 KB"),
    (61234, "59.8 KB"),
    (18_400_000, "17.5 MB"),
    (1_073_741_824, "1.0 GB"),
])
def test_format_size(num_bytes, expected):
    assert main.format_size(num_bytes) == expected


def test_format_size_never_shows_a_negative():
    assert main.format_size(-1) == "0 KB"


# ── 抓取進行中不可以邊寫邊刪 ────────────────────────────────────────────────

def test_cache_buttons_are_locked_while_a_fetch_is_running():
    """Tab1／批次／跨公司比較任一個 worker thread 還在跑時，兩顆清除按鈕都要
    disable——不然會邊寫邊刪同一個 ticker 的資料夾。沿用專案既有的
    「執行中鎖住相關按鈕」慣例。"""
    assert main.cache_buttons_state(True) == "disabled"
    assert main.cache_buttons_state(False) == "normal"
```

`tests/test_i18n.py` 既有的三條測試會自動擋掉「四個 locale 沒補齊」與
「`src/` 寫死中日文」，不需要另外寫。

- [ ] **Step 2: 跑測試確認失敗**

Run: `./venv/Scripts/python.exe -m pytest tests/test_gui_helpers.py -q`
Expected: FAIL — `AttributeError: module 'main' has no attribute 'format_size'`

- [ ] **Step 3: 寫最小實作**

3a. `src/main.py`（放在 `format_elapsed()` 旁邊）：

```python
def format_size(num_bytes: int) -> str:
    """位元組 → 一眼看得懂的容量。單位符號與語言無關，跟 `format_elapsed()`
    一樣畫面與 log 共用同一個格式，不維護兩套。"""
    n = max(0, int(num_bytes or 0))
    if n >= 1024 ** 3:
        return f"{n / 1024 ** 3:.1f} GB"
    if n >= 1024 ** 2:
        return f"{n / 1024 ** 2:.1f} MB"
    if n == 0:
        return "0 KB"
    return f"{n / 1024:.1f} KB"


def cache_buttons_state(is_running: bool) -> str:
    """抓取進行中鎖住兩顆清除鈕，不然會邊寫邊刪同一個資料夾。"""
    return "disabled" if is_running else "normal"
```

3b. 四個 locale 各補 9 條。`src/locales/zh_tw.py`（接在既有 `gui.frame.*` / `gui.btn.*` 群組後面即可，檔案本身沒有排序要求）：

```python
    "gui.frame.filing_cache": '本地資料快取',
    "gui.lbl.cache_total": '總容量：{size}',
    "gui.lbl.cache_filings": '{count} 份 filing',
    "gui.lbl.cache_empty": '尚無快取資料',
    "gui.btn.cache_open_folder": '開啟資料夾',
    "gui.btn.cache_clear": '清除',
    "gui.btn.cache_clear_all": '全部清除',
    "gui.dlg.cache_clear_all_title": '清除全部快取',
    "gui.msg.cache_clear_all_body": '將刪除全部 {n} 家公司的本地快取（{size}）。\n\n下次抓取這些公司要重新解析 20 年份資料，可能要好幾分鐘。確定要清除嗎？',
```

`src/locales/zh_cn.py`：

```python
    "gui.frame.filing_cache": '本地数据缓存',
    "gui.lbl.cache_total": '总容量：{size}',
    "gui.lbl.cache_filings": '{count} 份 filing',
    "gui.lbl.cache_empty": '尚无缓存数据',
    "gui.btn.cache_open_folder": '打开文件夹',
    "gui.btn.cache_clear": '清除',
    "gui.btn.cache_clear_all": '全部清除',
    "gui.dlg.cache_clear_all_title": '清除全部缓存',
    "gui.msg.cache_clear_all_body": '将删除全部 {n} 家公司的本地缓存（{size}）。\n\n下次抓取这些公司要重新解析 20 年份数据，可能要好几分钟。确定要清除吗？',
```

`src/locales/en.py`：

```python
    'gui.frame.filing_cache': 'Local data cache',
    'gui.lbl.cache_total': 'Total: {size}',
    'gui.lbl.cache_filings': '{count} filings',
    'gui.lbl.cache_empty': 'Nothing cached yet',
    'gui.btn.cache_open_folder': 'Open folder',
    'gui.btn.cache_clear': 'Clear',
    'gui.btn.cache_clear_all': 'Clear all',
    'gui.dlg.cache_clear_all_title': 'Clear all cached data',
    'gui.msg.cache_clear_all_body': 'This deletes the local cache for all {n} companies ({size}).\n\nFetching them again means re-parsing 20 years of filings, which can take several minutes. Continue?',
```

`src/locales/ja.py`：

```python
    'gui.frame.filing_cache': 'ローカルデータキャッシュ',
    'gui.lbl.cache_total': '合計容量：{size}',
    'gui.lbl.cache_filings': '{count} 件のfiling',
    'gui.lbl.cache_empty': 'キャッシュはまだありません',
    'gui.btn.cache_open_folder': 'フォルダを開く',
    'gui.btn.cache_clear': '削除',
    'gui.btn.cache_clear_all': 'すべて削除',
    'gui.dlg.cache_clear_all_title': 'すべてのキャッシュを削除',
    'gui.msg.cache_clear_all_body': '{n} 社分のローカルキャッシュ（{size}）をすべて削除します。\n\n次回取得時は20年分を解析し直すため、数分かかることがあります。よろしいですか？',
```

- [ ] **Step 4: 跑測試確認通過**

Run: `./venv/Scripts/python.exe -m pytest tests/test_gui_helpers.py tests/test_i18n.py -q`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/main.py src/locales/ tests/test_gui_helpers.py
git commit -m "feat(cache): 快取區塊的四語系字串與容量/鎖定純函式

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01BADSq3w2c82y1Ky5yWXD8M"
```

---

## Task 10: Tab3「本地資料快取」區塊

版面：
```
本地資料快取                          總容量：39.2 MB  [開啟資料夾]
┌─────────────────────────────────────────┐ ← 自帶固定高度捲動，約 4~5 列
│ NVDA    102 份 filing    18.4 MB   [清除] │
└─────────────────────────────────────────┘
                              [全部清除]
```

**Files:**
- Modify: `src/main.py`（`_build_settings_panel()` 尾端新增 row 4；新增 `_build_cache_panel()` / `_refresh_cache_panel()` / `_open_cache_folder()` / `_clear_cache_ticker()` / `_clear_all_cache()`；`_run_in_thread()` 與 `done` 分支加刷新與鎖定）
- Test: 手動驗證（見步驟）＋ Task 9 的純函式測試

**Interfaces:**
- Consumes: `filing_cache.list_cached_tickers()` / `total_size_bytes()` / `clear_ticker()` / `clear_all()` / `cache_root()`（Task 6）、`main.format_size()` / `cache_buttons_state()`（Task 9）
- Produces: `SECFetcherApp._refresh_cache_panel()`（可重複呼叫，切到 Tab3 時、清除後、任一次抓取完成後都會呼叫）

- [ ] **Step 1: 寫實作**

1a. `src/main.py` 頂部 import 區加 `import filing_cache`。

1b. `_build_settings_panel()` 最後（`self._on_template_mode_change()` 之前）加一行：

```python
        self._build_cache_panel(popup)
```

1c. 新增方法（放在 `_build_settings_panel()` 之後）：

```python
    def _build_cache_panel(self, popup):
        """Tab3 的「本地資料快取」區塊。

        ⚠ 這塊**自帶第二層固定高度捲動**。Tab3 整頁本來就是靠
        `_build_fixed_height_scrollable(tab, height=self._TAB3_HEIGHT)` 撐住
        （見 `_build_tab3`），快取常駐個位數到二三十家公司，清單直接攤開會把
        Tab3 撐爆、擠掉上面 SEC identity／AI 設定的可視範圍。

        沒有「立即更新」按鈕——每次抓取本來就會自動查新 filing，不需要手動觸發。
        """
        frame = ttk.LabelFrame(popup, text=t("gui.frame.filing_cache"), padding=8)
        frame.grid(row=4, column=0, sticky="ew", padx=12, pady=4)
        frame.columnconfigure(0, weight=1)

        header = ttk.Frame(frame)
        header.grid(row=0, column=0, sticky="ew")
        self._cache_total_label = ttk.Label(header, text="")
        self._cache_total_label.pack(side="left")
        ttk.Button(header, text=t("gui.btn.cache_open_folder"),
                   command=self._open_cache_folder).pack(side="right")

        list_host = ttk.Frame(frame)
        list_host.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        _, self._cache_list_inner = _build_fixed_height_scrollable(list_host, height=110)

        footer = ttk.Frame(frame)
        footer.grid(row=2, column=0, sticky="e", pady=(4, 0))
        self._cache_clear_all_btn = ttk.Button(
            footer, text=t("gui.btn.cache_clear_all"), command=self._clear_all_cache)
        self._cache_clear_all_btn.pack(side="right")

        self._cache_clear_btns = []
        self._refresh_cache_panel()

    def _refresh_cache_panel(self):
        """重畫快取清單。刷新時機：切到 Tab3、任一次清除之後、任一次抓取
        （Tab1／批次／跨公司比較）完成之後——不輪詢。"""
        if not hasattr(self, "_cache_list_inner"):
            return
        for child in self._cache_list_inner.winfo_children():
            child.destroy()
        self._cache_clear_btns = []

        rows = filing_cache.list_cached_tickers()
        total = sum(r["size_bytes"] for r in rows)
        self._cache_total_label.config(
            text=t("gui.lbl.cache_total", size=format_size(total)))

        if not rows:
            ttk.Label(self._cache_list_inner, text=t("gui.lbl.cache_empty"),
                      foreground="#555555").pack(anchor="w")
        for row in rows:
            line = ttk.Frame(self._cache_list_inner)
            line.pack(fill="x", pady=1)
            ttk.Label(line, text=row["ticker"], width=8).pack(side="left")
            ttk.Label(line, text=t("gui.lbl.cache_filings", count=row["count"]),
                      width=16).pack(side="left")
            ttk.Label(line, text=format_size(row["size_bytes"]),
                      width=10).pack(side="left")
            btn = ttk.Button(line, text=t("gui.btn.cache_clear"), width=6,
                             command=lambda tk_=row["ticker"]: self._clear_cache_ticker(tk_))
            btn.pack(side="right")
            self._cache_clear_btns.append(btn)
        self._sync_cache_buttons()

    def _sync_cache_buttons(self):
        """抓取進行中鎖住兩顆清除鈕——不然會邊寫邊刪同一個 ticker 的資料夾。"""
        state = cache_buttons_state(bool(getattr(self, "is_running", False)))
        for btn in getattr(self, "_cache_clear_btns", []):
            btn.config(state=state)
        if getattr(self, "_cache_clear_all_btn", None):
            self._cache_clear_all_btn.config(state=state)

    def _open_cache_folder(self):
        """讓使用者自己用檔案總管進一步查看／處理，不用我們另外做細部管理 UI。"""
        root = filing_cache.cache_root()
        try:
            root.mkdir(parents=True, exist_ok=True)
            os.startfile(str(root))
        except OSError as exc:
            _write_log(f"cannot open cache folder: {type(exc).__name__}", "ERROR")

    def _clear_cache_ticker(self, ticker: str):
        """整個刪掉那家公司的資料夾。下次抓這家會當作全新開始。
        單一公司不做二次確認——重抓一家的代價有限，跳確認反而礙事。"""
        filing_cache.clear_ticker(ticker)
        _write_log(f"cache cleared for {ticker}")
        self._refresh_cache_panel()

    def _clear_all_cache(self):
        """唯一不可逆的破壞性操作，要二次確認——雖然只是快取，
        重抓 20 年份是好幾分鐘的代價，值得防手滑。"""
        rows = filing_cache.list_cached_tickers()
        if not rows:
            return
        total = format_size(sum(r["size_bytes"] for r in rows))
        if not messagebox.askyesno(
                t("gui.dlg.cache_clear_all_title"),
                t("gui.msg.cache_clear_all_body", n=len(rows), size=total)):
            return
        removed = filing_cache.clear_all()
        _write_log(f"cache cleared for all {removed} companies")
        self._refresh_cache_panel()
```

1d. 抓取開始／結束時同步狀態與內容：
- `_run_in_thread()` 裡 `self.is_running = True` 之後加 `self._sync_cache_buttons()`
- `_poll_queue()` 的 `done` 分支 `self.is_running = False` 之後加
  `self._sync_cache_buttons()` 與 `self._refresh_cache_panel()`
- `compare_done` 與 `compare_error` 兩個分支各加 `self._refresh_cache_panel()`
- Notebook 切頁事件（`self.notebook.bind("<<NotebookTabChanged>>", ...)`）：切到 Tab3 時呼叫
  `self._refresh_cache_panel()`。該檔若還沒有這個 binding 就新增一個，
  handler 內用 `self.notebook.index("current") == 2` 判斷

- [ ] **Step 2: 跑既有測試確認沒壞**

Run: `./venv/Scripts/python.exe -m pytest tests/ -q`
Expected: PASS（`test_i18n.py` 會驗四個 locale 補齊、`src/` 沒有寫死中日文）

- [ ] **Step 3: 重新量測 `_TAB3_HEIGHT`**

`docs/ARCHITECTURE.md`「視窗擺放」的既有規則：改任何一頁的版面之後要重量。
在 scratchpad 寫一支 Tk 探針（**不要放進 `scripts/`**）：

```python
# C:\Users\CTH\AppData\Local\Temp\claude\...\scratchpad\probe_tab_heights.py
import sys
from pathlib import Path
sys.path.insert(0, str(Path(r"C:\Users\CTH\Documents\Code\SEC Financial Tools") / "src"))
import tkinter as tk
import main

root = tk.Tk()
app = main.SECFetcherApp(root)
root.update_idletasks()
for i in range(app.notebook.index("end")):
    tab = app.notebook.nametowidget(app.notebook.tabs()[i])
    print(i, app.notebook.tab(i, "text"), tab.winfo_reqheight())
print("notebook", app.notebook.winfo_reqheight())
root.destroy()
```

Run: `./venv/Scripts/python.exe <scratchpad>/probe_tab_heights.py`

判準：Tab3 的 `winfo_reqheight()` 要跟 Tab1 相差在 ±5px 內、Notebook 整體高度
與加這塊之前一樣。差太多就調 `_TAB3_HEIGHT`（現值 355）重跑，直到貼齊，
並把新的實測值與日期寫進 `_TAB3_HEIGHT` 上方的註解。

- [ ] **Step 4: 手動驗收 GUI**

雙擊 `啟動器.bat`（或 `./venv/Scripts/python.exe src/main.py`），逐項確認：

1. Tab3 出現「本地資料快取」區塊，清單超過 5 家時**這塊自己捲動**，Tab3 整頁
   高度與其他分頁不變（上面 SEC identity 沒被擠掉）
2. 抓一家公司（例如 ARLO），回到 Tab3 → 出現該公司一列，份數與容量合理
3. 抓取進行中切到 Tab3 → 兩顆清除鈕都是灰的
4. 按單一公司「清除」→ 該列消失，資料夾真的不見了；再抓同一家 → 會重新解析
5. 按「全部清除」→ 跳確認對話框，取消不刪、確定才刪
6. 「開啟資料夾」開得起檔案總管
7. 切四種語言各看一次，沒有殘留的 `{size}` / `{count}` 未格式化字樣

- [ ] **Step 5: Commit**

```bash
git add src/main.py
git commit -m "feat(cache): Tab3 新增本地資料快取區塊（自帶捲動、執行中鎖清除鈕）

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01BADSq3w2c82y1Ky5yWXD8M"
```

---

## Task 11: 驗收（0 格不同 + ARLO 效能基準）＋ 文件

驗收方法 spec 裡寫死了：`scripts/excel_golden.py` 的「清快取重抓」vs「讀快取」
**0 格不同**，加上 ARLO 的 3~5 倍效能基準（**不是「秒開」**——查清單那約 18s
的網路往返快取消不掉）。

**Files:**
- Modify: `docs/ARCHITECTURE.md`、`docs/TODO.md`
- Test: `scripts/excel_golden.py`（既有工具，不改）

**Interfaces:**
- Consumes: 前面所有任務
- Produces: 無新程式碼

- [ ] **Step 1: 冷跑（清快取重抓）並計時**

```bash
./venv/Scripts/python.exe -c "import sys; sys.path.insert(0,'src'); import filing_cache; print(filing_cache.clear_ticker('ARLO'))"
```
用 GUI 或 `src/cli.py` 抓 ARLO（GAAP、10-Q + 10-K、預設 max_filings），
記下 `logs/app.log` 的 elapsed，把產出的 xlsx 放進 `output/_final/`，然後：

```bash
./venv/Scripts/python.exe scripts/excel_golden.py make output/_golden_cache_base
```

- [ ] **Step 2: 熱跑（全部讀快取）並計時**

同樣參數再抓一次 ARLO（不清快取），記下 elapsed 與 `logs/app.log` 的
`ARLO cache X/Y` 那一行（X 應該等於 Y），然後：

```bash
./venv/Scripts/python.exe scripts/excel_golden.py make  output/_golden_cache_new
./venv/Scripts/python.exe scripts/excel_golden.py check output/_golden_cache_base output/_golden_cache_new
```

Expected: exit code 0，**0 格不同**。有任何一格不同就是 bug，不是可接受的差異
——快取只能改變耗時，不能改變任何輸出內容。

- [ ] **Step 3: 對照效能基準**

判準：熱跑 elapsed 落在冷跑的 **1/3 ~ 1/5**（ARLO 冷跑約 66s → 熱跑 <15s 之內）。
沒達標先看 `cache X/Y` 是不是真的全命中；全命中還是不夠快，量一下
`_list_filings()` 與 `_fetch_shares_outstanding()` 這兩段**本來就不受快取影響**
的網路時間再判斷。

同時記錄**冷跑有沒有變慢**：miss 時會多做一次三張表的 `to_dataframe()`（為了落檔），
理論上會讓冷跑比改動前慢一些。若冷跑比改動前慢超過 30%，記進 ARCHITECTURE 的
已知取捨（可行的優化是 miss 之後改回傳替身物件重用那三張表，但那會讓
golden 比對失去意義，所以不預設採用）。

- [ ] **Step 4: 手動驗證「新一期財報會被抓到」**

在 SEC 上確認某家公司最新一期 10-Q 已經在快取裡；等該公司真的發了新一期
（或改用一家已知剛發財報的公司）再抓一次，確認新那一期有進來、不是繼續讀舊資料。
這條沒辦法自動化，做過就記在 ARCHITECTURE 那節。

- [ ] **Step 5: 寫文件**

5a. `docs/ARCHITECTURE.md` 在「解析快取（`_parse_cache_scope()`，2026-08-22）」
那節後面新增一節，內容要涵蓋：

- 快取卡在解析層與比對層之間，比對規則改版不會讓快取失效；**但 edgartools
  升版會**，靠 `edgartools_version` 欄位擋
- 儲存位置與兩種檔案（`<accession>.json` 是事實來源、`_manifest.json` 是衍生索引）
- 四道閘（schema / cik / edgartools 版本 / JSON 損毀）
- 負向快取（`has_financials: false`）**與網路失敗的分野**
- 替身物件的兩條隱性規則（未知屬性拋 `AttributeError`、`None` ≠ 空 DataFrame）
- 這次量到的冷跑／熱跑實測值與 golden 0 格不同的結果
- **不涵蓋**什麼：Non-GAAP（`fetcher_nongaap.py` 有自己的 `nongaap_cache.json`）、
  `company.get_facts()`（流通股數）、`_list_filings()` 的清單查詢
- 修正案（10-Q/A、10-K/A）現況 `_list_filings(amendments=False)` 本來就不抓，
  跟這個快取無關

5b. 更新 `_TAB3_HEIGHT` 註解旁的實測記錄（Task 10 Step 3 量到的值與日期）。

5c. `docs/TODO.md` 把「本地 filing 快取」條目移到已完成，並另開一條
「抓取修正案（10-Q/A、10-K/A）」的獨立待辦（spec 明講這是另一個議題）。

- [ ] **Step 6: 跑全套測試**

Run: `./venv/Scripts/python.exe -m pytest tests/ -q`
Expected: PASS，總數 = 原本 701 + 本計畫新增（約 55 條）

- [ ] **Step 7: Commit**

```bash
git add docs/ARCHITECTURE.md docs/TODO.md src/main.py
git commit -m "docs(cache): 本地 filing 快取的架構說明、實測值與 _TAB3_HEIGHT 重量結果

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01BADSq3w2c82y1Ky5yWXD8M"
```

---

## Spec 覆蓋對照

| Spec 章節／要求 | 對應 Task |
|---|---|
| 一、快取卡在解析層與比對層之間；掛勾點 `_filing_obj()` | 7 |
| 一、替身物件三個約束（`.financials`、未知屬性 AttributeError、`None` ≠ 空表） | 2 |
| 一、G9 記憶體快取不用改 | 7（`test_a_cache_hit_still_goes_through_the_in_memory_parse_cache`） |
| 二、儲存位置、ticker 資料夾、accession 檔名 | 3 |
| 二、`_manifest.json` 是衍生索引 | 5 |
| 二、`cik` / `edgartools_version` / `schema_version` 欄位 | 4 |
| 二、存物件不存字串、`dtypes` 必要、`importlib.metadata`、原子寫入帶 PID | 1、3 |
| 三、每次都查清單、accession 比對、逐份即時落檔、負向快取 | 7 |
| 三、log 記快取命中數 | 8（位置有偏離，見上方說明） |
| 四、GUI Tab3 區塊、自帶捲動、重量 `_TAB3_HEIGHT`、清除／全部清除／確認框／執行中鎖定／四語系 | 9、10 |
| 五、錯誤處理（損毀、寫入失敗、manifest 重建、原子寫入） | 3、4、5 |
| 六、測試逐條 | 1~7、9 |
| 驗收（golden 0 格不同、ARLO 3~5 倍、新一期財報、`_TAB3_HEIGHT`） | 10、11 |
| 不做的事（無 DB、無自動清理、不改底層抓取、不影響呼叫端） | 全案（Global Constraints） |
