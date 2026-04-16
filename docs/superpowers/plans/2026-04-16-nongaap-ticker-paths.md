# Non-GAAP Fetching + Per-Ticker Output Path Memory Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 實作兩個功能：(1) 記住每間公司的輸出資料夾路徑；(2) 從 8-K press release 用 edgartools + AI 抓取 Non-GAAP 財報，輸出至 `Data_EPS_Recon` 和 `Data_NonGAAP` sheet，並以本機 JSON 快取支援增量更新。

**Architecture:** Per-ticker 路徑存於 `config.json["ticker_paths"]`，UI 確認公司時自動帶出，瀏覽後自動儲存。Non-GAAP 以 `fetcher_nongaap.py` 實作，從 8-K Item 2.02 取 EPS reconciliation（edgartools 原生）與 AI 提取指標（EX-99.1 press release → markdown → AI → JSON），結果快取於各公司輸出資料夾的 `nongaap_cache.json`，只對新季度呼叫 AI。

**Tech Stack:** Python 3.11+、edgartools v5.29、openpyxl、tkinter、google-generativeai / openai / anthropic（依使用者設定）

---

## 檔案結構

| 檔案 | 變動類型 | 說明 |
|------|----------|------|
| `config.py` | Modify | 新增 `ticker_paths: {}` 預設值 |
| `fetcher_nongaap.py` | Rewrite | 完整實作（取代 placeholder） |
| `main.py` | Modify | `_build_output_path`、`_browse_output_dir`、`_poll_queue`、`_worker_single` |
| `tests/test_fetcher_nongaap.py` | Create | Non-GAAP 單元測試 |
| `tests/test_config.py` | Create | config ticker_paths 測試 |
| `CHANGELOG.md` | Modify | 記錄新功能 |
| `ARCHITECTURE.md` | Modify | 更新資料流與新 sheet 說明 |

---

## Group A：Per-Ticker Output Path Memory

### Task 1：config.py 新增 ticker_paths 預設值

**Files:**
- Modify: `config.py`
- Create: `tests/test_config.py`

- [ ] **Step 1：寫失敗測試**

```python
# tests/test_config.py
import json, tempfile
from pathlib import Path
from config import load_config, save_config

def test_load_config_has_ticker_paths_default():
    cfg = load_config(Path("/nonexistent/config.json"))
    assert "ticker_paths" in cfg
    assert cfg["ticker_paths"] == {}

def test_ticker_paths_persists_through_save_load():
    with tempfile.NamedTemporaryFile(suffix=".json", delete=False) as f:
        path = Path(f.name)
    cfg = load_config(path)
    cfg["ticker_paths"]["TSLA"] = "C:\\Work\\TSLA"
    save_config(cfg, path)
    loaded = load_config(path)
    assert loaded["ticker_paths"]["TSLA"] == "C:\\Work\\TSLA"
```

- [ ] **Step 2：確認測試失敗**

```
pytest tests/test_config.py -v
```
預期：`test_load_config_has_ticker_paths_default` FAIL（KeyError: ticker_paths）

- [ ] **Step 3：修改 config.py**

```python
DEFAULT_CONFIG: dict = {
    "identity": "",
    "output_dir": "output",
    "ticker_paths": {},          # ← 新增
    "watchlist": [],
    "filename_format": "ticker_name",
    "filename_custom": "",
    "max_filings": 80,
    "ai": {
        "provider": "google",
        "model": "gemini-flash-latest",
        "api_key": "",
    },
}
```

`load_config` 的合併邏輯已處理 dict 型別，`ticker_paths` 會被正確 merge，不需額外修改。

- [ ] **Step 4：確認測試通過**

```
pytest tests/test_config.py -v
```
預期：2 passed

- [ ] **Step 5：Commit**

```bash
git add config.py tests/test_config.py
git commit -m "feat: add ticker_paths to config defaults"
```

---

### Task 2：main.py — _build_output_path 優先查 ticker_paths

**Files:**
- Modify: `main.py:669-682`

- [ ] **Step 1：寫失敗測試（手動驗證，無法 unittest main.py GUI，改用 smoke check）**

在 Python shell 驗證（不用 pytest，因為 GUI 需要 display）：

```python
# 在 main.py 底部暫時加這段確認邏輯（測完刪掉）
# cfg["ticker_paths"]["TSLA"] = "C:\\Work\\TSLA"
# app._build_output_path("TSLA") 應回傳 C:\Work\TSLA\TSLA Tesla, Inc. data.xlsx
```

- [ ] **Step 2：修改 `_build_output_path`**

```python
def _build_output_path(self, ticker: str) -> Path:
    """Build output file path. Per-ticker path takes priority over output_dir."""
    ticker_dir = self.cfg.get("ticker_paths", {}).get(ticker)
    if ticker_dir:
        output_dir = Path(ticker_dir)
    else:
        output_dir = SCRIPT_DIR / self.cfg.get("output_dir", "output")

    fmt = self.cfg.get("filename_format", "ticker_name")
    if fmt == "ticker_name":
        name = self._lookup_company_name(ticker)
        safe_name = re.sub(r'[\\/:*?"<>|]', "", name).strip()
        filename = f"{ticker} {safe_name} data.xlsx"
    elif fmt == "custom":
        custom = re.sub(r'[\\/:*?"<>|]', "", self.cfg.get("filename_custom", "")).strip()
        filename = f"{custom}.xlsx" if custom else f"{ticker}.xlsx"
    else:
        filename = f"{ticker}.xlsx"
    return output_dir / filename
```

- [ ] **Step 3：Commit**

```bash
git add main.py
git commit -m "feat: _build_output_path checks ticker_paths first"
```

---

### Task 3：main.py — 瀏覽後自動儲存 ticker_paths；確認公司後自動帶出路徑

**Files:**
- Modify: `main.py:323-339`（`_browse_output_dir`、`_save_tab1_output_settings`）
- Modify: `main.py:842-851`（`_poll_queue` tab1_name_result 段）

- [ ] **Step 1：修改 `_browse_output_dir`，選完後存進 ticker_paths**

```python
def _browse_output_dir(self):
    from tkinter import filedialog
    current = self.tab1_outdir_var.get().strip() if self.tab1_outdir_var else "output"
    initial = str(SCRIPT_DIR / current) if not os.path.isabs(current) else current
    folder = filedialog.askdirectory(title="選擇儲存位置", initialdir=initial)
    if folder:
        self.tab1_outdir_var.set(folder)
        # 記住這個 ticker 的路徑
        ticker = self._get_ph_value(self.ticker_var, self.TICKER_PH).upper()
        if ticker:
            if "ticker_paths" not in self.cfg:
                self.cfg["ticker_paths"] = {}
            self.cfg["ticker_paths"][ticker] = folder
        self._save_tab1_output_settings()
```

- [ ] **Step 2：修改 `_poll_queue` tab1_name_result，確認公司後自動帶出已存路徑**

找到 `elif msg_type == "tab1_name_result":` 段，在 `self._update_tab1_preview()` 後加入：

```python
elif msg_type == "tab1_name_result":
    status, looked_ticker, name = data
    current = self._get_ph_value(self.ticker_var, self.TICKER_PH).upper()
    if self.tab1_name_label and current == looked_ticker:
        if status == "ok":
            self.tab1_name_label.config(text=f"　{name}", foreground="#1a7a34")
            # 自動帶出已記憶的路徑
            saved_path = self.cfg.get("ticker_paths", {}).get(looked_ticker)
            if saved_path and self.tab1_outdir_var:
                self.tab1_outdir_var.set(saved_path)
        else:
            self.tab1_name_label.config(text="　查無此 Ticker，請確認後再試", foreground="orange")
        self._update_tab1_preview()
    if self.btn_confirm_company:
        self.btn_confirm_company.config(state="normal")
```

- [ ] **Step 3：啟動程式手動測試**

```
雙擊 啟動器.bat
1. 輸入 TSLA → 確認公司 → 「儲存位置」應為空（首次）
2. 點「瀏覽」選 C:\Users\CTH\Documents\Work\...\美股\TSLA
3. 關閉程式，重新開啟
4. 輸入 TSLA → 確認公司 → 「儲存位置」應自動帶出 TSLA 資料夾
```

- [ ] **Step 4：Commit**

```bash
git add main.py
git commit -m "feat: remember and restore per-ticker output path"
```

---

## Group B：Non-GAAP Fetching

### Task 4：fetcher_nongaap.py — 基礎工具函式（cache I/O + quarter label）

**Files:**
- Rewrite: `fetcher_nongaap.py`
- Create: `tests/test_fetcher_nongaap.py`

- [ ] **Step 1：寫失敗測試**

```python
# tests/test_fetcher_nongaap.py
import json, tempfile
from pathlib import Path
from fetcher_nongaap import _load_cache, _save_cache, _period_to_quarter_label

def test_period_to_quarter_label_q1():
    assert _period_to_quarter_label("20240331") == "FY2024Q1"

def test_period_to_quarter_label_q2():
    assert _period_to_quarter_label("20240630") == "FY2024Q2"

def test_period_to_quarter_label_q3():
    assert _period_to_quarter_label("20240930") == "FY2024Q3"

def test_period_to_quarter_label_q4():
    assert _period_to_quarter_label("20241231") == "FY2024Q4"

def test_period_with_dashes():
    assert _period_to_quarter_label("2024-03-31") == "FY2024Q1"

def test_load_cache_missing_file():
    result = _load_cache(Path("/nonexistent/nongaap_cache.json"))
    assert result == {}

def test_save_and_load_cache(tmp_path):
    cache_path = tmp_path / "nongaap_cache.json"
    data = {"FY2024Q1": {"metrics": {"Non-GAAP EPS": 0.71}}}
    _save_cache(cache_path, data)
    loaded = _load_cache(cache_path)
    assert loaded["FY2024Q1"]["metrics"]["Non-GAAP EPS"] == 0.71
```

- [ ] **Step 2：確認測試失敗**

```
pytest tests/test_fetcher_nongaap.py -v
```
預期：ImportError（函式尚未存在）

- [ ] **Step 3：實作基礎工具函式（rewrite fetcher_nongaap.py）**

```python
"""
fetcher_nongaap.py — Non-GAAP data extraction from 8-K press releases via AI.

Flow:
  8-K (Item 2.02) → edgartools eps_reconciliation + AI on EX-99.1 press release
  → nongaap_cache.json (per-ticker, incremental) → StatementTable list
"""

import json
import sys
import unicodedata
from pathlib import Path
from typing import Any

from edgar import Company, set_identity

from fetcher_gaap import StatementTable

CACHE_FILENAME = "nongaap_cache.json"


# ── Cache I/O ───────────────────────────────────────────────────────────────

def _load_cache(cache_path: Path) -> dict:
    """Load nongaap_cache.json. Returns {} if missing or malformed."""
    if not cache_path.exists():
        return {}
    try:
        with open(cache_path, encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return {}


def _save_cache(cache_path: Path, data: dict) -> None:
    """Save cache dict to JSON. Creates parent dirs if needed."""
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    with open(cache_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ── Period helpers ───────────────────────────────────────────────────────────

def _period_to_quarter_label(period_of_report: str) -> str:
    """Convert '20240331' or '2024-03-31' to 'FY2024Q1'."""
    period = period_of_report.replace("-", "")
    year = period[:4]
    month = int(period[4:6])
    if month <= 3:
        suffix = "Q1"
    elif month <= 6:
        suffix = "Q2"
    elif month <= 9:
        suffix = "Q3"
    else:
        suffix = "Q4"
    return f"FY{year}{suffix}"
```

- [ ] **Step 4：確認測試通過**

```
pytest tests/test_fetcher_nongaap.py -v
```
預期：7 passed

- [ ] **Step 5：Commit**

```bash
git add fetcher_nongaap.py tests/test_fetcher_nongaap.py
git commit -m "feat: nongaap cache I/O and period label helpers"
```

---

### Task 5：fetcher_nongaap.py — StatementTable 建構（從 cache 重建兩張 sheet）

**Files:**
- Modify: `fetcher_nongaap.py`（新增 `_build_eps_recon_table`、`_build_nongaap_table`）
- Modify: `tests/test_fetcher_nongaap.py`

- [ ] **Step 1：寫失敗測試**

```python
# 加入 tests/test_fetcher_nongaap.py
from fetcher_nongaap import _build_eps_recon_table, _build_nongaap_table

SAMPLE_CACHE = {
    "FY2024Q1": {
        "filing_date": "2024-01-24",
        "eps_recon": {"GAAP EPS": 0.53, "SBC": -0.12, "Non-GAAP EPS": 0.65},
        "metrics": {"Non-GAAP Net Income": 2513000000.0, "Adjusted EBITDA": 3800000000.0},
    },
    "FY2024Q2": {
        "filing_date": "2024-04-23",
        "eps_recon": {"GAAP EPS": 0.42, "SBC": -0.10, "Non-GAAP EPS": 0.52},
        "metrics": {"Non-GAAP Net Income": 1800000000.0},
    },
}

def test_build_eps_recon_table_sheet_name():
    tbl = _build_eps_recon_table("TSLA", SAMPLE_CACHE)
    assert tbl is not None
    assert tbl.sheet_name == "Data_EPS_Recon"

def test_build_eps_recon_table_quarters_oldest_to_newest():
    tbl = _build_eps_recon_table("TSLA", SAMPLE_CACHE)
    assert tbl.quarter_labels == ["FY2024Q1", "FY2024Q2"]

def test_build_eps_recon_table_concepts():
    tbl = _build_eps_recon_table("TSLA", SAMPLE_CACHE)
    assert "GAAP EPS" in tbl.concepts
    assert "Non-GAAP EPS" in tbl.concepts

def test_build_eps_recon_table_values():
    tbl = _build_eps_recon_table("TSLA", SAMPLE_CACHE)
    gaap_idx = tbl.concepts.index("GAAP EPS")
    assert tbl.values[gaap_idx] == [0.53, 0.42]

def test_build_nongaap_table_sheet_name():
    tbl = _build_nongaap_table("TSLA", SAMPLE_CACHE)
    assert tbl is not None
    assert tbl.sheet_name == "Data_NonGAAP"

def test_build_nongaap_table_union_of_metrics():
    tbl = _build_nongaap_table("TSLA", SAMPLE_CACHE)
    # Adjusted EBITDA only in Q1, should still appear with None in Q2
    assert "Adjusted EBITDA" in tbl.concepts
    ebitda_idx = tbl.concepts.index("Adjusted EBITDA")
    assert tbl.values[ebitda_idx] == [3800000000.0, None]

def test_build_eps_recon_table_empty_cache():
    assert _build_eps_recon_table("TSLA", {}) is None

def test_build_nongaap_table_empty_cache():
    assert _build_nongaap_table("TSLA", {}) is None
```

- [ ] **Step 2：確認測試失敗**

```
pytest tests/test_fetcher_nongaap.py -v
```
預期：ImportError（函式尚未存在）

- [ ] **Step 3：實作 `_build_eps_recon_table` 和 `_build_nongaap_table`**

在 `fetcher_nongaap.py` 的快取函式後加入：

```python
# ── StatementTable builders ──────────────────────────────────────────────────

def _build_eps_recon_table(ticker: str, cache: dict) -> StatementTable | None:
    """Build Data_EPS_Recon StatementTable from cache. Returns None if cache empty."""
    if not cache:
        return None

    sorted_qs = sorted(cache.keys())
    filing_dates = [cache[q].get("filing_date", "") for q in sorted_qs]

    # Collect all EPS recon keys (union across quarters)
    all_keys: list[str] = []
    seen: set[str] = set()
    for q in sorted_qs:
        for key in cache[q].get("eps_recon", {}):
            if key not in seen:
                all_keys.append(key)
                seen.add(key)

    if not all_keys:
        return None

    values: list[list[Any]] = []
    for key in all_keys:
        values.append([cache[q].get("eps_recon", {}).get(key) for q in sorted_qs])

    return StatementTable(
        sheet_name="Data_EPS_Recon",
        quarter_labels=sorted_qs,
        filing_dates=filing_dates,
        concepts=all_keys,
        values=values,
        ticker=ticker,
        labels=[""] * len(all_keys),
    )


def _build_nongaap_table(ticker: str, cache: dict) -> StatementTable | None:
    """Build Data_NonGAAP StatementTable from cache. Returns None if cache empty."""
    if not cache:
        return None

    sorted_qs = sorted(cache.keys())
    filing_dates = [cache[q].get("filing_date", "") for q in sorted_qs]

    # Union of all metric names
    all_metrics: list[str] = []
    seen: set[str] = set()
    for q in sorted_qs:
        for key in cache[q].get("metrics", {}):
            if key not in seen:
                all_metrics.append(key)
                seen.add(key)

    if not all_metrics:
        return None

    values: list[list[Any]] = []
    for metric in all_metrics:
        values.append([cache[q].get("metrics", {}).get(metric) for q in sorted_qs])

    return StatementTable(
        sheet_name="Data_NonGAAP",
        quarter_labels=sorted_qs,
        filing_dates=filing_dates,
        concepts=all_metrics,
        values=values,
        ticker=ticker,
        labels=[""] * len(all_metrics),
    )
```

- [ ] **Step 4：確認測試通過**

```
pytest tests/test_fetcher_nongaap.py -v
```
預期：15 passed

- [ ] **Step 5：Commit**

```bash
git add fetcher_nongaap.py tests/test_fetcher_nongaap.py
git commit -m "feat: build EPS recon and NonGAAP StatementTables from cache"
```

---

### Task 6：fetcher_nongaap.py — EPS reconciliation 提取（edgartools）

**Files:**
- Modify: `fetcher_nongaap.py`（新增 `_extract_eps_recon`）

- [ ] **Step 1：實作 `_extract_eps_recon`**

```python
# 加入 fetcher_nongaap.py

def _extract_eps_recon(eight_k) -> dict[str, float]:
    """Extract EPS reconciliation using edgartools native support.

    Returns dict like {"GAAP EPS": 0.53, "SBC": -0.12, "Non-GAAP EPS": 0.65}.
    Returns {} if not available.
    """
    try:
        earnings = getattr(eight_k, "earnings", None)
        if earnings is None:
            return {}
        recon = getattr(earnings, "eps_reconciliation", None)
        if recon is None:
            return {}
        df = recon.dataframe
        if df is None or df.empty:
            return {}

        result: dict[str, float] = {}
        # DataFrame has variable structure; take first numeric column as value
        value_cols = [c for c in df.columns if c not in {"label", "concept", "description"}]
        if not value_cols:
            return {}
        val_col = value_cols[0]

        label_col = "label" if "label" in df.columns else df.columns[0]
        for _, row in df.iterrows():
            label = str(row.get(label_col, "") or "").strip()
            val = row.get(val_col)
            if label and val is not None:
                try:
                    result[label] = float(val)
                except (ValueError, TypeError):
                    pass
        return result
    except Exception as exc:
        print(f"[fetcher_nongaap] eps_recon warning: {exc!r}", file=sys.stderr)
        return {}
```

- [ ] **Step 2：手動 smoke test（需要 EDGAR identity 設定）**

若本機已設定 identity，在 Python shell 跑：

```python
from edgar import Company, set_identity
set_identity("Your Name your@email.com")
company = Company("AAPL")
filing = list(company.get_filings(form="8-K"))[0]
eight_k = filing.obj()
from fetcher_nongaap import _extract_eps_recon
print(_extract_eps_recon(eight_k))
# 預期：dict 或 {}（若該 8-K 非 earnings release）
```

- [ ] **Step 3：Commit**

```bash
git add fetcher_nongaap.py
git commit -m "feat: extract EPS recon via edgartools"
```

---

### Task 7：fetcher_nongaap.py — AI 提取 Non-GAAP 指標

**Files:**
- Modify: `fetcher_nongaap.py`（新增 `_call_ai`、`_extract_nongaap_metrics`）

- [ ] **Step 1：實作 `_call_ai`（provider-agnostic）**

```python
# 加入 fetcher_nongaap.py

_NONGAAP_PROMPT = """\
你是財務分析師。以下是一份公司季度財報新聞稿（Markdown 格式）。
請提取所有 Non-GAAP 財務指標，回傳 JSON 格式：

{{"指標名稱": 數值（純數字，不含貨幣符號或逗號）}}

規則：
- 只取 Non-GAAP / Adjusted / Excluding 相關指標
- 金額若以百萬為單位則乘以 1000000，以十億為單位則乘以 1000000000
- 百分比直接回傳數字（如 17.6%→17.6）
- 若找不到任何 Non-GAAP 指標，回傳空 JSON {{}}
- 只回傳 JSON，不要說明文字

新聞稿內容：
{press_release_text}
"""


def _call_ai(text: str, ai_config: dict) -> dict[str, Any]:
    """Call configured AI provider with press release text. Returns parsed JSON dict."""
    provider = ai_config.get("provider", "google")
    model    = ai_config.get("model", "")
    api_key  = ai_config.get("api_key", "")
    prompt   = _NONGAAP_PROMPT.format(press_release_text=text[:12000])  # token guard

    try:
        if provider == "google":
            import google.generativeai as genai
            genai.configure(api_key=api_key)
            response = genai.GenerativeModel(model).generate_content(prompt)
            raw = response.text
        elif provider == "openai":
            from openai import OpenAI
            response = OpenAI(api_key=api_key).chat.completions.create(
                model=model,
                messages=[{"role": "user", "content": prompt}],
                max_tokens=1024,
            )
            raw = response.choices[0].message.content
        elif provider == "anthropic":
            import anthropic
            response = anthropic.Anthropic(api_key=api_key).messages.create(
                model=model, max_tokens=1024,
                messages=[{"role": "user", "content": prompt}],
            )
            raw = response.content[0].text
        else:
            return {}

        # Strip markdown code fences if present
        raw = raw.strip()
        if raw.startswith("```"):
            raw = "\n".join(raw.split("\n")[1:])
        if raw.endswith("```"):
            raw = raw.rsplit("```", 1)[0]

        parsed = json.loads(raw.strip())
        return {k: float(v) for k, v in parsed.items() if v is not None}

    except Exception as exc:
        print(f"[fetcher_nongaap] AI call failed: {exc!r}", file=sys.stderr)
        return {}
```

- [ ] **Step 2：實作 `_extract_nongaap_metrics`**

```python
def _extract_nongaap_metrics(eight_k, ai_config: dict) -> dict[str, Any]:
    """Get press release text and call AI to extract Non-GAAP metrics.

    Returns dict of {metric_name: value}. Returns {} on any failure.
    """
    try:
        # Try edgartools press_releases first
        press_releases = getattr(eight_k, "press_releases", None)
        text = None

        if press_releases:
            for pr in press_releases:
                try:
                    text = pr.markdown() or pr.text()
                    if text:
                        break
                except Exception:
                    continue

        # Fallback: search attachments for EX-99.1
        if not text:
            try:
                attachments = eight_k._filing.attachments
                for att in attachments:
                    doc_type = str(getattr(att, "document_type", "") or "")
                    if "EX-99" in doc_type.upper():
                        text = att.markdown() if hasattr(att, "markdown") else att.text()
                        if text:
                            break
            except Exception:
                pass

        if not text:
            return {}

        text = unicodedata.normalize("NFKC", text)
        return _call_ai(text, ai_config)

    except Exception as exc:
        print(f"[fetcher_nongaap] metrics extraction failed: {exc!r}", file=sys.stderr)
        return {}
```

- [ ] **Step 3：Commit**

```bash
git add fetcher_nongaap.py
git commit -m "feat: AI extraction of Non-GAAP metrics from 8-K press release"
```

---

### Task 8：fetcher_nongaap.py — 8-K discovery + 增量抓取主流程

**Files:**
- Modify: `fetcher_nongaap.py`（新增 `_get_earnings_filings`、`fetch_nongaap_statements`）

- [ ] **Step 1：實作 `_get_earnings_filings`**

```python
# 加入 fetcher_nongaap.py

def _get_earnings_filings(company) -> list[tuple[str, Any]]:
    """Return list of (quarter_label, filing) for 8-K filings with Item 2.02.

    Sorted oldest → newest.
    """
    results = []
    for filing in company.get_filings(form="8-K", amendments=False):
        try:
            eight_k = filing.obj()
            items = getattr(eight_k, "items", []) or []
            has_202 = any("2.02" in str(item) for item in items)
            if not has_202:
                # Also check has_earnings as fallback
                if not getattr(eight_k, "has_earnings", False):
                    continue
            period = str(filing.period_of_report or "").replace("-", "")
            if len(period) < 8:
                continue
            label = _period_to_quarter_label(period)
            results.append((label, filing))
        except Exception as exc:
            print(f"[fetcher_nongaap] 8-K scan warning: {exc!r}", file=sys.stderr)
            continue

    # Sort oldest first, deduplicate by quarter_label (keep first = most recent filing for that period)
    seen: set[str] = set()
    deduped = []
    for label, filing in reversed(results):
        if label not in seen:
            seen.add(label)
            deduped.append((label, filing))
    return list(reversed(deduped))
```

- [ ] **Step 2：實作 `fetch_nongaap_statements`（公開 API）**

```python
def fetch_nongaap_statements(
    ticker: str,
    identity: str,
    ai_config: dict,
    output_dir: Path,
    progress_cb=None,
) -> list[StatementTable]:
    """Fetch Non-GAAP statements from 8-K filings for a ticker.

    Args:
        ticker:      Stock ticker, e.g. "AAPL"
        identity:    SEC EDGAR identity string
        ai_config:   {"provider": ..., "model": ..., "api_key": ...}
        output_dir:  Directory where nongaap_cache.json will be stored
        progress_cb: Optional callable(current, total, label) for progress updates

    Returns:
        List of StatementTable: [Data_EPS_Recon, Data_NonGAAP] (omits None tables)
    """
    set_identity(identity)
    company = Company(ticker)
    cache_path = Path(output_dir) / CACHE_FILENAME

    cache = _load_cache(cache_path)
    filings = _get_earnings_filings(company)

    new_filings = [(lbl, f) for lbl, f in filings if lbl not in cache]
    total = len(new_filings)

    for i, (quarter_label, filing) in enumerate(new_filings, 1):
        if progress_cb:
            progress_cb(i, total, f"Non-GAAP {ticker} {quarter_label} ({i}/{total})")

        try:
            eight_k = filing.obj()
            eps_recon = _extract_eps_recon(eight_k)
            metrics   = _extract_nongaap_metrics(eight_k, ai_config)
            cache[quarter_label] = {
                "filing_date": str(filing.filing_date),
                "eps_recon":   eps_recon,
                "metrics":     metrics,
            }
            # Save after each quarter (crash-safe incremental)
            _save_cache(cache_path, cache)
        except Exception as exc:
            print(f"[fetcher_nongaap] {quarter_label} failed: {exc!r}", file=sys.stderr)

    tables: list[StatementTable] = []
    eps_tbl = _build_eps_recon_table(ticker, cache)
    ng_tbl  = _build_nongaap_table(ticker, cache)
    if eps_tbl:
        tables.append(eps_tbl)
    if ng_tbl:
        tables.append(ng_tbl)
    return tables
```

- [ ] **Step 3：確認全部 Non-GAAP 測試仍通過**

```
pytest tests/test_fetcher_nongaap.py -v
```
預期：全部 passed

- [ ] **Step 4：Commit**

```bash
git add fetcher_nongaap.py
git commit -m "feat: 8-K discovery and incremental Non-GAAP fetch orchestration"
```

---

### Task 9：main.py — 串接 Non-GAAP 到執行流程

**Files:**
- Modify: `main.py:737-771`（`_worker_single`）

- [ ] **Step 1：在 `_worker_single` 中加入 Non-GAAP 呼叫**

```python
def _worker_single(self, ticker: str, fetch_gaap: bool, fetch_nongaap: bool, max_filings: int = 80):
    try:
        identity = self.cfg.get("identity", "")
        if not identity:
            self._log("[ERROR] 請先在進階設定填入 Identity（姓名 + 信箱）")
            self._done(False)
            return

        tables = []
        output_path = self._build_output_path(ticker)
        output_dir  = output_path.parent
        output_dir.mkdir(parents=True, exist_ok=True)

        if fetch_gaap:
            self._log(f"[{ticker}] 抓取 GAAP 財報中...")
            self._set_progress(0, 3, "抓取 GAAP...")
            gaap_tables = fetch_gaap_statements(ticker, identity, max_filings=max_filings)
            tables.extend(gaap_tables)
            self._log(f"[{ticker}] GAAP：取得 {len(gaap_tables)} 份財報")

        if fetch_nongaap:
            ai_config = self.cfg.get("ai", {})
            self._log(f"[{ticker}] 抓取 Non-GAAP 財報中...")
            self._set_progress(1, 3, "抓取 Non-GAAP...")

            from fetcher_nongaap import fetch_nongaap_statements

            def _ng_progress(current, total, label):
                self._log(f"[{ticker}] {label}")
                self._set_progress(current, total, label)

            ng_tables = fetch_nongaap_statements(
                ticker, identity, ai_config,
                output_dir=output_dir,
                progress_cb=_ng_progress,
            )
            tables.extend(ng_tables)
            self._log(f"[{ticker}] Non-GAAP：{len(ng_tables)} 張 sheet")

        if not tables:
            self._log("[WARNING] 無資料可寫入")
            self._done(False)
            return

        self._log(f"[{ticker}] 寫入 Excel...")
        self._set_progress(2, 3, "寫入 Excel...")
        write_statements(tables, output_path)
        self._log(f"[{ticker}] 完成 → {output_path.name}")
        self._set_progress(3, 3, "完成！")
        self._done(True)

    except Exception as e:
        self._log(f"[ERROR] {e}")
        self._done(False)
```

- [ ] **Step 2：端對端測試（需要 EDGAR identity + AI API key）**

```
1. 啟動程式
2. 輸入 AAPL，勾選 GAAP + Non-GAAP
3. 按執行
4. 確認 output 資料夾下產生 nongaap_cache.json
5. 確認 Excel 有 Data_EPS_Recon 和 Data_NonGAAP sheet
6. 再跑一次 AAPL → log 應顯示「Non-GAAP: 0 季新增」（全部從快取）
```

- [ ] **Step 3：Commit**

```bash
git add main.py
git commit -m "feat: wire Non-GAAP fetch into single-company worker"
```

---

### Task 10：全套測試 + 文件更新

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `ARCHITECTURE.md`

- [ ] **Step 1：跑全套測試**

```
pytest tests/ -v
```
預期：所有 tests passed

- [ ] **Step 2：更新 CHANGELOG.md**

在「已完成」清單加入：

```markdown
- [x] Per-ticker output path memory（ticker_paths in config.json）
- [x] Non-GAAP fetching from 8-K press releases（Data_EPS_Recon + Data_NonGAAP）
- [x] nongaap_cache.json 增量快取（每季 AI 呼叫結果本機快取）
```

在待辦移除：
```markdown
- [ ] Non-GAAP 抓取（Phase 2）  ← 刪除
```

在更新記錄加入：

```markdown
### 2026-04-16（Session 3）

**Per-Ticker Output Path Memory**
- config.json 新增 ticker_paths 欄位
- 確認公司後自動帶出已記憶路徑
- 瀏覽選資料夾後自動儲存至 ticker_paths

**Non-GAAP Fetching（Phase 2）**
- fetcher_nongaap.py 完整實作
- 8-K Item 2.02 篩選，EPS reconciliation（edgartools 原生）
- AI 從 EX-99.1 press release 提取 Non-GAAP 指標
- nongaap_cache.json 增量快取，只對新季度呼叫 AI
- 輸出：Data_EPS_Recon + Data_NonGAAP sheet
```

- [ ] **Step 3：更新 ARCHITECTURE.md**

在 File Map 表格加入：
```
| nongaap_cache.json | 各公司輸出資料夾內，Non-GAAP 快取（runtime，非 git） |
```

在 Data Flow 更新：
```
fetcher_nongaap.py
    ├─ _get_earnings_filings()  → 8-K Item 2.02 清單
    ├─ _extract_eps_recon()     → edgartools eps_reconciliation
    ├─ _extract_nongaap_metrics() → AI 解析 EX-99.1 press release
    ├─ _build_eps_recon_table() → Data_EPS_Recon
    └─ _build_nongaap_table()   → Data_NonGAAP
```

在 Key Config Variables 加入：
```
| ticker_paths | {TICKER: absolute_path} 各公司輸出資料夾記憶 |
```

- [ ] **Step 4：Commit**

```bash
git add CHANGELOG.md ARCHITECTURE.md
git commit -m "docs: update CHANGELOG and ARCHITECTURE for Phase 2"
```

---

## Self-Review Checklist

- [x] **Spec coverage**
  - Per-ticker path memory → Task 1-3 ✓
  - ticker_paths in config.json → Task 1 ✓
  - UI auto-fill on ticker confirm → Task 3 ✓
  - Save on browse → Task 3 ✓
  - Non-GAAP cache in company folder → Task 8 (`output_dir / CACHE_FILENAME`) ✓
  - Data_EPS_Recon sheet → Task 5 ✓
  - Data_NonGAAP sheet → Task 5 ✓
  - Incremental fetch → Task 8 ✓
  - AI extraction prompt → Task 7 ✓
  - Error handling for missing press release / AI failure → Task 7 ✓
  - ARCHITECTURE.md + CHANGELOG.md → Task 10 ✓

- [x] **Type consistency**
  - `StatementTable` 用法與 fetcher_gaap.py 完全一致
  - `_build_eps_recon_table` / `_build_nongaap_table` 回傳 `StatementTable | None`，Task 8 正確過濾 None

- [x] **No placeholders** — 所有 step 都有完整程式碼
