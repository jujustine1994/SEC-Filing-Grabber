# Known Pitfalls

遇到問題時記錄在這裡。

---

## 地雷一：edgartools v5.29 DataFrame 結構與舊文件不符

**問題：** 網路上的 edgartools 範例（包括舊版文件）假設 `stmt.to_dataframe()` 的 index 是概念名稱（concept names）。實際 v5.29 回傳的是 RangeIndex，concept/label 都是普通欄位（columns），不是 index。

**原因：** edgartools 改版後改了 API，文件沒有跟上。

**解法：**
```python
META_COLS = {'concept', 'label', 'standard_concept', 'level', 'abstract', ...}
df = stmt.to_dataframe()
period_cols = [c for c in df.columns if c not in META_COLS]
concepts = df['label'].fillna(df.get('concept', '')).tolist()
```

**禁止：** 不要用 `df.index` 取概念名稱，也不要假設 DataFrame 有特定的 named index。

---

## 地雷二：edgartools 期間欄位格式為 `"2023-03-31 (Q1)"`，不是 `"FY2023Q1"`

**問題：** `stmt.to_dataframe()` 的期間欄位名稱格式是 `"2023-03-31 (Q1)"` 或 `"2024-12-31 (FY)"`，直接用做 Excel 標頭會顯示原始字串。

**解法：** 用 regex 轉換：
```python
import re
def _col_to_quarter_label(col_name: str) -> str:
    m = re.match(r'(\d{4})-\d{2}-\d{2}\s+\((\w+)\)', col_name)
    if not m:
        return col_name
    year, period = m.group(1), m.group(2)
    return f"FY{year}" if period == "FY" else f"FY{year}{period}"
```

---

## 地雷三：Tkinter BooleanVar 不能在 background thread 呼叫

**問題：** `self.fetch_gaap_var.get()` 在 `threading.Thread` 裡執行時，違反 Tkinter 的 thread safety 規範（所有 widget 操作必須在主執行緒）。在 Windows 通常不即時崩潰，但屬於未定義行為。

**解法：** 在主執行緒用 `.get()` 讀出 bool 值，透過參數傳入 worker：
```python
def _run_single(self):
    fetch_gaap    = self.fetch_gaap_var.get()  # 主執行緒讀
    fetch_nongaap = self.fetch_nongaap_var.get()
    self._start_worker(lambda: self._worker_single(ticker, fetch_gaap, fetch_nongaap))
```

**禁止：** 不要在 daemon thread 裡呼叫任何 `tk.Variable.get()` 或 widget 操作。

---

## 地雷四：company_cache.json 損毀會讓 Watchlist popup 無法開啟

**問題：** `_wl_cache_status()` 在 popup 開啟時同步執行（主執行緒），若 `company_cache.json` 是無效 JSON，`json.load()` 拋 `JSONDecodeError`，整個 popup 無法開啟。

**解法：** 所有讀取 JSON 檔案的地方都要 try/except：
```python
try:
    with open(CACHE_PATH, encoding="utf-8") as f:
        data = json.load(f)
except (json.JSONDecodeError, OSError):
    return "名稱庫：檔案損毀"
```

**適用範圍：** `_wl_cache_status`、`_wl_lookup_worker` 的 cache 讀取都需要保護。
