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

---

## 地雷五：多家公司的 Net Income 用 `ProfitLoss` 而非 `NetIncome`

**問題：** BA、TSLA、XOM、WMT 在 XBRL 裡 Net Income 的 `standard_concept` 是 `ProfitLoss`（含少數股東損益），不是 `NetIncome`。直接查 `NetIncome` 會得到 None。

**解法：** `_build_is_table` post-processing：
```python
if row_vals.get(_NET_INCOME_IDX) is None:
    idx = _match_is_row(df, "ProfitLoss", "ProfitLoss")
    if idx is not None:
        row_vals[_NET_INCOME_IDX] = _to_python_val(df.loc[idx, q_col])
```

同樣的 fallback 在 CF 的 Net Income 行也需要（_build_cf_table 目前未處理，待補）。

---

## 地雷六：TSLA D&A 的 `standard_concept` 為 nan

**問題：** TSLA 的 CF 「Depreciation, amortization and impairment」行，edgartools 的 `standard_concept` 是 `nan`（未標準化）。用 `DepreciationExpense` 比對失敗，fallback_suffix 也可能比對不到自訂的 concept 名稱。

**解法：** `_match_is_row` 第三層 label fallback：
```python
idx = _match_is_row(cf_df, None, "", label_fallback="depreciation")
```

---

## 地雷七：GOOGL BS 含非 ASCII 字元導致 cp950 編碼錯誤

**問題：** GOOGL 某些 BS label 含有 `\xa0`（non-breaking space）。在 Windows 中文環境（cp950 terminal），`print()` 呼叫嘗試用 cp950 編碼時失敗。

**解法：** 存 label 時先做 NFKC normalize，將 `\xa0` 等相容字元轉為一般 ASCII：
```python
import unicodedata
concept_labels[key] = unicodedata.normalize("NFKC", raw_label)
```

**位置：** `_build_dynamic_table` 和所有存 XBRL label 的地方。

---

## 地雷八：CF 彙總行有多個相同 standard_concept

**問題：** `NetCashFromOperatingActivities` 在部分公司（BA 4次、AMD 3次）會出現多次，對應中間小計和最終合計。取 first 會拿到錯誤的中間值。

**解法：** CF 彙總行（Op/Inv/Fin CF）使用 `match="last"`：
```python
("Operating Cash Flow", "NetCashFromOperatingActivities", "...", "CF", "last", None),
```

同樣適用 `CashAndCashEquivalents`（期初 + 期末，要取 last = 期末）。

---

## 地雷九：openpyxl 寫入空字串後讀回來是 None

**問題：** `ws.cell(value="")` 寫入空字串，`load_workbook` 後讀回來是 `None`，不是 `""`。

**影響：** test 不能用 `== ""` 斷言空的 label cell，要用 `is None`。

**解法：**
```python
assert ws["B5"].value is None   # 空 label
```
