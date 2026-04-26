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

---

## 地雷十：XOM 等公司的 CF 彙總行末尾有 noncash 項目共用同一 standard_concept

**問題：** XOM 的 `NetCashFromOperatingActivities` 在 edgartools 中出現 4 次：正確的彙總行（"Net cash provided by operating activities"）在 index 26，但末尾還有兩行 ROU lease noncash 調整項目（index 55, 56）也用同樣的 `standard_concept`。`match="last"` 會拿到 index 56（$6M），而非正確的 $12.95B。

**解法：** CF 三大彙總行改加 `label_hint="net cash"`，讓 `_match_is_row` 先縮到 label 含 "net cash" 的候選行再取 last：
```python
("Operating Cash Flow", "NetCashFromOperatingActivities", ..., "CF", "last", "net cash"),
("Investing Cash Flow", "NetCashFromInvestingActivities", ..., "CF", "last", "net cash"),
("Financing Cash Flow", "NetCashFromFinancingActivities", ..., "CF", "last", "net cash"),
```

**已修復（2026-04-18）。**

---

## 地雷十一：部分公司 Capex 以負數回報（現金流出）導致 FCF 計算錯誤

**問題：** US GAAP XBRL 中，Capex（`PaymentsToAcquirePropertyPlantAndEquipment`）可以是正數（大多數公司）或負數（XOM 等）。若 Capex = -5,898，FCF = OCF - capex = OCF + 5,898，結果偏高。

**解法：** FCF 計算改用 `abs(capex)`：
```python
tbl.values[_CF_FCF_IDX][j] = op_cf - abs(capex)
```
兩種符號均正確。**已修復（2026-04-18）。**

---

## 地雷十二：10-Q Q2/Q3 的 CF 欄位是 YTD，`_current_q_col` 會跳過

**問題：** edgartools 對 Q2/Q3 的 10-Q CF 表回傳 YTD（年初至今）欄位，如 `2025-06-30 (YTD)` 和 `2025-09-30 (YTD)`。`_is_q_col` 只認 `Q1`/`FY` 格式，故 `_current_q_col` 對 Q2/Q3 回傳 None，整個 filing 被跳過。

**影響：** TSLA、BA、XOM 等公司每年 CF 資料只有 Q1 有值，Q2/Q3 全為 None。這是美國 GAAP 規定 interim CF 用累計方式呈現的結果。

**正確修法（未實作）：** Q2 quarterly = Q2 YTD − Q1 YTD；Q3 quarterly = Q3 YTD − Q2 YTD。需要跨 filing 減法，屬於較大重構。目前列為已知限制。

---

## 地雷十三：`label_hint` 不符合時不應 fallback 到未過濾候選行

**問題：** 舊 `_match_is_row` 在 `label_hint` 找不到符合行時，仍從候選行取 first/last（等同忽略 hint）。e.g. COHR 的 `ChangeInReceivables` std_concept 被 edgartools 貼在 "Income taxes" 行；若加 `label_hint="receivable"`，理想是跳過並試下一優先級，但舊邏輯會忽略 hint 直接取 "Income taxes"。

**解法：** `_pick()` 當 label_hint 不符時傳回 `None`（不是 fallback 取未過濾行），使呼叫端進入下一優先級（std_concept → fallback_suffix → label_fallback）：
```python
def _pick(rows):
    if rows.empty:
        return None
    if label_hint:
        hinted = rows[rows["label"].astype(str).str.contains(label_hint, case=False, na=False)]
        if hinted.empty:
            return None  # hint 不符 → 跳至下一優先級
        rows = hinted
    return rows.index[-1] if match == "last" else rows.index[0]
```
**已修復（2026-04-23）。**

---

## 地雷十四：CF 三大彙總行的 label_hint 用 `"net cash"` 無法匹配 AAPL

**問題：** AAPL CF 使用 "Cash generated by operating activities"（開頭是 "Cash"），不含 "net cash"。搭配 Session 10 的 cascading 修復後，`label_hint="net cash"` 找不到 → 回傳 None → OCF/ICF/FCF 全空。

**解法：** label_hint 改為正則 `"^net cash|^cash"`（starts-with "net cash" 或 starts-with "cash"）：
- 匹配 AAPL "Cash generated by operating activities" ✅
- 匹配標準 "Net cash provided by operating activities" ✅
- 排除 XOM "Noncash right of use assets..." ✅（starts with "Noncash"）

**注意：** `str.contains` 預設 `regex=True`，`^` 有效。
**已修復（2026-04-23）。**

---

## 地雷十五：部分公司（XOM / COHR）IS 缺少 Gross Profit 或 Operating Income 行

**問題：** 
- XOM（石油整合公司）：IS 結構為 Revenue → Pre-tax Income，沒有獨立的 Gross Profit 或 Operating Income XBRL 行。
- COHR：IS 結構為 Revenue → 各項費用 → Total Costs → EBT，沒有 Operating Income 行。

**影響：** Gross Profit / Operating Income 在這些公司的輸出為 None（非 bug，是原始資料缺失）。

**部分緩解：** Gross Profit 新增 DERIVED fallback（Revenue − COGS），有 COGS 的公司（如 COHR）可衍生出 Gross Profit。Operating Income 目前無 DERIVED 邏輯。

---

## 地雷十六：金融股（JPM）的 CF Capex 全部為 None

**問題：** JPM 的 CF 中，Capex（PaymentsToAcquirePropertyPlantAndEquipment）全為 None。銀行類公司對不動產購置使用不同的 XBRL 概念（如 `PurchasesOfPremisesAndEquipment`），不符合我們模板的 fallback 條件。

**影響：** JPM 的 `Capex` row 在 live snapshot test 中被標記為 structural_absence。GS 沒有此問題（GS 的 Capex 可正常匹配）。

**已知限制，不修：** `test_snapshot_cf` 的金融股測試已將 `"Capex"` 加入 `allowed_missing`，與 IS 的 `"Operating Income"` 同等處理。
