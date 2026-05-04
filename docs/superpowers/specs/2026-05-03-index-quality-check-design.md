# Index Sheet 品質檢測設計

**目標：** 在 Excel 輸出的 Index sheet 加入報表完成度指標，讓分析師打開檔案就能立即判斷資料品質。

**Tech Stack：** Python 3.11, openpyxl, pytest

---

## 需求摘要

- Index sheet 表格加第 5 欄「完成度」，只有 `Data_Financials(Q)` 顯示實際值，其餘顯示「—」
- 表格下方加「品質明細」區塊，逐行列出 9 個 key rows 的 ✓/✗ 狀態
- 缺失行底色淺橘，方便視覺辨識
- 全部通過時完成度欄綠色「9/9 ✓」；有缺失時橘色「N/9 ⚠」

---

## 檔案異動

| 檔案 | 動作 | 說明 |
|------|------|------|
| `excel_formatter.py` | 修改 | `_build_index_sheet()` 加入完成度欄與品質明細區塊 |
| `tests/test_excel_formatter.py` | 修改（若存在）或新增 | 新增品質欄與明細區塊的 unit tests |

其他檔案（`excel_writer.py`、`fetcher_gaap.py`、`main.py`）不需異動。

---

## 架構設計

### Key rows 來源

沿用 `override_engine.check_key_rows()` 現有邏輯，檢查以下 9 個指標在近 4 季是否全為 None：

| 指標 | 報表 |
|------|------|
| Revenue | IS |
| Operating Income | IS |
| Net Income | IS |
| Diluted EPS | IS |
| Total Assets | BS |
| Total Liabilities | BS |
| Total Equity — Parent | BS |
| Operating Cash Flow | CF |
| Capex | CF |

### 完成度計算

```python
from override_engine import check_key_rows

q_tbl = next((t for t in tables if t.sheet_name == "Data_Financials(Q)"), None)
if q_tbl:
    missing_is = check_key_rows(q_tbl.concepts, q_tbl.values, "IS")
    missing_bs = check_key_rows(q_tbl.concepts, q_tbl.values, "BS")
    missing_cf = check_key_rows(q_tbl.concepts, q_tbl.values, "CF")
    all_missing = missing_is + missing_bs + missing_cf
    total = 9
    score = total - len(all_missing)   # e.g. 7
```

---

## Index Sheet 版面變更

### 現況（4 欄）

```
Row 1:  [公司名稱 header — merged A1:D1]
Row 2:  [抓取日期 metadata — merged A2:D2]
Row 3:  (blank)
Row 4:  Sheet | 說明 | 最早期間 | 最新期間
Row 5+: 各 sheet 資料列
```

### 修改後（5 欄）

```
Row 1:  [公司名稱 header — merged A1:E1]
Row 2:  [抓取日期 metadata — merged A2:E2]
Row 3:  (blank)
Row 4:  Sheet | 說明 | 最早期間 | 最新期間 | 完成度
Row 5+: 各 sheet 資料列（E 欄：Data_Financials(Q) 顯示分數，其餘「—」）
Row N:  (blank separator)
Row N+1: [品質明細 section header]
Row N+2+: 各 key row 狀態列
```

### 完成度欄顯示規則

- 全部通過：`"9/9 ✓"`，字色綠色（`#1A7A34`）
- 有缺失：`"N/9 ⚠"`，字色橘色（`#C25C00`）
- 其他 sheet：`"—"`，字色灰色

### 品質明細區塊

- Section header：`"品質明細 — Data_Financials(Q)"`，深藍底白字，合併 A:B 欄
- 每個 key row 一行：A 欄 = 指標名稱，B 欄 = `"✓"` 或 `"✗  缺失"`
- 缺失行底色：淺橘（`#FFF0E0`），✗ 字色橘紅（`#C00000`）
- 通過行底色：白色
- 若 `Data_Financials(Q)` 不存在（例如只抓年報），不顯示品質明細區塊
- 行順序固定為 IS → BS → CF，共 9 行：

```python
ALL_KEY_ROWS = [
    # IS
    "Revenue", "Operating Income", "Net Income", "Diluted EPS",
    # BS
    "Total Assets", "Total Liabilities", "Total Equity — Parent",
    # CF
    "Operating Cash Flow", "Capex",
]
```

實作時用 `all_missing`（set）查每行是否缺失，而非只迭代缺失清單：

```python
for row_name in ALL_KEY_ROWS:
    is_missing = row_name in all_missing
    # 寫入 ✓ 或 ✗，套對應底色
```

---

## 顏色常數（沿用 excel_formatter.py 現有慣例）

```python
QUALITY_GREEN  = "FF1A7A34"   # 完成度全 OK 字色
QUALITY_ORANGE = "FFC25C00"   # 完成度有缺失字色
QUALITY_MISS_BG = "FFFFF0E0"  # 缺失行底色（淺橘）
QUALITY_MISS_FG = "FFC00000"  # 缺失行字色（橘紅）
```

---

## 測試計畫

| 測試 | 說明 |
|------|------|
| `test_index_quality_col_all_ok` | 9 個 key rows 全有值 → E 欄顯示「9/9 ✓」 |
| `test_index_quality_col_missing` | 2 個 key rows 缺失 → E 欄顯示「7/9 ⚠」 |
| `test_index_quality_col_other_sheets` | 非 Q 表的 sheet 在 E 欄顯示「—」 |
| `test_index_quality_detail_section_present` | 品質明細區塊出現在表格下方 |
| `test_index_quality_detail_missing_row_highlighted` | 缺失行有橘色底色 |
| `test_index_quality_no_q_table` | 無 Data_Financials(Q) 時不顯示明細區塊 |

---

## 不在範圍內

- `Data_Financials(Y)` 的品質檢查（年報 key rows 邏輯相同但使用者只要 Q）
- Non-GAAP sheet 的品質檢查
- UI 端的品質提示（僅 Excel 內顯示）
