# Auto-Repair System for Missing Financial Data

**Date:** 2026-04-23  
**Status:** Design complete — ready for implementation  
**Goal:** 當新 ticker 的關鍵財務指標為 None 時，系統自動診斷原因並修復，不需人工干預。診斷結果永久儲存，未來同一公司不重複診斷。

---

## 背景

現行 `_match_is_row()` 採三層查找（std_concept → fallback_suffix → label_fallback），對已測試的 4 家公司（AAPL, TSLA, BA, XOM）運作正常。但新 ticker 可能因以下原因導致欄位為 None：

1. **概念對應錯誤**（E1）：EDGAR 用不同的 `standard_concept` 標記同一指標
2. **概念不存在**（E2）：公司真的沒有該行，或需要從其他行衍生

---

## 設計原則

- **全自動**：fetch 完成後自動檢查 key rows，有缺立即診斷，不用人工
- **AI 最小化**：E1（rule-based fuzzy match）先跑，E2（LLM）只在 E1 失敗時才呼叫
- **AI 只跑一次**：只對最新一期 filing 診斷，不對 80 期每期都呼叫 API
- **永久記憶**：診斷結果寫入全域 override 檔，下次同 ticker 直接套用不重診斷
- **fetch 時檢查**：不等使用者打開 Excel 才發現問題

---

## Key Rows（診斷範圍）

只檢查以下 ~9 個關鍵指標（避免對 89 個 template row 全部診斷產生大量假陽性）：

| 報表 | Std Name | 說明 |
|------|----------|------|
| IS | Revenue | 總營收，幾乎所有公司都有 |
| IS | Operating Income | 營業利益 |
| IS | Net Income | 淨利 |
| IS | EPS Diluted | 稀釋每股盈餘 |
| BS | Total Assets | 總資產 |
| BS | Total Liabilities | 總負債 |
| BS | Total Equity | 股東權益 |
| CF | Operating Cash Flow | 營運現金流 |
| CF | Capital Expenditures | 資本支出 |

FCF 為衍生值，不直接診斷（若 OCF 和 Capex 正確，FCF 自然正確）。

---

## Override 資料模型

### 儲存位置

```
%APPDATA%\SEC_Financial_Tools\ticker_overrides.json
```

全域（跨 output_dir），一個 ticker 診斷過一次就永久記錄。

### 結構

```json
{
  "TICKER": {
    "IS": {
      "Operating Income": {
        "fix_type": "concept_override",
        "std_concept": "OperatingIncomeLoss",
        "diagnosed_at": "2026-04-23",
        "source": "E1"
      },
      "Revenue": {
        "fix_type": "structural_absence",
        "confirmed_absent": true,
        "diagnosed_at": "2026-04-23",
        "source": "E2"
      }
    },
    "CF": {
      "Capital Expenditures": {
        "fix_type": "derived",
        "formula": "PaymentsToAcquirePropertyPlantAndEquipment",
        "diagnosed_at": "2026-04-23",
        "source": "E1"
      }
    }
  }
}
```

### fix_type 說明

| fix_type | 意義 | 行動 |
|----------|------|------|
| `concept_override` | 用另一個 std_concept 取代原本查法 | 改用 override 的 std_concept 查 |
| `derived` | 從其他欄位計算 | 套用指定公式 |
| `structural_absence` | 公司本身就沒有這個指標 | 填 None，不再嘗試，不視為錯誤 |

### 區分「AI 未跑」vs「AI 確認缺失」

`structural_absence` 必須有 `"confirmed_absent": true`，這個欄位只在 AI（E2）明確確認後才寫入。  
若 AI API key 未設定 → 只會做 E1，E2 跳過，`structural_absence` **不會**被寫入，下次 fetch 還會再試。

---

## 診斷流程（Override Engine Pipeline）

```
fetch_ticker(ticker)
    ↓
所有 filing 跑完 → 得到 values[row][quarter]
    ↓
check_key_rows(ticker, statement, values, quarter_labels)
    找出「最近 4 期全為 None」的 key rows
    ↓（若有缺失）
load_overrides(ticker)  ← 已有 override 就直接套用，不重診斷
    ↓（無 override 的缺失 rows）
E1: fuzzy_match(df_latest_filing, missing_row)
    在最新一期的 EDGAR DataFrame 中，對所有 standard_concept 做 fuzzy 比對
    命中 → 寫 concept_override override，重跑該 row
    ↓（E1 未命中）
E2: llm_diagnose(df_latest_filing, missing_row, ticker)
    把 EDGAR DataFrame 的 concept/label 清單發給 LLM
    LLM 回傳：找到（提供 std_concept）/ 確認缺失（structural_absence）
    寫入 override
    ↓
重跑缺失 rows（套用新 override）
    ↓
save_overrides(ticker, new_overrides)
    ↓
繼續 excel_writer
```

### 關鍵設計決策：Override 在 filing loop 開頭套用

**問題**：若在所有 filing 跑完後才診斷，原始 DataFrame 已不在記憶體（需重新 fetch 80 期）。  
**解法**：override 在**每個 filing 的 row_vals 計算前**套用：

```python
# fetcher_gaap.py - _build_is_table() 內部 filing loop
for filing in filings:
    df = filing.get_dataframe(...)
    
    # ① 先套用既有 overrides（若有）
    active_overrides = overrides.get(ticker, {}).get("IS", {})
    
    for row_idx, (label, std_concept, ...) in enumerate(IS_TEMPLATE):
        if label in active_overrides:
            ov = active_overrides[label]
            if ov["fix_type"] == "concept_override":
                # 用 override 的 std_concept 取代 template 的
                val = _get_by_concept(df, ov["std_concept"], ...)
            elif ov["fix_type"] == "derived":
                val = _compute_derived(df, ov["formula"])
            elif ov["fix_type"] == "structural_absence":
                val = None  # 不嘗試，直接 None
        else:
            val = _match_is_row(df, std_concept, ...)  # 正常查找
        row_vals[row_idx] = val
```

### AI 只跑一次（最新一期）

```python
# 診斷時，只對最新一期的 df 跑 E1/E2
latest_df = _get_latest_filing_df(ticker, statement)
missing_rows = [r for r in KEY_ROWS if all_quarters_none(r)]
for row in missing_rows:
    result = diagnose_row(latest_df, row, ticker)  # E1 → E2
    write_override(ticker, statement, row, result)
```

---

## E1：Rule-Based Fuzzy Match

在 `latest_df` 中遍歷所有 rows，對每個 row 的 `standard_concept` 和 `label` 做模糊比對：

```python
def e1_fuzzy_match(df, target_std_name: str) -> str | None:
    """
    target_std_name: e.g. "Operating Income"
    回傳找到的 standard_concept，或 None
    """
    synonyms = SYNONYM_MAP.get(target_std_name, [])
    for _, row in df.iterrows():
        sc = str(row.get("standard_concept", ""))
        lb = str(row.get("label", ""))
        for syn in synonyms:
            if syn.lower() in sc.lower() or syn.lower() in lb.lower():
                return sc  # 命中
    return None
```

`SYNONYM_MAP` 預定義每個 key row 的備選關鍵字，例如：

```python
SYNONYM_MAP = {
    "Operating Income": ["OperatingIncome", "IncomeLossFromOperations", "OperatingProfit"],
    "Revenue": ["Revenues", "SalesRevenue", "NetRevenue", "TotalRevenues"],
    "Net Income": ["NetIncome", "ProfitLoss", "NetEarnings"],
    ...
}
```

---

## E2：LLM Diagnosis

### 觸發條件
- E1 未命中
- AI API key 已設定（`config.json` 的 `ai.api_key` 非空）

### Prompt 設計

```
你是財務資料工程師。以下是公司 {TICKER} 的 {STATEMENT} EDGAR XBRL 概念清單：

{concept_list}  ← standard_concept + label，每行一個

我要找的指標是："{TARGET_ROW}"（例如：Operating Income / 營業利益）

請問：
1. 清單中哪個 standard_concept 最符合這個指標？直接給 standard_concept 字串。
2. 如果清單中根本沒有，回答 "ABSENT"。

只回答一行：standard_concept 或 ABSENT。
```

### 解析與寫入

```python
response = llm_call(prompt)
if response.strip() == "ABSENT":
    write_override(ticker, stmt, row, {"fix_type": "structural_absence", "confirmed_absent": True, "source": "E2"})
else:
    write_override(ticker, stmt, row, {"fix_type": "concept_override", "std_concept": response.strip(), "source": "E2"})
```

---

## UI 整合

### AI API 使用警告

在 **Settings / 進階設定** 區塊加入說明文字：

```
AI API 用途：
・Non-GAAP 數據提取（從 8-K press release 解析）
・財務指標診斷（新 ticker 首次 fetch 時，若關鍵欄位為空）

未設定 API Key 的影響：
・Non-GAAP 功能完全停用
・財務指標診斷僅執行 rule-based 修復（E1），無法確認真正缺失的指標
```

### Fetch 完成後提示

若本次 fetch 觸發了診斷（有 None key rows），在 GUI log 區顯示：

```
[AAPL] 診斷完成：找到 2 項缺失指標的修復方案，已儲存至 override。
[COHR] Operating Income 確認為 structural absence（公司 IS 無此項目）。
```

若 AI 未設定而跳過 E2：

```
[NEWCO] 警告：Revenue 未找到對應概念，E2 診斷需要 AI API key（設定 → AI 設定）。
```

---

## 實作計畫

### 新增檔案

- `override_engine.py`：`load_overrides()`, `save_overrides()`, `check_key_rows()`, `e1_fuzzy_match()`, `e2_llm_diagnose()`, `run_diagnosis()`

### 修改現有檔案

| 檔案 | 變更 |
|------|------|
| `fetcher_gaap.py` | `_build_is_table`, `_build_bs_table`, `_build_cf_table` 各加 override 套用邏輯；fetch 結束後呼叫 `run_diagnosis()` |
| `main.py` | Settings 頁加 AI API 用途說明；fetch worker 顯示診斷 log |

### 不修改

- `fetcher_nongaap.py`（Non-GAAP 已有自己的 AI 路徑）
- `excel_writer.py`（只管寫，不管診斷）
- `config.py`（override 存 APPDATA，不存 config.json）

---

## 已知限制

- E2 只對最新一期診斷；若最新 filing 異常但歷史有數據，可能誤判
- LLM 回傳格式不保證乾淨，需加防禦性解析
- 金融股（GS/JPM）的 BS 結構完全不同，key rows 的 structural_absence 比例高，E2 會頻繁回 ABSENT（這是正確行為，不是 bug）
