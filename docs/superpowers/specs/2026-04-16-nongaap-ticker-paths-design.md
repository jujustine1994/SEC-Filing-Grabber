# Design Spec: Non-GAAP Fetching + Per-Ticker Output Path Memory

Date: 2026-04-16

---

## 功能概述

兩個獨立但相關的功能，合併在同一次實作：

1. **Per-Ticker Output Path Memory** — 記住每間公司的輸出資料夾，不需每次重選
2. **Non-GAAP Fetching (Phase 2)** — 從 8-K press release 抓 Non-GAAP 指標，寫入新 sheet

---

## 功能一：Per-Ticker Output Path Memory

### 需求

使用者有既有資料夾結構：
```
C:\Users\CTH\Documents\Work\...\美股\
  TSLA\   AAPL\   NVDA\   ...（每間公司一個資料夾）
```

希望 SEC Fetcher 的輸出直接放進對應資料夾，不需每次手動選路徑。

### config.json 新增欄位

```json
{
  "ticker_paths": {
    "TSLA": "C:\\Users\\CTH\\Documents\\Work\\...\\美股\\TSLA",
    "AAPL": "C:\\Users\\CTH\\Documents\\Work\\...\\美股\\AAPL"
  }
}
```

`ticker_paths` 獨立於現有的 `output_dir`——`output_dir` 作為沒有設定 per-ticker 路徑時的預設值。

### UI 行為（Tab 1，單一公司）

1. **自動帶出路徑**：輸入 ticker 並確認公司名稱後，若 `ticker_paths[ticker]` 存在，自動將「儲存位置」欄位更新為該路徑。
2. **手動更改**：使用者可透過「瀏覽」按鈕選新路徑。選完後，路徑**自動存入** `ticker_paths[ticker]`（不需額外按「記住」）。
3. **首次使用 ticker**：顯示預設 `output_dir`，使用者瀏覽後自動記憶。

### _build_output_path 邏輯變更

```
1. 查 config["ticker_paths"].get(ticker) → 有值則用此路徑作為 output_dir
2. 否則用 config["output_dir"]（現有預設）
3. 檔名邏輯不變（ticker_name / ticker_only / custom）
```

---

## 功能二：Non-GAAP Fetching

### 資料來源

- **8-K Exhibit 99.1**（earnings press release）— 每季財報發布時附的 HTML 新聞稿
- 篩選條件：8-K 含 Item 2.02（Results of Operations）

### 輸出 Sheets（新增兩張）

#### `Data_EPS_Recon`
- 來源：edgartools 原生 `eight_k.earnings.eps_reconciliation`（結構化，不需 AI）
- 行：EPS 調和項目（GAAP EPS、SBC adjustment、Tax adjustment、Non-GAAP EPS 等）
- 欄：季度（最舊→最新）
- B 欄存在但值為 None（沿用現有 writer 格式，EPS Recon 沒有 XBRL labels）

#### `Data_NonGAAP`
- 來源：AI 從 press release 全文提取
- 行：AI 抓到的所有 Non-GAAP 指標，跨季取聯集（缺的季填 None）
- 欄：季度（最舊→最新）
- 格式同 `Data_Seg_*`

### 本機快取（增量更新）

路徑：`{ticker_output_dir}/nongaap_cache.json`

```json
{
  "FY2024Q1": {
    "8k_date": "2024-01-24",
    "filing_date": "2024-01-24",
    "eps_recon": {
      "GAAP EPS Diluted": 0.53,
      "SBC": 0.12,
      "Non-GAAP EPS Diluted": 0.65
    },
    "metrics": {
      "Non-GAAP Net Income": 2513000000,
      "Non-GAAP Gross Margin %": 17.6,
      "Adjusted EBITDA": 3800000000
    }
  }
}
```

快取存在 ticker 的輸出資料夾內（與 Excel 並排）。若輸出資料夾在專案外（如使用者的 Work 資料夾），不需額外 gitignore；若輸出資料夾在專案內的 `output/` 下，`output/` 已在 gitignore，自動涵蓋。

### 增量抓取邏輯

```
① 讀 nongaap_cache.json → 得知已處理的季度
② 取得全部 8-K filings（含 Item 2.02）→ 轉換為 quarter_label
③ 計算差集：未處理的季度
④ 對每個新季度：
   a. edgartools 取 eps_reconciliation → 存入 eps_recon
   b. 取 EX-99.1 HTML → 轉 markdown → 呼叫 AI
   c. AI 回傳 JSON metrics → 存入 metrics
   d. 即時寫回 nongaap_cache.json（crash-safe）
⑤ 從完整 cache 重建兩個 StatementTable
```

### AI 提取 Prompt

```
你是財務分析師。以下是一份公司季度財報新聞稿（Markdown 格式）。
請提取所有 Non-GAAP 財務指標，回傳 JSON 格式：

{
  "指標名稱": 數值（純數字，不含單位）,
  ...
}

規則：
- 只取 Non-GAAP / Adjusted / Excluding 相關的指標
- 數值單位若為百萬則乘以 1,000,000，億則乘以 1,000,000,000
- 百分比直接回傳小數（如 17.6%  → 17.6）
- 若找不到任何 Non-GAAP 指標，回傳空 JSON {}

新聞稿內容：
{press_release_markdown}
```

### 錯誤處理

| 情境 | 處理方式 |
|------|---------|
| 8-K 無 Item 2.02 | 跳過，不計入快取 |
| EX-99.1 不存在 | 跳過，記錄警告 |
| AI 回傳非 JSON | 記錄警告，該季 metrics 存 `{}` |
| eps_reconciliation 為 None | eps_recon 存 `{}` |
| AI API 失敗 | 停止該 ticker，保留已處理的快取 |

---

## 檔案變動摘要

| 檔案 | 變動 |
|------|------|
| `config.py` | `load_config` 補 `ticker_paths: {}` 預設值 |
| `config.json` | 新增 `ticker_paths` 欄位 |
| `fetcher_nongaap.py` | 實作完整 Non-GAAP 抓取邏輯 |
| `main.py` | Tab 1：ticker 確認後自動帶路徑；瀏覽後自動記憶 |
| `main.py` | `_build_output_path`：優先查 `ticker_paths` |
| `main.py` | `_worker_single`：呼叫 `fetch_nongaap_statements` |
| `.gitignore` | 無需修改（輸出在 Work 資料夾外，或在已 gitignore 的 `output/` 內）|
| `ARCHITECTURE.md` | 更新資料流、新增 sheet 說明 |
| `CHANGELOG.md` | 記錄新功能 |

---

## 不在本次範圍

- 批量更新（Tab 2）的 Non-GAAP 支援：留待後續
- Non-GAAP 歷史回填進度條：留待後續
- 金融股模板：獨立票
