# UI 日期區間 + Sheet 預覽設計

> 狀態：設計確認，待實作
> 日期：2026-04-29

## 背景

使用者希望在主介面加入：
1. 時間區間選擇（起訖年份）
2. 報表類型選擇（季報 / 年報 / 兩者）
3. 執行前可預覽有哪些 sheet，並選擇性跳過

## 確認需求

### 1. 日期區間（起訖年份）

- 預設：留空 = 全部可用年份（現有行為不變）
- 可選：填入起始年 / 截止年 → 只抓該區間的 filings
- UI：兩個 Spinbox，起 / 迄年份，都可留空
- Tab 1（單一公司）和 Tab 2（批量更新）都要有

### 2. 報表類型（季報 / 年報）

- `[v] 季報 (10-Q)`  `[v] 年報 (10-K)`，預設兩者都勾（現有行為不變）
- 只勾季報 → 跳過 `fetch_gaap_statements` 的 10-K 部分（不產生 Data_Financials(Y)）
- 只勾年報 → 跳過 10-Q，速度快很多，適合做長期財務模型
- Tab 1 和 Tab 2 都有此選項
- 至少一個要勾，否則擋下執行

### 3. Sheet 快速掃描 + 選擇性跳過

- 輸入 ticker 後，可點「快速掃描 ▶」按鈕
- 快速掃描：只抓最新一份 10-Q（~5–15 秒），偵測有哪些 Data_Seg_* sheet
- 顯示預覽清單，使用者可取消不需要的 sheet
- **不點掃描直接執行 = 全抓（預設路徑，現有行為不變）**
- 固定 sheet（Data_Financials(Q/Y)、Data_Meta）預設勾選且不可取消

## Sheet 類型說明

| Sheet | 類型 | 說明 |
|-------|------|------|
| Data_Financials(Q) | 固定 | 季報 IS+BS+CF 三合一 |
| Data_Financials(Y) | 固定 | 年報 IS+BS+CF 三合一 |
| Data_Meta | 固定 | 後設資料 |
| Data_Financials_NG(Q/Y) | 條件 | 有 overflow NG rows 才出現 |
| Data_Seg_* | 動態 | 每家公司分部不同，1–10+ 張不等 |
| Data_NonGAAP | Non-GAAP | 8-K press release AI 解析 |
| Data_EPS_Recon | Non-GAAP | EPS reconciliation（多數為空） |

## UI 設計

### Tab 1（單一公司）

```
┌────────────────────────────────────────────┐
│  單一公司  │  批量更新                       │
├────────────────────────────────────────────┤
│  Ticker: [AAPL      ]  Apple Inc.          │
│                                    [快速掃描 ▶] │
│  [v] GAAP 財報  [ ] Non-GAAP（需設定API）  │
│                                            │
│  報表類型：[v] 季報(10-Q)  [v] 年報(10-K)  │
│                                            │
│  日期區間：起 [    ] 迄 [    ] 年           │
│            （留空 = 全部可用年份）          │
│                                            │
│  ── 掃描後才出現 ─────────────────────     │
│  [v] Data_Financials(Q)  [固定，不可取消]  │
│  [v] Data_Financials(Y)  [固定，不可取消]  │
│  [v] Data_Meta           [固定，不可取消]  │
│  [v] Data_Seg_Revenue                      │
│  [v] Data_Seg_Americas                     │
│  [v] Data_Financials_NG(Q)                 │
│  ─────────────────────────────────────     │
│                                            │
│  ▼ 輸出設定                                │
│  ┌──────────────────────────────────────┐  │
│  │  儲存位置：[output          ] [瀏覽] │  │
│  │  檔名格式：○ Ticker+名稱  ○ Ticker   │  │
│  │           ● 自訂：[        ] .xlsx   │  │
│  │  預覽：（請輸入檔名）                │  │
│  └──────────────────────────────────────┘  │
│                                            │
│                 [▶  執行]                  │
└────────────────────────────────────────────┘
```

### Tab 2（批量更新）

```
┌────────────────────────────────────────────┐
│  Watchlist （現有 checkbox 列表）           │
│  [全選] [全不選]                            │
│                                            │
│  [ ] 同時抓取 Non-GAAP                     │
│                                            │
│  報表類型：[v] 季報(10-Q)  [v] 年報(10-K)  │
│  日期區間：起 [    ] 迄 [    ] 年           │
│            （留空 = 全部可用年份）          │
│                                            │
│          [▶  開始批量更新]                  │
└────────────────────────────────────────────┘
```

> 批量模式不做 sheet 預覽（多 ticker 不實際），只提供日期區間。

## 後端影響

### fetch_gaap_statements() 修改

新增參數：
- `start_year: int | None = None`
- `end_year: int | None = None`
- `fetch_quarterly: bool = True`
- `fetch_annual: bool = True`
- `excluded_sheets: set[str] | None = None`

過濾邏輯（在 `list(company.get_filings(...))` 之後）：
```python
if start_year or end_year:
    filings_q = [f for f in filings_q if _in_year_range(f, start_year, end_year)]
    filings_k = [f for f in filings_k if _in_year_range(f, start_year, end_year)]
```

sheet 跳過邏輯：在 `_merge_financials` / `_build_seg_tables` 回傳前檢查 `excluded_sheets`。

### 新增 preview_sheets() 函式

```python
def preview_sheets(ticker: str, identity: str) -> list[str]:
    """快速掃描：只抓最新一份 10-Q，回傳預期 sheet 名稱清單。"""
```

只讀取 `filings_q[0]`，偵測 segment concepts，回傳 sheet 名稱列表。

### fetch_nongaap_statements() 修改

同樣加入 `start_year` / `end_year` 參數。

## 不在此次範圍

- 批量模式的 sheet 預覽：技術上可行但實際效益低
