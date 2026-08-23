# 模板體檢：102 家公司的逐列覆蓋率（2026-08-23 產出）

**這份是自動產出的基線，不是手寫的。** 資料來源 `output/_spike/`（102 家的
companyfacts JSON 與現行路徑答案卷快取），重跑不用打網路。

**⚠ 答案卷的抓取窗不一致**：AAPL/ADBE/AMD/AVGO/COST/GOOGL/INTC/META/MSFT/
NVDA/TSLA/WMT 這 12 家是全部 filing（44~69 期），其餘都是 `max_filings=16`
（約 21 期）。重建時要沿用同樣的參數，不然逐列覆蓋率沒得比。

公司清單刻意涵蓋大中小型 × 跨產業，**包含金融股（JPM/GS/BAC/SCHW）與 REIT（PLD）**
——它們的報表結構跟製造業差很多，是檢驗模板通不通用最有效的一群。

## 零、這份文件怎麼讀（先看這段，不然數字會誤導）

### 「達標列數」是什麼

一列要「達標」必須**同時**滿足兩個條件：

```
有值的公司數 >= 87 家（85% 的樣本）   這一列在絕大多數公司都抓得到
填滿率中位數 > 90%      抓得到的那些公司，幾乎每一季都有值
```

兩個缺一不可。只滿足前者代表「大家都有、但常常缺季」；只滿足後者代表
「少數公司很完整、多數抓不到」。兩種都不能算穩。

### 不要追求 97/97，那個目標本身是錯的

達標門檻假設「這一列應該人人都有」，但有些列天生就不該。**達不到標不等於有 bug**：

| 列 | 為什麼永遠達不了標 |
|---|---|
| `Preferred Stock` | 多數公司根本沒發特別股 |
| `Minority Interest` / `Noncontrolling Interests` | 沒有非控制權益的公司就是沒有 |
| `Finance Lease Liabilities, LT` | 多數公司只有營業租賃 |
| `R&D Expense` | 零售、餐飲、能源業不在損益表單獨揭露 |
| `Pension & Retirement Oblig.` | 大多數公司沒有確定給付制退休金 |

**真正該當 KPI 的是第六節那兩個數字**：〔真缺口〕該抓到卻沒抓到幾列、
〔假警報〕Index 標紅裡有幾個是誤判。達標列數只是一支粗略的體溫計。

## 一、每家公司的缺漏判斷

`data_quality.assess()` 的四個判斷。缺季那一欄幾乎每家都是 1，要打折看：
答案卷是 `max_filings=16` 抓的，最舊那一年的 Q4 合成材料不足。

| ticker | 期數 | 缺季 | 稀疏欄 | 有洞列 | 矛盾 | 模板不適用 |
|---|---|---|---|---|---|---|
| COST | 67 | 0 | 25 | 26 | 1 |  |
| BAC | 21 | 1 | 21 | 0 | 0 | **是** |
| AXP | 21 | 1 | 18 | 0 | 0 | **是** |
| GS | 21 | 1 | 17 | 0 | 1 | **是** |
| PEP | 27 | 1 | 12 | 1 | 0 |  |
| SCHW | 21 | 1 | 10 | 1 | 0 |  |
| PLD | 21 | 1 | 4 | 2 | 2 |  |
| ADBE | 69 | 1 | 3 | 7 | 0 |  |
| INTC | 69 | 1 | 3 | 10 | 0 |  |
| ONTO | 21 | 2 | 2 | 3 | 4 |  |
| SNOW | 21 | 2 | 2 | 1 | 0 |  |
| AAPL | 69 | 1 | 1 | 3 | 0 |  |
| ABBV | 21 | 1 | 1 | 3 | 0 |  |
| ADI | 21 | 1 | 1 | 4 | 1 |  |
| AMAT | 21 | 1 | 1 | 2 | 0 |  |
| AMD | 65 | 1 | 1 | 11 | 1 |  |
| AMGN | 21 | 1 | 1 | 4 | 0 |  |
| AMT | 21 | 1 | 1 | 6 | 0 |  |
| AMZN | 21 | 1 | 1 | 0 | 3 |  |
| ARLO | 21 | 1 | 1 | 1 | 1 |  |
| AVGO | 33 | 1 | 1 | 6 | 0 |  |
| BA | 21 | 1 | 1 | 3 | 1 |  |
| BLK | 8 | 1 | 1 | 1 | 2 |  |
| BMY | 21 | 1 | 1 | 3 | 0 |  |
| C | 21 | 1 | 1 | 1 | 0 |  |
| CAT | 21 | 1 | 1 | 1 | 0 |  |
| CDNS | 21 | 1 | 1 | 3 | 0 |  |
| CL | 21 | 1 | 1 | 3 | 0 |  |
| COP | 21 | 1 | 1 | 4 | 1 |  |
| CVX | 21 | 1 | 1 | 4 | 2 |  |
| DDOG | 21 | 1 | 1 | 3 | 0 |  |
| DE | 21 | 1 | 1 | 3 | 0 |  |
| DHR | 21 | 1 | 1 | 5 | 1 |  |
| DIS | 21 | 1 | 1 | 4 | 0 |  |
| DUK | 21 | 1 | 1 | 3 | 1 |  |
| EMR | 21 | 1 | 1 | 5 | 0 |  |
| EOG | 21 | 1 | 1 | 1 | 0 |  |
| EQIX | 21 | 1 | 1 | 6 | 1 |  |
| ETN | 21 | 1 | 1 | 6 | 2 |  |
| FORM | 21 | 1 | 1 | 5 | 0 |  |
| GE | 21 | 1 | 1 | 2 | 4 |  |
| GILD | 21 | 1 | 1 | 2 | 1 |  |
| GOOGL | 44 | 1 | 1 | 4 | 2 |  |
| HON | 21 | 1 | 1 | 4 | 0 |  |
| INTU | 21 | 1 | 1 | 3 | 1 |  |
| ISRG | 21 | 1 | 1 | 3 | 3 |  |
| JNJ | 21 | 1 | 1 | 0 | 1 |  |
| JPM | 21 | 1 | 1 | 0 | 2 |  |
| KMB | 21 | 1 | 1 | 4 | 1 |  |
| KO | 21 | 1 | 1 | 0 | 1 |  |
| LLY | 21 | 1 | 1 | 4 | 2 |  |
| LMT | 21 | 1 | 1 | 0 | 1 |  |
| MA | 21 | 1 | 1 | 3 | 1 |  |
| MCD | 21 | 1 | 1 | 8 | 1 |  |
| MDLZ | 21 | 1 | 1 | 0 | 1 |  |
| META | 57 | 1 | 1 | 4 | 1 |  |
| MMM | 21 | 1 | 1 | 0 | 0 |  |
| MRK | 21 | 1 | 1 | 7 | 1 |  |
| MS | 21 | 1 | 1 | 3 | 3 |  |
| MSFT | 67 | 2 | 1 | 4 | 1 |  |
| MU | 21 | 1 | 1 | 3 | 2 |  |
| NEE | 21 | 1 | 1 | 0 | 1 |  |
| NFLX | 21 | 1 | 1 | 2 | 0 |  |
| NOW | 21 | 1 | 1 | 2 | 1 |  |
| NVDA | 68 | 1 | 1 | 7 | 1 |  |
| NXPI | 21 | 1 | 1 | 4 | 0 |  |
| ON | 21 | 1 | 1 | 6 | 0 |  |
| PANW | 21 | 1 | 1 | 4 | 2 |  |
| PFE | 21 | 1 | 1 | 3 | 1 |  |
| PSX | 21 | 1 | 1 | 4 | 0 |  |
| PYPL | 21 | 1 | 1 | 3 | 1 |  |
| QCOM | 21 | 1 | 1 | 4 | 0 |  |
| RTX | 21 | 1 | 1 | 5 | 1 |  |
| SBUX | 21 | 1 | 1 | 5 | 0 |  |
| SLB | 21 | 1 | 1 | 4 | 1 |  |
| SNPS | 21 | 1 | 1 | 3 | 0 |  |
| SO | 21 | 1 | 1 | 1 | 2 |  |
| SWKS | 21 | 1 | 1 | 5 | 1 |  |
| T | 21 | 1 | 1 | 9 | 2 |  |
| TMO | 21 | 1 | 1 | 2 | 1 |  |
| TSLA | 61 | 1 | 1 | 8 | 2 |  |
| TXN | 21 | 1 | 1 | 2 | 0 |  |
| UNH | 21 | 1 | 1 | 1 | 0 |  |
| UNP | 21 | 1 | 1 | 2 | 1 |  |
| UPS | 21 | 1 | 1 | 1 | 1 |  |
| V | 21 | 1 | 1 | 5 | 0 |  |
| WFC | 21 | 1 | 1 | 1 | 0 |  |
| WMT | 68 | 1 | 1 | 9 | 0 |  |
| XOM | 21 | 1 | 1 | 2 | 1 |  |
| COHR | 21 | 1 | 0 | 3 | 0 |  |
| CRM | 21 | 0 | 0 | 4 | 0 |  |
| HD | 21 | 0 | 0 | 5 | 0 |  |
| KLAC | 21 | 1 | 0 | 1 | 0 |  |
| LITE | 21 | 1 | 0 | 2 | 0 |  |
| LOW | 21 | 0 | 0 | 5 | 0 |  |
| LRCX | 21 | 1 | 0 | 3 | 1 |  |
| MDT | 21 | 1 | 0 | 1 | 0 |  |
| MRVL | 21 | 0 | 0 | 4 | 0 |  |
| NKE | 21 | 1 | 0 | 1 | 1 |  |
| ORCL | 21 | 1 | 0 | 1 | 1 |  |
| PG | 21 | 1 | 0 | 1 | 0 |  |
| TGT | 21 | 0 | 0 | 4 | 1 |  |

**觸發「模板不適用」的 3 家：AXP, BAC, GS**——全是金融股。IS/BS/CF 模板是為製造業設計的，銀行／券商的報表結構完全不同（存款、放款、備抵呆帳…），這是 TODO D8 記錄的已知限制，現在有量化證據。

## 二、最常出問題的列

### 中間有洞（同一列有些期有、有些沒有——一定是漏抓）

| 列名 | 幾家中招 |
|---|---|
| Acquisitions | 27 / 102 |
| Debt Repayments | 26 / 102 |
| Debt Proceeds | 18 / 102 |
| Short-term Debt | 16 / 102 |
| Share Repurchases | 15 / 102 |
| Other Working Capital | 13 / 102 |
| Cash Taxes Paid | 11 / 102 |
| Operating Income | 11 / 102 |
| Investment Purchases | 10 / 102 |
| Current Portion of LT Debt | 9 / 102 |
| Investment Proceeds | 9 / 102 |
| Shares Outstanding | 9 / 102 |
| Ending Cash | 8 / 102 |
| Amortization of Intangibles | 7 / 102 |
| Cash Interest Paid | 7 / 102 |

### 零星有值（填滿率 <70%，多半是公司本來就沒這項活動，不是漏抓）

2026-08-23（H3-2）從「中間有洞」拆出來的一類。當時拿 companyfacts 當真值驗 52 家、
2,906 個洞：填滿率 70% 以下的那 1,526 個洞**只有 18% 是真的漏抓**，70% 以上才
到 53%。門檻的完整證據見 `data_quality._SPORADIC_FILL_RATIO`。

| 列名 | 幾家中招 |
|---|---|
| Debt Proceeds | 23 / 102 |
| Acquisitions | 20 / 102 |
| Preferred Stock | 10 / 102 |
| Debt Repayments | 8 / 102 |
| D&A (CF memo) | 7 / 102 |
| Short-term Debt | 6 / 102 |
| Treasury Stock | 6 / 102 |
| Investment Purchases | 5 / 102 |
| Investment Proceeds | 4 / 102 |
| Additional Paid-in Capital | 4 / 102 |
| Intangible Assets, net | 3 / 102 |
| Short-term Investments | 3 / 102 |
| Dividends Paid | 3 / 102 |
| Deferred Tax Liability, LT | 2 / 102 |
| Current Portion of LT Debt | 2 / 102 |

### 被判矛盾（整列空白，但同一家公司的相關欄位顯示應該要有）

| 列名 | 幾家中招 |
|---|---|
| Op. Lease Liabilities, current | 22 / 102 |
| Change in Inventories | 17 / 102 |
| Minority Interest | 8 / 102 |
| Debt Proceeds | 8 / 102 |
| Current Portion of LT Debt | 6 / 102 |
| Share Repurchases | 6 / 102 |
| Op. Lease Liabilities, LT | 5 / 102 |
| Debt Repayments | 5 / 102 |
| Noncontrolling Interests | 2 / 102 |

**中招家數多 ≠ concept 對照錯。** 仍在榜上的 `Op. Lease Liabilities, current`
等列，實測多數是**公司沒有在報表表面單獨列出**（金額併在「其他流動負債」裡，
只在附註拆開），現行逐份解 filing 的路徑結構上拿不到。動 concept 對照之前，
先把那份 filing 的報表 dataframe 印出來確認這一列到底在不在。

## 三、逐列覆蓋率：現行路徑 vs companyfacts

「有值公司數」＝ 102 家裡有幾家這一列拿得到資料。兩邊差 8 家以上的標 ⚠。

| 表 | 列名 | 現行 | facts | 差 |
|---|---|---|---|---|
| IS | Revenue | 102 | 101 | -1 |
| IS | Cost of Revenue | 79 | 79 | +0 |
| IS | Gross Profit | 79 | 59 | -20 ⚠ |
| IS | R&D Expense | 60 | 56 | -4 |
| IS | SG&A Expense | 90 | 87 | -3 |
| IS | D&A (CF memo) | 102 | 101 | -1 |
| IS | Other Operating Expense | 27 | 34 | +7 |
| IS | Total Operating Expense | 37 | 43 | +6 |
| IS | Total Costs and Expenses | 35 | 42 | +7 |
| IS | Operating Income | 94 | 86 | -8 ⚠ |
| IS | Interest Expense | 89 | 90 | +1 |
| IS | Interest Income | 31 | 33 | +2 |
| IS | Other Non-op Inc/(Exp) | 62 | 77 | +15 ⚠ |
| IS | Total Non-op Income/(Loss) | 98 | 77 | -21 ⚠ |
| IS | Pre-tax Income | 97 | 99 | +2 |
| IS | Income Tax | 102 | 102 | +0 |
| IS | Net Income | 102 | 102 | +0 |
| IS | Minority Interest | 54 | 63 | +9 ⚠ |
| IS | Net Income incl. NCI | 57 | 73 | +16 ⚠ |
| IS | SBC | 89 | 91 | +2 |
| IS | Basic EPS | 102 | 101 | -1 |
| IS | Diluted EPS | 102 | 101 | -1 |
| IS | Basic Shares | 86 | 101 | +15 ⚠ |
| IS | Diluted Shares | 87 | 101 | +14 ⚠ |
| BS | Cash | 92 | 102 | +10 ⚠ |
| BS | Short-term Investments | 64 | 85 | +21 ⚠ |
| BS | Accounts Receivable | 92 | 86 | -6 |
| BS | Inventories | 73 | 76 | +3 |
| BS | Other Current Assets | 100 | 90 | -10 ⚠ |
| BS | Total Current Assets | 91 | 91 | +0 |
| BS | PP&E, net | 98 | 102 | +4 |
| BS | Operating Lease ROU Assets | 43 | 102 | +59 ⚠ |
| BS | Long-term Investments | 38 | 43 | +5 |
| BS | Goodwill | 92 | 98 | +6 |
| BS | Intangible Assets, net | 82 | 94 | +12 ⚠ |
| BS | Deferred Tax Assets | 51 | 59 | +8 ⚠ |
| BS | Other Non-current Assets | 89 | 91 | +2 |
| BS | Total Non-current Assets | 91 | 11 | -80 ⚠ |
| BS | Total Assets | 102 | 102 | +0 |
| BS | Accounts Payable | 97 | 89 | -8 ⚠ |
| BS | Short-term Debt | 82 | 87 | +5 |
| BS | Current Portion of LT Debt | 49 | 78 | +29 ⚠ |
| BS | Op. Lease Liabilities, current | 23 | 85 | +62 ⚠ |
| BS | Accrued Compensation | 33 | 76 | +43 ⚠ |
| BS | Deferred Revenue, current | 44 | 55 | +11 ⚠ |
| BS | Income Tax Payable | 34 | 59 | +25 ⚠ |
| BS | Other Current Liabilities | 76 | 73 | -3 |
| BS | Total Current Liabilities | 97 | 91 | -6 |
| BS | Long-term Debt | 99 | 97 | -2 |
| BS | Op. Lease Liabilities, LT | 39 | 86 | +47 ⚠ |
| BS | Finance Lease Liabilities, LT | 2 | 41 | +39 ⚠ |
| BS | Deferred Revenue, LT | 15 | 34 | +19 ⚠ |
| BS | Deferred Tax Liability, LT | 70 | 82 | +12 ⚠ |
| BS | Pension & Retirement Oblig. | 22 | 27 | +5 |
| BS | Other Non-current Liabilities | 89 | 86 | -3 |
| BS | Total Non-current Liabilities | 97 | 17 | -80 ⚠ |
| BS | Total Liabilities | 102 | 102 | +0 |
| BS | Preferred Stock | 59 | 89 | +30 ⚠ |
| BS | Common Stock & APIC | 98 | 95 | -3 |
| BS | Additional Paid-in Capital | 83 | 86 | +3 |
| BS | Treasury Stock | 58 | 63 | +5 |
| BS | Retained Earnings | 101 | 101 | +0 |
| BS | AOCI | 102 | 102 | +0 |
| BS | Total Equity — Parent | 102 | 99 | -3 |
| BS | Noncontrolling Interests | 60 | 67 | +7 |
| BS | Total Equity incl. NCI | 63 | 76 | +13 ⚠ |
| BS | Total Liabilities & Equity | 102 | 102 | +0 |
| BS | Shares Outstanding | 97 | 99 | +2 |
| CF | Net Income | 102 | 102 | +0 |
| CF | D&A | 99 | 101 | +2 |
| CF | SBC | 89 | 91 | +2 |
| CF | Amortization of Intangibles | 25 | 72 | +47 ⚠ |
| CF | Change in Receivables | 74 | 76 | +2 |
| CF | Change in Inventories | 57 | 63 | +6 |
| CF | Change in Accounts Payable | 80 | 80 | +0 |
| CF | Change in Prepaid & Other Assets | 27 | 33 | +6 |
| CF | Change in Other Operating Assets | 36 | 51 | +15 ⚠ |
| CF | Change in Deferred Revenue | 35 | 42 | +7 |
| CF | Other Working Capital | 61 | 45 | -16 ⚠ |
| CF | Other Non-cash Items | 64 | 63 | -1 |
| CF | Operating Cash Flow | 102 | 102 | +0 |
| CF | Capex | 92 | 86 | -6 |
| CF | Acquisitions | 83 | 89 | +6 |
| CF | Investment Purchases | 76 | 77 | +1 |
| CF | Investment Proceeds | 58 | 65 | +7 |
| CF | Investing Cash Flow | 102 | 102 | +0 |
| CF | Debt Proceeds | 92 | 72 | -20 ⚠ |
| CF | Debt Repayments | 95 | 69 | -26 ⚠ |
| CF | Share Repurchases | 92 | 98 | +6 |
| CF | Dividends Paid | 86 | 87 | +1 |
| CF | Financing Cash Flow | 102 | 102 | +0 |
| CF | FX Effect on Cash | 82 | 85 | +3 |
| CF | Net Change in Cash | 102 | 102 | +0 |
| CF | Ending Cash | 98 | 102 | +4 |
| CF | Cash Taxes Paid | 39 | 46 | +7 |
| CF | Cash Interest Paid | 36 | 48 | +12 ⚠ |
| CF | Free Cash Flow | 92 | 0 | -92 ⚠ |

**現行路徑達到「>=87 家（85%）有值且填滿率 >90%」的列：46 / 97**
（這個數字不該以 97/97 為目標，理由見第零節。）

## 四、哪些數字是直接讀 XBRL、哪些是推理出來的

**不是每一格都是從財報直接讀出來的。** 下面這些是程式算出來的，來源在
`fetcher_gaap.py` 的後處理段落。看數字有疑問時先確認它屬於哪一類。

### A. 整列都是算的

| 列 | 算式 |
|---|---|
| `Free Cash Flow` | 營運現金流 − 資本支出取絕對值。**XBRL 沒有這個 tag**，本來就只能算 |

### B. 抓不到才用算的（抓得到就用公司報的）

| 列 | 算式 | 什麼情況會用到 |
|---|---|---|
| `Gross Profit` | 營收 − 銷貨成本 | GOOGL／AMZN 等損益表沒有毛利小計行的公司 |
| `Total Non-current Assets` | 總資產 − 流動資產 | 多數公司不標 `AssetsNoncurrent` |
| `Total Non-current Liabilities` | 總負債 − 流動負債 | 多數公司不標 `LiabilitiesNoncurrent` |
| `Total Non-op Income/(Loss)` | 稅前淨利 − 營業利益 | 沒有營業外損益合計行的公司 |

### C. 多列加總（不是挑一條）

| 列 | 加總範圍 |
|---|---|
| `Debt Proceeds` | 所有借款流入（長期、短期、商業本票、可轉債…），排除淨額列 |
| `Debt Repayments` | 所有還款流出，排除淨額列 |
| `Investment Proceeds` | 所有投資處分／到期流入 |

### D. 期間換算（影響範圍最大，最容易被忽略）

| 什麼 | 怎麼算 |
|---|---|
| **現金流量表的每一個單季值** | 公司多半只 tag 年初至今累計 → 本季 YTD − 上季 YTD |
| **每一個 Q4 欄** | 10-Q 只有 Q1~Q3，Q4 由年報 − Q1 − Q2 − Q3 合成（餘額列直接取年報值） |

本次 102 家共 2620 個期間欄，其中 **629 欄是 Q4**（24%）——
這些欄的流量列全部是合成的。Q1~Q3 不齊全時合成會失敗，那一整欄會空掉，
`data_quality` 的「整欄稀疏」就是用來抓這件事的。

## 五、XBRL 裡到底有沒有模板要的數字

把「97 個模板列 × 102 家公司」每一格分成三類。**判斷「有沒有」靠
companyfacts**（它讀得到公司 tag 過的全部 fact，含附註層），比只看報表表面準。

| 分類 | 格數 | 佔比 | 意思 |
|---|---|---|---|
| 我們抓到了 | 7145 | 73% | 正常 |
| **真缺口** | 879 | 9% | 公司有 tag，我們沒抓到 → 見下面 KPI 1 |
| 公司真的沒有 | 1768 | 18% | **不是問題**，這家公司就是沒報這個科目 |

另有 102 格不列入分類：那些模板列在 `facts_mapping` 裡沒有對應
concept（例如 `Free Cash Flow`，XBRL 本來就沒有這個 tag），無從判斷「有沒有」。

**所以答案是：不是每一格都存在。** 「公司真的沒有」那一類佔了相當比例，而且
**那是正常的**——沒發特別股、沒有非控制權益、不揭露 R&D 的公司本來就不該有值。
值得追的只有中間那一類。

## 六、兩個真正的 KPI

### KPI 1 — 真缺口：該抓到卻沒抓到

判準：**這家公司確實 tag 過**（companyfacts 讀得到），我們卻整列空白。
兩邊都沒有的不算——那是公司真的沒報，不是我們的問題。

| 列名 | 幾家真缺 | 哪幾家 |
|---|---|---|
| Op. Lease Liabilities, current | 63 / 102 | AAPL, ABBV, ADI, AMAT, AMD, AMGN, AMZN, ARLO … |
| Operating Lease ROU Assets | 59 / 102 | AAPL, ABBV, ADI, AMAT, AMGN, AVGO, AXP, BA … |
| Amortization of Intangibles | 48 / 102 | ADBE, AMAT, AMT, ARLO, BA, BAC, BMY, CAT … |
| Op. Lease Liabilities, LT | 47 / 102 | AAPL, ABBV, ADI, AMAT, AMGN, AMZN, AVGO, BA … |
| Accrued Compensation | 45 / 102 | AAPL, ABBV, ADBE, ADI, AMAT, AMD, AMGN, AMT … |
| Finance Lease Liabilities, LT | 39 / 102 | AAPL, ABBV, AMAT, AMT, AMZN, ARLO, AVGO, COST … |
| Current Portion of LT Debt | 31 / 102 | AMAT, AMZN, BMY, CDNS, CRM, CVX, DHR, EMR … |
| Income Tax Payable | 30 / 102 | AAPL, AMAT, AMD, AMGN, AMT, BA, BMY, CDNS … |
| Preferred Stock | 30 / 102 | AAPL, ABBV, AMT, BAC, CL, COP, DE, EMR … |
| Short-term Investments | 21 / 102 | AMT, ARLO, C, CDNS, CL, DE, DHR, DUK … |
| Deferred Revenue, LT | 19 / 102 | AMT, AMZN, AVGO, CAT, COHR, COP, DHR, DIS … |
| Deferred Revenue, current | 17 / 102 | AVGO, AXP, COHR, COP, DHR, EMR, EQIX, ETN … |
| Deferred Tax Liability, LT | 16 / 102 | AAPL, AMZN, AVGO, CAT, CDNS, CRM, DDOG, DHR … |
| Net Income incl. NCI | 16 / 102 | ADI, BAC, CRM, DDOG, DHR, FORM, GS, JNJ … |
| Basic Shares | 16 / 102 | BA, BMY, CL, COHR, GE, GOOGL, HON, KMB … |
| Other Non-op Inc/(Exp) | 15 / 102 | AAPL, CVX, DIS, EMR, GILD, GOOGL, HD, LLY … |
| Change in Other Operating Assets | 15 / 102 | AMT, COHR, CRM, DUK, EMR, HD, HON, MA … |
| Diluted Shares | 15 / 102 | BA, BMY, CL, COHR, GE, GOOGL, HON, KMB … |
| Deferred Tax Assets | 13 / 102 | AAPL, ARLO, AVGO, AXP, CRM, CVX, EQIX, INTC … |
| Long-term Investments | 13 / 102 | ADBE, CDNS, COST, DUK, EQIX, GILD, KLAC, MDT … |

**真缺口總計：879 個（列 × 公司）組合，分布在 70 個模板列。**

榜首那幾列全部是 TODO D10（只寫在附註、沒印在報表表面）——這是**已知的暫時性
限制**，不是新 bug。要壓低這個數字只有兩條路：接一條讀附註的路徑，或接受它。

### KPI 2 — 假警報：Index 標紅裡有幾個是誤判

標紅只有兩類：〔矛盾〕整列空白但相關欄位顯示該有、〔中間有洞〕。
「零星有值」刻意不標紅（H3-2），所以不算在內。

| | 家次 |
|---|---|
| 標紅：矛盾 | 79 |
| 標紅：中間有洞 | 352 |
| **標紅合計** | **431** |
| 降級為零星有值（不標紅） | 143 |

**要壓低的是標紅合計裡的誤判比例**，不是把標紅壓到 0——真缺口該標就要標。
驗證方式：對標紅的列抽樣，走 ARCHITECTURE「三步排查順序」確認是哪一類。

## 七、怎麼重跑

```
venv/Scripts/python.exe scripts/spike_derive_mapping.py    # 需要答案卷，慢
venv/Scripts/python.exe scripts/spike_verify_mapping.py    # 用快取，幾秒
```