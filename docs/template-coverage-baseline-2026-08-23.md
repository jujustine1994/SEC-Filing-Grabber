# 模板體檢：52 家公司的逐列覆蓋率（2026-08-23 產出）

**這份是自動產出的基線，不是手寫的。** 資料來源 `output/_spike/`（52 家的
companyfacts JSON 與現行路徑答案卷快取），重跑不用打網路。

公司清單刻意涵蓋大中小型 × 跨產業，**包含金融股（JPM/GS/BAC/SCHW）與 REIT（PLD）**
——它們的報表結構跟製造業差很多，是檢驗模板通不通用最有效的一群。

## 一、每家公司的缺漏判斷

`data_quality.assess()` 的四個判斷。缺季那一欄幾乎每家都是 1，要打折看：
答案卷是 `max_filings=16` 抓的，最舊那一年的 Q4 合成材料不足。

| ticker | 期數 | 缺季 | 稀疏欄 | 有洞列 | 矛盾 | 模板不適用 |
|---|---|---|---|---|---|---|
| COST | 67 | 0 | 25 | 26 | 1 |  |
| BAC | 21 | 1 | 21 | 0 | 0 | **是** |
| GS | 21 | 1 | 17 | 0 | 1 | **是** |
| SCHW | 21 | 1 | 10 | 1 | 0 |  |
| PLD | 21 | 1 | 4 | 2 | 2 |  |
| ADBE | 69 | 1 | 3 | 7 | 0 |  |
| INTC | 69 | 1 | 3 | 10 | 0 |  |
| ONTO | 21 | 2 | 2 | 3 | 4 |  |
| SNOW | 21 | 2 | 2 | 1 | 0 |  |
| AAPL | 69 | 1 | 1 | 3 | 0 |  |
| ADI | 21 | 1 | 1 | 4 | 1 |  |
| AMAT | 21 | 1 | 1 | 2 | 0 |  |
| AMD | 65 | 1 | 1 | 11 | 1 |  |
| AMZN | 21 | 1 | 1 | 0 | 3 |  |
| ARLO | 21 | 1 | 1 | 1 | 1 |  |
| AVGO | 33 | 1 | 1 | 6 | 0 |  |
| CAT | 21 | 1 | 1 | 1 | 0 |  |
| CVX | 21 | 1 | 1 | 4 | 2 |  |
| DDOG | 21 | 1 | 1 | 3 | 0 |  |
| FORM | 21 | 1 | 1 | 5 | 0 |  |
| GE | 21 | 1 | 1 | 2 | 4 |  |
| GOOGL | 44 | 1 | 1 | 4 | 2 |  |
| JNJ | 21 | 1 | 1 | 0 | 1 |  |
| JPM | 21 | 1 | 1 | 0 | 2 |  |
| KO | 21 | 1 | 1 | 0 | 1 |  |
| MCD | 21 | 1 | 1 | 8 | 1 |  |
| META | 57 | 1 | 1 | 4 | 1 |  |
| MSFT | 67 | 2 | 1 | 4 | 1 |  |
| MU | 21 | 1 | 1 | 3 | 2 |  |
| NEE | 21 | 1 | 1 | 0 | 1 |  |
| NOW | 21 | 1 | 1 | 2 | 1 |  |
| NVDA | 68 | 1 | 1 | 7 | 1 |  |
| NXPI | 21 | 1 | 1 | 4 | 0 |  |
| ON | 21 | 1 | 1 | 6 | 0 |  |
| PANW | 21 | 1 | 1 | 4 | 2 |  |
| PFE | 21 | 1 | 1 | 3 | 1 |  |
| QCOM | 21 | 1 | 1 | 4 | 0 |  |
| SWKS | 21 | 1 | 1 | 5 | 1 |  |
| TSLA | 61 | 1 | 1 | 8 | 2 |  |
| TXN | 21 | 1 | 1 | 2 | 0 |  |
| UNH | 21 | 1 | 1 | 1 | 0 |  |
| WMT | 68 | 1 | 1 | 9 | 0 |  |
| XOM | 21 | 1 | 1 | 2 | 1 |  |
| COHR | 21 | 1 | 0 | 3 | 0 |  |
| CRM | 21 | 0 | 0 | 4 | 0 |  |
| KLAC | 21 | 1 | 0 | 1 | 0 |  |
| LITE | 21 | 1 | 0 | 2 | 0 |  |
| LRCX | 21 | 1 | 0 | 3 | 1 |  |
| MRVL | 21 | 0 | 0 | 4 | 0 |  |
| NKE | 21 | 1 | 0 | 1 | 1 |  |
| ORCL | 21 | 1 | 0 | 1 | 1 |  |
| PG | 21 | 1 | 0 | 1 | 0 |  |

**金融股與 REIT 全數觸發「模板不適用」**（稀疏欄佔 90~100%）。這是 TODO D8
記錄的已知限制，現在有量化證據。

## 二、最常出問題的列

### 中間有洞（同一列有些期有、有些沒有——一定是漏抓）

| 列名 | 幾家中招 |
|---|---|
| Acquisitions | 13 / 52 |
| Debt Repayments | 12 / 52 |
| Debt Proceeds | 11 / 52 |
| Short-term Debt | 11 / 52 |
| Other Working Capital | 8 / 52 |
| Share Repurchases | 8 / 52 |
| Current Portion of LT Debt | 6 / 52 |
| Investment Proceeds | 6 / 52 |
| Other Non-cash Items | 5 / 52 |
| Preferred Stock | 5 / 52 |
| Cash Taxes Paid | 5 / 52 |
| Ending Cash | 5 / 52 |
| Investment Purchases | 5 / 52 |
| Shares Outstanding | 5 / 52 |
| Amortization of Intangibles | 4 / 52 |

### 零星有值（填滿率 <70%，多半是公司本來就沒這項活動，不是漏抓）

2026-08-23（H3-2）從「中間有洞」拆出來的一類。拿 companyfacts 當真值驗 52 家、
2,906 個洞：填滿率 70% 以下的那 1,526 個洞**只有 18% 是真的漏抓**，70% 以上才
到 53%。門檻的完整證據見 `data_quality._SPORADIC_FILL_RATIO`。

| 列名 | 幾家中招 |
|---|---|
| Debt Proceeds | 13 / 52 |
| Acquisitions | 12 / 52 |
| Preferred Stock | 7 / 52 |
| Debt Repayments | 4 / 52 |
| Short-term Debt | 4 / 52 |
| D&A (CF memo) | 4 / 52 |
| Investment Purchases | 3 / 52 |
| Intangible Assets, net | 3 / 52 |
| Short-term Investments | 2 / 52 |
| Deferred Tax Liability, LT | 2 / 52 |
| Current Portion of LT Debt | 2 / 52 |
| Treasury Stock | 2 / 52 |
| Additional Paid-in Capital | 2 / 52 |
| Long-term Debt | 2 / 52 |
| Goodwill | 1 / 52 |

### 被判矛盾（整列空白，但同一家公司的相關欄位顯示應該要有）

| 列名 | 幾家中招 |
|---|---|
| Op. Lease Liabilities, current | 14 / 52 |
| Change in Inventories | 10 / 52 |
| Debt Proceeds | 6 / 52 |
| Op. Lease Liabilities, LT | 4 / 52 |
| Debt Repayments | 4 / 52 |
| Current Portion of LT Debt | 3 / 52 |
| Share Repurchases | 1 / 52 |
| Minority Interest | 1 / 52 |
| Noncontrolling Interests | 1 / 52 |

**中招家數多 ≠ concept 對照錯。** 仍在榜上的 `Op. Lease Liabilities, current`
等列，實測多數是**公司沒有在報表表面單獨列出**（金額併在「其他流動負債」裡，
只在附註拆開），現行逐份解 filing 的路徑結構上拿不到。動 concept 對照之前，
先把那份 filing 的報表 dataframe 印出來確認這一列到底在不在。

## 三、逐列覆蓋率：現行路徑 vs companyfacts

「有值公司數」＝ 52 家裡有幾家這一列拿得到資料。兩邊差 ≥8 家的標 ⚠。

| 表 | 列名 | 現行 | facts | 差 |
|---|---|---|---|---|
| IS | Revenue | 52 | 51 | -1 |
| IS | Cost of Revenue | 42 | 45 | +3 |
| IS | Gross Profit | 42 | 35 | -7 |
| IS | R&D Expense | 38 | 35 | -3 |
| IS | SG&A Expense | 48 | 47 | -1 |
| IS | D&A (CF memo) | 52 | 51 | -1 |
| IS | Other Operating Expense | 13 | 17 | +4 |
| IS | Total Operating Expense | 26 | 27 | +1 |
| IS | Total Costs and Expenses | 14 | 17 | +3 |
| IS | Operating Income | 48 | 44 | -4 |
| IS | Interest Expense | 46 | 46 | +0 |
| IS | Interest Income | 15 | 14 | -1 |
| IS | Other Non-op Inc/(Exp) | 35 | 42 | +7 |
| IS | Total Non-op Income/(Loss) | 50 | 42 | -8 ⚠ |
| IS | Pre-tax Income | 49 | 49 | +0 |
| IS | Income Tax | 52 | 52 | +0 |
| IS | Net Income | 52 | 104 | +52 ⚠ |
| IS | Minority Interest | 22 | 26 | +4 |
| IS | Net Income incl. NCI | 22 | 33 | +11 ⚠ |
| IS | SBC | 45 | 94 | +49 ⚠ |
| IS | Basic EPS | 52 | 52 | +0 |
| IS | Diluted EPS | 52 | 52 | +0 |
| IS | Basic Shares | 46 | 52 | +6 |
| IS | Diluted Shares | 47 | 52 | +5 |
| BS | Cash | 48 | 52 | +4 |
| BS | Short-term Investments | 37 | 44 | +7 |
| BS | Accounts Receivable | 48 | 46 | -2 |
| BS | Inventories | 38 | 41 | +3 |
| BS | Other Current Assets | 50 | 46 | -4 |
| BS | Total Current Assets | 47 | 47 | +0 |
| BS | PP&E, net | 50 | 52 | +2 |
| BS | Operating Lease ROU Assets | 25 | 52 | +27 ⚠ |
| BS | Long-term Investments | 22 | 24 | +2 |
| BS | Goodwill | 48 | 51 | +3 |
| BS | Intangible Assets, net | 41 | 47 | +6 |
| BS | Deferred Tax Assets | 27 | 31 | +4 |
| BS | Other Non-current Assets | 47 | 46 | -1 |
| BS | Total Non-current Assets | 47 | 8 | -39 ⚠ |
| BS | Total Assets | 52 | 52 | +0 |
| BS | Accounts Payable | 49 | 46 | -3 |
| BS | Short-term Debt | 43 | 44 | +1 |
| BS | Current Portion of LT Debt | 23 | 39 | +16 ⚠ |
| BS | Op. Lease Liabilities, current | 12 | 44 | +32 ⚠ |
| BS | Accrued Compensation | 16 | 38 | +22 ⚠ |
| BS | Deferred Revenue, current | 28 | 31 | +3 |
| BS | Income Tax Payable | 16 | 29 | +13 ⚠ |
| BS | Other Current Liabilities | 35 | 37 | +2 |
| BS | Total Current Liabilities | 50 | 47 | -3 |
| BS | Long-term Debt | 49 | 49 | +0 |
| BS | Op. Lease Liabilities, LT | 22 | 44 | +22 ⚠ |
| BS | Finance Lease Liabilities, LT | 1 | 19 | +18 ⚠ |
| BS | Deferred Revenue, LT | 12 | 21 | +9 ⚠ |
| BS | Deferred Tax Liability, LT | 34 | 43 | +9 ⚠ |
| BS | Pension & Retirement Oblig. | 6 | 11 | +5 |
| BS | Other Non-current Liabilities | 46 | 43 | -3 |
| BS | Total Non-current Liabilities | 50 | 7 | -43 ⚠ |
| BS | Total Liabilities | 52 | 52 | +0 |
| BS | Preferred Stock | 36 | 45 | +9 ⚠ |
| BS | Common Stock & APIC | 52 | 50 | -2 |
| BS | Additional Paid-in Capital | 40 | 42 | +2 |
| BS | Treasury Stock | 23 | 25 | +2 |
| BS | Retained Earnings | 52 | 52 | +0 |
| BS | AOCI | 52 | 52 | +0 |
| BS | Total Equity — Parent | 52 | 50 | -2 |
| BS | Noncontrolling Interests | 22 | 26 | +4 |
| BS | Total Equity incl. NCI | 22 | 32 | +10 ⚠ |
| BS | Total Liabilities & Equity | 52 | 52 | +0 |
| BS | Shares Outstanding | 50 | 49 | -1 |
| CF | Net Income | 52 | 104 | +52 ⚠ |
| CF | D&A | 49 | 51 | +2 |
| CF | SBC | 45 | 94 | +49 ⚠ |
| CF | Amortization of Intangibles | 14 | 38 | +24 ⚠ |
| CF | Change in Receivables | 36 | 37 | +1 |
| CF | Change in Inventories | 28 | 33 | +5 |
| CF | Change in Accounts Payable | 39 | 40 | +1 |
| CF | Change in Prepaid & Other Assets | 17 | 20 | +3 |
| CF | Change in Other Operating Assets | 19 | 22 | +3 |
| CF | Change in Deferred Revenue | 22 | 26 | +4 |
| CF | Other Working Capital | 25 | 17 | -8 ⚠ |
| CF | Other Non-cash Items | 32 | 32 | +0 |
| CF | Operating Cash Flow | 52 | 52 | +0 |
| CF | Capex | 49 | 45 | -4 |
| CF | Acquisitions | 41 | 43 | +2 |
| CF | Investment Purchases | 37 | 38 | +1 |
| CF | Investment Proceeds | 37 | 39 | +2 |
| CF | Investing Cash Flow | 52 | 52 | +0 |
| CF | Debt Proceeds | 44 | 35 | -9 ⚠ |
| CF | Debt Repayments | 46 | 34 | -12 ⚠ |
| CF | Share Repurchases | 48 | 49 | +1 |
| CF | Dividends Paid | 40 | 40 | +0 |
| CF | Financing Cash Flow | 52 | 52 | +0 |
| CF | FX Effect on Cash | 40 | 41 | +1 |
| CF | Net Change in Cash | 52 | 52 | +0 |
| CF | Ending Cash | 50 | 52 | +2 |
| CF | Cash Taxes Paid | 25 | 27 | +2 |
| CF | Cash Interest Paid | 22 | 26 | +4 |
| CF | Free Cash Flow | 49 | 0 | -49 ⚠ |

**現行路徑達到「≥45 家有值且填滿率 >90%」的列：47 / 97**

## 四、怎麼重跑

```
venv/Scripts/python.exe scripts/spike_derive_mapping.py    # 需要答案卷，慢
venv/Scripts/python.exe scripts/spike_verify_mapping.py    # 用快取，幾秒
```