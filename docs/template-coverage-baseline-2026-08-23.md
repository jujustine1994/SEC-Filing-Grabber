# 模板體檢：56 家公司的逐列覆蓋率（2026-08-23 產出）

**這份是自動產出的基線，不是手寫的。** 資料來源 `output/_spike/`（52 家的
companyfacts JSON 與現行路徑答案卷快取），重跑不用打網路。

公司清單刻意涵蓋大中小型 × 跨產業，**包含金融股（JPM/GS/BAC/SCHW）與 REIT（PLD）**
——它們的報表結構跟製造業差很多，是檢驗模板通不通用最有效的一群。

## 零、這份文件怎麼讀（先看這段，不然數字會誤導）

### 「達標列數」是什麼

一列要「達標」必須**同時**滿足兩個條件：

```
有值的公司數 >= 45      這一列在絕大多數公司都抓得到
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

**真正該當 KPI 的是第五節那兩個數字**：〔真缺口〕該抓到卻沒抓到幾列、
〔假警報〕Index 標紅裡有幾個是誤判。達標列數只是一支粗略的體溫計。

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
| ABBV | 21 | 1 | 1 | 3 | 0 |  |
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
| LLY | 21 | 1 | 1 | 4 | 2 |  |
| MCD | 21 | 1 | 1 | 8 | 1 |  |
| META | 57 | 1 | 1 | 4 | 1 |  |
| MRK | 21 | 1 | 1 | 7 | 1 |  |
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
| TMO | 21 | 1 | 1 | 2 | 1 |  |
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
| Acquisitions | 15 / 56 |
| Debt Repayments | 13 / 56 |
| Debt Proceeds | 12 / 56 |
| Short-term Debt | 12 / 56 |
| Other Working Capital | 9 / 56 |
| Share Repurchases | 9 / 56 |
| Current Portion of LT Debt | 6 / 56 |
| Preferred Stock | 6 / 56 |
| Investment Proceeds | 6 / 56 |
| Other Non-cash Items | 5 / 56 |
| Cash Taxes Paid | 5 / 56 |
| Amortization of Intangibles | 5 / 56 |
| Ending Cash | 5 / 56 |
| Investment Purchases | 5 / 56 |
| Shares Outstanding | 5 / 56 |

### 零星有值（填滿率 <70%，多半是公司本來就沒這項活動，不是漏抓）

2026-08-23（H3-2）從「中間有洞」拆出來的一類。拿 companyfacts 當真值驗 52 家、
2,906 個洞：填滿率 70% 以下的那 1,526 個洞**只有 18% 是真的漏抓**，70% 以上才
到 53%。門檻的完整證據見 `data_quality._SPORADIC_FILL_RATIO`。

| 列名 | 幾家中招 |
|---|---|
| Debt Proceeds | 14 / 56 |
| Acquisitions | 13 / 56 |
| Preferred Stock | 7 / 56 |
| Debt Repayments | 4 / 56 |
| Short-term Debt | 4 / 56 |
| D&A (CF memo) | 4 / 56 |
| Investment Purchases | 3 / 56 |
| Intangible Assets, net | 3 / 56 |
| Short-term Investments | 2 / 56 |
| Deferred Tax Liability, LT | 2 / 56 |
| Current Portion of LT Debt | 2 / 56 |
| Treasury Stock | 2 / 56 |
| Additional Paid-in Capital | 2 / 56 |
| Long-term Debt | 2 / 56 |
| Goodwill | 1 / 56 |

### 被判矛盾（整列空白，但同一家公司的相關欄位顯示應該要有）

| 列名 | 幾家中招 |
|---|---|
| Op. Lease Liabilities, current | 14 / 56 |
| Change in Inventories | 13 / 56 |
| Debt Proceeds | 6 / 56 |
| Op. Lease Liabilities, LT | 4 / 56 |
| Debt Repayments | 4 / 56 |
| Current Portion of LT Debt | 3 / 56 |
| Minority Interest | 2 / 56 |
| Share Repurchases | 1 / 56 |
| Noncontrolling Interests | 1 / 56 |

**中招家數多 ≠ concept 對照錯。** 仍在榜上的 `Op. Lease Liabilities, current`
等列，實測多數是**公司沒有在報表表面單獨列出**（金額併在「其他流動負債」裡，
只在附註拆開），現行逐份解 filing 的路徑結構上拿不到。動 concept 對照之前，
先把那份 filing 的報表 dataframe 印出來確認這一列到底在不在。

## 三、逐列覆蓋率：現行路徑 vs companyfacts

「有值公司數」＝ 52 家裡有幾家這一列拿得到資料。兩邊差 ≥8 家的標 ⚠。

| 表 | 列名 | 現行 | facts | 差 |
|---|---|---|---|---|
| IS | Revenue | 56 | 55 | -1 |
| IS | Cost of Revenue | 46 | 48 | +2 |
| IS | Gross Profit | 46 | 37 | -9 ⚠ |
| IS | R&D Expense | 42 | 39 | -3 |
| IS | SG&A Expense | 52 | 51 | -1 |
| IS | D&A (CF memo) | 56 | 55 | -1 |
| IS | Other Operating Expense | 14 | 18 | +4 |
| IS | Total Operating Expense | 27 | 28 | +1 |
| IS | Total Costs and Expenses | 16 | 19 | +3 |
| IS | Operating Income | 52 | 46 | -6 |
| IS | Interest Expense | 48 | 50 | +2 |
| IS | Interest Income | 16 | 15 | -1 |
| IS | Other Non-op Inc/(Exp) | 38 | 46 | +8 ⚠ |
| IS | Total Non-op Income/(Loss) | 54 | 46 | -8 ⚠ |
| IS | Pre-tax Income | 53 | 53 | +0 |
| IS | Income Tax | 56 | 56 | +0 |
| IS | Net Income | 56 | 112 | +56 ⚠ |
| IS | Minority Interest | 25 | 29 | +4 |
| IS | Net Income incl. NCI | 25 | 37 | +12 ⚠ |
| IS | SBC | 49 | 102 | +53 ⚠ |
| IS | Basic EPS | 56 | 56 | +0 |
| IS | Diluted EPS | 56 | 56 | +0 |
| IS | Basic Shares | 49 | 56 | +7 |
| IS | Diluted Shares | 50 | 56 | +6 |
| BS | Cash | 52 | 56 | +4 |
| BS | Short-term Investments | 41 | 48 | +7 |
| BS | Accounts Receivable | 52 | 50 | -2 |
| BS | Inventories | 42 | 45 | +3 |
| BS | Other Current Assets | 54 | 50 | -4 |
| BS | Total Current Assets | 51 | 51 | +0 |
| BS | PP&E, net | 54 | 56 | +2 |
| BS | Operating Lease ROU Assets | 25 | 56 | +31 ⚠ |
| BS | Long-term Investments | 24 | 26 | +2 |
| BS | Goodwill | 52 | 55 | +3 |
| BS | Intangible Assets, net | 45 | 51 | +6 |
| BS | Deferred Tax Assets | 29 | 32 | +3 |
| BS | Other Non-current Assets | 51 | 50 | -1 |
| BS | Total Non-current Assets | 51 | 9 | -42 ⚠ |
| BS | Total Assets | 56 | 56 | +0 |
| BS | Accounts Payable | 53 | 50 | -3 |
| BS | Short-term Debt | 47 | 48 | +1 |
| BS | Current Portion of LT Debt | 24 | 42 | +18 ⚠ |
| BS | Op. Lease Liabilities, current | 12 | 48 | +36 ⚠ |
| BS | Accrued Compensation | 18 | 42 | +24 ⚠ |
| BS | Deferred Revenue, current | 29 | 32 | +3 |
| BS | Income Tax Payable | 18 | 32 | +14 ⚠ |
| BS | Other Current Liabilities | 38 | 40 | +2 |
| BS | Total Current Liabilities | 54 | 51 | -3 |
| BS | Long-term Debt | 53 | 53 | +0 |
| BS | Op. Lease Liabilities, LT | 22 | 48 | +26 ⚠ |
| BS | Finance Lease Liabilities, LT | 1 | 20 | +19 ⚠ |
| BS | Deferred Revenue, LT | 12 | 22 | +10 ⚠ |
| BS | Deferred Tax Liability, LT | 38 | 47 | +9 ⚠ |
| BS | Pension & Retirement Oblig. | 7 | 13 | +6 |
| BS | Other Non-current Liabilities | 50 | 47 | -3 |
| BS | Total Non-current Liabilities | 54 | 9 | -45 ⚠ |
| BS | Total Liabilities | 56 | 56 | +0 |
| BS | Preferred Stock | 37 | 48 | +11 ⚠ |
| BS | Common Stock & APIC | 56 | 53 | -3 |
| BS | Additional Paid-in Capital | 44 | 46 | +2 |
| BS | Treasury Stock | 26 | 29 | +3 |
| BS | Retained Earnings | 56 | 56 | +0 |
| BS | AOCI | 56 | 56 | +0 |
| BS | Total Equity — Parent | 56 | 54 | -2 |
| BS | Noncontrolling Interests | 26 | 30 | +4 |
| BS | Total Equity incl. NCI | 26 | 36 | +10 ⚠ |
| BS | Total Liabilities & Equity | 56 | 56 | +0 |
| BS | Shares Outstanding | 54 | 53 | -1 |
| CF | Net Income | 56 | 112 | +56 ⚠ |
| CF | D&A | 53 | 55 | +2 |
| CF | SBC | 49 | 102 | +53 ⚠ |
| CF | Amortization of Intangibles | 17 | 41 | +24 ⚠ |
| CF | Change in Receivables | 37 | 40 | +3 |
| CF | Change in Inventories | 29 | 36 | +7 |
| CF | Change in Accounts Payable | 40 | 42 | +2 |
| CF | Change in Prepaid & Other Assets | 18 | 22 | +4 |
| CF | Change in Other Operating Assets | 19 | 23 | +4 |
| CF | Change in Deferred Revenue | 22 | 26 | +4 |
| CF | Other Working Capital | 27 | 20 | -7 |
| CF | Other Non-cash Items | 35 | 35 | +0 |
| CF | Operating Cash Flow | 56 | 56 | +0 |
| CF | Capex | 53 | 48 | -5 |
| CF | Acquisitions | 45 | 47 | +2 |
| CF | Investment Purchases | 41 | 42 | +1 |
| CF | Investment Proceeds | 38 | 39 | +1 |
| CF | Investing Cash Flow | 56 | 56 | +0 |
| CF | Debt Proceeds | 48 | 39 | -9 ⚠ |
| CF | Debt Repayments | 50 | 37 | -13 ⚠ |
| CF | Share Repurchases | 52 | 53 | +1 |
| CF | Dividends Paid | 44 | 44 | +0 |
| CF | Financing Cash Flow | 56 | 56 | +0 |
| CF | FX Effect on Cash | 44 | 45 | +1 |
| CF | Net Change in Cash | 56 | 56 | +0 |
| CF | Ending Cash | 53 | 56 | +3 |
| CF | Cash Taxes Paid | 25 | 27 | +2 |
| CF | Cash Interest Paid | 22 | 27 | +5 |
| CF | Free Cash Flow | 53 | 0 | -53 ⚠ |

**現行路徑達到「>=45 家有值且填滿率 >90%」的列：51 / 97**
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

本次 56 家共 1661 個期間欄，其中 **402 欄是 Q4**（24%）——
這些欄的流量列全部是合成的。Q1~Q3 不齊全時合成會失敗，那一整欄會空掉，
`data_quality` 的「整欄稀疏」就是用來抓這件事的。

## 五、兩個真正的 KPI

### KPI 1 — 真缺口：該抓到卻沒抓到

判準：**這家公司確實 tag 過**（companyfacts 讀得到），我們卻整列空白。
兩邊都沒有的不算——那是公司真的沒報，不是我們的問題。

| 列名 | 幾家真缺 | 哪幾家 |
|---|---|---|
| Op. Lease Liabilities, current | 36 / 56 | AAPL, ABBV, ADI, AMAT, AMD, AMZN, ARLO, AVGO … |
| Operating Lease ROU Assets | 31 / 56 | AAPL, ABBV, ADI, AMAT, AVGO, BAC, CAT, COHR … |
| Op. Lease Liabilities, LT | 26 / 56 | AAPL, ABBV, ADI, AMAT, AMZN, AVGO, CAT, CVX … |
| Amortization of Intangibles | 25 / 56 | ADBE, AMAT, ARLO, BAC, CAT, CRM, DDOG, FORM … |
| Accrued Compensation | 24 / 56 | AAPL, ABBV, ADBE, ADI, AMAT, AMD, AMZN, ARLO … |
| Finance Lease Liabilities, LT | 19 / 56 | AAPL, ABBV, AMAT, AMZN, ARLO, AVGO, COST, CRM … |
| Current Portion of LT Debt | 19 / 56 | AMAT, AMZN, CRM, CVX, GE, GOOGL, INTC, JNJ … |
| Income Tax Payable | 15 / 56 | AAPL, AMAT, AMD, CRM, DDOG, FORM, LITE, META … |
| Net Income incl. NCI | 12 / 56 | ADI, BAC, CRM, DDOG, FORM, GS, JNJ, JPM … |
| Preferred Stock | 11 / 56 | AAPL, ABBV, BAC, GOOGL, JPM, MRK, MSFT, MU … |
| Deferred Tax Liability, LT | 10 / 56 | AAPL, AMZN, AVGO, CAT, CRM, DDOG, MRVL, MU … |
| Total Equity incl. NCI | 10 / 56 | AMAT, CRM, FORM, JPM, LITE, MU, NKE, PANW … |
| Deferred Revenue, LT | 10 / 56 | AMZN, AVGO, CAT, COHR, LITE, NVDA, ON, ONTO … |
| Interest Expense | 8 / 56 | AAPL, GOOGL, INTC, LLY, LRCX, MRK, MSFT, PFE |
| Other Non-op Inc/(Exp) | 8 / 56 | AAPL, CVX, GOOGL, LLY, MSFT, NXPI, ORCL, QCOM |
| Other Current Liabilities | 8 / 56 | ADI, CRM, LRCX, META, NKE, PANW, PG, XOM |
| Change in Inventories | 8 / 56 | CVX, MRK, NEE, ONTO, ORCL, PFE, TMO, XOM |
| Pension & Retirement Oblig. | 7 / 56 | ABBV, ADI, AMAT, INTC, LITE, NXPI, ON |
| Long-term Investments | 7 / 56 | ADBE, COST, KLAC, MRVL, QCOM, SWKS, TXN |
| Short-term Investments | 7 / 56 | ARLO, GS, JPM, LRCX, MCD, SCHW, XOM |

**真缺口總計：455 個（列 × 公司）組合，分布在 65 個模板列。**

榜首那幾列全部是 TODO D10（只寫在附註、沒印在報表表面）——這是**已知的暫時性
限制**，不是新 bug。要壓低這個數字只有兩條路：接一條讀附註的路徑，或接受它。

### KPI 2 — 假警報：Index 標紅裡有幾個是誤判

標紅只有兩類：〔矛盾〕整列空白但相關欄位顯示該有、〔中間有洞〕。
「零星有值」刻意不標紅（H3-2），所以不算在內。

| | 家次 |
|---|---|
| 標紅：矛盾 | 48 |
| 標紅：中間有洞 | 208 |
| **標紅合計** | **256** |
| 降級為零星有值（不標紅） | 78 |

**要壓低的是標紅合計裡的誤判比例**，不是把標紅壓到 0——真缺口該標就要標。
驗證方式：對標紅的列抽樣，走 ARCHITECTURE「三步排查順序」確認是哪一類。

## 六、怎麼重跑

```
venv/Scripts/python.exe scripts/spike_derive_mapping.py    # 需要答案卷，慢
venv/Scripts/python.exe scripts/spike_verify_mapping.py    # 用快取，幾秒
```