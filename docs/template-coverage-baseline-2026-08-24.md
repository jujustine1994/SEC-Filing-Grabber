# 模板體檢：201 家公司的逐列覆蓋率（2026-08-24 產出）

**這份是自動產出的基線，不是手寫的。** 資料來源 `output/_spike/`（201 家的
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
有值的公司數 >= 171 家（85% 的樣本）   這一列在絕大多數公司都抓得到
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
| AFL | 21 | 1 | 21 | 0 | 0 | **是** |
| AMP | 21 | 1 | 21 | 0 | 0 | **是** |
| BAC | 21 | 1 | 21 | 0 | 0 | **是** |
| COF | 21 | 1 | 21 | 0 | 2 | **是** |
| AXP | 21 | 1 | 18 | 0 | 0 | **是** |
| MET | 21 | 1 | 18 | 0 | 0 | **是** |
| GS | 21 | 1 | 17 | 0 | 1 | **是** |
| AZO | 26 | 1 | 15 | 3 | 1 | **是** |
| AIG | 21 | 1 | 12 | 3 | 1 | **是** |
| PEP | 27 | 1 | 12 | 1 | 0 |  |
| SCHW | 21 | 1 | 10 | 1 | 0 |  |
| HPQ | 22 | 1 | 6 | 0 | 0 |  |
| KR | 16 | 7 | 6 | 0 | 1 |  |
| CSX | 23 | 1 | 5 | 24 | 2 |  |
| ALL | 21 | 1 | 4 | 3 | 1 |  |
| PLD | 21 | 1 | 4 | 2 | 2 |  |
| ADBE | 69 | 1 | 3 | 7 | 0 |  |
| INTC | 69 | 1 | 3 | 10 | 0 |  |
| CB | 21 | 1 | 2 | 5 | 0 |  |
| HIG | 21 | 1 | 2 | 1 | 1 |  |
| LHX | 21 | 2 | 2 | 11 | 1 |  |
| ONTO | 21 | 2 | 2 | 3 | 4 |  |
| SNOW | 21 | 2 | 2 | 1 | 0 |  |
| AAPL | 69 | 1 | 1 | 3 | 0 |  |
| ABBV | 21 | 1 | 1 | 3 | 0 |  |
| ABT | 21 | 1 | 1 | 2 | 1 |  |
| ACN | 21 | 1 | 1 | 0 | 1 |  |
| ADI | 21 | 1 | 1 | 4 | 1 |  |
| AEP | 21 | 1 | 1 | 5 | 2 |  |
| AMAT | 21 | 1 | 1 | 2 | 0 |  |
| AMD | 65 | 1 | 1 | 11 | 1 |  |
| AMGN | 21 | 1 | 1 | 4 | 0 |  |
| AMT | 21 | 1 | 1 | 6 | 0 |  |
| AMZN | 21 | 1 | 1 | 0 | 3 |  |
| ANET | 21 | 1 | 1 | 1 | 1 |  |
| AON | 21 | 1 | 1 | 2 | 1 |  |
| APD | 21 | 1 | 1 | 1 | 2 |  |
| APH | 21 | 1 | 1 | 3 | 1 |  |
| ARLO | 21 | 1 | 1 | 1 | 1 |  |
| AVGO | 33 | 1 | 1 | 6 | 0 |  |
| BA | 21 | 1 | 1 | 3 | 1 |  |
| BDX | 21 | 1 | 1 | 0 | 1 |  |
| BK | 21 | 1 | 1 | 3 | 0 |  |
| BKNG | 21 | 1 | 1 | 5 | 1 |  |
| BLK | 8 | 1 | 1 | 1 | 2 |  |
| BMY | 21 | 1 | 1 | 2 | 0 |  |
| BSX | 21 | 1 | 1 | 0 | 1 |  |
| C | 21 | 1 | 1 | 0 | 0 |  |
| CAT | 21 | 1 | 1 | 1 | 0 |  |
| CCI | 21 | 1 | 1 | 7 | 0 |  |
| CDNS | 21 | 1 | 1 | 3 | 0 |  |
| CDW | 21 | 1 | 1 | 3 | 1 |  |
| CHTR | 21 | 1 | 1 | 1 | 0 |  |
| CI | 21 | 1 | 1 | 2 | 0 |  |
| CL | 21 | 1 | 1 | 3 | 0 |  |
| CMCSA | 21 | 1 | 1 | 0 | 0 |  |
| CME | 21 | 1 | 1 | 4 | 0 |  |
| CMG | 21 | 1 | 1 | 0 | 0 |  |
| CNC | 21 | 1 | 1 | 1 | 0 |  |
| COP | 21 | 1 | 1 | 3 | 1 |  |
| CSCO | 21 | 1 | 1 | 0 | 0 |  |
| CTSH | 21 | 1 | 1 | 1 | 0 |  |
| CVX | 21 | 1 | 1 | 4 | 2 |  |
| D | 21 | 1 | 1 | 4 | 0 |  |
| DAL | 21 | 1 | 1 | 1 | 1 |  |
| DD | 21 | 1 | 1 | 5 | 0 |  |
| DDOG | 21 | 1 | 1 | 3 | 0 |  |
| DE | 21 | 1 | 1 | 3 | 0 |  |
| DHR | 21 | 1 | 1 | 5 | 1 |  |
| DIS | 21 | 1 | 1 | 4 | 0 |  |
| DLR | 21 | 1 | 1 | 2 | 1 |  |
| DOV | 21 | 1 | 1 | 4 | 2 |  |
| DOW | 21 | 1 | 1 | 3 | 0 |  |
| DUK | 21 | 1 | 1 | 1 | 1 |  |
| DXCM | 21 | 1 | 1 | 3 | 1 |  |
| ECL | 21 | 1 | 1 | 5 | 1 |  |
| EMR | 21 | 1 | 1 | 5 | 0 |  |
| EOG | 21 | 1 | 1 | 1 | 0 |  |
| EQIX | 21 | 1 | 1 | 6 | 1 |  |
| ETN | 21 | 1 | 1 | 5 | 2 |  |
| EW | 21 | 1 | 1 | 7 | 1 |  |
| EXC | 21 | 1 | 1 | 2 | 1 |  |
| F | 21 | 1 | 1 | 2 | 0 |  |
| FAST | 21 | 1 | 1 | 0 | 0 |  |
| FICO | 21 | 1 | 1 | 0 | 1 |  |
| FIS | 21 | 1 | 1 | 1 | 0 |  |
| FORM | 21 | 1 | 1 | 5 | 0 |  |
| FTNT | 21 | 1 | 1 | 3 | 0 |  |
| GD | 21 | 1 | 1 | 2 | 0 |  |
| GE | 21 | 1 | 1 | 2 | 4 |  |
| GILD | 21 | 1 | 1 | 2 | 1 |  |
| GLW | 21 | 1 | 1 | 6 | 0 |  |
| GM | 21 | 1 | 1 | 0 | 3 |  |
| GOOGL | 44 | 1 | 1 | 4 | 2 |  |
| GWW | 21 | 1 | 1 | 5 | 0 |  |
| HAL | 21 | 1 | 1 | 4 | 0 |  |
| HCA | 21 | 1 | 1 | 0 | 1 |  |
| HLT | 21 | 1 | 1 | 3 | 1 |  |
| HON | 21 | 1 | 1 | 3 | 0 |  |
| HSY | 21 | 1 | 1 | 1 | 0 |  |
| HUM | 21 | 1 | 1 | 2 | 0 |  |
| IBM | 21 | 1 | 1 | 0 | 2 |  |
| ICE | 21 | 1 | 1 | 1 | 2 |  |
| IDXX | 21 | 1 | 1 | 2 | 1 |  |
| INTU | 21 | 1 | 1 | 3 | 1 |  |
| IP | 21 | 1 | 1 | 4 | 1 |  |
| IQV | 21 | 1 | 1 | 3 | 1 |  |
| IR | 21 | 1 | 1 | 0 | 0 |  |
| ISRG | 21 | 1 | 1 | 3 | 3 |  |
| ITW | 21 | 1 | 1 | 2 | 1 |  |
| JCI | 21 | 1 | 1 | 2 | 0 |  |
| JNJ | 21 | 1 | 1 | 0 | 1 |  |
| JPM | 21 | 1 | 1 | 0 | 2 |  |
| KDP | 21 | 1 | 1 | 5 | 0 |  |
| KEYS | 21 | 1 | 1 | 5 | 0 |  |
| KHC | 21 | 1 | 1 | 5 | 0 |  |
| KMB | 21 | 1 | 1 | 4 | 1 |  |
| KO | 21 | 1 | 1 | 0 | 1 |  |
| LIN | 21 | 1 | 1 | 4 | 0 |  |
| LLY | 21 | 1 | 1 | 3 | 2 |  |
| LMT | 21 | 1 | 1 | 0 | 1 |  |
| LVS | 21 | 1 | 1 | 4 | 1 |  |
| MA | 21 | 1 | 1 | 3 | 1 |  |
| MAR | 21 | 1 | 1 | 2 | 1 |  |
| MCD | 21 | 1 | 1 | 8 | 1 |  |
| MCO | 21 | 1 | 1 | 1 | 0 |  |
| MDLZ | 21 | 1 | 1 | 0 | 1 |  |
| META | 57 | 1 | 1 | 4 | 1 |  |
| MMM | 21 | 1 | 1 | 0 | 0 |  |
| MNST | 21 | 1 | 1 | 7 | 1 |  |
| MPC | 21 | 1 | 1 | 1 | 1 |  |
| MRK | 21 | 1 | 1 | 6 | 1 |  |
| MS | 21 | 1 | 1 | 1 | 3 |  |
| MSCI | 21 | 1 | 1 | 1 | 1 |  |
| MSFT | 67 | 2 | 1 | 4 | 1 |  |
| MSI | 21 | 1 | 1 | 1 | 1 |  |
| MU | 21 | 1 | 1 | 3 | 2 |  |
| NDAQ | 21 | 1 | 1 | 3 | 1 |  |
| NEE | 21 | 1 | 1 | 0 | 1 |  |
| NEM | 21 | 1 | 1 | 3 | 0 |  |
| NFLX | 21 | 1 | 1 | 2 | 0 |  |
| NOC | 21 | 1 | 1 | 2 | 2 |  |
| NOW | 21 | 1 | 1 | 2 | 1 |  |
| NSC | 21 | 1 | 1 | 0 | 0 |  |
| NUE | 21 | 1 | 1 | 5 | 0 |  |
| NVDA | 68 | 1 | 1 | 7 | 1 |  |
| NXPI | 21 | 1 | 1 | 4 | 0 |  |
| ODFL | 21 | 1 | 1 | 1 | 1 |  |
| OKE | 21 | 1 | 1 | 3 | 0 |  |
| OMC | 21 | 1 | 1 | 2 | 1 |  |
| ON | 21 | 1 | 1 | 6 | 0 |  |
| ORLY | 21 | 1 | 1 | 4 | 1 |  |
| OTIS | 21 | 1 | 1 | 2 | 1 |  |
| OXY | 21 | 1 | 1 | 4 | 1 |  |
| PANW | 21 | 1 | 1 | 4 | 2 |  |
| PFE | 21 | 1 | 1 | 3 | 1 |  |
| PSX | 21 | 1 | 1 | 3 | 0 |  |
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
| WFC | 21 | 1 | 1 | 0 | 0 |  |
| WMT | 68 | 1 | 1 | 9 | 0 |  |
| XOM | 21 | 1 | 1 | 2 | 1 |  |
| ADP | 21 | 1 | 0 | 5 | 1 |  |
| COHR | 21 | 1 | 0 | 3 | 0 |  |
| CRM | 21 | 0 | 0 | 4 | 0 |  |
| CTAS | 21 | 1 | 0 | 3 | 1 |  |
| EA | 21 | 0 | 0 | 2 | 2 |  |
| EL | 21 | 1 | 0 | 3 | 0 |  |
| FDX | 21 | 1 | 0 | 4 | 0 |  |
| GEHC | 14 | 0 | 0 | 2 | 2 |  |
| GIS | 21 | 1 | 0 | 1 | 0 |  |
| HD | 21 | 0 | 0 | 5 | 0 |  |
| KLAC | 21 | 1 | 0 | 1 | 0 |  |
| LITE | 21 | 1 | 0 | 2 | 0 |  |
| LOW | 21 | 0 | 0 | 5 | 0 |  |
| LRCX | 21 | 1 | 0 | 3 | 1 |  |
| LULU | 21 | 0 | 0 | 1 | 0 |  |
| MCHP | 21 | 0 | 0 | 1 | 0 |  |
| MCK | 21 | 0 | 0 | 3 | 0 |  |
| MDT | 21 | 1 | 0 | 1 | 0 |  |
| MRVL | 21 | 0 | 0 | 4 | 0 |  |
| NKE | 21 | 1 | 0 | 1 | 1 |  |
| ORCL | 21 | 1 | 0 | 1 | 1 |  |
| PAYX | 21 | 1 | 0 | 2 | 2 |  |
| PG | 21 | 1 | 0 | 1 | 0 |  |
| TGT | 21 | 0 | 0 | 4 | 1 |  |

**觸發「模板不適用」的 9 家：AFL, AIG, AMP, AXP, AZO, BAC, COF, GS, MET**——全是金融股。IS/BS/CF 模板是為製造業設計的，銀行／券商的報表結構完全不同（存款、放款、備抵呆帳…），這是 TODO D8 記錄的已知限制，現在有量化證據。

## 二、最常出問題的列

### 中間有洞（同一列有些期有、有些沒有——一定是漏抓）

| 列名 | 幾家中招 |
|---|---|
| Acquisitions | 43 / 201 |
| Debt Repayments | 37 / 201 |
| Debt Proceeds | 35 / 201 |
| Share Repurchases | 26 / 201 |
| Short-term Debt | 22 / 201 |
| Investment Purchases | 20 / 201 |
| Cash Taxes Paid | 18 / 201 |
| Other Working Capital | 17 / 201 |
| Current Portion of LT Debt | 16 / 201 |
| Minority Interest | 16 / 201 |
| Shares Outstanding | 14 / 201 |
| Ending Cash | 13 / 201 |
| Interest Income | 13 / 201 |
| Investment Proceeds | 12 / 201 |
| Operating Income | 12 / 201 |

### 零星有值（填滿率 <70%，多半是公司本來就沒這項活動，不是漏抓）

2026-08-23（H3-2）從「中間有洞」拆出來的一類。當時拿 companyfacts 當真值驗 52 家、
2,906 個洞：填滿率 70% 以下的那 1,526 個洞**只有 18% 是真的漏抓**，70% 以上才
到 53%。門檻的完整證據見 `data_quality._SPORADIC_FILL_RATIO`。

| 列名 | 幾家中招 |
|---|---|
| Debt Proceeds | 45 / 201 |
| Acquisitions | 32 / 201 |
| Debt Repayments | 23 / 201 |
| D&A (CF memo) | 13 / 201 |
| Preferred Stock | 13 / 201 |
| Treasury Stock | 11 / 201 |
| Investment Purchases | 10 / 201 |
| Short-term Debt | 9 / 201 |
| Share Repurchases | 8 / 201 |
| Income Tax Payable | 7 / 201 |
| Current Portion of LT Debt | 6 / 201 |
| Investment Proceeds | 6 / 201 |
| Additional Paid-in Capital | 6 / 201 |
| Other Current Liabilities | 5 / 201 |
| Dividends Paid | 5 / 201 |

### 被判矛盾（整列空白，但同一家公司的相關欄位顯示應該要有）

| 列名 | 幾家中招 |
|---|---|
| Op. Lease Liabilities, current | 48 / 201 |
| Change in Inventories | 25 / 201 |
| Minority Interest | 15 / 201 |
| Debt Proceeds | 15 / 201 |
| Current Portion of LT Debt | 14 / 201 |
| Share Repurchases | 10 / 201 |
| Debt Repayments | 9 / 201 |
| Op. Lease Liabilities, LT | 8 / 201 |
| Noncontrolling Interests | 2 / 201 |

**中招家數多 ≠ concept 對照錯。** 仍在榜上的 `Op. Lease Liabilities, current`
等列，實測多數是**公司沒有在報表表面單獨列出**（金額併在「其他流動負債」裡，
只在附註拆開），現行逐份解 filing 的路徑結構上拿不到。動 concept 對照之前，
先把那份 filing 的報表 dataframe 印出來確認這一列到底在不在。

## 三、逐列覆蓋率：現行路徑 vs companyfacts

「有值公司數」＝ 201 家裡有幾家這一列拿得到資料。兩邊差 8 家以上的標 ⚠。

| 表 | 列名 | 現行 | facts | 差 |
|---|---|---|---|---|
| IS | Revenue | 200 | 199 | -1 |
| IS | Cost of Revenue | 149 | 144 | -5 |
| IS | Gross Profit | 150 | 111 | -39 ⚠ |
| IS | R&D Expense | 88 | 84 | -4 |
| IS | SG&A Expense | 173 | 170 | -3 |
| IS | D&A (CF memo) | 197 | 197 | +0 |
| IS | Other Operating Expense | 51 | 73 | +22 ⚠ |
| IS | Total Operating Expense | 62 | 79 | +17 ⚠ |
| IS | Total Costs and Expenses | 71 | 84 | +13 ⚠ |
| IS | Operating Income | 186 | 171 | -15 ⚠ |
| IS | Interest Expense | 178 | 176 | -2 |
| IS | Interest Income | 64 | 67 | +3 |
| IS | Other Non-op Inc/(Exp) | 122 | 152 | +30 ⚠ |
| IS | Total Non-op Income/(Loss) | 191 | 152 | -39 ⚠ |
| IS | Pre-tax Income | 193 | 196 | +3 |
| IS | Income Tax | 201 | 201 | +0 |
| IS | Net Income | 201 | 200 | -1 |
| IS | Minority Interest | 111 | 131 | +20 ⚠ |
| IS | Net Income incl. NCI | 120 | 154 | +34 ⚠ |
| IS | SBC | 169 | 169 | +0 |
| IS | Basic EPS | 198 | 200 | +2 |
| IS | Diluted EPS | 198 | 200 | +2 |
| IS | Basic Shares | 159 | 199 | +40 ⚠ |
| IS | Diluted Shares | 161 | 200 | +39 ⚠ |
| BS | Cash | 178 | 201 | +23 ⚠ |
| BS | Short-term Investments | 106 | 155 | +49 ⚠ |
| BS | Accounts Receivable | 183 | 170 | -13 ⚠ |
| BS | Inventories | 134 | 137 | +3 |
| BS | Other Current Assets | 197 | 179 | -18 ⚠ |
| BS | Total Current Assets | 180 | 180 | +0 |
| BS | PP&E, net | 191 | 199 | +8 ⚠ |
| BS | Operating Lease ROU Assets | 95 | 200 | +105 ⚠ |
| BS | Long-term Investments | 73 | 89 | +16 ⚠ |
| BS | Goodwill | 184 | 195 | +11 ⚠ |
| BS | Intangible Assets, net | 160 | 186 | +26 ⚠ |
| BS | Deferred Tax Assets | 99 | 120 | +21 ⚠ |
| BS | Other Non-current Assets | 177 | 179 | +2 |
| BS | Total Non-current Assets | 180 | 29 | -151 ⚠ |
| BS | Total Assets | 201 | 201 | +0 |
| BS | Accounts Payable | 189 | 173 | -16 ⚠ |
| BS | Short-term Debt | 145 | 161 | +16 ⚠ |
| BS | Current Portion of LT Debt | 101 | 156 | +55 ⚠ |
| BS | Op. Lease Liabilities, current | 50 | 173 | +123 ⚠ |
| BS | Accrued Compensation | 72 | 141 | +69 ⚠ |
| BS | Deferred Revenue, current | 84 | 106 | +22 ⚠ |
| BS | Income Tax Payable | 78 | 114 | +36 ⚠ |
| BS | Other Current Liabilities | 148 | 145 | -3 |
| BS | Total Current Liabilities | 190 | 180 | -10 ⚠ |
| BS | Long-term Debt | 194 | 192 | -2 |
| BS | Op. Lease Liabilities, LT | 89 | 174 | +85 ⚠ |
| BS | Finance Lease Liabilities, LT | 2 | 87 | +85 ⚠ |
| BS | Deferred Revenue, LT | 26 | 67 | +41 ⚠ |
| BS | Deferred Tax Liability, LT | 146 | 167 | +21 ⚠ |
| BS | Pension & Retirement Oblig. | 47 | 66 | +19 ⚠ |
| BS | Other Non-current Liabilities | 175 | 170 | -5 |
| BS | Total Non-current Liabilities | 190 | 37 | -153 ⚠ |
| BS | Total Liabilities | 201 | 201 | +0 |
| BS | Preferred Stock | 110 | 174 | +64 ⚠ |
| BS | Common Stock & APIC | 189 | 188 | -1 |
| BS | Additional Paid-in Capital | 177 | 180 | +3 |
| BS | Treasury Stock | 122 | 140 | +18 ⚠ |
| BS | Retained Earnings | 200 | 200 | +0 |
| BS | AOCI | 199 | 200 | +1 |
| BS | Total Equity — Parent | 201 | 198 | -3 |
| BS | Noncontrolling Interests | 124 | 139 | +15 ⚠ |
| BS | Total Equity incl. NCI | 132 | 158 | +26 ⚠ |
| BS | Total Liabilities & Equity | 201 | 201 | +0 |
| BS | Shares Outstanding | 190 | 197 | +7 |
| CF | Net Income | 201 | 200 | -1 |
| CF | D&A | 188 | 197 | +9 ⚠ |
| CF | SBC | 169 | 169 | +0 |
| CF | Amortization of Intangibles | 45 | 145 | +100 ⚠ |
| CF | Change in Receivables | 152 | 149 | -3 |
| CF | Change in Inventories | 111 | 118 | +7 |
| CF | Change in Accounts Payable | 157 | 149 | -8 ⚠ |
| CF | Change in Prepaid & Other Assets | 52 | 62 | +10 ⚠ |
| CF | Change in Other Operating Assets | 61 | 87 | +26 ⚠ |
| CF | Change in Deferred Revenue | 59 | 78 | +19 ⚠ |
| CF | Other Working Capital | 136 | 105 | -31 ⚠ |
| CF | Other Non-cash Items | 107 | 115 | +8 ⚠ |
| CF | Operating Cash Flow | 201 | 201 | +0 |
| CF | Capex | 172 | 167 | -5 |
| CF | Acquisitions | 149 | 170 | +21 ⚠ |
| CF | Investment Purchases | 139 | 135 | -4 |
| CF | Investment Proceeds | 93 | 111 | +18 ⚠ |
| CF | Investing Cash Flow | 201 | 201 | +0 |
| CF | Debt Proceeds | 181 | 148 | -33 ⚠ |
| CF | Debt Repayments | 187 | 141 | -46 ⚠ |
| CF | Share Repurchases | 182 | 195 | +13 ⚠ |
| CF | Dividends Paid | 170 | 173 | +3 |
| CF | Financing Cash Flow | 201 | 201 | +0 |
| CF | FX Effect on Cash | 157 | 162 | +5 |
| CF | Net Change in Cash | 201 | 201 | +0 |
| CF | Ending Cash | 193 | 201 | +8 ⚠ |
| CF | Cash Taxes Paid | 86 | 99 | +13 ⚠ |
| CF | Cash Interest Paid | 80 | 104 | +24 ⚠ |
| CF | Free Cash Flow | 172 | 0 | -172 ⚠ |

**現行路徑達到「>=171 家（85%）有值且填滿率 >90%」的列：44 / 97**
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

本次 201 家共 4695 個期間欄，其中 **1123 欄是 Q4**（24%）——
這些欄的流量列全部是合成的。Q1~Q3 不齊全時合成會失敗，那一整欄會空掉，
`data_quality` 的「整欄稀疏」就是用來抓這件事的。

## 五、XBRL 裡到底有沒有模板要的數字

把「97 個模板列 × 201 家公司」每一格分成三類。**判斷「有沒有」靠
companyfacts**（它讀得到公司 tag 過的全部 fact，含附註層），比只看報表表面準。

| 分類 | 格數 | 佔比 | 意思 |
|---|---|---|---|
| 我們抓到了 | 13944 | 72% | 正常 |
| **真缺口** | 1824 | 9% | 公司有 tag，我們沒抓到 → 見下面 KPI 1 |
| 公司真的沒有 | 3528 | 18% | **不是問題**，這家公司就是沒報這個科目 |

另有 201 格不列入分類：那些模板列在 `facts_mapping` 裡沒有對應
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
| Op. Lease Liabilities, current | 125 / 201 | AAPL, ABBV, ABT, ADI, ADP, AMAT, AMD, AMGN … |
| Operating Lease ROU Assets | 105 / 201 | AAPL, ABBV, ABT, ADI, AIG, ALL, AMAT, AMGN … |
| Amortization of Intangibles | 100 / 201 | ACN, ADBE, ADP, ALL, AMAT, AMT, ANET, APH … |
| Op. Lease Liabilities, LT | 85 / 201 | AAPL, ABBV, ABT, ADI, AMAT, AMGN, AMZN, APH … |
| Finance Lease Liabilities, LT | 85 / 201 | AAPL, ABBV, AEP, AMAT, AMT, AMZN, ANET, AON … |
| Accrued Compensation | 74 / 201 | AAPL, ABBV, ADBE, ADI, AMAT, AMD, AMGN, AMT … |
| Preferred Stock | 64 / 201 | AAPL, ABBV, AFL, ALL, AMP, AMT, AON, AZO … |
| Current Portion of LT Debt | 57 / 201 | ADP, AMAT, AMZN, ANET, BDX, BMY, BSX, CB … |
| Short-term Investments | 51 / 201 | AFL, AMT, ARLO, AZO, BK, BSX, C, CDNS … |
| Income Tax Payable | 50 / 201 | AAPL, ALL, AMAT, AMD, AMGN, AMT, AON, BA … |
| Deferred Revenue, LT | 44 / 201 | AMT, AMZN, AON, APD, AVGO, CAT, COHR, COP … |
| Basic Shares | 41 / 201 | AMP, BA, BDX, BMY, CB, CL, CMCSA, COF … |
| Diluted Shares | 40 / 201 | AMP, BA, BDX, BMY, CB, CI, CL, CMCSA … |
| Deferred Revenue, current | 36 / 201 | AON, APD, AVGO, AXP, CHTR, COHR, COP, DD … |
| Net Income incl. NCI | 34 / 201 | ABT, ADI, AMP, AZO, BAC, BDX, BKNG, CRM … |
| Other Non-op Inc/(Exp) | 31 / 201 | AAPL, ADP, AFL, ANET, CB, CHTR, CMCSA, CNC … |
| Deferred Tax Assets | 29 / 201 | AAPL, APD, APH, ARLO, AVGO, AXP, CMG, CRM … |
| Pension & Retirement Oblig. | 29 / 201 | ABBV, ABT, ADI, AMAT, AZO, BMY, CHTR, CME … |
| Deferred Tax Liability, LT | 28 / 201 | AAPL, AMZN, AVGO, CAT, CDNS, CRM, DDOG, DHR … |
| Long-term Investments | 28 / 201 | ADBE, AON, APD, CDNS, CME, COST, DAL, DUK … |

**真缺口總計：1824 個（列 × 公司）組合，分布在 78 個模板列。**

榜首那幾列全部是 TODO D10（只寫在附註、沒印在報表表面）——這是**已知的暫時性
限制**，不是新 bug。要壓低這個數字只有兩條路：接一條讀附註的路徑，或接受它。

### KPI 2 — 假警報：Index 標紅裡有幾個是誤判

標紅只有兩類：〔矛盾〕整列空白但相關欄位顯示該有、〔中間有洞〕。
「零星有值」刻意不標紅（H3-2），所以不算在內。

| | 家次 |
|---|---|
| 標紅：矛盾 | 146 |
| 標紅：中間有洞 | 601 |
| **標紅合計** | **747** |
| 降級為零星有值（不標紅） | 270 |

**要壓低的是標紅合計裡的誤判比例**，不是把標紅壓到 0——真缺口該標就要標。
驗證方式：對標紅的列抽樣，走 ARCHITECTURE「三步排查順序」確認是哪一類。

## 七、怎麼重跑

```
venv/Scripts/python.exe scripts/spike_derive_mapping.py    # 需要答案卷，慢
venv/Scripts/python.exe scripts/spike_verify_mapping.py    # 用快取，幾秒
```