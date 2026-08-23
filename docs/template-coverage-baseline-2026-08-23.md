# 模板體檢：59 家公司的逐列覆蓋率（2026-08-23 產出）

**這份是自動產出的基線，不是手寫的。** 資料來源 `output/_spike/`（59 家的
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

**真正該當 KPI 的是第六節那兩個數字**：〔真缺口〕該抓到卻沒抓到幾列、
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
| DHR | 21 | 1 | 1 | 5 | 1 |  |
| FORM | 21 | 1 | 1 | 5 | 0 |  |
| GE | 21 | 1 | 1 | 2 | 4 |  |
| GOOGL | 44 | 1 | 1 | 4 | 2 |  |
| ISRG | 21 | 1 | 1 | 3 | 3 |  |
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
| MDT | 21 | 1 | 0 | 1 | 0 |  |
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
| Acquisitions | 17 / 59 |
| Debt Repayments | 14 / 59 |
| Debt Proceeds | 12 / 59 |
| Short-term Debt | 12 / 59 |
| Share Repurchases | 11 / 59 |
| Other Working Capital | 9 / 59 |
| Current Portion of LT Debt | 6 / 59 |
| Preferred Stock | 6 / 59 |
| Investment Proceeds | 6 / 59 |
| Cash Taxes Paid | 6 / 59 |
| Other Non-cash Items | 5 / 59 |
| Amortization of Intangibles | 5 / 59 |
| Ending Cash | 5 / 59 |
| Investment Purchases | 5 / 59 |
| Shares Outstanding | 5 / 59 |

### 零星有值（填滿率 <70%，多半是公司本來就沒這項活動，不是漏抓）

2026-08-23（H3-2）從「中間有洞」拆出來的一類。當時拿 companyfacts 當真值驗 52 家、
2,906 個洞：填滿率 70% 以下的那 1,526 個洞**只有 18% 是真的漏抓**，70% 以上才
到 53%。門檻的完整證據見 `data_quality._SPORADIC_FILL_RATIO`。

| 列名 | 幾家中招 |
|---|---|
| Debt Proceeds | 15 / 59 |
| Acquisitions | 14 / 59 |
| Preferred Stock | 7 / 59 |
| Debt Repayments | 5 / 59 |
| D&A (CF memo) | 5 / 59 |
| Short-term Debt | 4 / 59 |
| Investment Purchases | 3 / 59 |
| Intangible Assets, net | 3 / 59 |
| Short-term Investments | 2 / 59 |
| Deferred Tax Liability, LT | 2 / 59 |
| Current Portion of LT Debt | 2 / 59 |
| Treasury Stock | 2 / 59 |
| Additional Paid-in Capital | 2 / 59 |
| Long-term Debt | 2 / 59 |
| Goodwill | 1 / 59 |

### 被判矛盾（整列空白，但同一家公司的相關欄位顯示應該要有）

| 列名 | 幾家中招 |
|---|---|
| Op. Lease Liabilities, current | 14 / 59 |
| Change in Inventories | 13 / 59 |
| Debt Proceeds | 7 / 59 |
| Debt Repayments | 5 / 59 |
| Current Portion of LT Debt | 4 / 59 |
| Op. Lease Liabilities, LT | 4 / 59 |
| Minority Interest | 3 / 59 |
| Share Repurchases | 1 / 59 |
| Noncontrolling Interests | 1 / 59 |

**中招家數多 ≠ concept 對照錯。** 仍在榜上的 `Op. Lease Liabilities, current`
等列，實測多數是**公司沒有在報表表面單獨列出**（金額併在「其他流動負債」裡，
只在附註拆開），現行逐份解 filing 的路徑結構上拿不到。動 concept 對照之前，
先把那份 filing 的報表 dataframe 印出來確認這一列到底在不在。

## 三、逐列覆蓋率：現行路徑 vs companyfacts

「有值公司數」＝ 59 家裡有幾家這一列拿得到資料。兩邊差 8 家以上的標 ⚠。

| 表 | 列名 | 現行 | facts | 差 |
|---|---|---|---|---|
| IS | Revenue | 59 | 58 | -1 |
| IS | Cost of Revenue | 49 | 51 | +2 |
| IS | Gross Profit | 49 | 40 | -9 ⚠ |
| IS | R&D Expense | 45 | 42 | -3 |
| IS | SG&A Expense | 55 | 54 | -1 |
| IS | D&A (CF memo) | 59 | 58 | -1 |
| IS | Other Operating Expense | 16 | 20 | +4 |
| IS | Total Operating Expense | 28 | 30 | +2 |
| IS | Total Costs and Expenses | 16 | 20 | +4 |
| IS | Operating Income | 55 | 49 | -6 |
| IS | Interest Expense | 50 | 52 | +2 |
| IS | Interest Income | 16 | 16 | +0 |
| IS | Other Non-op Inc/(Exp) | 40 | 48 | +8 ⚠ |
| IS | Total Non-op Income/(Loss) | 57 | 48 | -9 ⚠ |
| IS | Pre-tax Income | 56 | 56 | +0 |
| IS | Income Tax | 59 | 59 | +0 |
| IS | Net Income | 59 | 118 | +59 ⚠ |
| IS | Minority Interest | 27 | 32 | +5 |
| IS | Net Income incl. NCI | 27 | 40 | +13 ⚠ |
| IS | SBC | 52 | 108 | +56 ⚠ |
| IS | Basic EPS | 59 | 59 | +0 |
| IS | Diluted EPS | 59 | 59 | +0 |
| IS | Basic Shares | 52 | 59 | +7 |
| IS | Diluted Shares | 53 | 59 | +6 |
| BS | Cash | 54 | 59 | +5 |
| BS | Short-term Investments | 43 | 51 | +8 ⚠ |
| BS | Accounts Receivable | 55 | 53 | -2 |
| BS | Inventories | 45 | 48 | +3 |
| BS | Other Current Assets | 57 | 53 | -4 |
| BS | Total Current Assets | 54 | 54 | +0 |
| BS | PP&E, net | 57 | 59 | +2 |
| BS | Operating Lease ROU Assets | 25 | 59 | +34 ⚠ |
| BS | Long-term Investments | 24 | 27 | +3 |
| BS | Goodwill | 55 | 58 | +3 |
| BS | Intangible Assets, net | 47 | 54 | +7 |
| BS | Deferred Tax Assets | 31 | 34 | +3 |
| BS | Other Non-current Assets | 53 | 53 | +0 |
| BS | Total Non-current Assets | 54 | 9 | -45 ⚠ |
| BS | Total Assets | 59 | 59 | +0 |
| BS | Accounts Payable | 56 | 52 | -4 |
| BS | Short-term Debt | 49 | 50 | +1 |
| BS | Current Portion of LT Debt | 25 | 44 | +19 ⚠ |
| BS | Op. Lease Liabilities, current | 12 | 51 | +39 ⚠ |
| BS | Accrued Compensation | 19 | 44 | +25 ⚠ |
| BS | Deferred Revenue, current | 30 | 33 | +3 |
| BS | Income Tax Payable | 19 | 34 | +15 ⚠ |
| BS | Other Current Liabilities | 41 | 43 | +2 |
| BS | Total Current Liabilities | 57 | 54 | -3 |
| BS | Long-term Debt | 56 | 55 | -1 |
| BS | Op. Lease Liabilities, LT | 22 | 51 | +29 ⚠ |
| BS | Finance Lease Liabilities, LT | 1 | 21 | +20 ⚠ |
| BS | Deferred Revenue, LT | 12 | 23 | +11 ⚠ |
| BS | Deferred Tax Liability, LT | 39 | 50 | +11 ⚠ |
| BS | Pension & Retirement Oblig. | 8 | 15 | +7 |
| BS | Other Non-current Liabilities | 53 | 50 | -3 |
| BS | Total Non-current Liabilities | 57 | 10 | -47 ⚠ |
| BS | Total Liabilities | 59 | 59 | +0 |
| BS | Preferred Stock | 39 | 51 | +12 ⚠ |
| BS | Common Stock & APIC | 58 | 56 | -2 |
| BS | Additional Paid-in Capital | 47 | 49 | +2 |
| BS | Treasury Stock | 26 | 29 | +3 |
| BS | Retained Earnings | 59 | 59 | +0 |
| BS | AOCI | 59 | 59 | +0 |
| BS | Total Equity — Parent | 59 | 57 | -2 |
| BS | Noncontrolling Interests | 29 | 33 | +4 |
| BS | Total Equity incl. NCI | 29 | 39 | +10 ⚠ |
| BS | Total Liabilities & Equity | 59 | 59 | +0 |
| BS | Shares Outstanding | 57 | 56 | -1 |
| CF | Net Income | 59 | 118 | +59 ⚠ |
| CF | D&A | 56 | 58 | +2 |
| CF | SBC | 52 | 108 | +56 ⚠ |
| CF | Amortization of Intangibles | 19 | 44 | +25 ⚠ |
| CF | Change in Receivables | 40 | 42 | +2 |
| CF | Change in Inventories | 32 | 39 | +7 |
| CF | Change in Accounts Payable | 43 | 45 | +2 |
| CF | Change in Prepaid & Other Assets | 20 | 24 | +4 |
| CF | Change in Other Operating Assets | 19 | 23 | +4 |
| CF | Change in Deferred Revenue | 23 | 27 | +4 |
| CF | Other Working Capital | 29 | 21 | -8 ⚠ |
| CF | Other Non-cash Items | 36 | 36 | +0 |
| CF | Operating Cash Flow | 59 | 59 | +0 |
| CF | Capex | 56 | 51 | -5 |
| CF | Acquisitions | 48 | 50 | +2 |
| CF | Investment Purchases | 44 | 45 | +1 |
| CF | Investment Proceeds | 39 | 42 | +3 |
| CF | Investing Cash Flow | 59 | 59 | +0 |
| CF | Debt Proceeds | 50 | 41 | -9 ⚠ |
| CF | Debt Repayments | 52 | 39 | -13 ⚠ |
| CF | Share Repurchases | 55 | 56 | +1 |
| CF | Dividends Paid | 47 | 47 | +0 |
| CF | Financing Cash Flow | 59 | 59 | +0 |
| CF | FX Effect on Cash | 47 | 48 | +1 |
| CF | Net Change in Cash | 59 | 59 | +0 |
| CF | Ending Cash | 56 | 59 | +3 |
| CF | Cash Taxes Paid | 27 | 29 | +2 |
| CF | Cash Interest Paid | 24 | 29 | +5 |
| CF | Free Cash Flow | 56 | 0 | -56 ⚠ |

**現行路徑達到「>=45 家有值且填滿率 >90%」的列：56 / 97**
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

本次 59 家共 1724 個期間欄，其中 **417 欄是 Q4**（24%）——
這些欄的流量列全部是合成的。Q1~Q3 不齊全時合成會失敗，那一整欄會空掉，
`data_quality` 的「整欄稀疏」就是用來抓這件事的。

## 五、XBRL 裡到底有沒有模板要的數字

把「97 個模板列 × 59 家公司」每一格分成三類。**判斷「有沒有」靠
companyfacts**（它讀得到公司 tag 過的全部 fact，含附註層），比只看報表表面準。

| 分類 | 格數 | 佔比 | 意思 |
|---|---|---|---|
| 我們抓到了 | 4155 | 73% | 正常 |
| **真缺口** | 488 | 9% | 公司有 tag，我們沒抓到 → 見下面 KPI 1 |
| 公司真的沒有 | 1021 | 18% | **不是問題**，這家公司就是沒報這個科目 |

另有 59 格不列入分類：那些模板列在 `facts_mapping` 裡沒有對應
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
| Op. Lease Liabilities, current | 39 / 59 | AAPL, ABBV, ADI, AMAT, AMD, AMZN, ARLO, AVGO … |
| Operating Lease ROU Assets | 34 / 59 | AAPL, ABBV, ADI, AMAT, AVGO, BAC, CAT, COHR … |
| Op. Lease Liabilities, LT | 29 / 59 | AAPL, ABBV, ADI, AMAT, AMZN, AVGO, CAT, CVX … |
| Amortization of Intangibles | 26 / 59 | ADBE, AMAT, ARLO, BAC, CAT, CRM, DDOG, FORM … |
| Accrued Compensation | 25 / 59 | AAPL, ABBV, ADBE, ADI, AMAT, AMD, AMZN, ARLO … |
| Finance Lease Liabilities, LT | 20 / 59 | AAPL, ABBV, AMAT, AMZN, ARLO, AVGO, COST, CRM … |
| Current Portion of LT Debt | 20 / 59 | AMAT, AMZN, CRM, CVX, DHR, GE, GOOGL, INTC … |
| Income Tax Payable | 16 / 59 | AAPL, AMAT, AMD, CRM, DDOG, DHR, FORM, LITE … |
| Net Income incl. NCI | 13 / 59 | ADI, BAC, CRM, DDOG, DHR, FORM, GS, JNJ … |
| Deferred Tax Liability, LT | 12 / 59 | AAPL, AMZN, AVGO, CAT, CRM, DDOG, DHR, ISRG … |
| Preferred Stock | 12 / 59 | AAPL, ABBV, BAC, GOOGL, JPM, MDT, MRK, MSFT … |
| Deferred Revenue, LT | 11 / 59 | AMZN, AVGO, CAT, COHR, DHR, LITE, NVDA, ON … |
| Total Equity incl. NCI | 10 / 59 | AMAT, CRM, FORM, JPM, LITE, MU, NKE, PANW … |
| Interest Expense | 8 / 59 | AAPL, GOOGL, INTC, LLY, LRCX, MRK, MSFT, PFE |
| Other Non-op Inc/(Exp) | 8 / 59 | AAPL, CVX, GOOGL, LLY, MSFT, NXPI, ORCL, QCOM |
| Pension & Retirement Oblig. | 8 / 59 | ABBV, ADI, AMAT, DHR, INTC, LITE, NXPI, ON |
| Long-term Investments | 8 / 59 | ADBE, COST, KLAC, MDT, MRVL, QCOM, SWKS, TXN |
| Other Current Liabilities | 8 / 59 | ADI, CRM, LRCX, META, NKE, PANW, PG, XOM |
| Short-term Investments | 8 / 59 | ARLO, DHR, GS, JPM, LRCX, MCD, SCHW, XOM |
| Deferred Revenue, current | 8 / 59 | AVGO, COHR, DHR, LITE, MRVL, MU, NVDA, ON |

**真缺口總計：488 個（列 × 公司）組合，分布在 67 個模板列。**

榜首那幾列全部是 TODO D10（只寫在附註、沒印在報表表面）——這是**已知的暫時性
限制**，不是新 bug。要壓低這個數字只有兩條路：接一條讀附註的路徑，或接受它。

### KPI 2 — 假警報：Index 標紅裡有幾個是誤判

標紅只有兩類：〔矛盾〕整列空白但相關欄位顯示該有、〔中間有洞〕。
「零星有值」刻意不標紅（H3-2），所以不算在內。

| | 家次 |
|---|---|
| 標紅：矛盾 | 52 |
| 標紅：中間有洞 | 217 |
| **標紅合計** | **269** |
| 降級為零星有值（不標紅） | 82 |

**要壓低的是標紅合計裡的誤判比例**，不是把標紅壓到 0——真缺口該標就要標。
驗證方式：對標紅的列抽樣，走 ARCHITECTURE「三步排查順序」確認是哪一類。

## 七、怎麼重跑

```
venv/Scripts/python.exe scripts/spike_derive_mapping.py    # 需要答案卷，慢
venv/Scripts/python.exe scripts/spike_verify_mapping.py    # 用快取，幾秒
```