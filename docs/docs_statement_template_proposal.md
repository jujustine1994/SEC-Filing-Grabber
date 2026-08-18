# 三表固定模板提案（依 49 家實測資料）

> 產生日期：2026-08-03　　資料來源：`scripts/survey_statement_concepts.py`（純 EDGAR，零 API 額度）
>
> ⚠ **歷史提案文件，模板行數已過時**：BS 後續於 2026-08-12 新增
> Total Non-current Assets/Liabilities，現行行數見 `docs/ARCHITECTURE.md`
> 「Template 行數摘要」（IS 22／BS 44／CF 26）。本檔保留供回溯當初取捨依據。

## 方法

抓 50 家公司最新 10-Q 的 XBRL 三表，統計每個科目的**跨公司出現率**。成功 49 家（HON 只有 10-Q/A 無完整 XBRL）。金融股（JPM／BAC／GS）三表結構本就不同，**單獨分組**，下方百分比一律以**非金融 46 家**為分母。

樣本涵蓋：大型科技 9、軟體 SaaS 5、半導體 7、中小型硬體 6、消費 6、工業 3、能源公用 3、醫療 4、電信 REIT 3、金融 3。

排除 abstract 標題列與帶 dimension 的明細列（後者屬 segment，不是三表本體）。

**判準**（你定的）：多數公司都有的 → 進固定模板，某些公司沒有就留空白列；只有極少數特例才有的 → 落 overflow。

---

## 結論摘要

現行模板體質比預期好。三表共 90 列，實測命中率高，**漏收的高覆蓋率科目只有 6 個**。

| | 模板列數 | 實際出現科目數 | 建議新增 | 建議檢討 | 每家 overflow 中位數 |
|---|---|---|---|---|---|
| 損益表 Income Statement | 22 | 213 | 2 | 5 | 4 |
| 資產負債表 Balance Sheet | 42 | 281 | 1 | 5 | 2 |
| 現金流量表 Cash Flow | 26 | 451 | 3 | 1 | 10 |

**overflow 規模很小**：每家中位數 IS 4 / BS 2 / CF 10 個科目，合計約 16 列。放到 sheet 底部完全可接受，不會干擾閱讀，也不會再把 BS/CF 的列位往下推。

---

## 損益表 Income Statement

### 現行模板逐列命中率

| 命中率 | 列名 | 判定 |
|---:|---|---|
| 98% | Revenue | 核心，保留 |
| 91% | Cost of Revenue | 核心，保留 |
| 50% | Gross Profit | 常見，保留 |
| 74% | R&D Expense | 核心，保留 |
| 100% | SG&A Expense | 核心，保留 |
| 15% | D&A (CF memo) | 從 CF 表取值（memo 列），此處百分比不適用 |
| 0% | Other Operating Expense | **建議檢討** |
| 48% | Total Operating Expense | 常見，保留 |
| 83% | Operating Income | 核心，保留 |
| 74% | Interest Expense | 核心，保留 |
| 9% | Interest Income | **建議檢討** |
| 4% | Other Non-op Inc/(Exp) | **建議檢討** |
| 80% | Total Non-op Income/(Loss) | 核心，保留 |
| 91% | Pre-tax Income | 核心，保留 |
| 100% | Income Tax | 核心，保留 |
| 96% | Net Income | 核心，保留 |
| 41% | Minority Interest | 常見，保留 |
| 0% | SBC | 從 CF 表取值（memo 列），此處百分比不適用 |
| 100% | Basic EPS | 核心，保留 |
| 100% | Diluted EPS | 核心，保留 |
| 85% | Basic Shares | 核心，保留 |
| 85% | Diluted Shares | 核心，保留 |

### 模板沒收、覆蓋率 ≥25% 的科目

| 覆蓋率 | XBRL 概念 | 標籤 |
|---:|---|---|
| 46% | `us-gaap_ProfitLoss` | Net income |
| 33% | `us-gaap_CostsAndExpenses` | Total costs and expenses |

## 資產負債表 Balance Sheet

### 現行模板逐列命中率

| 命中率 | 列名 | 判定 |
|---:|---|---|
| 93% | Cash | 核心，保留 |
| 54% | Short-term Investments | 常見，保留 |
| 96% | Accounts Receivable | 核心，保留 |
| 80% | Inventories | 核心，保留 |
| 98% | Other Current Assets | 核心，保留 |
| 96% | Total Current Assets | 核心，保留 |
| 96% | PP&E, net | 核心，保留 |
| 50% | Operating Lease ROU Assets | 常見，保留 |
| 30% | Long-term Investments | 偏低但屬正常差異，保留 |
| 87% | Goodwill | 核心，保留 |
| 78% | Intangible Assets, net | 核心，保留 |
| 48% | Deferred Tax Assets | 常見，保留 |
| 93% | Other Non-current Assets | 核心，保留 |
| 100% | Total Assets | 核心，保留 |
| 96% | Accounts Payable | 核心，保留 |
| 57% | Short-term Debt | 常見，保留 |
| 41% | Current Portion of LT Debt | 常見，保留 |
| 28% | Op. Lease Liabilities, current | 偏低但屬正常差異，保留 |
| 37% | Accrued Compensation | 偏低但屬正常差異，保留 |
| 70% | Deferred Revenue, current | 常見，保留 |
| 26% | Income Tax Payable | 偏低但屬正常差異，保留 |
| 46% | Other Current Liabilities | 常見，保留 |
| 100% | Total Current Liabilities | 核心，保留 |
| 93% | Long-term Debt | 核心，保留 |
| 50% | Op. Lease Liabilities, LT | 常見，保留 |
| 0% | Finance Lease Liabilities, LT | **建議檢討** |
| 17% | Deferred Revenue, LT | **建議檢討** |
| 63% | Deferred Tax Liability, LT | 常見，保留 |
| 17% | Pension & Retirement Oblig. | **建議檢討** |
| 87% | Other Non-current Liabilities | 核心，保留 |
| 83% | Total Liabilities | 核心，保留 |
| 46% | Preferred Stock | 常見，保留 |
| 100% | Common Stock & APIC | 核心，保留 |
| 72% | Additional Paid-in Capital | 核心，保留 |
| 13% | Treasury Stock | **建議檢討** |
| 100% | Retained Earnings | 核心，保留 |
| 100% | AOCI | 核心，保留 |
| 87% | Total Equity — Parent | 核心，保留 |
| 43% | Noncontrolling Interests | 常見，保留 |
| 46% | Total Equity incl. NCI | 常見，保留 |
| 100% | Total Liabilities & Equity | 核心，保留 |
| 4% | Shares Outstanding | **建議檢討** |

### 模板沒收、覆蓋率 ≥25% 的科目

| 覆蓋率 | XBRL 概念 | 標籤 |
|---:|---|---|
| 70% | `us-gaap_CommitmentsAndContingencies` | Commitments and contingencies |

## 現金流量表 Cash Flow

### 現行模板逐列命中率

| 命中率 | 列名 | 判定 |
|---:|---|---|
| 98% | Net Income | 核心，保留 |
| 91% | D&A | 核心，保留 |
| 85% | SBC | 核心，保留 |
| 17% | Amortization of Intangibles | **建議檢討** |
| 76% | Change in Receivables | 核心，保留 |
| 59% | Change in Inventories | 常見，保留 |
| 37% | Change in Deferred Revenue | 偏低但屬正常差異，保留 |
| 41% | Other Working Capital | 常見，保留 |
| 50% | Other Non-cash Items | 常見，保留 |
| 100% | Operating Cash Flow | 核心，保留 |
| 98% | Capex | 核心，保留 |
| 46% | Acquisitions | 常見，保留 |
| 54% | Investment Purchases | 常見，保留 |
| 50% | Investment Proceeds | 常見，保留 |
| 100% | Investing Cash Flow | 核心，保留 |
| 63% | Debt Proceeds | 常見，保留 |
| 74% | Debt Repayments | 核心，保留 |
| 78% | Share Repurchases | 核心，保留 |
| 39% | Dividends Paid | 偏低但屬正常差異，保留 |
| 100% | Financing Cash Flow | 核心，保留 |
| 74% | FX Effect on Cash | 核心，保留 |
| 100% | Net Change in Cash | 核心，保留 |
| 83% | Ending Cash | 核心，保留 |
| 76% | Cash Taxes Paid | 核心，保留 |
| 41% | Cash Interest Paid | 常見，保留 |
| — | Free Cash Flow | 衍生計算，不比對 XBRL |

### 模板沒收、覆蓋率 ≥25% 的科目

| 覆蓋率 | XBRL 概念 | 標籤 |
|---:|---|---|
| 52% | `us-gaap_IncreaseDecreaseInAccountsPayable` | Accounts payable |
| 35% | `us-gaap_IncreaseDecreaseInPrepaidDeferredExpenseAndOtherAssets` | Prepaid expenses and other assets |
| 26% | `us-gaap_IncreaseDecreaseInOtherOperatingAssets` | Other current and non-current assets |

---

## 金融股（JPM / BAC / GS）

三家的科目數遠高於一般公司（BS 分別 168 / 244 / 110 個科目，一般公司中位數約 50）。現行模板對它們的命中率：

| 表 | 模板命中列數（3 家皆有） | 說明 |
|---|---|---|
| 損益表 Income Statement | 10 / 22 | — |
| 資產負債表 Balance Sheet | 12 / 42 | — |
| 現金流量表 Cash Flow | 13 / 25 | — |

金融股沒有存貨、沒有毛利、負債結構完全不同（存款、附買回、交易性負債）。**建議維持現況：用同一套模板 + overflow 承接**，不另開金融股模板。理由是三家就有三種樣貌（商業銀行 vs 投行），另開模板等於又一個追不完的特例；overflow 已經能完整承接。

---

## overflow 該怎麼放

1. **移到 sheet 最底部**，不要插在 IS／BS／CF 之間。這是目前 `Cash` 在第 28～56 列之間浮動的唯一原因。

2. 用一列分隔標題 `Other (as reported)` 隔開，A 欄放 XBRL 原始 label、B 欄放 concept。

3. overflow 不進機器鍵體系（不給 `IS.xxx` 這種鍵），因為它們本來就是公司特有、跨公司不可比。
