# 交接：接下來三件事（H5 / D11 / H4 第二步）

**貼這份給新對話即可開工。** 2026-08-24 產出，接續 H3 與 H4 第一步。

---

## 建議順序與工作量

| 順位 | 項目 | 工作量 | 風險 | 效益 |
|---|---|---|---|---|
| **1** | **H5：`Interest Income` fallback 太窄** | ~1.5 小時（1 小時是機器跑） | 低 | **可能 +84 家，單列最大** |
| **2** | **D11-A：Index 顯示「本次抓取失敗幾份」** | ~1 小時 | 極低 | 讓「抓壞了」跟「資料本身缺」分得開 |
| **3a** | **量化 H4 第二步還能救多少**（第 3 項的前置） | ~30 分鐘 | 無（純分析） | 決定 3b 做不做 |
| **3b** | H4 第二步：數值指紋 | ~5~7 小時 | 中（動取值主流程） | **未知，看 3a 的結果** |

**先做 1 和 2。3a 是純分析、不動程式，做完才決定 3b 值不值得。**
第一步的 `label_fallback` 效果超過 spec 預估（NVDA Capex 季表 spec 估上限 53/57、
實際做到 67/68），所以 3b 的邊際效益可能已經變小——**不要跳過 3a 直接做 3b**。

---

## 動手前務必先讀

1. `docs/ARCHITECTURE.md` 的 **「edgartools 到底是什麼」**、**「Template Matching Logic」**、**「缺漏判斷」**
2. `docs/template-coverage-baseline-2026-08-24.md` 的**第零節**（怎麼讀這份文件、為什麼不該追求 97/97）
3. `docs/TODO.md` 的 H5、D11、H4

---

## ⚠ 三個共用的坑（每一輪都會遇到）

### 1. 比對腳本一定要用 `(列名, 第幾次出現)` 當鍵

`Net Income` 與 `SBC` 在 IS 和 CF 模板**各有一列**，`concepts` 清單裡同名出現兩次。
用列名當字典鍵只會留最後一個，變成拿 IS 那列去比 CF 那列——2026-08-24 這樣憑空
生出 **3,659 個假異動**，一度誤判成嚴重回歸。overflow 區的 `Other`／
`Accrued expenses` 這類重複列名更嚴重。

現成的正確版本在 **`scripts/diag_celldiff2.py`**（2026-08-24 收進 repo），
直接用它做回歸驗收，不要自己重寫一個。

### 2. 答案卷的抓取窗不一致，重建務必沿用

`AAPL/ADBE/AMD/AVGO/COST/GOOGL/INTC/META/MSFT/NVDA/TSLA/WMT` 這 **12 家是全部
filing**（44~69 期），其餘 189 家是 `max_filings=16`（約 21 期）。
用同一個參數重建全部，會讓那 12 家從 69 期縮到 20 期，逐列覆蓋率整片假性下降。

### 3. 覆蓋門檻要用比例不能用絕對家數

`gen_template_coverage_baseline.py` 的 `MIN_CO_RATIO = 0.85`。曾經寫死「≥45 家」，
樣本從 52 擴到 102 時等於門檻悄悄鬆到 44%，達標列數從 47 假性跳到 74。

---

## 順位 1：H5 — `Interest Income` 的 fallback 太窄

### 證據（2026-08-24 實測 201 家）

```
Interest Income   整列空白 137 / 201 家

那批公司在 companyfacts 實際 tag 的：
    InvestmentIncomeInterest                84 家
    InterestIncomeOther                     21 家
    InvestmentIncomeInterestAndDividend     14 家
    InterestIncomeExpenseNet                12 家
```

模板的 fallback 是 `InterestIncome`，而 **`InvestmentIncomeInterest` 不含這個
子字串**（字序相反），兩層都比不到。實測 KO 的損益表表面就是
`us-gaap_InvestmentIncomeInterest`，label 寫「Interest income」。

**這跟 H4 是不同問題**——這些都是標準 us-gaap concept，不是公司自訂延伸 tag。
修法跟 2026-08-23 修 `Debt Proceeds`／`Other Operating Expense` 完全同型。

### 要注意的一個判斷

`InterestIncomeExpenseNet` 是**淨額**（利息收入減支出），收進來可能跟
`Interest Expense` 那一列重複計算。**建議不收**，只收純收入的
`InvestmentIncomeInterest`／`InterestIncomeOther`／`InvestmentIncomeInterestAndDividend`。
要收的話先跟 CTH 確認。

### 順帶一起查

`Amortization of Intangibles` 真缺口 **100/201 家**，同一輪分析發現 overflow 區
常出現「Amortization of (acquired/purchased) intangible assets」變體。同一類線索，
一起查省一次重建。

### 步驟

1. 抽 10 家 probe，確認那些 concept**在損益表表面**不是只在附註（20 分）
2. TDD 改 fallback 正則（20 分）
3. 重建 201 家 + 逐格比對（1 小時，背景）
4. 全套測試 + 重產基線 + commit（20 分）

---

## 順位 2：D11-A — Index 要講清楚「本次抓取失敗幾份」

### 問題

**抓取結果會被當下的 SEC 狀況影響。** 2026-08-24 實測：連續抓 50 家時，
AXP／BMY／C／COP／ETN／HON／LLY／MRK／MS／PSX／WFC 的 `Operating Income`
只有 16/21 期；隔幾小時單獨重抓 LLY **兩次都是 20/21 期**。

根因：偶發失敗 → `_note_gap()` 記帳後跳過那份 filing → 少了年報就
`_synthesize_q4()` 合成不出 Q4 → **靜默少掉 4 格**。

**程式本身是決定性的**（同碼重抓 AAPL 逐格完全相同），不確定性來自網路。

### 為什麼會漏看

帳本有記，但 Index 的完成度只會顯示成「整欄稀疏」或「中間有洞」，看起來像
**資料本身缺**，不像**這次抓壞了**。使用者不會想到要重抓。

### 三個選項（CTH 尚未決定）

- **A（建議）**：Index 上把「本次抓取有幾份 filing 失敗」單獨顯示，跟資料缺漏分開。
  帳本已經有資料，只是沒顯示。**約 1 小時，風險極低**，四個語系要加 key
- **B**：抓完自動偵測帳本有缺漏就重試那幾份。**約 3~4 小時，風險中**，動到主流程
  控制流，重試次數與間隔要設計，測試要模擬失敗
- **C**：降低連續抓取速率。**約 30 分鐘**，但每次抓取都變慢

A 和 B 不衝突，A 先做不會白做。

### 待查證的候選名單

2026-08-24 抓新公司那輪，log 出現「自動修復：發現 N 期缺失需要修復重疊」的有
**AIG／ALL／BK／CB／COF／DD／DOW／HIG／MET／NEM／OXY** 共 11 家。
**還沒確認那則訊息是否等同於「抓壞了」**（也可能是正常的重疊修復路徑）。
確認方式：單獨重抓一家，看 `Operating Income` 會不會從 16/21 變 20/21。

---

## 順位 3：H4 第二步 — 動手前先量化

### 設計已經定案

完整規格：`docs/superpowers/specs/2026-08-23-concept-rename-linking-design.md`
（含五道保險、實作順序、驗收標準）。一句話：**同一個期末日、完全相同的金額，
出現在兩個不同的 concept 底下，就證明它們是同一條科目。**

證據是硬的，實測 NVDA：

```
期間 2023-01-29 (FY)
  FY2025 10-K   us-gaap_PaymentsToAcquireProductiveAssets                 -1,833,000,000
  FY2023 10-K   nvda_PurchasesOfPropertyAndEquipmentAndIntangibleAssets   -1,833,000,000
```

### 但邊際效益可能已經變小

第一步（label_fallback）的效果**超過 spec 預估**：

```
spec 原本預估   NVDA Capex 季表上限 53/57
第一步實際做到   67/68
```

當初設計第二步是假設「label 也會改、光靠 label 救不了」。實際上 NVDA 那個案例
label 夠穩定，第一步就解決了。**第二步真正還能多救多少，目前沒有數字。**

### ⚠ 動手前先做這個量化（約 30 分鐘）——這是第 3 項的前置任務，不是選配

**要回答的問題**：模板列的空格裡，有多少是「那一期其實有一個公司自訂延伸 tag
帶著值，只是進了 overflow 區」？

- 數字大 → 值得花 5~7 小時做第二步
- 數字小 → 降級成長期項目，時間花在 H5 那類對照缺口上

#### 為什麼不能純離線做

`output/_spike/` 的 pkl 有四個欄位：`labels`（**期間標籤**，如 `FY2026Q2`）、
`ends`、`concepts`、`values`。overflow 列在 `concepts` 裡放的是**公司原文標籤**
（如「Purchases related to property and equipment and intangible assets」），
**concept key 沒有存進去**——所以離線分不出 `nvda_Xxx` 與 `us-gaap_Xxx`。

`StatementTable` 本身有 `labels` 欄位存 concept key（`_build_*_table` 裡
`labels_g.append(key)`），只是重建腳本沒存它。

#### 建議做法：抽樣 20 家用網路 probe（約 15 分鐘）

挑 20 家（**務必含 NVDA／TSLA／GE／PG，它們已知有自訂 tag**），對每家最新的
5~8 份 filing：

1. 取 IS/BS/CF 的 `to_dataframe()`
2. 跑一次現行的模板比對，記下哪些模板列在那一期沒命中
3. 在同一份 dataframe 的**未被消化列**裡，數有多少 concept **不是 `us-gaap_`
   開頭**且有值
4. 統計「模板列空格」中有多少對得上這種延伸 tag

現成的腳本已經收進 repo（2026-08-24 從 scratchpad 移入，scratchpad 是 session
專屬的、換對話就沒了）：

| 腳本 | 用途 |
|---|---|
| `scripts/diag_probe.py` | 印出某家某張報表裡符合正則的列（最常用） |
| `scripts/diag_rowprobe.py` | 某個模板列在多家公司的命中情況 + 所有候選 |
| `scripts/diag_hintsweep.py` | 掃出 `label_hint` 太窄殺掉正確答案的列 |
| `scripts/diag_celldiff2.py` | 兩份答案卷快取逐格比對（回歸驗收用，鍵已處理好） |

#### 順手做掉，讓下次免費

**重建腳本存 pkl 時多存一個 `"row_labels": list(q.labels)`。** 之後這類分析就完全
離線、幾秒跑完，不用再打網路。改一行的事，但要注意 `gen_template_coverage_baseline.py`
讀 pkl 的地方不會壞（它只讀 `labels`／`ends`／`concepts`／`values`，多一個鍵不影響）。

### 已經試過、不要重走的路

**文字相似度自動配對**：102 家掃描，加上「模板列已有 ≥3 期」「overflow 補 ≥3 期」
「完全互補」三個條件後仍有 6,910 組候選，強候選 458 組，**實測約一半是誤判**
（`ADBE Deferred Revenue, LT ← 「Total revenue」`）。而且 **NVDA 那個案例根本
偵測不到**——相似度是拿模板列名「Capex」比公司原文「Purchases of property and
equipment」，字面 0% 重疊。

**companyfacts 補值**：102 家實測，`facts` 底下只有 `us-gaap`／`dei`／`srt`／
`ffd`／`ecd`／`invest` 這些 SEC 標準 taxonomy，**沒有任何 `nvda_` 這種延伸 tag**。
第二資料源救不了這題，而且基線的〔真缺口〕KPI **低估了**這一類。

---

## 現況（2026-08-24）

```
分支          fix/period-alignment-and-companyfacts-spike（尚未併回 master，已 push）
測試          1172 passed / 7 skipped / 0 failed（含真連 SEC，2026-08-24）
驗證樣本      201 家，快取在 output/_spike/
基線          docs/template-coverage-baseline-2026-08-24.md
```

三次量測的結論高度穩定：

| | 52 家 | 102 家 | 201 家 |
|---|---|---|---|
| 達標列數 | 47/97 | 46/97 | 44/97 |
| 我們抓到／真缺口／公司真的沒有 | — | 73/9/18% | 72/9/18% |

---

## 不要順手做的事

- **不要改符號慣例**（CTH 2026-08-22：一律照公司原始申報，不做正規化）
- **不要動 `Data_Financials` 的列位或列名**（下游靠固定列位取值）
- **不要動 `fetcher_facts.py` / `facts_mapping.py`**（G11 已決議不切換，保留當第二資料源）
- **不要重開 G11**（2026-08-23 結案：segments 要 20 年 → 混合架構只快 1.9 倍；
  H3 之後重驗 83.96%／95.17%，比之前更低）
- **`Short-term Debt` 不用修**（已查證：64 個洞只有 11 個是真的漏抓，其餘是公司
  當季真的沒有短期借款）
- 發現新問題記進 `docs/TODO.md` 再問 CTH，不要擴大範圍
