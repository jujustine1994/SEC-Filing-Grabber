# 交接：H3 系統性 concept 對照問題

**貼這份給新對話即可開工。** 目標是把模板列的覆蓋率從 40/97 往上推。

---

## 任務

`docs/TODO.md` 的 **H3**：52 家實測發現，模板列抓不到的問題不是散落的個案，
而是同一個 concept 對照問題在多家公司重複出現。**修一次會讓很多公司一起變好。**

現況基線（`docs/template-coverage-baseline-2026-08-23.md`）：

```
模板列 97 個，達到「≥45 家有值且填滿率 >90%」的只有 40 個

最常被判「矛盾」（整列全空、但同一家公司的相關欄位顯示它應該要有）
    Current Portion of LT Debt        25 / 52 家
    Op. Lease Liabilities, current    14 / 52 家
    Change in Inventories             13 / 52 家
    Debt Proceeds                     11 / 52 家
    Share Repurchases                  9 / 52 家

最常「中間有洞」（同一列有些期有、有些期沒有——一定是漏抓）
    Shares Outstanding                43 / 52 家
    Acquisitions                      24 / 52 家
    Debt Proceeds                     24 / 52 家
    Debt Repayments                   16 / 52 家
    Short-term Debt                   15 / 52 家
```

`Current Portion of LT Debt` 25 家中招——不可能這麼多公司剛好都沒有一年內到期
負債，幾乎確定是 concept 對照有問題。

---

## 動手前務必先讀

1. `docs/ARCHITECTURE.md` 的 **「缺漏判斷」** 與 **「期間標籤與日曆季」** 兩節
2. `docs/template-coverage-baseline-2026-08-23.md`（完整基線，含逐列覆蓋率）
3. `src/data_quality.py` 的模組說明（四個判斷各自的細節與已知陷阱）

---

## ⚠ 最重要的一個坑（前一輪踩過兩次）

**不要看到「抓不到」就假設是 concept 名字錯。**

前一輪對 `Accrued Compensation` 那批七列，第一次判斷「名字錯」、第二次判斷
「在附註不在報表表面」，**兩次都不精確**。實際查證的結果是：

- 模板的 `fallback` 欄名字**本來就是對的**（`EmployeeRelatedLiabilitiesCurrent`
  等），改名字救不了
- AAPL 那份 10-Q 的**全部 784 個 fact 裡根本沒有這個 concept**——那份 filing
  沒 tag。但 52 家裡有 31 家的 10-Q 有 tag，所以不是「這個科目不存在」

**正確的排查順序**（每一列都要走完，不要跳）：

```
1. 這份 filing 的 XBRL 到底有沒有 tag 這個東西？
      filing.xbrl().facts.to_dataframe()  → 搜 concept 欄
   沒有 → 這家這期就是沒有，不是我們的問題，記錄後跳過

2. 有 tag，但報表 dataframe（income_statement()/balance_sheet()/
   cashflow_statement()）裡沒有？
      → 公司放在附註、presentation linkbase 沒收進那三張表
      → 現行路徑結構上拿不到，記進 TODO，等 G11 決策

3. 報表 dataframe 裡有，但 _match_is_row() 沒命中？
      → 這才是真正可以修的。看 standard_concept / concept / label 三欄，
        對照模板的 std_concept / fallback / label_hint
```

第 3 類已知的兩種成因：

- **edgartools 的 `standard_concept` 標錯**：實測 AMD/MRVL 的
  `Depreciation and amortization` 被標成 `NonoperatingIncomeExpense`。
  `_match_is_row()` 優先比對 std_concept 就漏掉，要靠 `label_hint` 退路
- **`label_hint` 太窄把整個優先層濾掉**：`Deferred Revenue, current` 的 hint 是
  `'unearned revenue'`，但多數公司寫 `Deferred revenue`。`_pick()` 濾空之後
  **整個優先層被跳過**，不會退回去用 concept 比對

---

## 工作方式

**TDD**（專案規則）。每一列的修改都要有測試釘住，測試裡寫清楚是哪一家公司、
哪一份 filing 的哪個 concept。

**每修一批就重跑基線**，看數字有沒有往上走：

```
venv/Scripts/python.exe scripts/gen_template_coverage_baseline.py
```

它不打網路，吃 `output/_spike/` 的 52 家快取，幾秒跑完。**但注意**：那些快取是
用舊的抓取邏輯產生的答案卷，改了 concept 對照之後要重新抓才會反映改善——
先用 `scripts/spike_derive_mapping.py` 重建幾家的快取再看。

**收尾**：跑全套 `venv/Scripts/python.exe -m pytest tests/ -q`（約 20 分鐘，
會真連 SEC），更新 `docs/CHANGELOG.md` 與 `docs/TODO.md`（做完的條目直接從
TODO 刪掉搬進 CHANGELOG，這是本專案的維護規則）。

---

## 建議的下手順序

1. **`Current Portion of LT Debt`**（25 家）——中招最多，先做這個練排查流程
2. **`Shares Outstanding`**（43 家有洞）——但注意有一部分是 D0-2 已知限制
   （多股別公司 Class A/B/C 分開標），要先分清楚哪些是可修的
3. **`Op. Lease Liabilities, current`**（14 家）與 **`Change in Inventories`**（13 家）
4. `Debt Proceeds` / `Debt Repayments` / `Acquisitions` / `Short-term Debt`

---

## 不要順手做的事

- **不要改符號慣例**。CTH 2026-08-22 決定「一律照公司原始申報，不做正規化」
- **不要動 `Data_Financials` 的列位或列名**。下游（`financial-assistant` 的
  `read_excel.py`、使用者自己的 Excel 公式）靠固定列位取值
- **不要碰 G11（companyfacts 切換）**，那是另一個決策中的項目
- 發現新問題記進 `docs/TODO.md` 再問 CTH，不要擴大範圍

---

## 環境

```
venv/Scripts/python.exe          專案的 Python
SEC Identity 從 config 讀        config.load_config()["identity"]
output/_spike/                   52 家的快取（facts JSON + 答案卷 pkl）
```

分支：`fix/period-alignment-and-companyfacts-spike`（尚未併回 master）
