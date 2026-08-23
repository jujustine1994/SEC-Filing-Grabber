# 交接：H4 第二步——用數值指紋自動接回改名的 concept

**貼這份給新對話即可開工。** 第一步（模板 tuple 第七欄）已完成並驗證，這輪做第二步。

---

## 任務

實作 `docs/superpowers/specs/2026-08-23-concept-rename-linking-design.md` 的**階段一到三**。

一句話：**同一個期末日、完全相同的金額，出現在兩個不同的 concept 底下，就證明
它們是同一條科目。** 用這個把公司改過名的 concept 自動接回模板列。

---

## 動手前務必先讀

1. **`docs/superpowers/specs/2026-08-23-concept-rename-linking-design.md`** —— 完整設計，含五道保險、實作順序、驗收標準。**這份是這輪的規格書，照著做。**
2. `docs/ARCHITECTURE.md` 的 **「edgartools 到底是什麼」**、**「Template Matching Logic」**、**「缺漏判斷」** 三節
3. `docs/TODO.md` 的 **H4**

---

## 已經完成的（第一步，2026-08-23）

模板 tuple 從 6 欄變 **7 欄**，第七欄是 `label_fallback`，接到 `_match_is_row()`
本來就有、但一直沒人餵的第三層（label 比對）。

```
_T = tuple[str, str | None, str, str, str, str | None, str | None]
                                                        ^^^^^^^^^^ 新增
```

- 97 列全部補了第七欄，目前只有 `Capex` 真的填了正則：`r"^purchases (?:of|related to).*propert"`
- 四個呼叫點（IS／IS-from-CF／BS／CF）都已接線
- 非 live 測試 **1114 passed**

**為什麼只填 Capex**：那是唯一有實證的案例（NVDA）。其他列要填之前先拿資料證明，
不要憑印象加——第三層很寬，後面沒有任何東西再擋它。

---

## ⚠ 這輪最大的風險

階段一要**讀每份 filing 的所有期間欄**，但現行邏輯刻意每份只讀一欄
（`_current_q_col()` / `bs_col` / `data_col`），而且有 `if label in periods: continue`
這個 dedup 防止同一期被兩份 filing 蓋來蓋去。

**控制方式（spec 第 5 節已寫死）**：蒐證與取值**完全分離**。階段一只往索引裡寫，
不碰 `row_vals`；真正的取值路徑一行都不改。所以「沒有任何連結成立」時輸出必須
逐格相同——這是驗收第一條。

---

## 三個已經驗證過的事實（不用再查）

**1. companyfacts 沒有公司自訂 tag。**
102 家實測，`facts` 底下只有 `us-gaap`／`dei`／`srt`／`ffd`／`ecd`／`invest` 這些
SEC 標準 taxonomy，**沒有任何 `nvda_` 這種**。所以：
- 第二資料源救不了這題
- 基線的〔真缺口〕KPI **低估了**這一類（那些格子被歸進「公司真的沒有」）

**2. 文字相似度那條路試過了，不要重走。**
102 家掃描，加上「模板列已有 ≥3 期」「overflow 補 ≥3 期」「完全互補」三個條件後
仍有 6,910 組候選，其中「名稱高度相似」的強候選 458 組，**實測約一半是誤判**
（`ADBE Deferred Revenue, LT ← 「Total revenue」`）。而且 **NVDA 那個案例根本
偵測不到**——相似度是拿模板列名「Capex」比公司原文「Purchases of property and
equipment」，字面 0% 重疊。

**3. 數值指紋這條路的證據是硬的。** 實測 NVDA：

```
期間 2023-01-29 (FY)
  FY2025 10-K   us-gaap_PaymentsToAcquireProductiveAssets                 -1,833,000,000
  FY2023 10-K   nvda_PurchasesOfPropertyAndEquipmentAndIntangibleAssets   -1,833,000,000
```

`2022-01-30 (FY)` 也同樣對得上（−976,000,000），滿足 spec 的「至少 2 個期間佐證」。

---

## NVDA 的完整實測資料（第一批測試就用這個）

10-K 的 Capex tag 逐份確認：

| 10-K 申報日 | concept | 命中？ |
|---|---|---|
| 2026-02、2025-02、2024-02 | `us-gaap_PaymentsToAcquireProductiveAssets` | ✅ |
| 2023-02 … 2013-03（11 份） | `nvda_PurchasesOfPropertyAndEquipmentAndIntangibleAssets` | ❌ |
| 2012-03 | `us-gaap_PaymentsToAcquirePropertyPlantAndEquipment` | ✅ |
| 2011-03 | `nvda_PurchasesOfPropertyAndEquipmentAndIntangibleAssets` | ❌ |

已知金額（可直接寫進測試）：

```
2023-01-29 (FY)  -1,833,000,000
2022-01-30 (FY)    -976,000,000
2021-01-31 (FY)  -1,128,000,000
2020-01-26 (FY)    -489,000,000
```

第一步做完之後，NVDA 的 Capex 應該已經改善（label 對得上）。**這輪開工前先重跑
一次 NVDA 確認現況**，不要沿用舊數字：

```
venv/Scripts/python.exe src/cli.py gaap NVDA --max-filings 40 --xlsx <某個暫存路徑>
```

改善前的基準（第一步之前）：季表 `Capex` 36/57 期、年表 17 年只有 4 年有值。

---

## 工作方式

**TDD**（專案規則）。每個修改的測試要釘住「哪一家公司、哪個 concept、哪個期間、
哪個金額」，字串從真實 `to_dataframe()` 逐字抄，不要編。

**驗收（spec 第 6 節）**：

1. **回歸最重要**：102 家逐格對照，**沒有任何一格從「有值」變成「不同的值」**
2. NVDA 年表 Capex 從 4 年補到 14 年以上；季表上限 53/57（最舊 4 個 Q4 缺的是
   合成材料不是 tag，不在這輪範圍）
3. 102 家統計：補上幾格、成立幾條連結、因衝突放棄幾條
4. 抽樣 20 條連結回原始 filing 確認
5. 全套測試（含 live，約 11 分鐘）+ 重產基線

**收尾**：更新 `docs/CHANGELOG.md` 與 `docs/TODO.md`（做完的條目直接從 TODO 刪掉
搬進 CHANGELOG，這是本專案的維護規則），commit + push。

---

## 環境與快取

```
venv/Scripts/python.exe                  專案的 Python
config.load_config()["identity"]         SEC Identity
output/_spike/                           102 家快取（facts JSON + 答案卷 pkl）
scripts/gen_template_coverage_baseline.py  重產基線，不打網路、幾秒
```

**⚠ 答案卷的抓取窗不一致，重建時務必沿用**：
`AAPL/ADBE/AMD/AVGO/COST/GOOGL/INTC/META/MSFT/NVDA/TSLA/WMT` 這 12 家是全部
filing（44~69 期），其餘 90 家是 `max_filings=16`（約 21 期）。
**用同一個參數重建全部會讓那 12 家從 69 期縮到 20 期，逐列覆蓋率整片假性下降**
——2026-08-23 踩過，白跑一輪。

---

## 不要順手做的事

- **不要改符號慣例**。CTH 2026-08-22 決定「一律照公司原始申報，不做正規化」
- **不要動 `Data_Financials` 的列位或列名**。下游（`financial-assistant` 的
  `read_excel.py`、使用者自己的 Excel 公式）靠固定列位取值
- **不要動 `fetcher_facts.py` / `facts_mapping.py`**。G11 已決議不切換，那條平行
  路徑保留當第二個獨立資料來源與交叉驗證工具
- **不要重開 G11**。2026-08-23 已結案：segments 要 20 年份 → 混合架構只快 1.9 倍
  （實測解 filing 佔 54% 時間，那步 segments 非做不可）；且 H3 之後重驗
  83.96%／95.17%，比之前更低（現行路徑變好了，差距反而拉開）
- 發現新問題記進 `docs/TODO.md` 再問 CTH，不要擴大範圍

---

## 這輪之後還剩什麼

- **`Data_Segments` 覆蓋率基線**（TODO，CTH 標「不急切，有空再做」）。完全沒量測過，
  而且公司會改營業項目申報，比三表複雜
- **KPI 2 誤判率抽樣**（CTH 標「之後再做」）。基線第六節有標紅總數，但真假比例未知
- **`Short-term Debt`**：已查證，64 個洞裡只有 11 個（17%）是真的漏抓，其餘是公司
  當季真的沒有短期借款。**不是 bug，不用修**

分支：`fix/period-alignment-and-companyfacts-spike`（尚未併回 master）
