# 設計：用數值指紋自動接回改名的 XBRL concept（TODO H4）

**狀態**：設計待實作
**日期**：2026-08-23
**起因**：端到端實跑 NVDA 的 Excel 發現 `Capex` 57 期只有 36 期有值

---

## 一句話

**同一個期末日、完全相同的金額，出現在兩個不同的 concept 底下，就證明它們是同一條科目。**
用這個當連結，把公司改過名的 concept 自動接回模板列——不靠文字相似度、不靠人工對照表。

---

## 1. 問題

### 症狀

NVDA 的 `Capex`：

```
季表   57 期只有 36 期有值
年表   17 年只有 2012、2024、2025、2026 四年有值
連帶   FY2014~FY2023 每一年的 Q4 都空（Q4 = 年報 − Q1 − Q2 − Q3，年報那格空就算不出來）
       Free Cash Flow 跟著一起空（OCF − Capex）
```

### 根因

NVDA 的 10-K 用**自己的延伸 tag**，實測逐份確認：

| 10-K 申報日 | concept | 命中？ |
|---|---|---|
| 2026-02、2025-02、2024-02 | `us-gaap_PaymentsToAcquireProductiveAssets` | ✅ |
| 2023-02 … 2013-03（11 份） | `nvda_PurchasesOfPropertyAndEquipmentAndIntangibleAssets` | ❌ |
| 2012-03 | `us-gaap_PaymentsToAcquirePropertyPlantAndEquipment` | ✅ |
| 2011-03 | `nvda_PurchasesOfPropertyAndEquipmentAndIntangibleAssets` | ❌ |

`_match_is_row()` 前兩層都比對 concept 名稱，**延伸 tag 的名字每家公司自己取**，比不到。

### 為什麼現有的招都救不了

| 招 | 為什麼不行 |
|---|---|
| 把延伸 tag 寫進 `fallback` 正則 | 每家公司名字不同，寫不完 |
| companyfacts 交叉補值 | **companyfacts 根本沒有延伸 tag**。102 家實測，`facts` 底下只有 `us-gaap`／`dei`／`srt`／`ffd`／`ecd`／`invest` 這些 SEC 標準 taxonomy，沒有任何 `nvda_` 這種。這也代表基線的〔真缺口〕KPI **低估了**這一類 |
| 放寬 `label_hint` | hint 是過濾**已經命中的候選**，concept 比不到就根本沒有候選可濾。今天把 Capex 的 hint 從 `property` 放寬到 `propert\|capital expenditure` 之後，NVDA 依然全空 |
| 文字相似度自動配對 | 實測 102 家，強候選 458 組裡約一半是誤判（`ADBE Deferred Revenue, LT ← 「Total revenue」`）。而且 NVDA 這個案例**根本偵測不到**——相似度是拿模板列名「Capex」去比公司原文「Purchases of property and equipment」，字面 0% 重疊 |

### 資料其實沒有掉

`nvda_PurchasesOfPropertyAndEquipmentAndIntangibleAssets` 現在落在 `Other (as reported)`
overflow 區。**問題是進錯列，不是抓不到。**

---

## 2. 關鍵發現：比較欄提供了「同一期、兩個 tag」的重疊

一份 10-K 的現金流量表有 3 個年度欄，**三個年度共用同一個 concept**。所以新舊 tag 在
交界處會對同一個期間各報一次：

```
期間 2023-01-29 (FY)

  FY2025 10-K（2025-02-26 申報）
      us-gaap_PaymentsToAcquireProductiveAssets            -1,833,000,000

  FY2023 10-K（2023-02-24 申報）
      nvda_PurchasesOfPropertyAndEquipmentAndIntangibleAssets  -1,833,000,000
```

**同一個期末日、完全相同的金額。** 這就是連結的證據。

現行程式**每份 filing 只讀當期那一欄**（`_current_q_col()` / `bs_col` / `data_col`），
其餘 2~3 個比較欄直接丟掉——證據一直都在 dataframe 裡，只是沒人看。
讀它們**不需要額外網路請求**。

---

## 3. 設計

### 三個階段

```
階段一  蒐證      每份 filing 的每一個期間欄都掃過，記下兩份索引
階段二  建立連結  兩份索引的交集 = (延伸 concept → 模板列) 的對照
階段三  補值      只填空格，永不覆蓋已有的值
```

### 階段一：蒐證（`_collect_fingerprints`）

處理每一份 filing 時，對**所有**期間欄（不只當期那欄）建立兩份索引：

```python
# 已命中的：這個模板列在這個期間拿到什麼數字
matched[(row_index, period_key)] = value

# 沒命中的：這個 concept 在這個期間報了什麼數字
unmatched[(concept, period_key)] = value
```

`period_key` 用「期末日 + 期間型態」，**不能只用期末日**——同一個期末日同時有
`2026-06-30 (Q2)` 與 `2026-06-30 (YTD)` 兩欄，數字完全不同，混在一起會亂配。
型態取欄名括號裡的 `Q1`~`Q4`／`YTD`／`FY`；**資產負債表的欄名是裸日期沒有括號**
（`2026-06-30`），型態一律記成 `INSTANT`。

**數值比對用完全相等**（轉成 `float` 之後 `==`）。edgartools 的 `to_dataframe()`
已經把 XBRL 的 `decimals`／`scale` 還原成實際金額，兩份 filing 報同一個事實會得到
同一個浮點數。**不做容差比對**——容差會把「相近但不同的科目」誤配，而這裡寧可漏
不可錯。若實作時發現同一事實在兩份 filing 出現極小差異（例如四捨五入位數不同），
那代表其中一份重編過，本來就不該連結。

### 階段二：建立連結（`_derive_concept_links`）

```
對每一個 (concept C, period_key P, value V) in unmatched：
    若存在 (row_index R, P) in matched 且 matched[(R, P)] == V：
        候選連結 C → R
```

**四道保險**（每一道都是實測踩出來或可預見的誤判來源）：

1. **跳過 0 與 None**。`0.0` 在一張現金流量表裡會出現很多次，配對毫無資訊量
2. **同一個 (P, V) 對到多個模板列 → 整組丟掉**。例如某期 `Net Income` 與
   `Net Income incl. NCI` 剛好相等（沒有少數股權的公司天天如此），無從分辨
3. **同一個 concept 被連到 2 個以上不同模板列 → 那個 concept 整個放棄**。寧可不補
4. **連結至少要有 2 個獨立期間佐證**。單一期間的巧合擋不住，兩個期間同時
   吻合的機率極低
5. **同一個模板列被 2 個以上 concept 連到時**（合法情況：一列歷史上換過兩次
   tag），補值取**佐證期間數最多**的那個；仍然並列就取 concept 名稱字典序，
   讓結果可重現。**兩個連結對同一個空格給出不同數字時整格放棄**，並記進
   `Data_Meta` 的衝突清單

第 4 點是精確度的主要來源。NVDA 的案例遠超過門檻——FY2025 10-K 與 FY2023 10-K
在 `2023-01-29` 與 `2022-01-30` 兩個期間都對得上。

### 階段三：補值（`_fill_from_links`）

```
對每一個模板列 R 的每一個空格 (R, period_label)：
    找 links 裡連到 R 的 concept C
    若 unmatched[(C, 對應的 period_key)] 有值 → 填進去，並標記來源
```

**只填空格。** 已經有值的格子完全不動，所以這個機制不可能改變任何現有正確的數字
——最壞情況是沒補到，不會弄壞。

### 生命週期

連結表**只活在一次抓取內**，跟 `_parse_cache_scope()` 同一個範圍、同樣的理由：
跨 ticker 殘留會拿到別家公司的連結，跨執行殘留會用到過期的對照。
**不落地存檔**，每次抓取重新從資料推導。

### 稽核軌跡

補進來的格子，C 欄（公司原文標籤）寫該延伸 concept 的 label，並在 `Data_Meta`
記一行「本次自動接回的 concept 對照」，列出 `concept → 模板列`、佐證期間數。
**不能讓補進來的數字看起來跟直接讀到的一樣**——這個專案的價值就在每一格查得到來源。

---

## 4. 這個設計救不了什麼（誠實的限制）

| 情況 | 為什麼救不了 |
|---|---|
| 公司從頭到尾只用延伸 tag，從沒用過 us-gaap 版本 | 沒有交界，建立不了連結。例如 TSLA 的 `tsla_LongTermDebtAndFinanceLeasesCurrent` |
| 抓取範圍沒涵蓋到新舊 tag 的交界 | 學不到關聯。但那種情況下舊期間本來就不在範圍內，實務上不影響 |
| 改 tag 的同時數字也重編過 | 數值對不上，連結不成立。這是**刻意的**——數字都不一樣了，本來就不該當成同一條 |

第一種要靠**模板 tuple 第七欄 `label_fallback`**（`_match_is_row()` 的第三層已經
支援，只是模板餵不進去）兜底。那是獨立的一小步，跟本設計互補：

- 本設計處理「**改過名**」——全自動、零維護
- `label_fallback` 處理「**從頭到尾都自訂**」——要人寫正則，但只需針對少數列

**兩者的順序**：`label_fallback` 是第三層比對，發生在階段一之前；數值連結是階段
二三，只處理第三層之後仍然空白的格子。

---

## 5. 動到哪些程式

| 檔案 | 改動 |
|---|---|
| `src/fetcher_gaap.py` | 新增 `_collect_fingerprints` / `_derive_concept_links` / `_fill_from_links`；三個 `_build_*_table` 各接一次蒐證與補值；模板 `_T` 加第七欄 `label_fallback`（97 列都要補一個 `None`） |
| `tests/test_fetcher_gaap.py` | 新增測試，釘住 NVDA 的真實 concept／期間／金額，以及四道保險各一個反向測試 |
| `docs/ARCHITECTURE.md` | 新增一節說明數值指紋連結；`_match_is_row` 那節補第七欄 |
| `scripts/gen_template_coverage_baseline.py` | 〔真缺口〕KPI 要加註：companyfacts 沒有延伸 tag，所以這個數字低估 |

**不動**：`facts_mapping.py`／`fetcher_facts.py`（第二資料源保持不變）、
`Data_Financials` 的列位與列名、符號慣例。

### 實作順序（兩步各自可獨立驗收，不要混在一起做）

**第一步：模板 tuple 加第七欄 `label_fallback`。** 純結構改動，97 列都補 `None`，
行為零變化（測試必須全綠且輸出逐格相同）；接著才針對少數已知需要的列填正則。
先做這步是因為它讓下一步的「還有哪些空格」定義清楚。

**第二步：數值指紋連結。** 蒐證 → 連結 → 補值三個函式，各自 TDD。

### 風險與控制

最大的風險是**階段一要讀所有期間欄**，而現行邏輯刻意只讀一欄
（`if label in periods: continue` 這個 dedup 就是為了避免同一期被兩份 filing 蓋來蓋去）。

控制方式：**蒐證與取值完全分離**。階段一只往索引裡寫，不碰 `row_vals`；
真正的取值路徑一行都不改。所以既有輸出在「沒有任何連結成立」時必須逐格相同——
這會是驗收的第一條。

---

## 6. 驗收

1. **回歸**：對既有 102 家跑前後逐格對照。**沒有任何一格從「有值」變成「不同的值」**
2. **NVDA 個案**：年表 `Capex` 從 4 年補到 14 年以上（2013~2023 那 11 年是這次的
   目標）；季表 `Capex` 從 36/57 往上走，**上限是 53/57**——最舊的 4 個 Q4
   （2010-01、2011-01、2012-01、2013-01）缺的是合成材料不是 tag，這次不處理；
   `Free Cash Flow` 隨 Capex 同步改善
3. **規模**：102 家統計「補上幾格」「成立幾條連結」「因衝突被放棄幾條」
4. **精確度**：抽樣 20 條連結，逐條回原始 filing 確認是同一條科目
5. 全套測試綠燈，重產基線

---

## 7. 為什麼不選其他做法

| 做法 | 否決理由 |
|---|---|
| 人工維護 alias 對照表 | CTH 2026-08-23 明確要求「不能用人工判斷，要全自動 + 高準確度」 |
| 文字相似度自動配對 | 實測約一半誤判，且 NVDA 這個案例偵測不到 |
| 只加 `label_fallback` | 救得了 label 穩定的案例，但要人寫正則，而且不知道還漏了什麼 |
| 改走 companyfacts | G11 已決議不換；而且 companyfacts 沒有延伸 tag，這題它根本解不了 |
| 只讀比較欄、不做連結 | 比較欄只往回兩年，NVDA 只補得到 FY2022／FY2023，FY2013~2021 仍然空 |
