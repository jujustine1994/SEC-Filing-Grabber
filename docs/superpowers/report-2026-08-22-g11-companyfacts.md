# G11 報告：改用 SEC companyfacts API 取數（2026-08-22）

**狀態**：平行路徑已建好並驗證完成，**尚未接上主流程**。切不切換要 CTH 決定。
**驗證規模**：52 家公司（大中小型 × 跨產業，含金融股與 REIT）逐格比對。

---

## 一句話結論

**快 215 倍、資料更早、能補上現行路徑抓不到的列，而且揭露了現行路徑三個既有錯誤。
逐格比對 92.8% 相同，符號慣例對齊後 95.4%。剩下的 4.6% 全部有明確解釋。**

---

## 1. 為什麼要做

現行路徑對每一份 10-Q/10-K 下載並解析 XBRL，而且**同一份被解析 4 次**
（IS/BS/CF/segments 各一次，override 觸發時最多 7 次）。實測：

```
下載        0.3 秒（之後被 edgartools 的 ~/.edgar/_tcache 快取成 0 秒）
XBRL 解析   1.3~2.1 秒，完全沒有快取
```

六家公司 ≈ 1,800 次解析 ≈ **45 分鐘**。

companyfacts 是 SEC 官方把「這家公司歷來 tag 過的所有 XBRL fact」整理成一包 JSON，
**一家一個 request**：

| | 現行 | companyfacts |
|---|---|---|
| 每家 | ~7.5 分鐘 | **0.34 秒** |
| 六家 | ~45 分鐘 | **~2 秒** |
| 最早涵蓋（NVDA） | 2009-07 | **2008-07** |

**解析完全不用 AI**，兩條路都是純程式解析。慢的是 CPU 解 XML，不是 API。

---

## 2. 做了什麼

| 檔案 | 內容 |
|---|---|
| `src/fetcher_facts.py` | companyfacts 取數路徑。40 個測試 |
| `src/facts_mapping.py` | 模板列 → us-gaap concept 對照表，**證據導出非手填** |
| `scripts/spike_companyfacts_diff.py` | 兩條路的逐格比對 |
| `scripts/spike_derive_mapping.py` | 反推對照表 + Excel 證據 |
| `scripts/spike_validate_facts.py` | 不依賴現行路徑的獨立驗證 |
| `scripts/spike_verify_mapping.py` | 用對照表實跑、逐格驗收 + Excel |

**現有程式一行都沒動。** 這條路是平行的。

### 對照表為什麼不是手填

模板的 `std_concept` 欄是 **edgartools 正規化過的名字**，不是原始 us-gaap element name：

```
模板寫的                            SEC 實際的
NetIncome                      →   NetIncomeLoss
ResearchAndDevelopmentExpenses →   ResearchAndDevelopmentExpense     ← 少一個 s
StockBasedCompensationExpense  →   ShareBasedCompensation
Revenue                        →   Revenues / RevenueFromContractWithCustomerExcludingAssessedTax
```

95 列憑印象填一定會錯，**而且錯了不會有人發現**——數字看起來都很像。

改成**反推**：拿現行路徑已知正確的數字當答案卷，對 52 家 × 每個 concept 算
「同一期末日數字對得上的比例」，命中率最高的就是正確對應。每一列都留下證據
（幾家命中、覆蓋率、命中率），寫在 `facts_mapping.py` 的行末註解裡。

**結果：95 列中 90 列自動對上，3 列人工補（concept 身分無疑義、只是證據家數不足），
2 列合理地沒有對應。**

---

## 3. 驗收數字（52 家逐格比對）

```
兩邊都有的格子      61,069
數字相同            56,686   → 92.82%
其中只差正負號       1,543   → 符號對齊後 95.35%
只有 facts 有       64,379 格（多抓到的）
只有現行路徑有      32,046 格
```

### 剩下的差異全部有解釋，分三類

**(a) 符號慣例（1,543 格）**——這些列符號對齊後就是 95~100%：

```
Treasury Stock      7.9% → 97.4%      Income Tax          85.5% → 100%
Acquisitions       33.8% → 100%       Investment Purchases 86.5% → 95.8%
Interest Expense   34.2% → 98.3%      Change in Inventories 88.9% → 99.6%
Minority Interest  69.3% → 100%       Capex                89.1% → 96.5%
```

**關鍵發現：符號沒辦法用「每列一個常數」表達，因為現行路徑自己的符號就不一致。**
`Income Tax` 精確命中 85.5%、含符號 100%——代表同一列在 85% 的公司是正數、
15% 是負數。那是**現行輸出的既有不一致**，改用 companyfacts 反而會變一致。

**(b) 現行路徑會加總、facts 是單一 concept（可修，但要另外寫加總邏輯）**

`Investment Proceeds`／`Debt Proceeds`／`Debt Repayments`／`Total Non-op` 這幾列，
現行路徑用 `_sum_matching_rows()` 把多個 XBRL 列加起來。facts 這邊目前只取單一
concept，所以對不上。實測 AAPL：

```
Investment Proceeds 2024-12-28   現行 3,492,000,000（多列加總）   facts 15,967,000,000（單一 concept）
```

**(c) 現行路徑本身算錯，facts 是對的（3 處）**

最明確的是 `Ending Cash`：

```
AAPL 2026-03-28   現行     255,000,000   facts 45,572,000,000
AAPL 2026-06-27   現行  -6,028,000,000   facts 39,544,000,000
```

Apple 的現金是 450 億，現行路徑給 2.55 億甚至負數。
**根因**：`Ending Cash` 是期末餘額（時點值），卻被 `_build_cf_table()` 的
YTD 拆算當成期間值去減上一季，減出來是「變動額」不是「餘額」。
這是現行路徑的真 bug，companyfacts 因為每筆 fact 自帶 instant/duration 標記，
結構上不可能犯。

---

## 4. 不依賴現行路徑的獨立驗證（24 家）

反推有個先天限制：**現行路徑錯的地方，比對也會跟著錯**。所以另外做了四項
不看現行路徑的檢查：

```
資產 = 負債 + 權益（含夾層權益）    710/744   95%
四季加總 = 年度                      79/83    95%
SEC 官方 frame vs 我們的期中點判準  59,564/59,564   100%
實質重編（非精度變更）              2,357/43,725   5.4%
```

**第三行是這次最有價值的一條**：SEC 自己在每筆 fact 上標的 `frame`
（如 `CY2025Q2`）是官方的日曆季正規化。它跟我們 F6/G2 決定的**期中點判準**
在 24 家、59,564 筆資料上**零例外全部一致**。這是對跨公司對齊決策的獨立背書。

第四行說明「重編取版」這個選擇影響 5.4% 的格子——不是可以忽略的比例，
所以 `prefer` 參數是必要的，預設取 `filed` 最早（當初申報值）。

---

## 5. companyfacts 拿不到的東西（限制，不是待辦）

fact 的欄位只有 `start`/`end`/`val`/`accn`/`fy`/`fp`/`form`/`filed`/`frame`
——**沒有任何維度欄位**。所以：

1. **`Data_Segments`（帶維度的分類細項）這條路拿不到**，非走解 filing 不可
2. **沒有 presentation linkbase** → 沒有公司自報的原文標籤。每個 concept 只有
   US-GAAP 官方標準標籤（`Share-based Payment Arrangement, Noncash Expense`）
3. **`Other (as reported)` overflow 區語意會整個改變**——現在是「這張報表裡有、
   模板沒收錄的列」，改用 facts 之後只剩「這家公司 tag 過的 concept 裡模板沒收的」

**建議架構是混合**：模板列走 facts，segments 仍解 filing（但只解最近 N 份）。

---

## 6. 我做的決定（含被我否決的做法）

| 決定 | 理由 |
|---|---|
| **對照表用反推，不手填** | 模板的 concept 欄是 edgartools 正規化名，手填 95 列必錯且不會被發現 |
| **50 家的門檻改覆蓋率加權，不用「全公司一致」** | 金融股報表結構本來就不同，小公司未必 tag 得到；分母改成「有答案卷的家數」 |
| **一列允許多個 concept 排 fallback 鏈** | 同一列在不同公司/年代會對到不同 concept |
| **一個 concept 只能當一列的主人** | 實測踩到：`OtherNonoperatingIncomeExpense` 被推成 `Operating Income` 的備援，在某幾家數字剛好對得上但語意完全不同，換一家就把營業外損益填進營業利益。純看命中率擋不掉，要靠結構性規則 |
| **`negate` 跟隨首選 concept，不用 `all()`** | 實測踩到：Capex 首選要反號、某備援不用，`all()` 變 False 導致 31 家的 Capex 全部正負相反（命中率 7%） |
| **備援符號慣例必須與首選一致，否則剔除** | 單一 `negate` 表達不了兩種慣例，落到那個備援會靜默生出相反數字 |
| **`unit` 與 `taxonomy` 進 spec，指錯一律回空不退回預設** | EPS 是 `USD/shares`、股數是 `shares`、流通股數在 `dei`。退回 USD 會讓每股盈餘抓到金額，數字看起來很合理不會有人發現 |
| **答案卷抓取深度砍到 16 份** | 評分只要 4 期重疊。50 家從 2 小時降到 25 分鐘，證據強度無實質損失 |
| **三表分開、列序照模板、共用同一條期間軸** | CTH 明確要求維持原架構。期間軸不統一的話 `_merge_financials()` 會錯位（有測試釘住） |
| **不改任何現有程式** | 這是平行路徑，要能跟現行輸出逐格對照才有決策依據 |

### 模板列的調整建議（**我沒有自己改**）

- `Free Cash Flow` — XBRL 沒有這個 tag，本來就該用算的（模板已標 `source=DERIVED`）。**不動**
- `Other Operating Expense` — 52 家裡**現行路徑一家都沒抓到**，facts 也沒有對應 concept。
  兩邊都沒資料，**建議刪除，但要 CTH 決定**
- 其餘 93 列都有對應，**不建議增刪**

---

## 7. 建議的下一步

**不要現在切換。** 剩下的 4.6% 差異雖然都有解釋，但有兩件事該先做完：

1. **符號慣例定案**——現行輸出自己就不一致（`Income Tax` 有 15% 的格子符號相反）。
   建議定義「每一列一個明確慣例」並在兩條路都強制執行，這會讓輸出比現在更好，
   但**是行為改變，要 CTH 拍板**
2. **加總型的列**（`Investment Proceeds`／`Debt Proceeds`／`Debt Repayments`／
   `Total Non-op`）要在 facts 這邊補上同樣的加總邏輯

做完這兩項再跑一次 `spike_verify_mapping.py`，目標 99%。到那時切換就是看數據，
不是賭。

**已經可以先動的**：`Ending Cash` 那個 bug 跟 G11 無關，現行路徑就該修
（時點值不該走 YTD 拆算）。已記進 TODO。

---

## 8. 檔案位置

```
證據 Excel（人眼看這兩份）
  output/_spike/mapping_evidence.xlsx    對照表 + 所有候選 + 每列×每家覆蓋矩陣
  output/_spike/verify_mapping.xlsx      每列×每家命中率，<80% 紅、<95% 黃

原始資料
  output/_spike/facts_<TICKER>.json      companyfacts 原始回應（52 家）
  output/_spike/gaap_<TICKER>.pkl        現行路徑的答案卷快取
  output/_spike/mapping.json             自動產出的對照表
  output/_spike/mapping_candidates.json  完整候選清單
```
