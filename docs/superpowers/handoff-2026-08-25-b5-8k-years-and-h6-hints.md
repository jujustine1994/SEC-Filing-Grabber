# 交接：B5（8-K `--years` 標錯年份）+ H6（hint 三項）

**貼這份給新對話即可開工。** 2026-08-25 產出，接續同日完成的 G2+G7+G6（已合併進 master）。

---

## 任務與順序（已定案，照做，不要重排）

1. ~~**B5**：8-K `--years` 篩選用的 label 有 31.5% 年份是錯的 → 換成已驗證 100% 的零下載規則~~
   ✅ **2026-08-25 完成**（CTH 中途指示「本對話先做完 B5 即可」，所以下面 2~4 項還沒做）。
   內容見 `docs/CHANGELOG.md` 的 B5 條目、`docs/8k-period-off-by-one.md`「零下載規則」一節
2. **H6 第一批**：Capex / Cash / Ordinary shares 的 label hint 太窄，擴充正則
3. **H6 第二批**：只**診斷**、不修（見下方「刻意不做」）
4. **順帶重新量 G10**：H6 修完重跑六家比較檔，看 D&A／Capex 的覆蓋率變多少

**規格全部寫在 `docs/TODO.md` 的 B5 與 H6 兩條**，含實測數字、原始證據、風險。
那兩條是這輪的規格書，動手前整條讀完。B5 另有完整報告
`docs/superpowers/report-2026-08-25-8k-years-zero-download-rule.md`。

---

## 資料已經抓好了，**不要重抓**

| 放在哪 | 是什麼 |
|---|---|
| `output/_8k_audit/8k_audit_html/` | 199 份新聞稿原文快取（12MB）。改規則重算**完全不必連 SEC** |
| `output/_8k_audit/*.json` | 三版規則各自的逐份比對結果 + 最終驗證 |
| `output/_8k_audit/check_*.py`、`impact.py` | 產生那些 json 的腳本，可直接改規則重跑 |
| `output/_hintsweep_201/hintsweep_201_result.txt` | 201 家逐列逐家的 hint 掃描原始輸出 |
| `output/_hintsweep_201/classification.md` | 人工判讀後的分類表（真缺口 vs hint 正常運作） |

`output/` 有 gitignore，資料在磁碟上但不進版控。**重跑一次 SEC 抓取要 40 分鐘以上，
而這些資料已經付過那個成本了。**

---

## 已經替你做完的決策（不要再問 CTH，照做）

### B5

| 決策 | 內容 |
|---|---|
| 用哪版規則 | **規則 C**（EDGAR `fiscal_year_end` 完整 MMDD 往前推 3/6/9 個月）。規則 A 實測 58%、規則 B 79.8%，都已否決，不要重試 |
| `tol` | **21**。掃過 0~40，3~30 之間命中率完全相同，高原 27 天寬 |
| 純函式放哪 | `fiscal_input.py`，命名 `quarter_label_from_announcement(announce_date, fiscal_year_end_mmdd, tol=21)`，零 I/O |
| 拿不到 `fiscal_year_end` 時 | 退回現行 `_period_to_quarter_label()`，**不要讓缺欄位變成整批失敗** |
| `fiscal_label` | **不動**。它是下載後從真實期末日算的，仍然最準 |
| `label_source` | 改成 `"announcement+fiscal_year_end"`，`_LABEL_WARNING` 那段過期警告改寫 |
| 對外行為改變 | 要同步改 `docs/CLI.md` |
| **`fiscal_year_end` 會不會隨時間變（最大風險）** | ⚠ **2026-08-25 訂正：本欄原本寫的決策是錯的。** 原文說「加一道 0~70 天的 sanity check 就會被接住」——**那是恆真式，攔不到任何東西**：候選季末永遠相隔 89~92 天，規則取「不晚於 發布日+tol 的最新候選」，選中的必然落在 `[-tol, 91-tol)`，tol=21 時上界剛好就是 70（200 份實測 lag 範圍 -2~58）。實作保留了那道檢查（`max_lag_days` 參數，擋參數被改壞與畸形輸入），但**FYE 漂移目前沒有對策**。唯一的偵測是下載後 label 與 `fiscal_label` 的比對旗標（`label_agrees_with_fiscal_label`），而它涵蓋不到「該被 `--years` 選進來卻被漏掉」的那一類。要真的驗有一條不打網路的路子（拿 `output/_spike/` 歷年期末日推出的 FYE 比對現值），寫在 `docs/TODO.md` 的 **B6** |

### H6

| 決策 | 內容 |
|---|---|
| 修哪些 | **Capex**（14 家）、**Cash**（5 家：ETN/APD/IP/KR/SLB）、**Common Stock & APIC 裡的 7 家 Ordinary shares**（ACN/AON/CB/ETN/JCI/LIN/MDT） |
| 不修哪些 | ① CS&APIC 另外 7 家（ABT/AMP/AXP/COP/KR/MPC/UNP）——那是 concept 層失守不是 hint 問題；② 銀行的 `CashAndDueFromBanks` 7 家——概念取捨題，要 CTH 決定；③ Cost of Revenue 那 29 家銀行/保險/鐵路——本來就沒有 COGS |
| Cost of Revenue 要修的 6 家 | CVX/COP/PSX（採購原油商品）、AEP/EXC（採購電力燃料）、CMG（食材包材）。**只放寬到吃得下這幾家，不要放寬到吃進 `LaborAndRelatedExpense`** |
| 怎麼確認沒有副作用 | 改完重跑 `scripts/diag_hintsweep.py`（指令見 TODO H6），比對「killed 清單有沒有變短、有沒有長出新的 DIFF」 |
| 抽查 | 分類表作者自己標了「只讀 concept/label 文字判斷，沒回頭核對原始 10-Q」。**每一類抽 1 家**去 SEC 原文確認那一列真的是我們要的科目，再改 |

---

## 刻意不做（不要順手做）

- **CS&APIC 那 7 家的 concept 層**：只用 `scripts/diag_rowprobe.py` 查「是哪一層先失守」，
  **把結果寫進 `docs/TODO.md` 就好，不要動 concept 對照**。那會影響所有公司，風險等級跟
  這輪不同
- **G8**（比較欄 fallback 補洞）：風險最高，會動所有既有期間的資料來源優先序
- **H1**（companyfacts 的 CF YTD 拆算）：G11 已定案不切換到 facts，現在做等於白做
- **H4 第二步**（數值指紋自動連結）：已量化，理論上限 15.3% 且多數是假候選，**已決定不做，
  不要重新評估**
- **D8／H2**（金融股模板）：產品決定，不是這輪的事

---

## 做法要求

- **TDD**：先寫測試看它紅，再實作。B5 的單元測試至少涵蓋 **WDC**（季末 1/2 跨年）、
  **COST**（Q3 季末 5/10 但 5/28 就發，需要 tol）、**MU**（FYE 0903，名目月底法會錯）、
  **NVDA/TGT**（1/2 月結算的財年編號慣例）——這四家正是三版規則的分水嶺
- **不要開平行 subagent**：B5 與 H6 都會跑真連 SEC 的驗收，併發會互搶 SEC 頻寬
  （實測兩份 20 分鐘的測試併著跑變 40 分鐘）
- **收尾測試**：`venv/Scripts/python.exe -m pytest -m "not slow" -q`（約 20~35 秒，
  現況 **1170 passed**）；再 `pytest -m "slow" -q`（約 31 分鐘，會真連 SEC，
  **背景跑不要乾等**，現況 **58 passed / 7 skipped**）
- **回歸**：B5 改完要跑 `scripts/verify_8k_fiscal_labels.py`，確認 `fiscal_label` 沒被動到
- **分支**：在 master 上開 feature branch 做，完工後照
  `superpowers:finishing-a-development-branch` 的三選一流程收尾

---

## 收尾一定要做的四件事

1. **更新 `docs/CHANGELOG.md`**：做完的內容搬進去，含實測數字與「為什麼這樣修」
2. **更新 `docs/TODO.md`**：B5、H6 做完的部分**整條刪掉**（專案既有規則：TODO 只留沒做的）；
   沒做的部分（CS&APIC concept 層、銀行 due from banks、Cost of Revenue 的 29 家）
   留下並更新成「已診斷、待決策」
3. **更新 `docs/ARCHITECTURE.md`**：`_list_earnings_filings()` 從「純 listing metadata」
   變成「listing metadata + 一次 company 層級查詢」，這條界線要改；測試分層表的數字要更新
4. **更新 `docs/CLI.md`**：`label_source` 與那段警告是對外行為
5. **`docs/8k-period-off-by-one.md`**：加一節「零下載規則（2026-08-25 驗證）」，
   並更新原本的「修法建議」三選項——當初 A 的成本評估是「要下載文件才知道期間」，
   現在證明不必下載，那個假設已經不成立

---

## 回報格式（跑完給 CTH 看的）

用中文，直接講重點，包含這四段：

1. **做完了什麼**：B5 / H6 各自的前後數字（例如「8-K label 命中率 X% → 100%」、
   「Capex 全損 14 家 → N 家」），要有數字不要只說「修好了」
2. **我自己做了哪些決定**：逐條列出。特別是規格沒寫死、你自己判斷的地方
   （正則怎麼寫、sanity check 的門檻、抽查抽了哪幾家、遇到什麼跟預期不同的）
3. **哪些沒做、為什麼**：包含刻意不做的，與做到一半發現不該做的
4. **測試與驗收證據**：非連線幾條過、slow 幾條過、實跑了哪幾家看了什麼

---

## 現況（2026-08-25）

```
分支      master（G1/G2/G3/G6/G7/G9 已完成，G11 已定案不切換，H0~H5/D11 都在裡面）
測試      1170 passed（不含 slow）／ slow 58 passed + 7 skipped
```

單一公司抓取那條線的體檢數字穩定在：達標列數 44/97、每格三分類
「我們抓到／真缺口／公司真的沒有」= 72~73% / 9% / 18%（201 家實測）。
B5 不影響這條線；H6 會讓「我們抓到」那格往上走一點，**修完值得重跑一次
`scripts/gen_template_coverage_baseline.py` 看變多少**（那支不打網路）。
