# 交接：2026-09-04 夜間無人值守

> **CTH 正在睡覺，不會回答任何問題。** 這份是給下一個 session 的作業指示。
> 遇到需要決定的事，**自己決定、記錄下來、繼續做**——停下來等回覆等於整夜浪費。
> 只有「不可逆或對外」的動作才停（見下方紅線）。

## 現在的狀態（都已驗證，不用重查）

- 分支 `master`，工作目錄乾淨，**本地領先 origin 29 個 commit，尚未 push**
- 測試：`./venv/Scripts/python.exe -m pytest tests/ -q -m "not slow"` → **1392 passed, 65 deselected**
- slow 層（真打 SEC，約 2~12 分鐘視快取冷熱）：58 passed / 7 skipped / 0 failed
- **本地 filing 快取已上線並生效**（`src/filing_cache.py`，掛在
  `fetcher_gaap._filing_obj()`）。目前快取裡有 12 家公司約 21MB。
  這件事直接影響今晚的作業：**重抓同一家公司只要十幾秒，失敗重跑很便宜**
- 剛完成 TODO I3（四顆抓取按鈕互鎖）與 I4（locale 重複 key 防線）

## 紅線（這幾件事不准做，做了 CTH 醒來會很不爽）

1. **不要 `git push`**。本地 commit 隨便，推遠端是對外動作，要等 CTH 醒來決定
2. **不要動 `master`**。今晚的工作全部開新分支（例如 `spike/h0-baseline-rebuild`）
3. **不要做下面這些「需要 CTH 決定」的 TODO 條目**——它們卡在產品判斷而不是技術：
   G4（overflow 標示方式）、G8（比較欄補洞，風險最高且有未決問題）、
   I5（要不要抓修正案）、D0-5、D8、H2、H6-1、E3／E5（要 CTH 在他自己那台機器重現）、
   B2（要跟 CTH 一起設計 skill）
4. **不要做 H4 第二步**（數值指紋自動連結）。已經量化過：理論上限只有 15.3%，
   人工抽查後多數是假候選，TODO 裡明確建議不做
5. **不要刪 `scripts/` 底下任何腳本**（專案鐵則）。過時的話在
   `scripts/README.md` 的表格標 `(停用)`，檔案保留
6. **不要碰 `output/_hintsweep_201/`**（一輪 12 分鐘的原始資料，TODO 明講不要重跑）

---

## 主任務 A：重建模板體檢基線（H0）——這是今晚的重點

### 為什麼要做

`docs/TODO.md` H0 記著：現行基線 `docs/template-coverage-baseline-2026-08-24.md`
是 **H6／INTC Cash／G10（2026-08-25）之前**的數字，那三輪在 201 家上合計多救回
**61 家次**，但基線沒跟著更新，所以現在沒有人知道達標列數到底往上走了多少。

**這件事今晚特別適合做，因為兩個前提剛好都成立**：
- 它是純量測，**完全不需要 CTH 做任何決定**
- 它要對 201 家公司真的抓 SEC，**要跑好幾個小時**——正好是他睡覺的這段時間
- 而且我們剛做完的本地快取讓「失敗重跑」從幾小時變成幾分鐘

### ⚠ 先看這個，不然你會白跑

我（上一個 session）已經驗證過：**只重跑
`gen_template_coverage_baseline.py` 不會有任何變化**。它吃的是
`output/_spike/gaap_*.pkl`（已經比對完的結果，不是原始 dataframe），
所以模板 hint 的改動反映不出來。我實測產出的檔案跟 2026-08-24 那份
**逐位元組相同**（只差它少了那個警告橫幅），已經刪掉了。

**真正要做的是先重建那 201 個 pkl**：`scripts/spike_derive_mapping.py` 的
`_load_gaap()` 在 pkl 存在時會直接回傳、不重抓（`spike_derive_mapping.py:81-83`），
所以要先刪掉才會真的重抓。

### 做法

```bash
cd "C:/Users/CTH/Documents/Code/SEC Financial Tools"
git checkout -b spike/h0-baseline-rebuild

# 1) 備份再刪 pkl（facts_*.json 不要動——那是 companyfacts，不受模板改動影響）
mkdir -p output/_spike_pkl_backup_20260904
cp output/_spike/gaap_*.pkl output/_spike_pkl_backup_20260904/
rm output/_spike/gaap_*.pkl

# 2) 重建答案卷（201 家、真的抓 SEC，估 2~4 小時）——**背景跑，不要前景卡住**
TICKERS=$(cat output/_hintsweep_201/tickers_joined.txt)
./venv/Scripts/python.exe scripts/spike_derive_mapping.py $TICKERS > output/_spike/rebuild_20260904.log 2>&1

# 3) 重產基線（不打網路，幾分鐘）
./venv/Scripts/python.exe scripts/gen_template_coverage_baseline.py
```

### 這一段的三個坑，動手前先讀

1. **D11：連續抓 201 家時 SEC 會偶發失敗，靜默少格。** 這正是
   `docs/TODO.md` D11 記錄的現象（`fetch50` 那輪有 11 家的 `Operating Income`
   只有 16/21 期，隔幾小時單獨重抓就正常）。而 `_load_gaap()` **不管有沒有缺漏
   都會把結果寫成 pkl**，所以壞掉的結果會被凍進去。
   **對策**：跑完後檢查 `rebuild_20260904.log` 裡有缺漏警告的 ticker，
   **只刪那幾家的 pkl** 再跑一次——第二輪會從我們的本地快取讀已經成功的部分，
   只重抓失敗那幾份，很快。這是新快取帶來的紅利，善用它
2. **答案卷的抓取窗會變得一致，這是刻意的但要講清楚。** 舊基線裡
   AAPL/ADBE/AMD/AVGO/COST/GOOGL/INTC/META/MSFT/NVDA/TSLA/WMT 這 12 家是用
   「全部 filing」抓的（44~69 期），其餘是 `max_filings=16`。腳本現在寫死
   `_ANSWER_KEY_FILINGS = 16`／`_ANSWER_KEY_ANNUALS = 5`，所以重建後 201 家
   會**統一**成 16/5。這比原本一致，但那 12 家的逐列覆蓋率會因此變動——
   **在新基線的開頭把這件事寫清楚**，不然下一個人會以為是回歸
3. **不要以 97/97 為目標。** 基線文件第零節講得很清楚：達標列數只是體溫計，
   真正的 KPI 是〔真缺口〕與〔假警報〕。報告時照那個框架講

### 驗收

- 新基線 `docs/template-coverage-baseline-2026-09-04.md` 產出，且**跟 08-24 那份
  不是逐位元組相同**（相同就代表 pkl 沒重建成功，回頭查）
- 達標列數與 08-24 的 **44 列**比較，說明變動（H6+Cash+G10 救回 61 家次，
  預期往上走，但**如果沒有往上走就照實說**，不要修飾）
- 把 `docs/TODO.md` H0 那條的數字更新，並拿掉「基線沒有跟著更新」那段警告
- `docs/CHANGELOG.md` 加一筆

---

## 次要任務 B：`docs/TODO.md` H1 那條是**過期的**，要改寫

我查過了：H1 說「**修法**：`fetcher_facts` 要加一層『同一財年內用本期 YTD −
上期 YTD 還原單季』」——**這件事已經做完了**：

- `quarterly_from_ytd()` 實作在 `src/fetcher_facts.py:164`，docstring 裡連
  AAPL Capex 的實測數字都有（52 期對 51 錯 1，那 1 個是孤單年度值，已擋掉）
- 已接進 `resolve_row()`（`fetcher_facts.py:232`）
- `tests/test_fetcher_facts.py` 有 5 條測試蓋住（相鄰相減、第一期直接採用、
  依 `start` 分財年、擋掉孤單年度值、concept 不存在回空）

**所以 H1 剩下的不是實作，是驗證**：那條聲稱「CF 流量列的填滿率中位數只有
25%」——修完之後到底變多少？這個數字要從**新的基線**（任務 A 的產出）裡撈，
兩件事天然接在一起。做完任務 A 之後：

1. 從新基線第五節（現行路徑 vs companyfacts 逐列對照）撈 CF 流量列的填滿率
2. 用實測數字改寫 H1，把「要加一層」改成「已完成於 XXX，實測填滿率從 25% →
   YY%」，或者如果沒改善，照實寫並說明為什麼
3. 如果新數字顯示 companyfacts 對 CF 已經堪用，順手更新 G11 的結論那段

---

## 有空再做（小、純調查、不需決策）

**G13：同一個期末日出現兩次。** 實測 SNOW 的 `Data_Financials(Q)` 有兩欄期末日
都是 `2022-01-31`（52 家、1,482 對相鄰期間裡只有這 1 筆）。要查的是：是同一期
被兩份 filing 用不同財季標籤收進來（label 沒撞號但期末日撞了），還是別的原因。
**只要查清楚成因並寫進 TODO 就算完成，不要順手改修法**——改法會動到期間去重
邏輯，那要 CTH 點頭。SNOW 現在抓一次很快（會進本地快取）。

---

## 工作方式

- 用 `superpowers:test-driven-development`（有寫程式的話）與
  `superpowers:verification-before-completion`（宣稱完成前一定要有指令輸出佐證）
- **數字要實測，不要推測。** 這個專案這兩天已經被「憑感覺講數字」咬過兩次：
  一次是效能倍數量錯（沒清 edgartools 自己的 HTTP 快取 `~/.edgar/_tcache`），
  一次是把 log 加速歸因錯。**沒量過就說沒量過**
- 長時間指令用背景執行，不要前景卡住整個 session
- 每個階段結束就 commit，訊息用繁中，結尾加：
  ```
  Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>
  ```
- 最後在這個檔案末尾追加一段「做了什麼、量到什麼、我自己決定了什麼」，
  CTH 起床第一件事會看這裡

## 給 CTH 的晨間摘要要包含

1. 達標列數：44 → ?（以及〔真缺口〕／〔假警報〕的變化）
2. 重建過程中有幾家撞到 D11 的抓取缺漏、怎麼處理的
3. H1 的實測結論（CF 填滿率到底改善了沒）
4. 你自己做了哪些決定、如果錯了代價是什麼
5. 還沒 push（29+ commits），要不要推是他的決定



---

# 執行記錄（2026-09-04 夜間）

## ⚠ 計畫中途改變：主任務 A（H0 基線重建）**沒有做**

CTH 半夜醒來兩次，第二次明確說「**先不要跑那 200 家，我改變主意**」。
所以 201 家答案卷重建**已停止並完全還原**：

- 抓到第 8 家（AEP）時中止，`output/_spike/` 的 201 個 pkl **已從備份還原成原狀**
  （驗證：AAPL 69 期＝舊的全 filing 版）
- 看門狗曾用當時只有 24 家的 pkl 誤產一份 `template-coverage-baseline-2026-09-04.md`，
  **已刪除**。`h0_summary.txt`／`rebuild_warnings.txt` 一併刪除
- 沒有殘留的 process，工作目錄乾淨
- 備份仍留在 `output/_spike_pkl_backup_20260904/`（跟 `output/_spike/` 現在內容相同，
  確認沒問題後可以刪）

**所以 H0 仍然是待辦**，TODO 裡那段「基線沒有跟著更新」的警告**沒有拿掉**，
因為它現在仍然成立。

## 實際完成的四件事

### 1. `perf(cache)` — TODO I7 第一刀（commit `f80f086`）

`_CachedStatement.to_dataframe()` 加 memo，`_CachedFinancials` 一併 memo
statement 物件（不然每次 new 一個，上層 memo 形同虛設）。

- **實測（ARLO，預設 80/20，各跑 5 次，範圍完全不重疊）**：
  端到端中位數 **7.07s → 6.52s（省 0.55s，約 9%）**；
  `payload_to_df()` **224 次／385ms → 99 次／160ms**
- **⚠ 順帶修正 TODO I7 記的「省下約 0.85s（6%）」**：秒數高估（解析總成本上限
  就只有 0.37~0.40s），百分比低估。原記錄的基準是「ARLO 一趟 10~17s」，
  這次熱跑量到 7s，**基準本身不同**
- 每次仍回傳 `.copy()`：全庫查過零 `inplace=True`、零欄位指派，共用「現在」安全，
  但未來有人寫 `df["x"] = ...` 會靜默污染別張表。深複製比重新解析便宜 9.8 倍，
  隔離幾乎免費
- **驗收**：`excel_golden.py` 驗的是 Excel 寫檔那段、跟這次改動不同軸，所以改用
  「抓取結果逐格比對」——5 家（ARLO／AAPL／GOOGL／META／JPM，含金融股）
  **23,859 格 0 格不同**；測試 1392 → **1396 passed**（+4）

### 2. `docs(todo)` — H1 實測驗收（commit `800f036`）

H1 記的「修法：要加一層 YTD 相減還原單季」**早就做完了**
（`quarterly_from_ytd()`，`fetcher_facts.py:164`，7 條測試），只是沒人回頭量過。

**關鍵發現：這個數字不需要重抓 201 家就能量。** 拿 `output/_spike/` 既有的
facts JSON 與答案卷，把 mapping 的 `from_ytd` 拿掉當「修法前」，做同一份資料上
的 A/B：

**27 個 `from_ytd` 列的填滿率中位數 25% → 100%**，完全重現 H1 原記錄的 25%。
沒到 100% 的六列（`Acquisitions` 61%、`Debt Proceeds` 66% 等）都是
「本來就不是每季都發生」的活動。G11 的 caveat 解除，但 G11「不切換」的決策不變。

### 3. `docs(todo)` — G13 成因查明（同 commit `800f036`）

原本猜「同一期被兩份 filing 用不同財季標籤收進來」，**猜錯了**。

逐份印 SNOW 的 16 份 10-Q，`0001640147-22-000044`（2022-06-03 申報，本該是
FY2023Q1、期末 2022-04-30）的損益表 dataframe **唯一的期間欄是
`2022-01-31 (FY)`**——那張表裡根本沒有 Q 欄。其餘 15 份都正常。

鏈路：`_is_q_col()`（`fetcher_gaap.py:875`）把 `"FY"` 也算成期間欄 →
`_current_q_col()` 回傳它 → `_col_to_quarter_label()` 回 `FY2022`。
**副作用比「看到兩個一樣的日期」嚴重**：dedup 是 `if label in periods: continue`，
那個假的 `FY2022` 佔掉位置後，**FY2023Q1 整季被靜默丟掉**。

201 家裡只有 SNOW，且「季表出現純年度標籤」與「期末日重複」一對一重合。
**修法未做**（會動到期間去重邏輯，要 CTH 點頭），三個候選方向記在 TODO G13。

### 4. 兩支基礎建設（commit `f142494`、`2b3e61e`）

- `gen_template_coverage_baseline.py` 加「現行填滿／facts填滿」兩欄＋一行
  `from_ytd` 列的中位數（H1 的驗收數字原本在文件裡撈不到）；
  「抓取窗不一致」那段改成從 pkl 實際期數動態算
- `scripts/watchdog_h0_baseline.sh`：H0 重建的無人值守看門狗（這次沒用上，
  但之後真的要跑那三小時時可以直接用）。`scripts/README.md` Index 已同步

## 給 CTH 的重點

1. **H0 沒做**，是你自己喊停的。TODO 那段警告仍然成立，沒有動它
2. **兩個「TODO 記錄跟現實不符」的案例**，都已用實測數字修正：
   I7 的效能數字高估、H1 的修法其實早就完成。**引用文件裡的效能數字前先確認基準**
3. **G13 查出來的東西比預期嚴重**：不只是「看到兩個一樣的日期」，是**整季資料被
   靜默丟掉**。修法要你點頭
4. **快取現在有 34 家、66.5 MB**（今晚抓 AAPL~AEP 那 8 家順便暖進去的）
5. **還沒 push（本地領先 origin 34 個 commit）**，要不要推是你的決定

## 我自己做的決定

1. **加填滿率欄位到基線 generator**：H1 的驗收數字原本在文件裡撈不到。
   代價是表格多兩欄、跟舊版不能逐列 diff。錯了回退 `f142494`
2. **I7 的 memo 每次回傳 `.copy()` 而不是共用物件**：實測隔離成本只有解析的
   1/9.8，幾乎免費，換掉一整類未來很難查的 bug
3. **I7 的驗收改用「抓取結果逐格比對」而不是 TODO 指定的 `excel_golden`**：
   後者驗的是 Excel 寫檔，跟這次改動不同軸，跑了也證明不了什麼
4. **H1 用「拿掉 `from_ytd`」模擬修法前**，而不是等 201 家重抓。
   同一份資料的 A/B 比跨時間比對更乾淨，而且零網路
5. **看門狗腳本保留不刪**（專案鐵則不刪 `scripts/`），下次跑 H0 直接用
