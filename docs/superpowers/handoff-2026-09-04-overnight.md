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

# 執行記錄（2026-09-04 夜間，隨時更新——斷線時看這裡接手）

> 這一段是**滾動更新**的，不是最後才寫。如果 AI session 中途因額度耗盡斷掉，
> 下一個 session 從這裡接手，**不要重跑那三小時的抓取**。

## 環境現況

- 分支：`spike/h0-baseline-rebuild`（從 `master` 開出來，master 沒動）
- 201 個舊 pkl 已備份到 `output/_spike_pkl_backup_20260904/`（含舊的
  `mapping.json`／`mapping_candidates.json`）。要回復原狀就複製回 `output/_spike/`
- **有一支 detached 看門狗在跑**：`scripts/watchdog_h0_baseline.sh`
  （log：`output/_spike/watchdog.log`）。它跟 AI 的額度無關，會自己等抓取跑完、
  掃警告、產基線、把驗收數字寫進 `output/_spike/h0_summary.txt`

## 斷線後怎麼接手（照這個順序）

1. 讀 `output/_spike/h0_summary.txt`——達標列數新舊對照、H1 的 `from_ytd` 填滿率、
   三分類、假警報都在裡面。**沒有這個檔就代表抓取還沒跑完**，看 `watchdog.log`
2. 讀 `output/_spike/rebuild_warnings.txt`——D11 的抓取缺漏。有可疑的 ticker 就
   **只刪那幾家的 pkl** 再跑一次 `spike_derive_mapping.py <那幾家>`
   （會走本地 filing 快取，很快），然後重跑 `gen_template_coverage_baseline.py`
3. 照主任務 A 的「驗收」四點寫報告、更新 `docs/TODO.md` H0、寫 CHANGELOG
4. 主任務 B（H1 改寫）：數字從新基線第三節的「facts填滿」欄與那行
   `from_ytd` 填滿率中位數撈

## 已完成

- **commit `f142494`**：`gen_template_coverage_baseline.py` 加「現行填滿／facts填滿」
  兩欄＋一行 `from_ytd` 列的 facts 填滿率中位數（H1 的驗收數字，原本文件裡撈不到）；
  「抓取窗不一致」那段改成從 pkl 實際期數動態算，並保留一段說明「那 12 家跟舊基線
  對不起來是抓取窗變了、不是回歸」
- **commit `2b3e61e`**：看門狗腳本 ＋ `scripts/README.md` Index 同步
- **5 家的煙霧測試**（重建到第 5 家時跑的）：`from_ytd` 那 29 列的 facts 填滿率
  中位數 **100%**，H1 記的原始症狀是 25%。201 家的正式數字要等跑完

## G13 已查到成因（用備份 pkl，零網路）

SNOW 那兩欄**不是**「兩份 filing 用不同財季標籤收進同一期」（TODO 原本的猜測），
而是**季表裡混進了一個純年度欄**：

```
FY2022     2022-01-31   ← 年度標籤，混進季表
FY2022Q4   2022-01-31   ← 真正的 Q4
FY2023Q1   （整欄不見了）
```

鏈路：`_is_q_col()`（`fetcher_gaap.py:875`）把 `(FY)` 也算成季度欄
→ 某份 10-Q 的 IS dataframe 被 `_current_q_col()` 挑到 `(FY)` 欄
→ `_col_to_quarter_label()`（`fetcher_gaap.py:819`）回傳 `FY2022`。

**副作用比重複欄嚴重**：`_build_is_table` 的 dedup 是 `if label in periods: continue`，
所以真正的 FY2023Q1 被那個假的 `FY2022` 佔掉位置、整欄被吃掉。

量化（掃 201 家備份 pkl）：「季表出現純年度標籤 `FY\d{4}`」與「期末日重複」
**都只有 SNOW 一家，兩者一對一重合**。

還沒做的最後一步：等 SNOW 進本地 filing 快取後，把那份 10-Q 的原始 dataframe
印出來確認欄名真的是 `2022-01-31 (FY)`。**確認完只寫進 TODO，不改修法**
（改法會動到期間去重邏輯，要 CTH 點頭）。

## 我自己做的決定

1. **加填滿率欄位到基線文件**（原本沒有）。理由：H1 的驗收數字「CF 流量列填滿率」
   在舊文件裡根本撈不到，交接指示要我從新基線撈，撈不到就得另外寫一支腳本。
   代價：文件表格多兩欄、跟舊版不能逐列 diff。錯了的話回退這個 commit 即可
2. **寫看門狗腳本**（交接沒要求）。理由：抓取要三小時，AI session 可能因額度耗盡
   中途斷掉，成果不該綁在 AI 活著。代價：多一支 `scripts/` 腳本要維護
3. **關掉 Monitor 的進度回報**，改由看門狗記在檔案裡。理由：每 15 分鐘叫醒 AI
   一次純燒額度，而進度資訊寫檔就夠
4. **ticker 清單用 `tr ',' ' '` 拆開**：`tickers_joined.txt` 是逗號串接的，
   交接文件寫的 `TICKERS=$(cat ...)` 直接帶進去會被當成單一 ticker
