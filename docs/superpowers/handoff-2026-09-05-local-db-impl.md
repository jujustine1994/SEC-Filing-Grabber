# 交接：實作本地財報資料庫（TODO J1–J4）

> **CTH 在忙，不會即時回答。** 需要決定的事**自己決定、記錄下來、繼續做**。
> 停下來等回覆等於白白浪費時間。只有「不可逆或對外」的動作才停（見紅線）。

## 任務

實作 `docs/superpowers/specs/2026-09-04-local-filing-db-design.md`。
設計**已經跟 CTH 走完 brainstorming 並定案**，不要重新設計、不要重開 brainstorming。
要做的條目是 `docs/TODO.md` 的 **J1、J2、J3、J4**。

先讀那兩份文件，再動手。設計書的〈十、刻意不做（YAGNI）〉那一節是硬性的，
不要自己加回來。

## 紅線

1. **不要 `git push`**。本地 commit 隨便，推遠端要等 CTH 決定
2. **不要動 `master`**。開新分支（例如 `feat/local-filing-db`）
3. **不要刪 `scripts/` 底下任何腳本**（專案鐵則）。過時的話在 `scripts/README.md`
   的表格標 `(停用)`，檔案保留
4. **不要碰 `output/_spike/`**（H0 體檢的答案卷，201 份 pkl + facts JSON）
   與 `output/_hintsweep_201/`（一輪 12 分鐘的原始資料）
5. **不要做設計書「刻意不做」那節的東西**：存原始 XBRL、容量上限／自動淘汰、
   「版本不符但照用舊快取」、動 `fetcher_gaap` 的抓取迴圈

## 工作方式

- `superpowers:test-driven-development`（先寫測試）
- `superpowers:verification-before-completion`（宣稱完成一定要有指令輸出佐證）
- **數字要實測，不要推測。** 這個專案已經被「憑感覺講數字」咬過三次
  （效能倍數量錯、log 加速歸因錯、TODO I7 的「省 0.85s」是高估）
- 每個階段結束就 commit，訊息用繁體中文，結尾加：
  ```
  Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>
  ```
- 長時間指令用背景執行

## 環境重點

```bash
cd "C:/Users/CTH/Documents/Code/SEC Financial Tools"
./venv/Scripts/python.exe -m pytest tests/ -q -m "not slow"
# 現在的基準：1396 passed, 65 deselected（2026-09-04 驗過）
```

快取現況：34 家、881 份、71.3 MB，位置
`C:\Users\CTH\AppData\Roaming\SEC Financial Tools\filing_cache\`

## 已知地雷（都是 2026-09-04 實際踩到的，不要再踩一次）

1. **`python - <<'PY'` 這種 heredoc 餵 stdin 的寫法，中文字串比對會失敗**
   ——Windows 上 python 讀 stdin 用 cp950 解碼。改成「把腳本寫成檔案，
   用 `./venv/Scripts/python.exe 檔案` 跑」，並加 `PYTHONIOENCODING=utf-8`
2. **`_disk_cache_scope()` 只在 `fetch_gaap_statements()` 裡開**
   （`fetcher_gaap.py:2619`）。直接呼叫 `_filing_obj()` 之類的內部函式
   **既不讀快取也不寫快取**——寫量測腳本時很容易白下載一堆
3. **`_dir_stats()`（`filing_cache.py:268`）算份數有 `ACCESSION_RE` 這道閘**，
   所以 `_meta.json` 不會被誤算成 filing——這點現成的就對。但 `size` 會加總
   所有 `*.json`。另外 `if count == 0 and size == 0: continue` 那道過濾，
   遇到「只剩 `_meta.json` 的資料夾」會失效，GUI 會出現一列「0 份」。
   設計書的建議是**清除時一併刪 meta**
4. **Bash 工具前景執行有 2 分鐘 timeout**，抓取類的指令一定要背景跑
5. `excel_golden.py` 驗的是 Excel 寫檔那段，**跟快取層不同軸**，
   拿它驗這次的改動證明不了什麼。要驗就驗「抓取結果逐格比對」

## 建議順序

```
1. J2 的純函式部分：_meta.json 的讀寫與自癒、reached_bottom 的推導
   （純邏輯，好測，不需要網路）
2. J1：更新名單的 config 讀寫 + 兩個便利動作
3. J3：「更新本地庫」的核心流程（先做 CLI，GUI 後接）
4. J4：requirements 鎖版本 + 版本不符的偵測與對話框
5. GUI：擴充既有的快取面板（main.py:2253），不開新分頁
6. 更新 docs/TODO.md（J 區改成已完成的部分搬進 CHANGELOG）與 docs/CHANGELOG.md
```

每個階段都要能獨立驗收。不要一路寫到底才跑測試。

## 驗收

- 非 slow 測試從 1396 往上（新增的測試要真的蓋到新邏輯，
  參考 `tests/test_filing_cache.py` 既有的風格）
- 「更新本地庫」對一家已經到底又沒有新財報的公司，**第二次執行要整家跳過**
  ——這是「不要每次全部重抓」的核心，要有實測輸出佐證（例如印出跳過了幾家）
- 用真實資料驗一次：拿快取裡已有的 3~5 家跑一輪，確認 meta 正確、沒有重複下載

## 可選（做完 J3 並驗證過才可以做）

TODO **J5：找一個晚上把全部公司的快取跑滿**（201 家、拓到底、約 1.4 GB、數小時）。
如果 J1–J4 都完成且驗收過，時間也還夠，可以用 J3 做出來的 CLI 跑起來。
跑之前先讀 TODO D11——連續大量抓取時 SEC 會偶發失敗、**靜默少格**，
跑完要檢查缺漏清單。

## 最後

在這個檔案末尾追加一段「做了什麼、量到什麼、我自己決定了什麼、哪些沒做完」。
CTH 回來第一件事會看這裡。


---

# 執行記錄（2026-09-04 由 Claude 追加）

分支 `feat/local-filing-db`，**沒有 push、沒有動 master**。

## 一、做了什麼

**J1-J4 全部完成。** 四個 commit：

| commit | 內容 |
|---|---|
| `3707050` | J1／J2／J4 狀態層（新檔 `src/local_db.py`） |
| `f6f77ae` | J3 的 CLI `update-db` ＋ GUI 快取面板擴充 |
| `23a7cf1` | TODO／CHANGELOG／ARCHITECTURE／CLI 四份文件 |
| `3793d28` | 抓取速率的數字修正（見下面第三節） |

- **J1 更新名單**：`config["local_db_tickers"]`，跟 `watchlist` 並列、語意分開。
  兩個便利動作（匯入 watchlist／匯入快取現況）GUI 與 CLI 都有
- **J2 `_meta.json`**：一家一份，分 form 記。「到底」在抓取迴圈**外面**推導
  （比對完整 filing 清單與已快取的 accession），**`fetcher_gaap` 一行都沒動**
- **J3 「更新本地庫」**：`local_db.update_local_db()`，GUI 按鈕（Tab3 既有面板，
  沒開新分頁）＋ CLI `src/cli.py update-db`
- **J4 版本鎖**：`requirements.txt` 改成 `edgartools==5.29.0`，
  啟動時偵測版本不符跳提醒（`main._warn_if_edgartools_changed()`）

測試 **1396 → 1446**（新增 50 條，全部離線）。GUI 用 Tk 探針驗過。
設計書〈十、刻意不做〉那節一條都沒加回來。

## 二、量到什麼（都是實測，不是推估）

**驗收條件（第二次執行要整家跳過）——過了：**

| | 耗時 | 結果 |
|---|---|---|
| 第一輪 AAPL／ARLO／META | 49.4s | AAPL、ARLO 整家跳過；META 新增 27 份（30→57） |
| 第二輪 同三家 | **1.0s** | 三家全部跳過，**零下載** |

`reached_bottom` 判定跟設計書那張實測對照表**完全一致**：
AAPL `xbrl_cutoff`（2008 撞 XBRL 起點）、ARLO／META `no_more_filings`
（2018／2012 才上市）。meta 的 count／oldest／newest 逐項核對過。

**抓取速率：冷跑 2.8 s/份**（連續抓 15 家沒抓過的公司，取中段 900 秒的窗
量到 321 份）。

## 三、我自己決定的事

1. **`SECONDS_PER_FILING` 用冷跑的 2.8，不用一開始量到的 1.8。**
   1.8 是對 META 量的，但那家在 `~/.edgar/_tcache`（edgartools 自己那層持久化
   HTTP 快取）裡已經是熱的——量到的是「本地重解析」不是「對 SEC 重新抓一次」。
   ARCHITECTURE.md 記過同一個坑讓第一次的快取效能量測整組作廢，這次差點又踩。
   估「重抓要幾小時」按最壞情況算。`3793d28` 把常數與兩份文件一起改了
2. **`plan_ticker()` 加了 `version_ok` 這個條件**（設計書沒寫）。版本不符時
   `load_filing()` 一律回 None，那些檔案等同不存在——這時如果照樣「整家跳過」，
   那家公司會**永遠停在失效狀態**。所以版本不符一律不跳過
3. **日期解析不出來時一律判「還沒到底」**。誤判「還沒到底」只是多查一次清單；
   誤判「到底」會讓那家公司永遠不再往下挖，而且完全沒有症狀
4. **`update_local_db()` 用 `load_meta()`（會自癒）不是 `read_meta()`**。
   既有的 34 家在這功能上線前沒有 meta，用 raw 讀會全部判成「版本不明→不可跳過」，
   第一輪就一定全部進抓取迴圈。實測結果：AAPL／ARLO **第一輪就跳過了**
5. **CLI 的名單維護做完就結束，不順便發動抓取。** 「改名單」跟「跑幾小時的抓取」
   混在同一次執行裡，手滑的代價差太多
6. **GUI 完成用新的 `db_done` 訊息，不顯示「開啟輸出資料夾」**——這條路不產 Excel，
   那顆按鈕會誤導
7. **順手修 `filing_cache._dir_stats()`**：容量不再把 `_meta.json` 算進去。
   不修的話「清空後只剩 meta」的資料夾會在 GUI 顯示成一列「0 份」
   （這就是交接文件地雷 #3 講的那件事，設計書建議「清除時一併刪 meta」——
   `clear_ticker()` 用的是 `rmtree`，本來就會一起刪，所以只要修統計那邊）
8. **沒有把 201 家寫進 CTH 的 `config.json`。** 規模實測用位置參數指定 ticker，
   不動使用者的設定

## 四、哪些沒做完

- **J5（把 201 家的快取跑滿）沒做。** 這是「可選」項目，而且照 2.8 s/份、
  201 家×約 80 份估算是 **12~13 小時**——那不是可以無人值守放著跑一整晚就算了的
  規模，SEC 的偶發失敗（D11）也還沒量過。所以改成先跑一個**有界的規模實測**
  （15 家、拓到底）來量缺漏率，讓 CTH 拿數字決定要不要先做 D11 (c) 降速。
  **這個實測在寫這段時還沒跑完**，結果會另外補在下面
- **D11 (c)「降低連續抓取的速率」沒做**——設計書明確說不在範圍內，只記一筆
- **`main.py` 那幾個新函式沒有自動測試**（`_open_local_db_popup`、
  `_local_db_worker`、`_warn_if_edgartools_changed`）。照專案現況，Tk 的部分
  用探針手動驗；純函式的部分（`local_db_row_text`）有 4 條自動測試蓋住

## 五、要跑 J5 的話

```bash
cd "C:/Users/CTH/Documents/Code/SEC Financial Tools"
./venv/Scripts/python.exe src/cli.py update-db --import-cached      # 先建名單
./venv/Scripts/python.exe src/cli.py update-db --add <其餘 ticker>  # 補到 201 家
./venv/Scripts/python.exe src/cli.py update-db --json out.json      # 開跑
```

中斷不會白費（逐份即時落檔），重跑會自動跳過已經到底的公司。
跑完看 `out.json` 的 `gap_tickers`——那幾家單獨重跑一次即可（會走本地快取，很快）。

---

## 追加（同一天稍晚）：15 家規模實測跑完了

第四節說「這個實測在寫這段時還沒跑完」——跑完了，結果在這裡。

| 輪 | 耗時 | 結果 |
|---|---|---|
| 第一輪 | 32m09s | 15 家全部更新，新增 **783 份**，0 失敗、**0 缺漏** |
| 第二輪 | 29s | 14 家跳過；ACN 補了 2 份 |
| 第三輪 | **12s** | **15/15 全部跳過** |

**三件事跟 J5 的決策直接相關：**

1. **D11 的缺漏率是 0。** 連續抓 32 分鐘、783 份，`gap_tickers` 空的。
   → **J5 可以直接跑，不必先做 D11 (c) 降速**
2. **推估更新**：約 14,000 份新的 × 2.46 s/份 ≈ **9.6 小時**、磁碟 **1.36 GB**
   （跟設計書估的 1.4 GB 對得上）
3. ⚠ **跑完要再跑第二輪。** ACN 第一輪沒抓齊，第二輪才補上，第三輪才跳過。
   「跑一輪不保證到底」——但第二輪只要 12 秒，而且會自動補齊

## 追加：自我複查修掉的三個問題

寫 GUI 探針時回頭複查自己的程式碼，抓到三個：

1. **啟動時的版本偵測會把整個本地庫開檔讀一遍**（`stale_cache_summary()` 對每家
   呼叫 `scan_filings()`）。201 家拓到底＝16,000 份，每次開程式都要跑完才畫得出
   主視窗。改走 meta，實測 1,565 份時 963ms → 220ms
2. **「整家跳過」之後還是照樣重建 meta**（把那家 75 份全部開一遍），
   跳過的意義少一半
3. 上面第 2 點的第一版有個洞：meta 殘缺（少一個 form）時沿用會生出沒有 `count`
   的條目，而且**份數對得上所以 `load_meta()` 不會自癒它，會一直錯下去**。
   加了 `_meta_is_reusable()` 這道閘

三個都先寫測試重現、確認 FAIL 再修。另外把 GUI 那條執行路徑的驗證補成
`scripts/probe_local_db_gui.py`（24 項，不打網路），README Index 已同步。

**測試 1396 → 1457。**
