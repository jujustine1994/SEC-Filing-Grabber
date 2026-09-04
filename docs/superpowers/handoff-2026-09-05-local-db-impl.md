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
