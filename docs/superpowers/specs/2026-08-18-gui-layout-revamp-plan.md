# GUI 版面/易用性整理計畫（規劃用，2026-08-18）

> 本檔案只做**盤點與分組規劃**，這個對話沒有動任何程式碼。下一個 session 依此執行。
> 涵蓋 TODO：E4、E6、E7、E8、E9、E10、E11、E12、E13、E14（`docs/TODO.md`）。
> **前提不變：不動功能，只動版面/視覺/回饋。**
>
> **✅ 組 A／組 B／組 C 都已於 2026-08-18 完成**，見 `docs/CHANGELOG.md`。
> 組 A：E8/E9/E10/E15。組 B：E4/E6/E11 footer/E16（E11 的「關閉前未儲存提示」
> 子問題沒做，範圍還沒問過 CTH）。組 C：E13 Watchlist 視窗三項可用性問題。
> 這個檔案的分組規劃全部執行完畢，剩下的都不是「版面重排」範圍內的項目了。

## 這次調查已查到的根因（省下次重查）

- **E9（抓取中重複按按鈕跳彈窗）**：`_run_single`／`_run_batch` 觸發的
  `_start_worker`（`main.py:1824-1837`）已經會在開跑時
  `btn_run_single.config(state="disabled")` / `btn_run_batch` 同步 disable，
  這條路徑本身沒問題。**真正跳彈窗的是「查可用期間」按鈕**：
  `_run_preview_scan`（`main.py:1734-1746`）在 `self.is_running` 為 True 時
  跳 `messagebox.showwarning(...wait_for_current_run...)`。`_scan_btn` 只在
  自己掃描時 disable 自己（`main.py:1747-1748`），不會因為「抓取正在跑」而
  disable。**修法方向**：`_start_worker` 開跑時順便 disable `_scan_btn`，
  收工時（現有收尾邏輯裡有把 `btn_run_single`/`btn_run_batch` 開回來的地方）
  一併把 `_scan_btn` 開回來，`_run_preview_scan` 裡那段 `is_running` 檢查與
  彈窗就可以整段刪掉。

- **E10（進階設定滾輪失效）**：全檔案搜尋 `<MouseWheel>` **零命中**——
  `_build_fixed_height_scrollable`（`main.py:34-59`，Tab3 設定頁用的就是這個
  容器，見 `main.py:1278`）只做了 Canvas + Scrollbar，從沒綁滑鼠滾輪事件。
  使用者只能拖 scrollbar，滾輪完全沒反應。**修法方向**：在
  `_build_fixed_height_scrollable` 裡幫 canvas 綁 `<MouseWheel>`（Windows 上
  `event.delta` 是 120 的倍數，除以 120 再乘負號去對應 `yview_scroll`）。
  這個容器目前有兩處在用（Tab3 設定頁 `main.py:1278`、Tab1 的 sheet 面板
  `main.py:524`），改共用函式兩處一起修好，不用分別改。

- **E11（存檔按鈕太下面／關掉沒提示）**：Tab3 設定頁的存檔按鈕在
  `_build_settings_panel` 最後一行（`main.py:1400-1405`，`btn_row` 是
  `popup`／即可捲動容器裡的最後一個 row），使用者要捲到底才看得到。註解
  `main.py:1398-1399` 說明「頁籤沒有可以關掉的視窗，取消鍵按了什麼都不會
  發生」——這代表**「關掉沒有提示」講的其實是關主視窗**（`self.root` 的
  `WM_DELETE_WINDOW`），不是這個 Tab。要查主視窗關閉時有沒有處理常式、有
  沒有檢查 Tab3 欄位是否跟 `self.cfg` 已存的值不同。
  **修法方向**：(1) 存檔按鈕移到 Tab3 的固定 footer（跳出
  `_build_fixed_height_scrollable` 的可捲動區，跟捲動內容分離，永遠可見）
  ——這個做法同時緩解 E10 沒修好時看不到存檔按鈕的問題；(2) 查主視窗關閉
  handler，若無，評估要不要加「有未存設定」的提示（範圍要問 CTH：只比對
  Tab3 欄位還是含 Tab1/2 的執行參數？後者通常不算「設定」，不用比）。

- **E13（Watchlist 視窗）**：群組列的圖示按鈕在 `main.py:1020-1030` 附近
  （`重新命名`/`刪除群組`旁的圖示按鈕），目前只有圖示沒有文字/tooltip，
  且按鈕排列導致文字 overflow（見原始截圖）。「名稱庫：尚未建立」那行文字
  待查其對應程式位置（在 `_open_watchlist_popup`，`main.py:908` 起）。

- **E12（GAAP 抓取進度回饋）**：`main.py` 已經有 `progress_bar` 元件與
  `progress_label`（`_start_worker` 裡 `main.py:1832-1833` 會重置
  `progress_bar["value"] = 0` 並顯示 `t("gui.status.preparing")`）——**進度條
  骨架已經存在，不是從零開始**。下一步要查的是：(1) 抓取過程中
  `progress_bar["value"]` 有沒有被實際更新過（估計是目前只在開始時歸零，
  中途沒人推進，所以「感覺像卡住」）；(2) `fetcher_gaap.py` 有沒有現成的
  「第幾份／共幾份」計數可以拿來當進度依據。這條**不是純版面調整**，會碰
  worker thread 與主執行緒的通訊（多半已經有機制在推 log，比照同一條路徑
  推進度即可，風險應該不大，但要先讀 `_worker_single`/`_worker_batch` 怎麼
  回報進度）。

- **E7（安裝元件不透明）**：屬於 `launcher.ps1`（打包/啟動流程），不在
  `main.py` GUI 範圍內，跟其他項目沒有版面耦合，**可以獨立排期**，不用綁在
  這次 GUI 重排的順序裡。

## 補充：設定架構盤點（2026-08-18 討論，非純版面問題）

CTH 提出「很多按鈕該怎麼放、有些設定不該藏太深、以及使用直覺性」，這條跟
E14 版面調整不是同一件事——版面是排列/視覺，這條是**架構**：同一類設定
散在不同地方、命名衝突、副作用不明顯。盤點結果：

### 現有設定去向（config.json 欄位 vs GUI 位置）

| 設定 | 是否全域持久化 | 現在在哪改 | 問題 |
|---|---|---|---|
| `language` | 是 | Tab3 | 正確位置 |
| `identity`（SEC EDGAR Identity，必填） | 是 | Tab3 最底層 | 藏太深，Tab1/Tab2 沒有任何提示 |
| `ai.provider` / `model` / `api_key` | 是 | Tab3 | 正確位置（Non-GAAP 停用中，優先度低） |
| `max_filings` / `template_path` | 是 | Tab3 | 正確位置 |
| `output_dir` / `filename_format` / `filename_custom` | 是 | **只有 Tab1** | Tab2 批次看不到也改不了，但批次照樣吃這個全域值 |
| ticker / GAAP·NonGAAP 勾選 / 期間範圍 / Q·K 勾選 | 否（單次執行參數） | Tab1、Tab2 各自一份 | 各自獨立跑的參數，重複合理，不用動 |

### 找到的三個問題與決定（已跟 CTH 對過，下次直接照做）

1. **「進階設定」命名衝突**：Tab1/Tab2 裡展開 Q/K 勾選的小按鈕，跟整個
   Tab3 頁籤用同一個名字，兩者完全不相關。
   **決定**：把 Tab1/Tab2 那個小按鈕改名（例如「報表類型」），
   `t("gui.btn.adv_collapsed")` / `t("gui.btn.adv_expanded")`
   （`main.py:489, 593, 596, 602, 605, 682`）目前是共用同一組 i18n key，
   改名要新增獨立的 key，不能直接沿用，否則 Tab3 的頁籤標題也會跟著變。
   四個 locale（zh_tw/zh_cn/en/ja）都要加。

2. **SEC Identity 沒有任何提示**：位置本身沒問題（全域系統設定就該在
   Tab3），問題是 Tab1/Tab2 使用者填了 ticker 按下去才會因為沒填 identity
   失敗。程式裡已有現成 pattern 可抄：Non-GAAP 沒填 API Key 時 Tab1 會跳
   一行橘字提示（`nongaap_warn_label`，`main.py:528-533`，由
   `_on_nongaap_toggle` 控制顯示/隱藏）。
   **決定**：比照做一行「⚠ 尚未設定 SEC Identity」提示，`cfg["identity"]`
   空的時候顯示，點了直接切到 Tab3。Tab1、Tab2 都要加（Tab2 目前完全沒有
   對應機制）。

3. **輸出資料夾/檔名格式被「悄悄」存成全域預設值**（這次盤點才發現，不是
   CTH 原本提的）：`main.py:832-838` 顯示 Tab1 每次按「執行」就把當下
   輸入框的值寫回 `config.json` 當全域預設，不是使用者主動存檔的副作用。
   Tab2 批次完全看不到這個全域值是什麼、也沒有地方改，抓出來的檔案卻是
   照著它存。
   **決定（CTH 選的方向）**：**位置不搬**，維持在 Tab1，但拿掉「按執行
   就悄悄存成全域預設」這個隱性行為——改成需要明確動作才會存成預設（例如
   旁邊加一顆「設為預設」按鈕）。Tab2 加一行唯讀文字顯示目前全域預設是
   什麼，讓批次使用者至少看得到抓出來的檔案會落在哪。
   **要查清楚的細節（動手前）**：目前「執行」的當下，`tab1_outdir_var`
   當次覆寫用的路徑（`_get_output_path` 的優先順序是
   watchlist item output_dir → ticker_paths → 全域 output_dir，
   `main.py:1567-1592`）跟「存成新全域預設」要拆成兩件事——按執行永遠
   用當下輸入框的值跑這一次，「設為預設」按鈕才動 `save_config`。

### 這三項跟組別的關係

三項都不碰 `self.root` 的 grid、也不是 Tab3 內部捲動容器的版面，風險低、
範圍侷限在各自的 Tab frame 內，**併入組 A**（跟 E8/E9/E10 一起，第一輪
就做，互不干擾）。i18n key 新增記得比照 `docs/ARCHITECTURE.md` 多語言章節
的鐵則：機器鍵/i18n key 用英文，四語都要補。

## 分組（哪些該一起做、哪些互相獨立）

### 組 A — 孤立小修，風險低，可以最先做且互不干擾
- E9 disable 掃描按鈕（刪掉彈窗那段）
- E10 幫 `_build_fixed_height_scrollable` 補滾輪綁定（兩處容器共用受益）
- E8 `ticker_entry` width 從 12 調大（`main.py:450`）
- 設定架構 1：Tab1/Tab2「進階設定」小按鈕改名，避免跟 Tab3 撞名
- 設定架構 2：Tab1/Tab2 加 SEC Identity 未設定的提示（仿 `nongaap_warn_label`）
- 設定架構 3：輸出資料夾/檔名格式改成明確「設為預設」，Tab2 加唯讀顯示

這幾項互相不碰同一塊版面，改動範圍都侷限在各自 Tab 內，建議第一輪一次做
完、一次驗證（含 E10 順便解掉「看不到存檔按鈕」的一半問題）。

### 組 B — 主視窗共用 grid/pack 版面，必須一次規劃、不要分批動
TODO 原文已點名 E4（視窗寬度亂跳）與 E6（log 區太扁）是同一片版面
（`self.root` 的 grid，`main.py:438` 附近），E14（整體重排視覺）性質相同。
這次盤點確認 **E11 的存檔按鈕搬到固定 footer** 也會動到 Tab3 的內部版面
（跳出捲動容器），雖然是 Tab3 局部、不碰 `self.root` 的 grid，但如果組 B
真的要大搬風，Tab3 footer 順手一起做，避免「Tab3 剛调完版面、隔週又因為
E4/E6/E14 重排一次」。

**執行前必讀**：`C:\Users\CTH\.claude\project-rules\windows-tool\tkinter-ui\INDEX.md`
（`main.py` 內部註解本身也這樣要求，見 `docs/TODO.md` E6）。

**已跟 CTH 對過的兩個方向（下次直接照做，不用重問）**：

1. **log 區太扁怎麼解**：**拉高視窗＋捲動**。不做拖曳分隔（PanedWindow）、
   不把 log 拉成獨立分頁——直接放寬固定高度限制，log 區高度不夠時允許
   捲動。CTH 另外提到：**順便評估整個視窗要不要加寬**，不是只有高度，
   要重新看現在 900px 寬是不是也卡到其他區塊（Tab2 的 Watchlist 列、
   Tab3 的 Entry 欄位這些目前寬度都是各自估的，加寬後可能都要重新排）。
   **待查**：目前 `__init__` 的 geometry 是寫死常數還是有算過 `SPI_GETWORKAREA`
   的動態邏輯（`docs/TODO.md` E4 提到用 Win32 API 算開窗座標），改高度/寬度
   常數前要先讀那段，不要只改數字沒看邏輯怎麼算的。

2. **Tab3 存檔按鈕搬固定 footer 後要不要加取消鍵**：**加「還原」按鈕**
   （不是「取消並關閉」，因為 Tab3 是頁籤沒有可關閉的視窗）。行為：把
   目前頁籤上的欄位值改回上次 `save_config` 存的值，讓使用者可以反悔
   本次編輯，不用真的離開頁籤或重啟程式。實作時要注意每個欄位現在的
   初始值是從 `self.cfg.get(...)` 讀的（`main.py:1310-1391`），「還原」
   等於把這段初始化邏輯再跑一次、把每個 `*_var.set()` 回填，逐一列清楚
   有哪些變數要還原（language / identity / ai 三欄 / max_filings /
   template 兩個變數），不要漏掉任何一個。

### 組 C — Watchlist 視窗（E13），獨立視窗、不碰組 A/B
可以跟組 A 同一輪做，也可以獨立排在組 B 前後，互不影響。範圍：
- 群組列圖示按鈕加文字說明或 tooltip
- 修 overflow（按鈕排列方式，或改 CTH 建議的「Windows 慣例、新增/編輯/刪除
  統一靠右」）
- 「名稱庫：尚未建立」那行文字加說明或改措辭

### 組 D — 進度回饋（E12 + E7），不是版面調整，是功能/回饋層級的活
跟組 A/B/C 完全獨立，可以另開時段處理，不用等版面整理完。E12 風險中等
（碰 worker thread 進度回報），E7 是打包腳本，兩者互相無關，各自獨立排期。

## 下一個 session 建議開場順序

1. 讀本檔案 + `docs/TODO.md` 對應項目確認沒有過時
2. 組 A（E8/E9/E10 + 設定架構 1/2/3，共六項）：改完立即用 `run` skill 或
   手動起 GUI 驗證（滾輪、重複按查詢、ticker 欄位、identity 提示、輸出
   設為預設都要實際點過，不能只看 code）
3. 組 C（E13）：獨立視窗，改完單獨驗證
4. 組 B（E4/E6/E11 footer/E14）：**先問 CTH 上面兩個討論題**，有共識再動手，
   一次規劃整個 `self.root` grid，不要分批
5. 組 D（E12/E7）：視時間，可另開 session
