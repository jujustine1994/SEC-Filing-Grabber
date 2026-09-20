# Scripts

開發與維護用的獨立腳本。不屬於主程式流程，可單獨執行。

## Index

| 腳本 | 說明 | 狀態 |
|------|------|------|
| `smoke_test_10.py` | 批次 live smoke test：10 間公司各抓 GAAP，自動檢查 Revenue/Gross Profit/Operating Income/Net Income/OCF/Capex/FCF 是否有值，輸出彙總表 | 啟用 |
| `audit_8k_period_labels.py` | 量化 Item 2.02 8-K 季度標籤的 off-by-one（TODO D4 前半）。用新聞稿內文判定實際財期，與 `_period_to_quarter_label()` 現行標籤比對；同時查 dedupe「保留最舊」丟掉哪些財報。**純文字比對，不呼叫 AI**。結論見 `docs/8k-period-off-by-one.md` | 啟用 |
| `verify_8k_fiscal_labels.py` | 驗證 `src/cli.py press-release` 的 `fiscal_label`（TODO D4 後半）。15 家 × 8 季，檢查期末日抓取率，並確認列清單的 `label` 與 `fiscal_label` **偏移全部是 0**（B5 之後兩條路本來就該對齊；B5 前每家各自是常數偏移 -3~+1，舊值留在腳本的 `LEGACY_OFFSETS`）。同時是 B5 的端對端驗收與 `fiscal_label` 的回歸。**不呼叫 AI**，約 3 分鐘。沒填過進階設定時第一個參數當 identity | 啟用 |
| `excel_golden.py` | **Excel 輸出的逐格回歸驗收。改 `excel_writer` / `excel_formatter` / `ratios` / `fiscal_input` 之前先 `make` 一份基準，改完再 `make` + `check`。** 把 `output/_final/*.xlsx` 讀回成 StatementTable 走真正的寫檔流程重產，比對值＋數字格式＋字型＋粗體＋底色。不打網路。2026-08-14 多語言遷移就是靠它確認 480 條字串搬完後繁中輸出逐格不變 | 啟用 |
| `gen_zh_cn.py` | 從 `src/locales/zh_tw.py` 重產 `zh_cn.py`（OpenCC tw2s + 自訂詞彙表）。**一次性工具，平常不跑**，且會覆蓋手改過的用詞。需要 `uv pip install opencc-python-reimplemented`（不在 requirements，執行期用不到） | 啟用 |
| `打包.bat` + `pack.ps1` | **打包散布用 zip，雙擊 `打包.bat` 即可，不必開 PowerShell、不必叫 AI。** 白名單複製 → 壓成 `dist\SEC-Financial-Fetcher-YYYYMMDD.zip` → 解到暫存目錄跑 **12 項自我驗證**（機敏檔、`.xlsx`、金鑰樣式、非預期 email、內部文件外流…）。**任一項沒過就刪掉 zip 並 exit 1**，不會產出一包不能傳的東西。是 `docs/PACKAGING.md` 的可執行版本，兩邊改動必須同步 | (停用，2026-08-18 CTH：GitHub 連得回來了，改走 clone/pull 發布，zip 打包暫不用；腳本保留待日後需要) |
| `gen_template_coverage_baseline.py` | **產出模板體檢基線** `docs/template-coverage-baseline-<日期>.md`：每家公司的缺漏判斷、最常出問題的列、逐列覆蓋率（現行路徑 vs companyfacts）。不打網路，吃 `output/_spike/` 的快取。改了模板 concept 對照或 `data_quality` 判斷規則之後重跑。**看數字前先讀產出文件的第零節**——達標列數只是體溫計，真正的 KPI 是〔真缺口〕與〔假警報〕，而且永遠不該以 97/97 為目標 | 啟用 |
| `watchdog_h0_baseline.sh` | **H0 基線重建的無人值守看門狗**（2026-09-04 夜間作業）。detached shell process，跟 AI session 的額度無關：等 `spike_derive_mapping.py` 那輪 201 家跑完（判定 log 出現最後一行「完整候選：」，或 log 20 分鐘沒更新視為 process 已死），接著掃抓取警告（D11）、自動跑 `gen_template_coverage_baseline.py`、把達標列數／H1 的 `from_ytd` 填滿率／三分類／假警報撈成 `output/_spike/h0_summary.txt`。**設計目的是讓成果不依賴 AI 還活著**——中途斷線的話下一個 session 直接讀那份摘要接手。`nohup bash scripts/watchdog_h0_baseline.sh > output/_spike/watchdog.log 2>&1 &` | 啟用 |
| `diag_probe.py` | **最常用的排查工具**：印出某家公司某張報表裡符合正則的列，含 `concept`／`standard_concept`／`label`／數值。`ARCHITECTURE` 那套「三步排查順序」的第 2、3 步就靠這支——先確認那一列在不在報表 dataframe 裡，再看 matcher 為什麼沒命中。2026-08-23 的 H3 修復全部從這裡開始 | 啟用 |
| `diag_rowprobe.py` | 某個**模板列**在多家公司的命中情況，同時列出 dataframe 裡所有長得像的候選。判斷一列的 concept 對照要不要改，看這支的輸出最快。**H6（2026-08-25）的抽查全靠它**：改 hint 前先用它回頭核對原始 10-Q，確認那一列真的是要的科目（UNP/IP/LIN/EXC），也是靠它推翻分類表「CS&APIC 那 7 家是 concept 層失守」的判讀 | 啟用 |
| `diag_hintsweep.py` | 掃出「`label_hint` 太窄把正確答案濾掉」的模板列——比較有 hint 與沒 hint 的命中差異。**2026-08-23（H3，22 家）與 2026-08-25（H6，201 家）兩輪修復都是從這支開始**。H6 用它量出前後：Capex 15→3、CS&APIC 14→2、Cash 20→15、Cost of Revenue 36→30，其餘 10 條有 hint 的列一筆沒變。201 家清單與原始輸出留在 `output/_hintsweep_201/`（**不要重跑，一輪 12 分鐘**）：`TICKERS=$(cat output/_hintsweep_201/tickers_joined.txt)`。⚠ 有些 hint 是必要的（擋現金流量表末尾的租賃補充揭露列、擋庫藏股列、擋銀行的人事費），不要看到被殺就拿掉——**改完一定要重跑一次比對 killed 清單有沒有長出新的** | 啟用 |
| `check_fye_drift.py` | **量「公司改過財年」對 8-K 零下載規則的風險（TODO B6）**。B5 的季度標籤是用 EDGAR 的 `fiscal_year_end` 現值回推的，公司改過財年的話，改制以前的申報會整段標錯。這支從 `output/_spike/facts_*.json` 取每份 10-K 的財年結束日，跟最新那年的月日比，超過門檻（預設 14 天，52/53 週制本來就會浮動 7 天）就判定改過財年。**零網路請求**，201 家跑幾秒。同時量「改制前那些季會標錯幾季」（發布日用「期末日 + 28 天」代入）。2026-08-25 首跑：201 家裡只有 **LHX**（2019 從 6 月底改成 12 月底，改制前 30 季**全錯**，一律差 2 季）與 **MSCI**（2010 從 11/30 改成 12/31，位移 31 天不跨季，0 季錯）兩家；其餘 199 家最大偏移都在 9 天內（52/53 週制的正常浮動） | 啟用 |
| `diag_celldiff2.py` | **改 concept 對照的回歸驗收**：兩份答案卷快取逐格比對。改之前先把 `output/_spike` 複製一份當基準，改完重建再比；驗收標準是「不能有任何一格從有值變成不同的值或空」。⚠ 鍵用 `(列名, 第幾次出現)`——`Net Income`／`SBC` 在 IS 和 CF 各有一列，用列名當鍵會拿 IS 那列比 CF 那列，2026-08-24 這樣憑空生出 3,659 個假異動 | 啟用 |
| `spike_verify_mapping.py` | **TODO G11 第四步（驗收）**：用 `src/facts_mapping.py` 實跑，跟現行路徑的快取答案卷逐格比對。**完全不打網路**，幾秒跑完 52 家。分開統計「數字相同」與「只差正負號」——那是慣例對不齊不是抓錯，處理方式完全不同。輸出 `output/_spike/verify_mapping.xlsx`（每列×每家命中率，<80% 紅、<95% 黃） | 啟用 |
| `spike_validate_facts.py` | **TODO G11 第三步**：不依賴現行路徑的獨立驗證（反推有個先天限制——現行路徑錯的地方比對也會跟著錯）。四項檢查：會計恆等式、四季加總=年度、**SEC 官方 `frame` vs 我們的期中點判準**、重編頻率。只打 companyfacts 所以能一次驗二三十家。實測 24 家：frame 對齊 59,564/59,564 全數一致 | 啟用 |
| `spike_companyfacts_diff.py` | **TODO G11 決策依據**：SEC companyfacts API 路徑 vs 現行「逐份解 filing」路徑的逐格比對。不改任何現有程式。輸出耗時對照、逐格異同、模板列覆蓋率、Q4 由 10-K 直接 tag 的比例。實測 NVDA：companyfacts 0.51s vs 解 filing 109s（**215 倍**） | 啟用 |
| `spike_derive_mapping.py` | **TODO G11 第二步**：用現行路徑已知正確的數字當答案卷，反推 companyfacts 的 us-gaap concept 對照表。模板的 `std_concept` 欄是 edgartools 正規化過的名字（`NetIncome`），不是原始 element name（`NetIncomeLoss`），憑印象填 75 列一定會錯。這支對每個 concept 算「同期末日數字對得上的比例」，命中率最高的就是正確 mapping，順便偵測正負號相反。結果存 `output/_spike/mapping_candidates.json`，並快取 facts JSON 與現行路徑結果避免重跑 | 啟用 |
| `survey_nongaap_metrics.py` | 調查 32 家（大中小型跨產業）8-K 新聞稿實際使用的 Non-GAAP 指標，統計跨公司覆蓋率，決定 `Data_NonGAAP` 固定模板要收哪些行。**不呼叫 AI**（純文字比對，不吃配額）。原文會存到快取目錄，調整比對規則後可重跑分析不必重新下載 | 啟用 |
| `run_localdb_batch.sh` | **分段跑「更新本地庫」，每段一個獨立 process**。2026-09-06 實測：一個 process 連跑 67 家會在第 16 家被系統因記憶體不足中止（edgartools 的內部快取跨公司累積，我們的 `_parse_cache_scope()` 只涵蓋單次抓取，擋不住）。切段之後每段結束記憶體整個還給系統。**J5 這種大批量一律用這支，不要一個 process 硬幹。** 中止不會白費——`save_filing()` 逐份即時落檔，已完成的下次整家跳過。`bash scripts/run_localdb_batch.sh <輸出前綴> <每段幾家> <ticker...>` | 啟用 |
| `audit_local_db.py` | **本地財報資料庫的體檢（TODO J5 的前置與收尾工具）**。回答「快取裡哪些公司已經完整且最新」——判斷邏輯直接用 `local_db.plan_ticker()`，所以結論跟 `update-db` 會不會跳過**保證一致**。三個條件：10-Q 與 10-K 都到底、SEC 上沒有本地缺的新 filing、快取是現在這個 edgartools 版本解的。**只列 filing 清單、不下載任何 filing**，34 家約 1 分鐘。`--plan-next N --universe <檔>` 會照字母序挑出「還沒抓過的前 N 家」給分批跑用（可重現，分批之間不漏不重）。回傳碼 0＝全部完整、1＝有不完整的（清單可直接餵給 `update-db`）。**J5 每跑完一批就用它收尾**——`update-db` 跑一輪不保證到底（實測 ACN 要兩輪） | 啟用 |
| `verify_local_db.py` | **本地財報資料庫的格式驗證（J5 每批跑完必跑）**。`audit_local_db.py` 答「份數夠不夠」，這支答**「這些檔案能不能用」**——一堆壞掉的檔案份數一樣是滿的。逐檔實際呼叫 `filing_cache.load_filing()` 走五道閘（回 None 就代表下次會整份重抓、那份等於白抓）。⚠ 第五道閘 `require_keys` 走**預設值**＝核心三張（IS/BS/CF），所以這支答的是「核心三張能不能用」；`statement_of_equity` 那類 extras 沒抓到**不會**被判成問題，再把命中的餵給 `cached_filing()` 真的取出三張 DataFrame，最後比對 `_meta.json` 與目錄。**直接用正式程式碼的函式判，不另外寫一套**，所以不可能跟實際行為分岔。不連網，9,233 份約 110 秒。⚠ 「空殼」（三張表全 None）**不算問題**——那是 SEC 的 XBRL 三階段強制時程造成的上游現實，已實測 7 份重抓結果完全一樣。首跑（batch 1，134 家 9,233 份）：四道閘 100%、0 家有問題 | 啟用 |
| `probe_local_db_gui.py` | **「更新本地庫」那條 GUI 路徑的 Tk 探針（TODO J1-J4）**。專案現況是 GUI 不寫自動測試，這支補上「動起來之後」的驗證：按鈕鎖 → 背景執行緒 → `msg_queue` → `db_done` → 按鈕解鎖 → log 內容，外加兩道前置檢查（名單空的、identity 沒填）、例外路徑也要解鎖、J4 版本提醒對話框的四個條件（有講幾家幾份、有回退指令、**沒有**「照用舊快取」的選項、版本相符時不跳）。`local_db.update_local_db` 整個換掉，**不打 SEC**，幾秒跑完，24 項全過才回 0。改 `main.py` 的 `_start_local_db_update`／`_local_db_worker`／`_poll_queue` 的 `db_done` 分支之後跑一次。不會寫你的 config.json | 啟用 |
| `probe_db_overview_gui.py` | **「資料庫總覽」分頁的 Tk 探針（TODO J7）**。純函式（`db_overview_cells`／`db_overview_csv`／`overview_rows`／排序篩選）在 `tests/test_local_db.py` 已經測過，這支補 Treeview 動起來之後的事：切頁觸發重讀 → 表格有列 → 點標題排序（同欄再點一次反向）→ 搜尋框即時篩選 → 匯出 CSV（有 BOM、只匯出篩選後的、標題寫「財報期間」）→ 空資料庫不炸。**特別釘一條**：`_on_tab_changed` 必須用分頁元件身分比對、不可寫死 index——原本那行是 `== 3`，插入這個新分頁時會靜默壞掉（快取面板再也不刷新且不報錯）。自己在 tmp 造假資料庫（`SEC_LOCAL_DB_ROOT` 導過去），**不碰你真正的 `local_db/`**，不連網，**26 項**全過才回 0。2026-09-18 擴充：沒選就按「更新選中的」只跳提示、只把選中的傳下去（不是傳 None＝跑整份名單 218 家）、「上次查」欄有值、財報期間欄顯示的是 03 季末而不是 05 收件月 | 啟用 |
| `probe_foreign_issuer.py` | **外國私人發行人到底抓不抓得到**（TODO D9）。給幾個 ticker，實測三件事：有沒有向 SEC 申報（ADR Level I 走 Rule 12g3-2(b) 豁免＝查無代號）、用 us-gaap 還是 IFRS（前者現有科目對照表直接能用）、6-K 裡有沒有季報（判準是 `R*.htm` 檔數量 > 1，ARM 實測 33/33 完全分離）。⚠ **不要抽樣一份就下結論**——查 ARM 時第一次剛好抽到沒財報的 6-K，差點做出錯誤結論，掃完 33 份才看到真相。2026-09-18 實測六家的結果見 `docs/TODO.md` D9 | 啟用 |
| `check_filer_type.py` | **加新公司進更新名單之前先跑這支**（TODO D9）。問 SEC 每家有幾份 10-Q／10-K／20-F，找出這個工具抓不到的公司。**2026-09-19 改成四分類**（D9 A 路線上線後）：有 10-Q＝正常／無 10-Q 但有含財報的 6-K＝FPI 但抓得到（ARM 這型，會多問一次 6-K 清單並用 R 檔過濾）／只有 20-F＝只抓得到年報／查無代號＝ADR Level I。⚠ **不可以靠註冊地或名稱猜，兩個方向都會錯**——2026-09-18 實測：NXPI（荷蘭）／LIN（英國）／ACN（愛爾蘭）／TEL（瑞士）註冊地都在國外但**全部報 10-Q**；反過來 ARM Holdings 看起來像正常美股，**實際只有 20-F**，當時憑印象列清單就漏掉了它。決定身分的是股東結構與管理層所在地不是註冊地，只能問 SEC。218 家約 5 分鐘。首跑結果：ARM／ASML／STM 只有 20-F、IFX 查無代號，四家已移出名單 | 啟用 |
| `overnight_update_db.ps1` + `過夜更新資料庫.bat` | **整夜重抓整份更新名單，雙擊就跑，全程不需要 AI**（跟 `watchdog_h0_baseline.sh` 同一個原則：成果不依賴 AI 還活著，也不吃 AI 額度）。分段跑（預設每段 8 家、每段一個獨立 process）→ 自動跑第二輪 → 呼叫 `overnight_summary.py` 出摘要。**跑兩輪是必要的**，`update-db` 一輪不保證抓齊（實測 ACN 要兩輪）。用 `SetThreadExecutionState` 擋系統睡眠（螢幕可以關），腳本結束自動解除。log 存 `output/_localdb/overnight_<時間戳>.log`。中斷不會白費，逐份即時落檔，重跑從斷點接。⚠ 三個 PowerShell 5.1 的坑已經踩過並註解在程式裡：`.ps1` **必須有 UTF-8 BOM**（否則中文全亂碼）、`0x80000001` 要明寫 `[uint32]`、原生執行檔的 stderr 合併**交給 cmd 做**不要在 PS 裡寫 `2>&1` | 啟用 |
| `overnight_summary.py` | 整夜更新的收尾摘要，`overnight_update_db.ps1` 呼叫，也可單獨跑：`overnight_summary.py <輸出目錄> <時間戳>`。回答「早上起來只想知道的那幾件事」：最後一輪跳過幾家（愈多愈好＝已收斂）、哪幾家失敗（**ASML／STM／IFX 是 20-F 外國申報人，失敗是預期的**，會分開列不混在紅字裡）、哪幾家有 D11 缺漏、資料庫多大、還有哪幾家沒到底 | 啟用 |
| `check_excel_repair.ps1` | **驗證一份 `.xlsx` 會不會被 Excel 判定內容毀損**（TODO A/F8 修復用）。用 Excel COM 開檔，比對 `%TEMP%` 底下 `error*.xml` 修復日誌開檔前後的變化，並清點 `Chart_*` 分頁還剩幾張圖。**實測發現**：壞檔會讓 `Workbooks.Open()` 直接丟 COM 例外（不是卡對話框，Open() 呼叫本身快速失敗），這個訊號比等修復日誌更乾淨，用來跟已知正常的 `.xlsx` 對照最快。`powershell -File scripts/check_excel_repair.ps1 -Path <絕對路徑>`，回傳碼 0=OK、2=REPRODUCED（判定毀損） | 啟用 |

---

## 測試方案對比：smoke_test_10.py vs tests/test_live_snapshots.py

| | `scripts/smoke_test_10.py` | `tests/test_live_snapshots.py` |
|---|---|---|
| **執行方式** | `python scripts/smoke_test_10.py` | `python -m pytest -m slow` |
| **用途** | 人工快查：最新季數值有沒有抓到 | 自動化迴歸：程式行為是否符合預期 |
| **公司** | AAPL/MSFT/TSLA/AMD/NVDA/GOOGL/META/WMT/COHR/AMZN（10 間） | MSFT/AMZN/META/GOOGL/NVDA/JPM/GS/JNJ（8 間）+ CF overflow 組（COHR/LITE/AAPL/NVDA/GOOGL） |
| **抓取筆數** | `max_filings=80`（完整，偵測長期資料） | `max_filings=8`（只抓最新 8 季，省時） |
| **輸出** | 彩色 terminal 表格，直接顯示數值 | pytest PASS/FAIL，失敗才顯示原因 |
| **判斷標準** | 最新季 7 個 key rows（Revenue/Gross Profit/Operating Income/Net Income/OCF/Capex/FCF）全非 None | key rows 近 4 季有 ≥1 非 None + B1 overflow 結構完整 + CF YTD 減法正確 |
| **耗時** | 較長（完整抓取） | 約 12 分鐘（8 季/ticker） |
| **適合場景** | 開發後手動驗收、懷疑某公司資料有問題時 | 改動 fetcher 後確認沒有迴歸 |
