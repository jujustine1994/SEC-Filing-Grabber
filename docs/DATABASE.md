# SEC 財報資料庫

程式與財報分開保存。預設資料庫位於 Windows 實際 Documents 目錄下的 `SEC財報資料庫`，不在程式專案、AppData、output 或快取目錄裡。換程式版本或重新下載專案後，連接同一份資料庫即可。

## 開始使用

GUI 的進階設定提供「連接資料庫」「新建資料庫」「搬遷舊資料庫」「建立快照」。新機器須明確新建或選取既有庫；程式不會在失聯時默默建立空庫。資料庫區顯示實際路徑與 UUID。

資料庫目錄要在程式目錄外；新建與搬遷目的地必須是空目錄。不要選整個 Documents 作為資料庫，要選其下獨立的新資料夾。遇到 symbolic link、junction 或 reparse point，操作會停止。

PowerShell 在程式專案內執行：

```powershell
$Py = Join-Path $env:USERPROFILE 'venvs\SEC Financial Tools\Scripts\python.exe'
& $Py src/cli.py db-connect 'C:\Users\CTH\Documents\SEC財報資料庫'
& $Py src/cli.py db-status --json -
& $Py src/cli.py update-db --list
& $Py src/cli.py update-db NVDA AAPL
```

`update-db` 沿用既有 SEC Identity 與固定深度，只新增缺少財報；名單保存於資料庫 `metadata/update_list.json`。移除公司只停止更新，不刪除已保存資料。`gaap`、跨公司比較、過夜更新與 revenue skill 共用同一個連接。

## 保存規則

```text
SEC財報資料庫/
  database.json                UUID、格式、建立時間
  filings/<TICKER>/            目前財報與公司管理資訊
  history/<TICKER>/<accession>/ 歷史財報，SHA-256 命名
  metadata/update_list.json    更新名單，包括還沒有財報的公司
  metadata/migration.json      搬遷校驗報告
  staging/                    暫存、跨程序寫入鎖、遷移清單
  snapshots/                  本機完整快照
```

正式程式沒有清除單一公司／全部公司的功能；舊 `clear_ticker`、`clear_all` API 也會拒絕。管理資訊可以更新；財報重新解析前必須先核對並保存舊檔，保存失敗不替換目前資料。多個程序透過同一個 OS 檔案鎖提交，程序中斷會釋放鎖。鎖檔本身不是財報，不要手動刪除正在使用的鎖檔。

自動清理只處理 staging 中超過一小時的 `.tmp`，且持有同一個寫入鎖。filings、history、snapshots 不做容量或到期清理。空名單與空庫是有效狀態；損毀／缺少 marker、名單或連接 ID 不符會報錯。

parser 版本相容檢查仍然嚴格。舊版本數據保留下來，但不能因此認為數字正確；不相容時需重新解析，舊內容留在 history。搬遷本身不改財報 schema、parser version 或內容，也不連 SEC。

**模板修正不回寫來源**：庫內 filing JSON 是 edgartools 的保存解析結果，與 SEC 原始申報及模板輸出分層。改選值、拆季、比率不修改已有 filing JSON；解析缺漏的恢復則另作來源核對、獨立驗證及有歷史保留的資料更新。完整界線見 [ARCHITECTURE.md](ARCHITECTURE.md)「來源、解析保存與模板的改動界線」。

## 搬遷

先關閉舊 GUI 與更新程序，再執行：

```powershell
& $Py src/cli.py db-migrate --source local_db/filing_cache --destination 'C:\Users\CTH\Documents\SEC財報資料庫'
```

來源所有檔案原樣複製，逐檔比對 size 與 SHA-256，並再次檢查來源未變動；通過後才登記連接。中斷可以用相同命令接續，目的地既有檔案若校驗不符會拒絕，不能任意覆蓋。已完成搬遷重跑不重新匯入舊更新名單。

舊 `local_db/filing_cache` 保留。`local_db/DATABASE-MOVED.txt` 記錄新位置。不要啟動舊版程式繼續寫舊位置；那份副本不會跟著新庫更新，不能當成長期備份。Git 不包含正式財報，也不能用來還原它。

## 快照與還原

```powershell
& $Py src/cli.py db-snapshot
# 或指定另一顆磁碟的空目的地：
& $Py src/cli.py db-snapshot --destination 'E:\SEC資料庫備份\2026-10-05'
& $Py src/cli.py db-verify-snapshot 'C:\Users\CTH\Documents\SEC財報資料庫\snapshots\實際快照名稱'
& $Py src/cli.py db-restore '快照完整路徑' --destination 'C:\Users\CTH\Documents\SEC財報資料庫_還原'
& $Py src/cli.py db-connect 'C:\Users\CTH\Documents\SEC財報資料庫_還原'
```

快照是目錄式完整副本，包含財報、歷史、UUID 與管理資訊；不包含 staging 或其他快照。建立期間短暫鎖住寫入，核對所有內容後才有有效 `snapshot.json`。建立與還原需要額外磁碟空間。

還原只允許新的空目的地；先驗證檔案清單、路徑與雜湊，再發布 marker，不覆蓋目前資料庫、不自動連接。還原完成後須明確 `db-connect`。

本機同磁碟快照能處理誤操作，不能處理整顆磁碟故障。本次不自動配置第二顆磁碟或雲端備份，也不修改 Google Drive 同步設定；同步不等於獨立可還原備份。

2026-10-05 已建立 `snapshots/20261005_173403_9332e679`，14,636 個檔案全部通過快照與新目錄還原校驗；詳見 [搬遷紀錄](local-db-runs.md)。

## 連接問題

設定在 AppData 的 config.json 保存 `database_path`、`database_id`，財報與名單留在庫中。路徑失效或 ID 不符時，重新選取有效資料庫；不改用舊庫或其他空庫。如果設定 JSON 損毀，只有明確的連接操作會先保留 `.corrupt-<SHA256>.bak`，再建立新連接；原本 SEC Identity／個人設定可能需重新填寫，原始位元組留在 backup。

一般偏好儲存與明確連接共用跨程序鎖；偏好儲存保留磁碟上最新的連接，不會用已開啟 GUI 的舊設定切回前一個資料庫。

CLI 管理命令與 `db-status` 支援 `--config-path`，用於另一份連接設定。`SEC_LOCAL_DB_ROOT` 環境覆寫也必須有有效 marker，不能當成刪除授權。測試建立隔離庫，正式財報永遠不拿來做刪除測試。
