# SEC 財報資料庫獨立保存 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 把現有 SEC 財報安全複製到独立資料庫，讓現有程式繼續更新，移除正式刪除能力並保留歷史內容。

**Architecture:** 保留 accession JSON 與既有 builders；新增集中連接、受保護寫入與搬遷／快照模組。所有入口透過有效資料庫 UUID 連接，更新先保存舊內容，再原子切換；不因搬遷升 schema 或重新下載 SEC。

**Tech Stack:** Python 3、標準函式庫、現有 pandas／Tk／pytest／edgartools；Windows 使用 Known Folder 與 `msvcrt` 檔案鎖，其他平台使用 `fcntl`。不新增套件。

**Spec:** `docs/superpowers/specs/2026-10-05-independent-sec-database-design.md`

## Global Constraints

- 正式資料不能被程式的快取清理功能刪除。
- 保存格式不升 `SCHEMA_VERSION`，不因搬遷觸發 SEC 重抓。
- 原 `local_db/filing_cache` 保留為搬遷前副本，實作腳本不得自動刪除。
- 正式財報、history、snapshots 不做容量／到期清理。
- 第二顆磁碟或獨立雲端目的地尚未由使用者指定，完成報告明確標為未配置。
- 不接受任意非空資料夾。遇到 reparse point／symlink 停止，不跨出來源範圍。
- 歷史內容不可在沒有歷史保存的情況下覆寫。
- 不宣稱任何本機資料夾能防止管理員刪除、硬碟故障或惡意程式。

## Review Focus

- AppData 設定損毀或資料庫 ID 不符：應失聯報錯，不能自動改接預設庫。Task 1。
- 同時開 GUI／CLI、history 損毀或磁碟滿：原財報仍完整，不能誤報更新成功。Task 2。
- 空更新名單與未匯入不同：停止追蹤全部公司後，不得再次匯入舊設定。Task 3。
- 搬遷中來源改變、重跑半成品、來源含 junction：不得登记未核對的目的地。Task 4。
- 快照含路徑穿越或 hash 不符：還原不得寫出目的地外，也不得覆蓋現用資料庫。Task 4。

## 檔案分工與公共介面

新增 `src/database.py`：位置、marker、連接與測試庫身分；不依賴 pandas 或 GUI。
新增 `src/database_io.py`：跨程序鎖、staging、歷史保存與受控管理狀態寫入。
新增 `src/database_transfer.py`：hash 清單、複製搬遷、快照與還原，不負責 UI。
保留 `filing_cache.py` 的讀取／DataFrame 介面作相容層；`local_db.py` 保持更新流程。
CLI 與 GUI 只做操作封裝。分析工具解析 repo 的程式模組，不自行拼資料目錄。

測試共用 `tests/database_helpers.py` 定義 `make_database(tmp_path, *, test_mode=True) -> Path`，明確呼叫 Task 1 的 `create_database`；既有 monkeypatch fixture 要一起建立測試 marker。不能只靠環境變數開啟清理權限。

PowerShell 執行 Python 時使用 `& 'C:/Users/CTH/venvs/SEC Financial Tools/Scripts/python.exe'`；測試統一 `-m pytest -m 'not slow' -p no:cacheprovider`，避免本機 `.pytest_cache` 權限問題。

### Task 1: 資料庫身分、位置與連接

**Files:** Create `src/database.py`, `tests/test_database.py`, `tests/database_helpers.py`; Modify `src/config.py`, `config.example.json`, `src/filing_cache.py` 路徑區、現有測試／探針的隔離 fixture。

**Interfaces:** `DatabaseError(RuntimeError)`；`default_database_path() -> Path`；`read_marker(root: Path) -> dict`；`create_database(root: Path, *, test_mode: bool = False) -> dict`；`connect_database(root: Path, *, config_path: Path | None = None) -> dict`；`database_root(*, config_path: Path | None = None) -> Path`；`is_test_database(root: Path) -> bool`。正式資料目錄為 root/filings；測試同樣採用此結構。

- [ ] **Step 1: 寫失聯／ID 測試與測試庫 helper。**

```python
def test_missing_registered_database_does_not_create_empty(tmp_path):
    import json, pytest
    from database import database_root, DatabaseError
    cfg = tmp_path / 'config.json'
    missing = tmp_path / 'missing'
    cfg.write_text(json.dumps({'database_path': str(missing),
                               'database_id': 'missing-id'}), encoding='utf-8')
    with pytest.raises(DatabaseError):
        database_root(config_path=cfg)
    assert not missing.exists()

def make_database(tmp_path, *, test_mode=True):
    from database import create_database
    root = tmp_path / 'database'
    create_database(root, test_mode=test_mode)
    return root
```

另測 corrupt config、marker UUID 不符、無 marker 環境覆寫、有效空庫、有效預設探索與 Documents 被重新導向。正式 create 不得位於 repo 下；測試庫必須位於測試建立的暫存路徑，禁止既有非空目錄被標成測試庫。

- [ ] **Step 2: 執行 `tests/test_database.py`，確認先因模組缺少失敗。**
- [ ] **Step 3: 實作連接。** marker 使用 `{'format_version': 1, 'database_id': str(uuid.uuid4()), 'created_at': ..., 'test_mode': False}`；嚴格讀取 config 的連接欄位，不沿用損毀 config 回傳 defaults 的寬鬆行為。`database_root()` 無副作用；探索與 connect 是分開操作。Windows `SHGetKnownFolderPath` 使用 Documents GUID `FDD39AD0-238F-46AF-ADB4-6C85480369C7`，釋放 COM 記憶體。連接設定原子保存且保留其他欄位；ID 必须合法 UUID。

```python
def cache_root():
    from database import database_root
    return database_root() / 'filings'
```

- [ ] **Step 4: fixture 先呼叫 create，再設定 `SEC_LOCAL_DB_ROOT`。** 更新原本返回 tmp/filing_cache 的測試為 tmp/database/filings；GUI probe 同樣初始化隔離庫。執行 database、config、filing_cache、local_db 的離線測試，區分新行為失敗與現有依賴問題。
- [ ] **Step 5: commit `feat: connect an independent SEC database by identity`。**

### Task 2: 正式刪除防護、歷史保存與寫入失敗回報

**Files:** Create `src/database_io.py`, `tests/test_database_io.py`; Modify `src/filing_cache.py`, `src/fetcher_gaap.py`, `src/local_db.py`, `src/fetch_ledger.py` 及相關測試。

**Interfaces:** `database_lock(root: Path)` 為可重入 context manager，跨程序取得 root/staging/database.lock；`commit_filing(root: Path, ticker: str, accession: str, entry: dict) -> None`；`write_metadata(root: Path, path: Path, value: dict) -> None`；`clean_staging(root: Path, *, older_than_seconds: int = 3600) -> int`。所有正式資料寫入使用同一個短期全庫鎖，fetch 網路請求在鎖外。無鎖外 metadata 寫入，因此快照能取得一致視圖。

- [ ] **Step 1: 寫入驗收測試。**

```python
def test_formal_database_cannot_be_cleared(tmp_path, monkeypatch):
    import pytest, filing_cache
    from database import create_database, DatabaseError
    root = tmp_path / 'formal'
    create_database(root)
    monkeypatch.setenv('SEC_LOCAL_DB_ROOT', str(root))
    with pytest.raises(DatabaseError):
        filing_cache.clear_all()
    assert (root / 'database.json').exists()
```

history 測試先透過 save_filing 寫一份，再修改數值重寫，斷言 history 的 SHA-256 與先前原始 bytes 一致。patch history 寫入丟 OSError，斷言正式檔 bytes 不變。另測既有 history hash 不符、損毀正式 JSON 被保留、兩個 subprocess 同時提交各自版本、重複內容忽略 cached_at 不產生額外歷史、清理不能碰正式區。

- [ ] **Step 2: 執行新測試，確認現有清除 API 與覆寫行為不符合預期。**
- [ ] **Step 3: 實作鎖與提交。** Windows 使用 `msvcrt.locking`，POSIX 使用 `fcntl.flock`，同程序使用 threading.RLock 及深度計數；OS 鎖在 crash 自動釋放，不採容易殘留的 mkdir 鎖。timeout 30 秒後回報 DatabaseError。持有锁时：序列化 staging UUID tmp、驗證 entry 與 path、先驗證／保存 history、fsync staging、os.replace 正式檔。ticker 使用 `[A-Z0-9][A-Z0-9.-]{0,19}`，禁止路徑跳脫；所有接觸路徑檢查 reparse／symlink。metadata 同樣 staging 原子切換。history 以 SHA-256 命名，用 exclusive create，既有內容必須核對；出錯不得切換。

```python
# save_filing 組裝 entry 後：
try:
    commit_filing(database_root(), ticker, accession, entry)
except (DatabaseError, OSError, ValueError) as exc:
    raise DatabaseError('SEC database write failed') from exc
return True
```

將 fetcher 廣泛吞掉的持久化錯誤分開記為 ledger `persistence_errors`；解析失敗仍遵守既有負向快取規則。`update-db` 只要持久化或 meta 寫入失敗就回報 failed／非零，不把記憶體抓取成功當作已保存。對正式庫的 clear_ticker/clear_all 一律先拒絕；若保留測試 cleanup，僅接受明確 test marker、驗證位於暫存區與有效 ticker。`clear_stale_tmp` 改委派 clean_staging，既有正式目錄 tmp 留存。

- [ ] **Step 4: 跑新測試與 GAAP cache／no-AI／ledger／local_db 離線測試。**
- [ ] **Step 5: commit `feat: preserve SEC filing history and reject database deletion`。**

### Task 3: 更新名單歸資料庫所有

**Files:** Modify `src/local_db.py` 更新名單、`src/main.py` 名單管理、`src/cli.py:cmd_update_db`；Create `tests/test_database_update_list.py`; Modify `tests/test_local_db.py`, `tests/test_cli.py`。

**Interfaces:** `read_update_list(root: Path) -> list[str]`；`write_update_list(root: Path, tickers: list[str]) -> list[str]` 放在 local_db；既有 `get_update_list(cfg)` 等介面保留，cfg 不再是唯一來源。update_list 文件含 format_version、database_id、tickers；缺檔／損毀報錯，不自動回填舊 config。Task 1 的 create_database 必須同步建立 UUID 相符的有效空名單；Task 4 在尚未登記的新庫中以舊名單取代這份空名單。

- [ ] **Step 1: 寫空名單與獨立設定測試。**

```python
def test_empty_database_list_does_not_reimport_old_config(tmp_path, monkeypatch):
    import local_db
    from tests.database_helpers import make_database
    root = make_database(tmp_path)
    monkeypatch.setenv('SEC_LOCAL_DB_ROOT', str(root))
    local_db.set_update_list({}, [])
    assert local_db.get_update_list({'local_db_tickers': ['NVDA']}) == []
```

另測換 config 後名單仍在、未下載 ticker 保留、移除不刪檔、寫入失敗不更新 GUI 記憶體狀態、連續 add/remove 以鎖保護 read-modify-write、文件 UUID 不符。

- [ ] **Step 2: 跑上述測試，確認現有 cfg-only 行為失敗。**
- [ ] **Step 3: 將名單操作改為在 database_lock 內讀取並更新 metadata。** cfg 的 local_db_tickers 僅兼容呼叫者顯示，正式 source of truth 是資料庫；搬遷工具在 Task 4 唯一匯入舊名單。GUI/CLI 不另行持久化名單副本。新建庫明確寫入有效空名單。

```python
def set_update_list(cfg, tickers):
    result = write_update_list(database_root(), normalize_tickers(tickers))
    cfg[UPDATE_LIST_KEY] = result
    return result
```

- [ ] **Step 4: 跑 update-list、local_db、CLI 測試。**
- [ ] **Step 5: commit `feat: store SEC update lists with the database`。**

### Task 4: 可核對搬遷、快照與還原

**Files:** Create `src/database_transfer.py`, `tests/test_database_transfer.py`; 由 Task 5 的 CLI 呼叫，不新增獨立 shell 搬資料實作。

**Interfaces:** `inventory(root: Path) -> dict[str, dict]`（相對 POSIX path → size/sha256）；`migrate_database(source: Path, destination: Path, tickers: list[str], *, config_path: Path | None = None) -> dict`；`create_snapshot(root: Path, destination: Path | None = None) -> Path`；`verify_snapshot(snapshot: Path) -> dict`；`restore_snapshot(snapshot: Path, destination: Path) -> dict`。snapshot 是有 manifest 的目錄式完整副本，避免 zip 大小／路徑擷取問題；目的地不在 snapshots 內時可用來做獨立磁碟副本。

- [ ] **Step 1: 寫搬遷成功與不切換測試。**

```python
def test_migration_preserves_bytes_and_source(tmp_path):
    import json
    from database_transfer import migrate_database
    source = tmp_path / 'legacy'
    (source / 'NVDA').mkdir(parents=True)
    path = source / 'NVDA' / '0001045810-25-000001.json'
    path.write_bytes(b'{"cik":1045810,"accession_no":"0001045810-25-000001"}')
    cfg = tmp_path / 'config.json'
    cfg.write_text('{}', encoding='utf-8')
    dest = tmp_path / 'separate'
    report = migrate_database(source, dest, ['NVDA', 'AAPL'], config_path=cfg)
    assert report['verified'] is True
    assert (dest / 'filings/NVDA' / path.name).read_bytes() == path.read_bytes()
    assert path.exists()
    assert json.loads(cfg.read_text(encoding='utf-8'))['database_path'] == str(dest.resolve())
```

另測 destination/source 重疊、repo 內目的地、非空未知目錄、symlink/junction、copy 後来源改變、hash 不符、磁碟滿、不相符 resume manifest。快照測試建立→驗證→還原至新庫→核對全部財報；manifest 中 `../escape`、絕對路徑、重复路徑、hash 不符與 marker UUID 不一致均拒絕。

- [ ] **Step 2: 跑新測試，確認模組缺少失敗。**
- [ ] **Step 3: 實作純 copy 與驗證。** 先來源 inventory，保存 destination/staging 的完整遷移清單；重跑只能匹配來源、原 hash 清單與目的地。已複製檔驗證後跳過，hash 不符不覆蓋。來源 lock 防本版寫入，另外重新 inventory 偵測舊版／外部寫入；遷移報告與原庫的醒目說明只有最後生成，不能污染先前核對清單。目的地 marker 在全核對完成後发布，connect 最後執行；source 保留。

```python
before = inventory(source)
for relative, expected in before.items():
    target = destination / 'filings' / relative
    target.parent.mkdir(parents=True, exist_ok=True)
    if not target.exists():
        shutil.copy2(source / relative, target)
    actual = {'size': target.stat().st_size,
              'sha256': hashlib.sha256(target.read_bytes()).hexdigest()}
    if actual != expected:
        raise DatabaseError('Copied file differs from migration manifest')
if inventory(source) != before:
    raise DatabaseError('Source changed during migration')
if inventory(destination / 'filings') != before:
    raise DatabaseError('Migration verification failed')
```

snapshot 在 database_lock 下複製 database.json、filings、history、metadata，不含 staging、snapshots 或 session lock；manifest 包含 UUID、時間、清單及清單摘要，完整 hash 核對後才標 verified。restore 拒絕現用／非空目的地，先在 destination staging 核對相同清單再发布 marker，不自動 connect。需要 connect 的操作由 CLI/GUI 明確發起。損毀源資料原 bytes 仍可保存，報告另外列為異常，不能把無法解析當作略過檔案的理由。

- [ ] **Step 4: 跑 transfer、database 與 I/O 測試。**
- [ ] **Step 5: commit `feat: verify SEC database migration and snapshot recovery`。**

### Task 5: GUI、CLI、過夜與分析入口整合

**Files:** Modify `src/main.py`, `src/cli.py`, `src/locales/{zh_tw,zh_cn,en,ja}.py`, `scripts/overnight_update_db.ps1`, `scripts/check_fy_labels.py`, `scripts/classify_holes.py`, `scripts/analyze_nvda_revenue_pilot.py`, `scripts/probe_nvda_revenue.py`, `skills/sec-revenue-breakdown/scripts/{local_context,build_nvda_history}.py`, GUI probe 及相關 tests；Create `tests/test_database_cli.py`。

**Interfaces:** CLI 子命令 `db-create PATH`、`db-connect PATH`、`db-migrate --source PATH --destination PATH`、`db-snapshot [--destination PATH]`、`db-verify-snapshot PATH`、`db-restore SNAPSHOT --destination PATH`；管理命令都接受 `--config-path`。API 調用分別對應 Tasks 1、4。`db-status` 顯示 root/UUID/connection/parser compatibility；未連接回非零並給 db-connect 指示。

- [ ] **Step 1: 寫 CLI 連接／失聯測試與 GUI 靜態入口測試。**

```python
def test_cli_connect_rejects_folder_without_marker(tmp_path, capsys):
    import cli
    assert cli.main(['db-connect', str(tmp_path), '--config-path',
                     str(tmp_path / 'config.json')]) != 0
    assert not (tmp_path / 'config.json').exists()
```

GUI probe 先建立隔離庫，測無清除按鈕、失聯仍能開 app、重新選取後正常顯示、執行中不能切庫、非同步更新 failure 有訊息。技能 export 測從設定的獨立庫讀取，而非 repo/local_db。

- [ ] **Step 2: 跑 CLI／skill 測試，確認目前缺指令與固定路徑失敗。**
- [ ] **Step 3: 整合入口。** CLI 捕捉 DatabaseError 印清楚的連接／保存錯誤；GUI 啟動保留未連接狀態，資料庫區加入新建、搬遷、連接與快照動作，使用 filedialog 選目錄，耗時搬遷／快照跑 worker。失聯時所有會寫庫的入口停用，但連接入口可用。移除清除按鈕及 handler；原先 button locking 改鎖更新／切庫／搬遷，probe 不再斷言已移除按鈕。

```python
# 分析脚本已有 sys.path 指向 repo/src 時：
import filing_cache
CACHE = filing_cache.cache_root()
# skill 使用 --repo 尋找程式 src，不能把 --repo 解釋成資料位置。
```

`build_nvda_history.py --cache-repo` 保留既有引數，意義改為找程式；可明確添加 --database-path，驗證 marker 后使用，不能隱藏 fallback。全面 rg 固定路徑。過夜 $Py 改為 USERPROFILE/venvs/SEC Financial Tools/Scripts/python.exe；開始前呼叫 db-status 驗證，檢查 native exit code 再取得資料庫更新名單。不得輸出 SEC identity 當連接結果；只記資料庫位置與 UUID。

- [ ] **Step 4: 跑 CLI、GUI helpers／labels／i18n、skill context、financial/segments 非 slow 測試；跑隔離 GUI probes 與 PowerShell Parser 語法檢查。**
- [ ] **Step 5: commit `feat: use the protected SEC database across all entry points`。**

### Task 6: 真實資料搬遷、驗證與文件交付

**Files:** Modify `README.md`, `docs/{ARCHITECTURE,CLI,PACKAGING,CHANGELOG,local-db-runs,TODO}.md`, `docs/RECIPIENT-README.txt`, `scripts/{README.md,pack.ps1}`, `skills/sec-revenue-breakdown/references/project-integration.md`；更新本 spec／plan 狀態與驗證紀錄。

- [ ] **Step 1: 執行全套非 slow 測試。**

```powershell
& 'C:/Users/CTH/venvs/SEC Financial Tools/Scripts/python.exe' -m pytest -m 'not slow' -p no:cacheprovider
git diff --check
```

若現有測試直接拿真庫路徑，先改為隔離 marker；正式驗證不打 SEC。只有已發現的新失敗才追加相關測試，不為改文件跑重複完整套件。

- [ ] **Step 2: 停止真庫更新、核對目標與磁碟空間後執行已測試的 db-migrate。** 只讀檢查 Python app processes 的 command line，停止本次建立的程序；使用者既有 GUI 必須先關閉，不隨意 kill。目標是 Windows Known Folder Documents 下 SEC財報資料庫，寫入此 repo 外路徑需按目前 sandbox 流程取得 escalation，授權來源是已確認搬遷任務。舊庫保留。

```powershell
& 'C:/Users/CTH/venvs/SEC Financial Tools/Scripts/python.exe' src/cli.py db-migrate --source local_db/filing_cache --destination 'C:/Users/CTH/Documents/SEC財報資料庫'
& 'C:/Users/CTH/venvs/SEC Financial Tools/Scripts/python.exe' src/cli.py db-status --json output/database-status-after-migration.json
```

- [ ] **Step 3: 核對遷移報告與輸出相同。** 所有來源財報、公司 metadata 及原有 tmp 的 size/hash 一致，count 以現場清單為準。以原始 bytes 載入 AAPL/NVDA，透過同一 `_CachedFinancials`／segments builder 比較搬前搬後 DataFrame／StatementTable，不使用連網 CLI gaap 作為离線證據。缺 marker 或 ID 不符都實際以隔離 config 驗證，不能對真庫做刪除測試。
- [ ] **Step 4: 建立真庫本機快照、verify，還原到新的隔離目的地，核對清單與讀取。** 還原庫不 connect；報告列出同磁碟限制。只刪隔離測試產物且確認路徑；正式舊庫、history、snapshots 不清理。
- [ ] **Step 5: 更新文件並 commit。** README 改新架構與首次連接；ARCHITECTURE 說明 ownership/UUID/鎖/歷史/失聯；CLI 列實際指令；PACKAGING/pack.ps1 保證不包財報或私人連接設定，收件人須新建／連接，不再自動空庫；local-db-runs 記實際 hash 驗證、公司數、快照還原；CHANGELOG 記具體行為與測試結果；TODO 只記獨立備份未配置；歷史 specs 添新設計链接、不改歷史決策。commit `docs: document independent SEC database migration and recovery`。
- [ ] **Step 6: 整體 code review 與交付。** 依使用者選定的執行方式完成獨立 review，修正真正問題，重跑相應測試。最後列實際資料庫位置、checkpoint/實作 commits、驗證結果、原庫保留與獨立備份未配置；不宣稱零刪除可能性。

## 自查與執行選擇

規格覆盖：位置／身分 Task 1；保存／版本／失敗 Task 2；名單 Task 3；搬遷／快照 Task 4；所有入口 Task 5；實庫／文件 Task 6。Review Focus 每項均已落在對應 task 測試；沒有用環境變數代替測試庫身分，也沒有將舊庫副本當成長期異地備份。

建議 Native：六項的連接與寫入介面彼此相依，由同一實作者維持一致、最後獨立 reviewer 檢查整體。Subagent-driven 也可依上述 task 邊界逐項實作／審查，成本較高。尚待使用者審閱本計畫並選擇執行方式；產品程式與真庫尚未更動。
