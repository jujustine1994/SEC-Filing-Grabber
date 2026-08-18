# PACKAGING.md — 打包散布作業指示

> ⚠ **暫停用（2026-08-18 CTH）**：GitHub 連得回來了，目前發布改走
> `git clone`，不需要壓 zip。這份流程與 `scripts/打包.bat` 都保留著，
> 之後又連不上 GitHub 時再撿回來用。

**這份是給 AI 照著執行的作業指示，不是給人看的教學。**
照著跑完，產出一包可以直接用通訊軟體傳給非技術使用者的 zip，並自我驗證。

情境：GitHub 連不上，不走 clone，直接壓一包傳過去。收件人解壓、雙擊
`啟動器.bat`、一路按 Enter 就能裝好並跑起來。

---

## 0. 前提檢查（不通過就停下來問人，不要自己往下做）

| 檢查 | 指令 | 期望 |
|---|---|---|
| 工作區乾淨 | `git status --porcelain` | 無輸出 |
| 測試全過 | `.\venv\Scripts\python.exe -m pytest -q -m "not slow"` | `N passed`，0 failed |
| launcher 語法正確 | 見 §4 的 parse 檢查 | `PARSE OK` |

---

## 1. 產出物規格

| 項目 | 規則 |
|---|---|
| 檔名 | `SEC-Financial-Fetcher-YYYYMMDD.zip`（YYYYMMDD＝打包當天，例：`SEC-Financial-Fetcher-20260817.zip`）。**全英數字，不含空白與中文**——經通訊軟體傳送時中文檔名會亂碼 |
| 放哪 | 專案根目錄的 `dist/`（已在 `.gitignore`，不會被 commit，也不會被下次打包捲進去） |
| 解壓後的頂層 | zip 內**要有一層資料夾** `SEC-Financial-Fetcher`，不要把檔案裸放在 zip 根。收件人常直接對著 zip 按「解壓縮到這裡」，裸放會把幾十個檔案倒進他的下載資料夾 |
| 大小 | 正常落在 150–400 KB（2026-08-17 排掉 `docs/` 後是 ~200 KB）。**超過 5 MB 一定是排除清單漏了東西**，回頭查 §2 |

---

## 2. 包含清單 / 排除清單

**採白名單複製，不要「複製全部再刪掉不要的」。** 黑名單漏一項就是把
機敏檔或 400MB 的 `venv/` 送出去，而且事後不一定看得出來。

### 一定要包含

| 項目 | 說明 |
|---|---|
| `啟動器.bat` | 唯一入口 |
| `launcher.ps1` | 全部安裝邏輯。**必須帶 UTF-8 BOM**，見 §4 |
| `README.md` | 專案說明（技術向，收件人不一定看，但留著） |
| `requirements.txt` | 套件清單，`launcher.ps1` 從根目錄讀 |
| `config.example.json` | 範例設定，內容是假值 |
| `src/` | 全部 `.py` 與 `src/locales/`（**排除 `__pycache__`**） |
| `docs/8k-period-off-by-one.md` | **只這一個 doc**，因為 `README.md` 內文有連到它。其餘 `docs/` 全部排除（見下） |
| `output/`（空目錄 + `.gitkeep`） | 預設輸出目錄（`src/config.py` 的 `output_dir` 預設值就是 `"output"`）。**只放空目錄，不放任何 `.xlsx`** |
| `先讀我.txt` | 從 `docs/RECIPIENT-README.txt` 複製過去並改名。給收件人看的那份，UTF-8 **含 BOM**（舊版記事本沒 BOM 會顯示亂碼） |

### 一定要排除

| 項目 | 為什麼 |
|---|---|
| `venv/` | 400 MB，`launcher.ps1` 會在收件人電腦上自己建 |
| `.git/` | 7.4 MB，且含完整開發史 |
| `company_cache.json` | 414 KB，程式會自己重建 |
| `output/*.xlsx`、`output/_final/` | CTH 的實測輸出檔（14 個），是個人資料 |
| `logs/` | 本機執行紀錄 |
| `20260814 sec tool.zip` | 舊的打包產物 |
| `dist/` | 本次與過去的打包產物 |
| `.pytest_cache/`、`__pycache__/`（所有層級）、`*.pyc` | 快取 |
| `.claude/`、`.superpowers/` | 開發環境設定 |
| `tests/`、`conftest.py`、`scripts/` | 對純使用者無用；`scripts/` 另含開發用工具 |
| `docs/`（除了 `8k-period-off-by-one.md`） | **內部開發紀錄**：`CHANGELOG.md` 一個就 84 KB，還有 `ARCHITECTURE` / `PITFALLS` / `TODO` / `superpowers/plans/` 的設計討論。收件人一個字都用不到，而且是內部資料。排掉之後包從 340 KB 降到約 200 KB（CTH 2026-08-17 決定） |

### 機敏資訊：天生不在專案內（已查證，不必額外處理）

`src/config.py:14-18` 的 `_default_config_path()` 把 `config.json` 放在
`%APPDATA%\SEC Financial Tools\`。**API Key、SEC EDGAR Identity、Watchlist、
個人輸出路徑全部在專案資料夾外**，壓 zip 碰不到。專案內只有
`config.example.json`，內容是假的範例值。

§5 的驗證步驟仍要實際 grep 一次，不要因為「理論上不會有」就跳過。

---

## 3. 打包步驟（PowerShell，逐段執行）

```powershell
# --- 3.1 變數 ---
$Root  = (Resolve-Path ".").Path              # 必須在專案根目錄執行
$Stamp = Get-Date -Format "yyyyMMdd"
$Name  = "SEC-Financial-Fetcher"
$Stage = Join-Path $env:TEMP "$Name-stage-$Stamp"
$Dist  = Join-Path $Root "dist"
$Zip   = Join-Path $Dist "$Name-$Stamp.zip"
$Pkg   = Join-Path $Stage $Name                # zip 內的那一層資料夾

# --- 3.2 清乾淨暫存區並重建 ---
if (Test-Path $Stage) { Remove-Item -Recurse -Force $Stage }
New-Item -ItemType Directory -Force $Pkg  | Out-Null
New-Item -ItemType Directory -Force $Dist | Out-Null

# --- 3.3 白名單複製：根目錄檔案 ---
foreach ($f in @("啟動器.bat","launcher.ps1","README.md","requirements.txt","config.example.json")) {
    Copy-Item (Join-Path $Root $f) $Pkg
}

# --- 3.4 白名單複製：src/（排除 __pycache__ / *.pyc）---
Copy-Item (Join-Path $Root "src") $Pkg -Recurse
# docs/ 只帶 README 有連到的那一份，其餘內部文件不外流
New-Item -ItemType Directory -Force (Join-Path $Pkg "docs") | Out-Null
Copy-Item (Join-Path $Root "docs\8k-period-off-by-one.md") (Join-Path $Pkg "docs")
# ⚠ 一定要寫成 ForEach-Object + -LiteralPath。直接 `| Remove-Item -Recurse -Force`
# 會被 Claude Code 的沙箱防護判成「刪除系統路徑 /」而整段拒絕執行（管線進來的
# 路徑它靜態分析不出來）。這不是 PowerShell 的問題，是工具側的保護。
Get-ChildItem $Pkg -Recurse -Directory -Filter "__pycache__" |
    ForEach-Object { Remove-Item -Recurse -Force -LiteralPath $_.FullName }
Get-ChildItem $Pkg -Recurse -File -Filter "*.pyc" |
    ForEach-Object { Remove-Item -Force -LiteralPath $_.FullName }

# --- 3.5 空的 output/ ---
New-Item -ItemType Directory -Force (Join-Path $Pkg "output") | Out-Null
New-Item -ItemType File -Force (Join-Path $Pkg "output\.gitkeep") | Out-Null

# --- 3.6 給收件人的說明（改名，保留 BOM）---
Copy-Item (Join-Path $Root "docs\RECIPIENT-README.txt") (Join-Path $Pkg "先讀我.txt")

# --- 3.7 壓縮 ---
if (Test-Path $Zip) { Remove-Item -Force $Zip }
Compress-Archive -Path $Pkg -DestinationPath $Zip
"產出：$Zip  ($([math]::Round((Get-Item $Zip).Length/1KB,1)) KB)"
```

> ⚠ `Compress-Archive -Path $Pkg`（指向資料夾本身，結尾**不加** `\*`）才會在
> zip 內保留 `SEC-Financial-Fetcher\` 那一層。加了 `\*` 會變成裸放。

---

## 4. launcher.ps1 的兩項硬性檢查

打包前跑，任一項不過就先修再打包。

```powershell
# BOM：必須是 239,187,191（地雷五，沒 BOM 中文訊息會變亂碼）
([System.IO.File]::ReadAllBytes("$Root\launcher.ps1")[0..2]) -join ','

# 語法
$e=$null
$null=[System.Management.Automation.Language.Parser]::ParseFile("$Root\launcher.ps1",[ref]$null,[ref]$e)
if($e.Count -eq 0){"PARSE OK"}else{$e|ForEach-Object{$_.Message}}
```

---

## 5. 自我驗證（解到另一個暫存目錄實際檢查，不可只看 §3 的輸出）

```powershell
$Verify = Join-Path $env:TEMP "$Name-verify-$Stamp"
if (Test-Path $Verify) { Remove-Item -Recurse -Force $Verify }
Expand-Archive -Path $Zip -DestinationPath $Verify
$V = Join-Path $Verify $Name
```

逐項比對，**全部要通過**：

| # | 檢查 | 指令 | 期望 |
|---|---|---|---|
| 1 | 頂層只有一層資料夾 | `Get-ChildItem $Verify -Name` | 只有 `SEC-Financial-Fetcher` |
| 2 | 必要檔案齊全 | `@("啟動器.bat","launcher.ps1","README.md","requirements.txt","config.example.json","先讀我.txt") \| ForEach-Object { "{0} {1}" -f $_, (Test-Path (Join-Path $V $_)) }` | 全 `True` |
| 3 | `src/` 完整 | `(Get-ChildItem "$V\src" -Filter *.py).Count` | 與 `(Get-ChildItem "$Root\src" -Filter *.py).Count` 相同 |
| 4 | 語言檔在 | `Get-ChildItem "$V\src\locales" -Name` | 四個語言的檔案都在 |
| 5 | **沒有機敏檔** | `Get-ChildItem $V -Recurse -Force -Include config.json,*.log,company_cache.json,.env` | **無輸出** |
| 6 | **沒有 xlsx** | `Get-ChildItem $V -Recurse -Force -Filter *.xlsx` | **無輸出** |
| 7 | 沒有開發目錄 | `Get-ChildItem $V -Recurse -Force -Directory \| Where-Object Name -in '__pycache__','.git','venv','tests','scripts','.pytest_cache','.claude','.superpowers'` | **無輸出** |
| 8 | 內容沒挾帶金鑰 | `Get-ChildItem $V -Recurse -File -Include *.py,*.json,*.md,*.ps1,*.bat,*.txt \| Select-String -Pattern 'sk-[A-Za-z0-9]{20}','AIza[A-Za-z0-9_-]{30}' -List` | **無輸出**（真金鑰的格式；`config.example.json` 的假值不該命中） |
| 9 | Identity 沒外洩 | `Get-ChildItem $V -Recurse -File \| Select-String -Pattern '[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}' -List` | 只該命中文件裡的範例信箱（如 `wangdaming@gmail.com`）。出現 CTH 的真實信箱就是漏了東西 |
| 10 | 大小合理 | `(Get-Item $Zip).Length/1MB` | < 5 |
| 11 | 空 `output/` | `Get-ChildItem "$V\output" -Force -Name` | 只有 `.gitkeep` |
| 12 | **內部文件沒外流** | `Get-ChildItem "$V\docs" -Name` | 只有 `8k-period-off-by-one.md`。出現 `CHANGELOG.md` / `TODO.md` / `superpowers` 就是 §3.4 寫錯了 |

驗證完清掉暫存（同樣避開沙箱防護，一次一個明確路徑）：

```powershell
Remove-Item -Recurse -Force -LiteralPath $Stage
Remove-Item -Recurse -Force -LiteralPath $Verify
```

> ⚠ 若 §5 的驗證是在解出來的複本上實跑過 `uv venv`，那裡會多出一個
> 幾百 MB 的 `venv/`，記得一起清掉。

---

## 6. 收件人端的已知狀況（不修，只在說明裡講）

- **Mark-of-the-Web**：zip 經網路或通訊軟體傳送，解壓出來的 `.bat` 帶
  Zone.Identifier，可能被 SmartScreen 攔一次。**CTH 2026-08-17 決定不處理**，
  程式端不加 `Unblock-File`。`先讀我.txt` 只寫「點『其他資訊』→『仍要執行』」，
  不寫成技術教學。往後不要再把它當 bug 提出來。
- **首次執行要下載**：uv（幾 MB）＋ Python（約 20 MB，只在收件人電腦沒有
  Python 時）＋ pip 套件。全程需要網路，約 3-5 分鐘。
- **不需要系統管理員權限**：uv 與 uv 管理的 Python 都裝在使用者目錄下。

## 7. 交出去之前

把 zip 傳給 CTH 之前，跟他確認兩件事：

1. `launcher.ps1` 改過的話，**必須由 CTH 親自雙擊 `啟動器.bat` 實測**
   （規則檔硬性要求：含互動安裝的流程，AI 不得只看程式碼就宣稱完成）。
   最有意義的測法是在一台沒有 Python 的電腦、或把 `venv/` 改名後重跑。
2. `dist/` 底下的舊 zip 要不要一起清掉。
