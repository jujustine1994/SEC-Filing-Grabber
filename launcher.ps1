# SEC Financial Fetcher 啟動器

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$host.UI.RawUI.WindowTitle = "SEC Financial Fetcher"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ScriptDir

# ======================================
# 執行紀錄（必加，須放在 trap 之前，閃退才記得到）
# 完整規則見 windows-tool.md「執行紀錄」。開檔→寫→關檔，不持有 handle（地雷十）；
# 用 UTF8Encoding($false) 不寫 BOM（地雷十一），不可用 Add-Content -Encoding UTF8。
# ======================================
$LogFile = Join-Path $ScriptDir "logs\app.log"
New-Item -ItemType Directory -Force (Split-Path $LogFile) | Out-Null
$Utf8NoBom = [System.Text.UTF8Encoding]::new($false)

function Write-Log {
    param([string]$Msg, [string]$Level = "INFO")
    $line = "[{0}] [{1,-5}] {2}`r`n" -f (Get-Date -Format "HH:mm:ss"), $Level, $Msg
    try { [System.IO.File]::AppendAllText($LogFile, $line, $Utf8NoBom) } catch {}
}

function Write-LogHeader {
    param([string]$Msg)
    $line = "=== {0} {1} ===`r`n" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Msg
    try { [System.IO.File]::AppendAllText($LogFile, $line, $Utf8NoBom) } catch {}
}

Write-LogHeader "啟動"

# 攔截所有未預期例外，防止視窗直接閃退
trap {
    Write-Log "[CRASH] $($_.Exception.Message) @ 第 $($_.InvocationInfo.ScriptLineNumber) 行" "FATAL"
    Write-Host ""
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Red
    Write-Host "[CRASH] 意外錯誤，程式無法繼續執行" -ForegroundColor Red
    Write-Host ""
    Write-Host "  錯誤訊息：$($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "  發生位置：$($_.InvocationInfo.ScriptLineNumber) 行" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  請截圖此畫面並回報給開發者。" -ForegroundColor White
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Red
    Read-Host "按 Enter 關閉"
    exit 1
}

Clear-Host
Write-Host "[INFO] Starting SEC Financial Fetcher..." -ForegroundColor Green
Write-Host ""

# ======================================
# [1/3] 檢查 Python
# ======================================
Write-Host "[1/3] 檢查 Python 環境..." -ForegroundColor Cyan
# 全新 Windows 電腦沒裝過 Python 時，PATH 裡常有內建的「App execution alias」
# python.exe 存根——Get-Command 找得到它、看起來像已安裝，但實際執行只會跳出
# Microsoft Store 頁面，`python --version` 不會印出版本號。用輸出內容二次確認，
# 不能只看命令存不存在。
$pythonReady = $false
if (Get-Command python -ErrorAction SilentlyContinue) {
    $verOutput = (python --version 2>&1 | Out-String).Trim()
    if ($verOutput -match "^Python \d+\.\d+") {
        $pythonReady = $true
    }
}
if (-not $pythonReady) {
    Write-Host "[WARNING] 未偵測到可用的 Python（若剛剛跳出 Microsoft Store，代表偵測到的是" -ForegroundColor Yellow
    Write-Host "          Windows 內建的假別名，不是真的 Python），本程式需要 Python 才能執行。" -ForegroundColor Yellow
    $ans = Read-Host "是否要立即安裝 Python？[Y/n] - 直接按 Enter 代表同意"
    if ($ans -eq "" -or $ans -ieq "Y") {
        if (Get-Command winget -ErrorAction SilentlyContinue) {
            Write-Host "[INFO] 透過 winget 安裝 Python，請稍候..." -ForegroundColor Gray
            winget install --id Python.Python.3 -e --silent --accept-source-agreements --accept-package-agreements --override "/quiet PrependPath=1 Include_pip=1"
        } else {
            Write-Log "找不到 winget，無法自動安裝 Python" "ERROR"
            Write-Host "[ERROR] 找不到 winget，請手動至 https://www.python.org/ 下載安裝後重新執行。" -ForegroundColor Red
            Read-Host "按 Enter 關閉"; exit 1
        }
        $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("PATH", "User")
        $verOutput = ""
        if (Get-Command python -ErrorAction SilentlyContinue) {
            $verOutput = (python --version 2>&1 | Out-String).Trim()
        }
        if ($verOutput -notmatch "^Python \d+\.\d+") {
            Write-Host "[INFO] 安裝完成，請關閉視窗後重新點兩下啟動檔。" -ForegroundColor Yellow
            Read-Host "按 Enter 關閉"; exit 0
        }
        $pyVer = $verOutput
        Write-Host "[OK] Python 安裝完成。" -ForegroundColor Green
    } else {
        Write-Host "已取消。" -ForegroundColor Gray; Read-Host "按 Enter 關閉"; exit 1
    }
} else {
    $pyVer = $verOutput
    Write-Host "[OK] $pyVer 已安裝。" -ForegroundColor Green
}

# ======================================
# [2/3] 檢查 uv
# ======================================
Write-Host "[2/3] 檢查 uv 套件管理工具..." -ForegroundColor Cyan
if (-not (Get-Command uv -ErrorAction SilentlyContinue)) {
    Write-Host "[WARNING] 找不到 uv，正在安裝..." -ForegroundColor Yellow
    Invoke-RestMethod https://astral.sh/uv/install.ps1 | Invoke-Expression
    $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH", "User") + ";" + $env:PATH
    if (-not (Get-Command uv -ErrorAction SilentlyContinue)) {
        Write-Log "uv 安裝失敗" "ERROR"
        Write-Host "[ERROR] uv 安裝失敗，請關閉視窗後重新點兩下啟動檔再試。" -ForegroundColor Red
        Read-Host "按 Enter 關閉"; exit 1
    }
    Write-Host "[OK] uv 安裝完成。" -ForegroundColor Green
} else {
    $uvVer = uv --version
    Write-Host "[OK] $uvVer 已安裝。" -ForegroundColor Green
}

# ======================================
# [3/3] 檢查虛擬環境
# ======================================
Write-Host "[3/3] 檢查虛擬環境..." -ForegroundColor Cyan
if (-not (Test-Path "venv")) {
    Write-Host ""
    Write-Host "  ============================================" -ForegroundColor Cyan
    Write-Host "    SEC Financial Fetcher - 首次安裝說明" -ForegroundColor Cyan
    Write-Host "  ============================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  接下來程式會自動幫你安裝以下東西：" -ForegroundColor White
    Write-Host ""
    Write-Host "    1. Python 虛擬環境（venv）" -ForegroundColor Yellow
    Write-Host "       讓這個工具有獨立乾淨的執行空間，不影響電腦其他程式" -ForegroundColor Gray
    Write-Host ""
    Write-Host "    2. edgartools" -ForegroundColor Yellow
    Write-Host "       從 SEC EDGAR 抓取上市公司財報的核心套件" -ForegroundColor Gray
    Write-Host ""
    Write-Host "    3. openpyxl" -ForegroundColor Yellow
    Write-Host "       讀寫 Excel 檔案" -ForegroundColor Gray
    Write-Host ""
    Write-Host "    4. AI 套件（google-generativeai / openai / anthropic）" -ForegroundColor Yellow
    Write-Host "       Non-GAAP 功能用，未設定 API Key 不影響 GAAP 功能" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  全程只需要一直按 Enter 同意即可。" -ForegroundColor Green
    Write-Host "  如果有任何疑問，可以把這段說明貼給 AI 詢問。" -ForegroundColor Green
    Write-Host ""
    Write-Host "  ============================================" -ForegroundColor Cyan
    Write-Host ""
    $ans = Read-Host "[WARNING] 找不到虛擬環境，現在建立並安裝套件？[Y/n] - 直接按 Enter 代表同意"
    if ($ans -eq "" -or $ans -ieq "Y") {
        Write-Host "[INFO] 建立虛擬環境中..." -ForegroundColor Gray
        uv venv venv
        Write-Host "[INFO] 安裝套件中（首次約需 2-3 分鐘）..." -ForegroundColor Gray
        uv pip install -r requirements.txt --python venv\Scripts\python.exe -q
        if ($LASTEXITCODE -ne 0) {
            Write-Log "套件安裝失敗（uv pip install 回傳 $LASTEXITCODE）" "ERROR"
            Write-Host "[ERROR] 套件安裝失敗，請確認網路連線後重新執行。" -ForegroundColor Red
            Read-Host "按 Enter 關閉"; exit 1
        }
        Write-Host "[OK] 套件安裝完成。" -ForegroundColor Green
    } else {
        Write-Host "已取消。" -ForegroundColor Gray; Read-Host "按 Enter 關閉"; exit 1
    }
} else {
    Write-Host "[OK] 虛擬環境已就緒，檢查套件更新..." -ForegroundColor Green
    # 清理損壞的 dist-info（METADATA 遺失時 uv 拒絕安裝）
    $broken = Get-ChildItem "venv\Lib\site-packages" -Directory -Filter "*dist-info" -ErrorAction SilentlyContinue | Where-Object {
        -not (Test-Path (Join-Path $_.FullName "METADATA"))
    }
    foreach ($dir in $broken) {
        Write-Host "[INFO] 清理損壞的套件資訊：$($dir.Name)" -ForegroundColor Yellow
        Remove-Item -Recurse -Force $dir.FullName
    }
    uv pip install -r requirements.txt --python venv\Scripts\python.exe -q
}

. ".\venv\Scripts\Activate.ps1"

Write-Log "環境就緒 | $pyVer | $uvVer"

Write-Host ""
Write-Host "[START] 啟動中，請保持此視窗開啟..." -ForegroundColor Green
Write-Host ""

# 主程式執行期間由它自己寫 log，launcher 不寫（避免搶 handle，地雷十）
python src\main.py
$exitCode = $LASTEXITCODE

# 2026-08-12 的目錄整理把進入點換成 src\main.py，bytecode 從此落在
# src\__pycache__ 與 src\locales\__pycache__，根目錄那行變成永久空操作。
foreach ($pc in @("__pycache__", "src\__pycache__", "src\locales\__pycache__")) {
    if (Test-Path $pc) {
        # 重啟語言時新行程可能正在用這些 .pyc，刪不掉不是錯誤，別嚇到使用者
        try { Remove-Item -Recurse -Force $pc -ErrorAction Stop } catch {}
    }
}

if ($exitCode -ne 0) {
    Write-Log "主程式異常結束（exit code $exitCode）" "ERROR"
    Write-Host ""
    Write-Host "[ERROR] 程式意外停止，請回報上方錯誤訊息。" -ForegroundColor Red
    Read-Host "按 Enter 關閉"
} else {
    Write-Host ""
    Write-Host "5 秒後自動關閉..." -ForegroundColor Gray
    Start-Sleep -Seconds 5
}
