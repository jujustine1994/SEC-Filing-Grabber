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

# 畫面上的「現在正在做什麼」提示，帶時間戳；不寫 log（log 已有自己的時間戳）。
# 只在會下載東西的步驟前呼叫（TODO E7：CTH 回報安裝過程沒有 timestamp，
# 不知道是卡住了還是正常在跑）。
function Write-Step {
    param([string]$Msg, [string]$Color = "Gray")
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] $Msg" -ForegroundColor $Color
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
# [1/2] 檢查 uv
#
# 2026-08-17：原本這裡是 [1/3] 檢查 Python、用 winget 裝系統 Python。
# 那條路已經壞了——`winget show --id Python.Python.3` 回 exit 20
# (No package found)，winget 上只剩 Python.Python.3.10 ~ 3.14。與其追著版號
# 改，整段拿掉：uv 自己就會下載 Python（python-build-standalone；Windows 版
# 內含 tkinter，實測 import + 開視窗都正常）。沒有系統 Python 也能跑，連帶
# 不需要 winget（舊 Win10 沒有）、不需要刷新 PATH、不需要叫使用者關掉視窗
# 再雙擊一次。真正做得到「一路按 Enter」。
# ======================================
Write-Host "[1/2] 檢查 uv 套件管理工具..." -ForegroundColor Cyan
if (-not (Get-Command uv -ErrorAction SilentlyContinue)) {
    Write-Log "找不到 uv，準備安裝" "WARN"
    Write-Step "找不到 uv，開始下載安裝程式（astral.sh）..." "Yellow"
    # 較舊的 Win10 上 PowerShell 5.1 可能仍預設 TLS 1.0/1.1，astral.sh 只收 1.2 以上，
    # 不先指定會在下載那行直接連線失敗
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
    try {
        Invoke-RestMethod https://astral.sh/uv/install.ps1 | Invoke-Expression
    } catch {
        Write-Log "uv 下載失敗 -> $($_.Exception.GetType().Name)" "ERROR"
    }
    $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH", "User") + ";" + $env:PATH
    if (-not (Get-Command uv -ErrorAction SilentlyContinue)) {
        Write-Log "uv 安裝失敗" "ERROR"
        Write-Host "[ERROR] uv 安裝失敗，多半是連不上網路。請確認網路連線後關閉視窗，" -ForegroundColor Red
        Write-Host "        重新點兩下啟動檔再試一次。" -ForegroundColor Red
        Read-Host "按 Enter 關閉"; exit 1
    }
    $uvVer = uv --version
    Write-Host "[OK] uv 安裝完成。" -ForegroundColor Green
} else {
    $uvVer = uv --version
    Write-Host "[OK] $uvVer 已安裝。" -ForegroundColor Green
}

# ======================================
# [2/2] 檢查虛擬環境
# ======================================
Write-Host "[2/2] 檢查虛擬環境..." -ForegroundColor Cyan
if (-not (Test-Path "venv")) {
    Write-Host ""
    Write-Host "  ============================================" -ForegroundColor Cyan
    Write-Host "    SEC Financial Fetcher - 首次安裝說明" -ForegroundColor Cyan
    Write-Host "  ============================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  接下來程式會自動幫你安裝以下東西：" -ForegroundColor White
    Write-Host ""
    Write-Host "    1. Python 本體與虛擬環境（venv）" -ForegroundColor Yellow
    Write-Host "       電腦沒裝過 Python 也沒關係，會自動下載（約 20MB）" -ForegroundColor Gray
    Write-Host "       虛擬環境讓這個工具有獨立乾淨的執行空間，不影響電腦其他程式" -ForegroundColor Gray
    Write-Host ""
    Write-Host "    2. edgartools" -ForegroundColor Yellow
    Write-Host "       從 SEC EDGAR 抓取上市公司財報的核心套件" -ForegroundColor Gray
    Write-Host ""
    Write-Host "    3. openpyxl" -ForegroundColor Yellow
    Write-Host "       讀寫 Excel 檔案" -ForegroundColor Gray
    Write-Host ""
    Write-Host "    4. AI 套件（google-genai / openai / anthropic）" -ForegroundColor Yellow
    Write-Host "       Non-GAAP 功能用，未設定 API Key 不影響 GAAP 功能" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  全程只需要一直按 Enter 同意即可。" -ForegroundColor Green
    Write-Host "  如果有任何疑問，可以把這段說明貼給 AI 詢問。" -ForegroundColor Green
    Write-Host ""
    Write-Host "  ============================================" -ForegroundColor Cyan
    Write-Host ""
    $ans = Read-Host "[WARNING] 找不到虛擬環境，現在建立並安裝套件？[Y/n] - 直接按 Enter 代表同意"
    if ($ans -eq "" -or $ans -ieq "Y") {
        Write-Step "下載/建立 Python 虛擬環境中（電腦若沒有 Python 會自動下載，約 20MB）..."
        # 指定 3.13：不寫版號時 uv 只會找現成直譯器，找不到就報錯而不會下載。
        # 寫了版號，本機有 3.13 就用本機的，沒有才下載 uv 自管版本。
        uv venv venv --python 3.13
        if ($LASTEXITCODE -ne 0) {
            Write-Log "建立虛擬環境失敗（uv venv 回傳 $LASTEXITCODE）" "ERROR"
            Write-Host "[ERROR] 建立虛擬環境失敗，多半是下載 Python 時連不上網路。" -ForegroundColor Red
            Write-Host "        請確認網路連線後關閉視窗，重新點兩下啟動檔再試一次。" -ForegroundColor Red
            Read-Host "按 Enter 關閉"; exit 1
        }
        Write-Step "安裝套件中（以下為 uv 逐一顯示的套件名稱與下載進度，首次約需 2-3 分鐘）..."
        # ⚠ 不加 -q/--quiet（TODO E7 修正：原本有 -q，這正是 CTH 回報「沒有進度條，
        #    不知道是卡住還是正常在跑」的根因）——uv 預設會逐一印出「正在下載哪個
        #    套件＋進度條」，這是使用者唯一看得懂「電腦沒當機、只是在裝東西」的
        #    畫面，不可靜音。
        uv pip install -r requirements.txt --python venv\Scripts\python.exe
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
    # ⚠ 不加 -q：平時沒有更新時 uv 幾乎瞬間印完「Audited N packages」，代價很小；
    #    一旦真的有更新要下載，使用者才看得到在裝什麼，不會誤以為卡住。
    uv pip install -r requirements.txt --python venv\Scripts\python.exe
}

. ".\venv\Scripts\Activate.ps1"

# $pyVer 改成問 venv 自己（原本問的是系統 Python，現在已經不一定存在）
$pyVer = (& ".\venv\Scripts\python.exe" --version 2>&1 | Out-String).Trim()
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
