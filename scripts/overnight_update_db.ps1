# overnight_update_db.ps1 — 整夜更新本地財報資料庫（TODO J8 的重抓）
#
# 雙擊 `過夜更新資料庫.bat` 執行，**全程不需要 AI**（跟 watchdog_h0_baseline.sh
# 同一個設計原則：成果不依賴 AI 還活著，也不吃額度）。
#
# 為什麼不直接跑 `cli.py update-db`（一個 process 跑完全部）：
#   2026-09-06 實測，一個 process 連跑 67 家會在第 16 家被系統因記憶體不足
#   中止——edgartools 的內部快取會跨公司累積，我們自己的 `_parse_cache_scope()`
#   只涵蓋單次抓取，擋不住。所以這裡切成每段 N 家、每段一個獨立 process，
#   段落結束記憶體整個還給系統。
#
# 為什麼跑兩輪：
#   `update-db` 跑一輪**不保證抓齊**（實測 ACN 要兩輪、2026-09-05 的 batch 1
#   有 16 家第二輪才補上）。第二輪很便宜（已完整的公司整家跳過，只花一次
#   filing 清單查詢），所以一律跑。
#
# 中斷不會白費：`save_filing()` 逐份即時落檔，重跑會自動跳過已完成的。
# 直接關視窗或按 Ctrl+C 都可以，隔天再跑一次會從斷點接下去。

param(
    [int]$ChunkSize = 8,
    [switch]$SkipSecondPass
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root

$Py = Join-Path $Root "venv\Scripts\python.exe"
$Stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$OutDir = Join-Path $Root "output\_localdb"
$LogFile = Join-Path $OutDir "overnight_$Stamp.log"
New-Item -ItemType Directory -Force $OutDir | Out-Null

$env:PYTHONIOENCODING = "utf-8"

function Log($msg) {
    $line = "[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $msg
    Write-Host $line
    Add-Content -Path $LogFile -Value $line -Encoding utf8
}

# ── 前置檢查：早點失敗，不要 218 家全部失敗才發現 ──────────────────────
if (-not (Test-Path $Py)) {
    Log "找不到 $Py——venv 還沒建，先跑一次 啟動器.bat"
    Read-Host "按 Enter 關閉"; exit 1
}

$identity = & $Py -c @"
import sys; sys.path.insert(0, 'src')
from config import load_config
print((load_config().get('identity') or '').strip())
"@
if (-not $identity) {
    Log "config.json 沒有 SEC EDGAR Identity——先開程式在「進階設定」填一次"
    Read-Host "按 Enter 關閉"; exit 1
}

$tickers = (& $Py -c @"
import sys; sys.path.insert(0, 'src')
from config import load_config
import local_db
print(' '.join(local_db.get_update_list(load_config())))
"@).Trim() -split '\s+' | Where-Object { $_ }

if ($tickers.Count -eq 0) {
    Log "更新名單是空的——先在「進階設定」→「更新名單」建一份"
    Read-Host "按 Enter 關閉"; exit 1
}

Log "更新名單 $($tickers.Count) 家，每段 $ChunkSize 家"
Log "身分：$identity"
Log "log：$LogFile"
Log "⚠ 這會跑很久（2026-09-18 實測 2.9 秒/份，14,000 份約 11 小時）。"
Log "   中途關掉不會白費，逐份即時落檔，下次從斷點接。"

# ── 跑整夜不要讓電腦睡著 ────────────────────────────────────────────────
# 睡著的話 process 會停，隔天醒來發現只跑了一半。ES_CONTINUOUS 讓這個
# 設定在 process 存活期間持續有效，process 結束自動失效（不會永久改設定）。
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class Sleepless {
  [DllImport("kernel32.dll")] public static extern uint SetThreadExecutionState(uint f);
}
"@
# ES_CONTINUOUS(0x80000000) | ES_SYSTEM_REQUIRED(0x1) — 螢幕可以關，系統不睡。
# ⚠ 一定要明寫 [uint32]：PowerShell 5.1 把 `0x80000001` 當有號 Int32 解成
# -2147483647，直接丟進去會 "Cannot convert ... to type System.UInt32"，
# 而且 $ErrorActionPreference="Stop" 會讓整支腳本當場中止（2026-09-18 實測）。
[void][Sleepless]::SetThreadExecutionState([uint32]2147483649)
Log "已要求系統不進入睡眠（這支結束就自動解除）"

function Invoke-Pass($tickerList, $passName) {
    # 抓取階段不要 Stop：單一公司失敗本來就不該中斷整批（設計書第六節），
    # 跑了 8 小時因為第 190 家壞掉而整個收工是最糟的結果。
    $ErrorActionPreference = "Continue"
    $total = $tickerList.Count
    $chunks = [Math]::Ceiling($total / $ChunkSize)
    Log "===== $passName 開始：$total 家、分 $chunks 段 ====="
    $started = Get-Date

    # 段落清單：哪一段有哪幾家。**這是給「整段 process 掛掉」用的**——
    # 被系統 OOM 砍掉的 process 不會寫它的 `--json`，收尾摘要只看 json 的話
    # 會靜默少報那一整段（8 家憑空消失，早上完全看不出來）。有這份清單，
    # 摘要才能比對出「第 N 段沒有產出，裡面是這幾家」。
    $manifest = @{}
    for ($i = 0; $i -lt $total; $i += $ChunkSize) {
        $manifest[([string]([int]($i / $ChunkSize) + 1))] =
            @($tickerList[$i..([Math]::Min($i + $ChunkSize - 1, $total - 1))])
    }
    $manifestPath = Join-Path $OutDir ("{0}_{1}_manifest.json" -f $Stamp, $passName)
    $manifest | ConvertTo-Json -Depth 4 |
        Set-Content -Path $manifestPath -Encoding utf8

    for ($i = 0; $i -lt $total; $i += $ChunkSize) {
        $part = $tickerList[$i..([Math]::Min($i + $ChunkSize - 1, $total - 1))]
        $n = [int]($i / $ChunkSize) + 1
        Log "--- $passName 第 $n/$chunks 段：$($part -join ' ') ---"
        $json = Join-Path $OutDir ("{0}_{1}_chunk{2:d3}.json" -f $Stamp, $passName, $n)

        # 每段一個獨立 process，跑完就結束——記憶體整個還給系統。
        #
        # ⚠ **合併 stderr 交給 cmd 做，不要在 PowerShell 裡寫 `2>&1`。**
        # `cli.py` 的逐家進度是印到 stderr 的（刻意的：stdout 要留給 --json）。
        # PowerShell 5.1 對原生執行檔用 `2>&1` 會把每一行 stderr 包成
        # NativeCommandError，配上 `$ErrorActionPreference="Stop"` 第一行進度
        # 就讓整支腳本中止（2026-09-18 實測）。讓 cmd 層合併，PowerShell 這邊
        # 只看到 stdout，這一整類問題就不存在。
        $cmdLine = '"{0}" -u src\cli.py update-db {1} --json "{2}" 2>&1' -f `
            $Py, ($part -join ' '), $json
        & cmd /c $cmdLine |
            ForEach-Object {
                $line = "$_"
                # edgartools 的雜訊，不是我們的錯誤
                if ($line -notmatch '^(No XBRL attachments|Failed to resolve)') {
                    Add-Content -Path $LogFile -Value $line -Encoding utf8
                    Write-Host $line
                }
            }
        Log "--- 第 $n 段結束（exit $LASTEXITCODE）---"
    }
    $elapsed = (Get-Date) - $started
    Log "===== $passName 結束，耗時 $($elapsed.ToString('hh\:mm\:ss')) ====="
}

Invoke-Pass $tickers "pass1"

if (-not $SkipSecondPass) {
    # 第二輪：已完整的整家跳過，只花一次清單查詢，很便宜但**必要**
    Invoke-Pass $tickers "pass2"
}

# ── 收尾摘要：早上起來看這幾行就好 ──────────────────────────────────────
Log ""
Log "================ 收尾摘要 ================"
# 摘要邏輯放獨立的 .py（scripts/overnight_summary.py），不塞 here-string：
# 那樣要同時應付 PowerShell 的 $ 插值與 Python 的 f-string，2026-09-18 實測
# 踩到語法錯誤，而且沒辦法單獨測試。stderr 一樣交給 cmd 合併。
$sumCmd = '"{0}" scripts\overnight_summary.py "{1}" "{2}" 2>&1' -f $Py, $OutDir, $Stamp
& cmd /c $sumCmd | ForEach-Object { Log "$_" }

Log "=========================================="
Log "完整 log：$LogFile"
Log ""
Log "接下來可以做的（不急，白天再說）："
Log "  1. 看資料庫內容：venv\Scripts\python.exe src\cli.py db-status"
Log "  2. 檢查完整度（會連網）：venv\Scripts\python.exe scripts\audit_local_db.py"
Log "  3. 驗證檔案能不能用：venv\Scripts\python.exe scripts\verify_local_db.py"

Read-Host "`n全部跑完，按 Enter 關閉"
