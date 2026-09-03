<#
.SYNOPSIS
    驗證一份 .xlsx 是否會觸發 Excel 的「發現無法讀取的內容，是否修復」機制。

.DESCRIPTION
    TODO A-5 步驟 0 的驗證手段。用 Excel COM 開檔（背景執行、不彈視窗），
    比對開檔前後 %TEMP% 底下的 error*.xml 修復日誌是否多了新檔案——多了就
    代表 Excel 判定內容毀損並自動修復過，跟 CTH 手動開檔看到的提示是同一套
    機制。同時清點檔案裡 Chart_* 分頁還剩幾張圖，方便跟修復前的預期數量比對。

.PARAMETER Path
    要檢查的 .xlsx 絕對路徑。

.EXAMPLE
    powershell -File scripts/check_excel_repair.ps1 -Path "C:\Users\CTH\Desktop\新增資料夾\MyComp_AMD_MSFT_NVDA_20260903.xlsx"
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$Path
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $Path)) {
    Write-Error "找不到檔案：$Path"
    exit 1
}
$Path = (Resolve-Path $Path).Path

$tempDir = $env:TEMP
$before = Get-ChildItem -Path $tempDir -Filter "error*.xml" -ErrorAction SilentlyContinue |
    Select-Object -ExpandProperty Name

Write-Host "開檔前 %TEMP% 既有 error*.xml：$($before.Count) 個"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$chartSheetInfo = @()
$openFailed = $false
$openError = $null
try {
    try {
        $wb = $excel.Workbooks.Open($Path)
    } catch {
        # 實測發現：Excel Automation 開一份會觸發「發現無法讀取的內容」的檔案時，
        # 不會像手動雙擊那樣跳修復對話框讓人按「是」，而是直接讓 Open() 這個
        # COM 呼叫丟例外（不管 Visible/DisplayAlerts 怎麼設都一樣）。這比等
        # error*.xml 修復日誌更直接、更能自動化判斷——同一支腳本開已知正常的
        # .xlsx 會成功，開這份壞檔就是穩定重現這個例外。
        $openFailed = $true
        $openError = $_.Exception.Message
    }
    if (-not $openFailed) {
        try {
            foreach ($ws in $wb.Worksheets) {
                if ($ws.Name -like "Chart_*") {
                    $chartCount = $ws.ChartObjects().Count
                    $chartSheetInfo += [PSCustomObject]@{
                        Sheet  = $ws.Name
                        Charts = $chartCount
                    }
                }
            }
        } finally {
            $wb.Close($false)
        }
    }
} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

if ($openFailed) {
    Write-Host ""
    Write-Host "=== Workbooks.Open() 失敗 ==="
    Write-Host "錯誤訊息：$openError"
    Write-Host ""
    Write-Host "結論：REPRODUCED（Excel Automation 無法開啟此檔，判定內容毀損）" -ForegroundColor Red
    exit 2
}

Start-Sleep -Milliseconds 500

$after = Get-ChildItem -Path $tempDir -Filter "error*.xml" -ErrorAction SilentlyContinue |
    Select-Object -ExpandProperty Name
$newLogs = $after | Where-Object { $before -notcontains $_ }

Write-Host ""
Write-Host "=== Chart_* 分頁清點 ==="
if ($chartSheetInfo.Count -eq 0) {
    Write-Host "沒有找到任何 Chart_* 分頁"
} else {
    $chartSheetInfo | Format-Table -AutoSize
    $zeroChartSheets = $chartSheetInfo | Where-Object { $_.Charts -eq 0 }
    Write-Host "Chart_* 分頁共 $($chartSheetInfo.Count) 張，其中 $($zeroChartSheets.Count) 張圖已被清空"
}

Write-Host ""
Write-Host "=== 修復日誌比對 ==="
if ($newLogs.Count -eq 0) {
    Write-Host "沒有產生新的 error*.xml -> 這次開檔沒有觸發 Excel 修復機制"
    $repaired = $false
} else {
    Write-Host "產生了 $($newLogs.Count) 個新的修復日誌："
    foreach ($log in $newLogs) {
        $full = Join-Path $tempDir $log
        Write-Host "--- $full ---"
        Get-Content -Path $full -Raw -Encoding UTF8
    }
    $repaired = $true
}

Write-Host ""
if ($repaired) {
    Write-Host "結論：REPRODUCED（Excel 判定內容毀損並自動修復）" -ForegroundColor Red
    exit 2
} else {
    Write-Host "結論：OK（Excel 開檔正常，未觸發修復）" -ForegroundColor Green
    exit 0
}
