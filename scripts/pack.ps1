# 打包散布用 zip —— docs/PACKAGING.md 的可執行版本
#
# 這支腳本是 docs/PACKAGING.md §3（打包）與 §5（自我驗證）的實作。
# 兩邊改動必須同步：說明書是給 AI 看的，這支是給 CTH 雙擊的，內容要一致。
#
# 使用方式：雙擊 scripts\打包.bat（不必開 PowerShell、不必叫 AI）

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$host.UI.RawUI.WindowTitle = "SEC Financial Fetcher - 打包"

$ErrorActionPreference = "Stop"
$Root  = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$Stamp = Get-Date -Format "yyyyMMdd"
$Name  = "SEC-Financial-Fetcher"
$Stage = Join-Path $env:TEMP "$Name-stage-$Stamp"
$Dist  = Join-Path $Root "dist"
$Zip   = Join-Path $Dist "$Name-$Stamp.zip"
$Pkg   = Join-Path $Stage $Name

$fail = @()      # 驗證失敗的項目，最後一次列出
function Check {
    param([string]$Name, [bool]$Ok, [string]$Detail = "")
    if ($Ok) {
        Write-Host ("  [OK]   {0}" -f $Name) -ForegroundColor Green
    } else {
        Write-Host ("  [FAIL] {0}  {1}" -f $Name, $Detail) -ForegroundColor Red
        $script:fail += $Name
    }
}

Clear-Host
Write-Host "SEC Financial Fetcher - 打包" -ForegroundColor Cyan
Write-Host "專案：$Root" -ForegroundColor Gray
Write-Host ""

# ======================================
# [1/4] 打包前檢查
# ======================================
Write-Host "[1/4] 打包前檢查..." -ForegroundColor Cyan

# launcher.ps1 的 UTF-8 BOM（地雷五：沒 BOM 中文訊息會變亂碼、甚至閃退）
$bom = [System.IO.File]::ReadAllBytes((Join-Path $Root "launcher.ps1"))[0..2] -join ','
Check "launcher.ps1 有 UTF-8 BOM" ($bom -eq '239,187,191') "實際為 $bom"

# launcher.ps1 語法
$perr = $null
$null = [System.Management.Automation.Language.Parser]::ParseFile(
    (Join-Path $Root "launcher.ps1"), [ref]$null, [ref]$perr)
Check "launcher.ps1 語法正確" ($perr.Count -eq 0) "$($perr.Count) 個語法錯誤"

if ($fail.Count -gt 0) {
    Write-Host ""
    Write-Host "[STOP] 打包前檢查沒過，不繼續。請先修好上面的問題。" -ForegroundColor Red
    Read-Host "按 Enter 關閉"; exit 1
}

# ======================================
# [2/4] 白名單複製
# ======================================
Write-Host ""
Write-Host "[2/4] 複製要打包的檔案..." -ForegroundColor Cyan

if (Test-Path $Stage) { Remove-Item -Recurse -Force -LiteralPath $Stage }
New-Item -ItemType Directory -Force $Pkg  | Out-Null
New-Item -ItemType Directory -Force $Dist | Out-Null

foreach ($f in @("啟動器.bat","launcher.ps1","README.md","requirements.txt","config.example.json")) {
    Copy-Item (Join-Path $Root $f) $Pkg
}
Copy-Item (Join-Path $Root "src") $Pkg -Recurse

# docs/ 只帶 README 有連到的那一份，其餘內部開發紀錄不外流
New-Item -ItemType Directory -Force (Join-Path $Pkg "docs") | Out-Null
Copy-Item (Join-Path $Root "docs\8k-period-off-by-one.md") (Join-Path $Pkg "docs")

# 快取不要跟著出去。管線進來的路徑一律用 -LiteralPath，
# 直接 `| Remove-Item -Recurse -Force` 會被 Claude Code 的沙箱防護擋掉。
Get-ChildItem $Pkg -Recurse -Directory -Filter "__pycache__" |
    ForEach-Object { Remove-Item -Recurse -Force -LiteralPath $_.FullName }
Get-ChildItem $Pkg -Recurse -File -Filter "*.pyc" |
    ForEach-Object { Remove-Item -Force -LiteralPath $_.FullName }

# 預設輸出目錄（src/config.py 的 output_dir 預設值就是 "output"），空的
New-Item -ItemType Directory -Force (Join-Path $Pkg "output") | Out-Null
New-Item -ItemType File -Force (Join-Path $Pkg "output\.gitkeep") | Out-Null

# 給收件人看的說明，改名放進包裡
Copy-Item (Join-Path $Root "docs\RECIPIENT-README.txt") (Join-Path $Pkg "先讀我.txt")

Write-Host "  複製完成" -ForegroundColor Gray

# ======================================
# [3/4] 壓縮
# ======================================
Write-Host ""
Write-Host "[3/4] 壓縮..." -ForegroundColor Cyan

if (Test-Path $Zip) { Remove-Item -Force -LiteralPath $Zip }
# -Path 指向資料夾本身（結尾不加 \*），zip 內才會保留 SEC-Financial-Fetcher\ 那一層。
# 收件人常直接對 zip 按「解壓縮到這裡」，裸放會把幾十個檔倒進他的下載資料夾。
Compress-Archive -Path $Pkg -DestinationPath $Zip
Write-Host ("  {0}  ({1:N1} KB)" -f (Split-Path $Zip -Leaf), ((Get-Item $Zip).Length/1KB)) -ForegroundColor Gray

# ======================================
# [4/4] 自我驗證：解到另一個暫存目錄實際檢查
# ======================================
Write-Host ""
Write-Host "[4/4] 自我驗證（解壓後逐項檢查）..." -ForegroundColor Cyan

$Verify = Join-Path $env:TEMP "$Name-verify-$Stamp"
if (Test-Path $Verify) { Remove-Item -Recurse -Force -LiteralPath $Verify }
Expand-Archive -Path $Zip -DestinationPath $Verify
$V = Join-Path $Verify $Name

# 1 頂層只有一層資料夾
$topCount = (Get-ChildItem $Verify).Count
Check "頂層只有一層資料夾" ($topCount -eq 1) "有 $topCount 個項目"

# 2 必要檔案齊全
$need = @("啟動器.bat","launcher.ps1","README.md","requirements.txt","config.example.json","先讀我.txt")
$missing = $need | Where-Object { -not (Test-Path (Join-Path $V $_)) }
Check "必要檔案齊全" ($missing.Count -eq 0) "缺：$($missing -join ', ')"

# 3 src/ 完整
$srcPkg  = (Get-ChildItem (Join-Path $V "src")    -Filter *.py).Count
$srcRoot = (Get-ChildItem (Join-Path $Root "src") -Filter *.py).Count
Check "src/ 的 .py 數量相符" ($srcPkg -eq $srcRoot) "包內 $srcPkg / 原始 $srcRoot"

# 4 四個語言檔都在
$loc = @("zh_tw.py","zh_cn.py","en.py","ja.py") | Where-Object { -not (Test-Path (Join-Path $V "src\locales\$_")) }
Check "四個語言檔都在" ($loc.Count -eq 0) "缺：$($loc -join ', ')"

# 5 沒有機敏檔
$sens = Get-ChildItem $V -Recurse -Force -Include config.json,*.log,company_cache.json,.env
Check "沒有機敏檔（config.json/log/快取/.env）" ($sens.Count -eq 0) "$(($sens|ForEach-Object Name) -join ', ')"

# 6 沒有你的 Excel 輸出
$xls = Get-ChildItem $V -Recurse -Force -Filter *.xlsx
Check "沒有 .xlsx 輸出檔" ($xls.Count -eq 0) "$(($xls|ForEach-Object Name) -join ', ')"

# 7 沒有開發用目錄
$devDirs = Get-ChildItem $V -Recurse -Force -Directory |
    Where-Object Name -in '__pycache__','.git','venv','tests','scripts','.pytest_cache','.claude','.superpowers'
Check "沒有開發用目錄" ($devDirs.Count -eq 0) "$(($devDirs|ForEach-Object Name) -join ', ')"

# 8 內容沒挾帶金鑰（真金鑰的格式；config.example.json 的假值不該命中）
$keys = Get-ChildItem $V -Recurse -File -Include *.py,*.json,*.md,*.ps1,*.bat,*.txt |
    Select-String -Pattern 'sk-[A-Za-z0-9]{20}','AIza[A-Za-z0-9_-]{30}' -List
Check "沒有 API 金鑰樣式" ($keys.Count -eq 0) "$(($keys|ForEach-Object Filename) -join ', ')"

# 9 沒有你的真實信箱（只該命中文件裡的範例值）
$mails = Get-ChildItem $V -Recurse -File |
    Select-String -Pattern '[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}' -AllMatches |
    ForEach-Object { $_.Matches.Value } | Sort-Object -Unique
$known = @('your@email.com','john@example.com','sec@example.com','wangdaming@gmail.com',
           'test@test.com','x@y.com')
$unknown = $mails | Where-Object { $_ -notin $known }
Check "沒有非預期的 email" ($unknown.Count -eq 0) "$($unknown -join ', ')"

# 10 大小合理
$mb = (Get-Item $Zip).Length/1MB
Check "大小合理（< 5 MB）" ($mb -lt 5) ("$([math]::Round($mb,2)) MB")

# 11 output/ 是空的
$outFiles = Get-ChildItem (Join-Path $V "output") -Force -Name
Check "output/ 是空的" (($outFiles -join ',') -eq '.gitkeep') "有：$($outFiles -join ', ')"

# 12 內部文件沒外流
$docs = Get-ChildItem (Join-Path $V "docs") -Name
Check "docs/ 只有 8k-period-off-by-one.md" (($docs -join ',') -eq '8k-period-off-by-one.md') "有：$($docs -join ', ')"

# 收尾
Remove-Item -Recurse -Force -LiteralPath $Stage
Remove-Item -Recurse -Force -LiteralPath $Verify

Write-Host ""
if ($fail.Count -gt 0) {
    Remove-Item -Force -LiteralPath $Zip
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Red
    Write-Host "[FAIL] $($fail.Count) 項驗證沒過，zip 已刪除，不要傳出去。" -ForegroundColor Red
    Write-Host "       沒過的項目：$($fail -join '、')" -ForegroundColor Yellow
    Write-Host "       把這個畫面貼給 AI 看，或查 docs\PACKAGING.md。" -ForegroundColor Yellow
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Red
    Read-Host "按 Enter 關閉"; exit 1
}

Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Green
Write-Host "[DONE] 12 項驗證全過，可以傳出去了。" -ForegroundColor Green
Write-Host ""
Write-Host ("  檔案：{0}" -f $Zip) -ForegroundColor White
Write-Host ("  大小：{0:N1} KB" -f ((Get-Item $Zip).Length/1KB)) -ForegroundColor White
Write-Host ""
Write-Host "  收件人解壓後雙擊「啟動器.bat」，一路按 Enter 即可。" -ForegroundColor Gray
Write-Host "  他要做的兩件事寫在包裡的「先讀我.txt」。" -ForegroundColor Gray
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Green
Write-Host ""

$open = Read-Host "要開啟 dist 資料夾嗎？[Y/n] - 直接按 Enter 代表要"
if ($open -eq "" -or $open -ieq "Y") { Start-Process explorer.exe $Dist }
