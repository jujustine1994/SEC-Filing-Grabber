param([string]$WorkbookFolder)
$ErrorActionPreference='Stop'
$auditApp=$null
$auditBook=$null
$auditResults=@()
try {
    $auditApp=New-Object -ComObject Excel.Application
    $auditApp.Visible=$false
    $auditApp.DisplayAlerts=$false
    $auditApp.AskToUpdateLinks=$false
    foreach($path in Get-ChildItem -LiteralPath $WorkbookFolder -Filter '*.xlsx') {
        $auditBook=$auditApp.Workbooks.Open($path.FullName,0,$false)
        $auditApp.CalculateFullRebuild()
        $auditBook.Save()
        $indexSheet=$auditBook.Worksheets.Item('Index')
        $originalMonth=$indexSheet.Range('B4').Value2
        $overrideMonth=if($originalMonth -eq 12){1}else{$originalMonth+1}
        $indexSheet.Range('B4').Value2=$overrideMonth
        $auditApp.CalculateFullRebuild()
        $overrideHeaders=@()
        foreach($sheetName in @('Data_Financials(Q)','Data_Financials(Y)')) {
            $sheet=$auditBook.Worksheets.Item($sheetName)
            for($col=4;$col -le $sheet.UsedRange.Columns.Count;$col++) {
                $overrideHeaders+=@{sheet=$sheetName;column=$col;label=$sheet.Cells.Item(1,$col).Value2;fiscal=$sheet.Cells.Item(3,$col).Value2;calendar=$sheet.Cells.Item(4,$col).Value2;end=$sheet.Cells.Item(5,$col).Value2}
            }
        }
        $auditResults+=@{ticker=$path.BaseName;defaultMonth=$originalMonth;overrideMonth=$overrideMonth;overrideHeaders=$overrideHeaders}
        $auditBook.Close($false)
        $auditBook=$null
        Write-Output "$($path.BaseName) recalculated"
    }
    $auditResults | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $WorkbookFolder 'excel-recalculation.json') -Encoding utf8
} finally {
    if($null -ne $auditBook){$auditBook.Close($false)}
    if($null -ne $auditApp){$auditApp.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($auditApp)}
}
