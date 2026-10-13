param(
    [Parameter(Mandatory=$true)][string]$Candidate,
    [Parameter(Mandatory=$true)][string]$PipelineJson,
    [Parameter(Mandatory=$true)][string]$ResultJson
)

$ErrorActionPreference = 'Stop'
$excel = $null
$book = $null
$ownsApp = $false
try {
    $candidatePath = (Resolve-Path -LiteralPath $Candidate).Path
    $pipeline = Get-Content -LiteralPath $PipelineJson -Raw -Encoding UTF8 | ConvertFrom-Json
    $excel = New-Object -ComObject Excel.Application
    if ($excel.Workbooks.Count -ne 0) { throw 'New Excel instance already contains books; ownership unclear' }
    $ownsApp = $true
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($candidatePath, 0, $false)
    if ($book.ReadOnly) { throw 'Candidate is read-only' }
    $excel.CalculateFullRebuild()
    $headers = 0
    foreach ($table in $pipeline.tables) {
        if ($table.sheet_name -notin @('Data_Financials(Q)', 'Data_Financials(Y)')) { continue }
        $sheet = $book.Worksheets.Item($table.sheet_name)
        for ($i=0; $i -lt $table.quarter_labels.Count; $i++) {
            $actual = $sheet.Cells.Item(1, $i+4).Value2
            if ($actual -ne $table.quarter_labels[$i]) { throw "Header mismatch in $($table.sheet_name), column $($i+4)" }
            $headers++
        }
    }
    $formulaErrors = @()
    foreach ($sheet in $book.Worksheets) {
        try { $errors = $sheet.UsedRange.SpecialCells(-4123, 16) }
        catch [System.Runtime.InteropServices.COMException] {
            if ($_.Exception.HResult -ne -2146827284) { throw }
            $errors = $null
        }
        if ($null -ne $errors) { $formulaErrors += "$($sheet.Name):$($errors.Address())" }
    }
    if ($formulaErrors.Count) { throw ('Excel formula errors: ' + ($formulaErrors -join ', ')) }
    $book.Save()
    $result = [ordered]@{ status='PASS'; candidate=$candidatePath; native_excel_recalculated=$true;
        period_headers_checked=$headers; formula_errors=0; sheets=$book.Worksheets.Count }
    $result | ConvertTo-Json | Set-Content -LiteralPath $ResultJson -Encoding UTF8
    $result | ConvertTo-Json
}
finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($ownsApp) { $excel.Quit() }
}
