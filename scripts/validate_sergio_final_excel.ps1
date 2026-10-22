param([string]$AuditPath = 'outputs/sergio_final_2026_10_17/excel_recalculation/audit.json')
$ErrorActionPreference = 'Stop'
$auditFile = (Resolve-Path -LiteralPath $AuditPath).Path
$entry = Get-Content -LiteralPath $auditFile -Raw | ConvertFrom-Json
$destination = Split-Path -Parent $auditFile
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 3
$excel.AskToUpdateLinks = $false
try {
    $book = $excel.Workbooks.Open($entry.cache_cleared_workbook, 0, $true)
    try {
        $book.ForceFullCalculation = $false
        $excel.Calculation = -4105
        $excel.CalculateFullRebuild()
        foreach ($sheet in $book.Worksheets) { $sheet.Calculate() }
        $book.SaveCopyAs((Join-Path $destination 'diesel_recalculated.xlsx'))
        $charts = $book.Worksheets.Item(1).ChartObjects()
        foreach ($number in @(3,4)) {
            $chart = $charts.Item($number)
            $chart.Activate()
            [void]$chart.Chart.Export((Join-Path $destination ('chart_' + $number + '.png')), 'PNG')
        }
        Write-Output ('Desktop Excel: rebuilt formulas and saved independent copy; charts=' + $charts.Count)
    } finally { $book.Close($false) }
} finally {
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
}
