param([Parameter(Mandatory = $true)][string]$AuditPath)
$ErrorActionPreference = 'Stop'
$auditFile = (Resolve-Path -LiteralPath $AuditPath).Path
$auditRoot = Split-Path -Parent $auditFile
$destination = Join-Path $auditRoot 'recalculated'
New-Item -ItemType Directory -Path $destination -Force | Out-Null
$excelApp = $null
try {
    $excelApp = New-Object -ComObject Excel.Application
    $excelApp.Visible = $false
    $excelApp.DisplayAlerts = $false
    $excelApp.AutomationSecurity = 3
    $excelApp.AskToUpdateLinks = $false
    foreach ($entry in (Get-Content -LiteralPath $auditFile -Raw | ConvertFrom-Json)) {
        $inputPath = $entry.cache_cleared_workbook
        if (-not $inputPath) { throw 'Run --prepare-recalculation first.' }
        $outputPath = Join-Path $destination (Split-Path -Leaf $inputPath)
        if (Test-Path -LiteralPath $outputPath) { throw ('Preserve existing recalculation evidence: ' + $outputPath) }
        $book = $excelApp.Workbooks.Open($inputPath, 0, $true)
        try {
            # The generated report requests full calculation on every opening.
            # Clear those persistent dirty flags only in this disposable copy,
            # then explicitly rebuild and calculate every dependency.
            $book.ForceFullCalculation = $false
            $excelApp.Calculation = -4105
            $excelApp.CalculateFullRebuild()
            foreach ($sheet in $book.Worksheets) { $sheet.Calculate() }
            $timer = [System.Diagnostics.Stopwatch]::StartNew()
            while ([int]$excelApp.CalculationState -eq 1 -and $timer.ElapsedMilliseconds -lt 30000) {
                Start-Sleep -Milliseconds 100
            }
            if ([int]$excelApp.CalculationState -eq 1) { throw 'Excel is still calculating.' }
            # Excel can retain xlPending for chart refresh/full-load flags.
            # Accept evidence only after Python checks every newly calculated
            # numeric value against the engine; every input cache was empty.
            $book.SaveCopyAs($outputPath)
            Write-Output ($entry.profile + ': CalculateFullRebuild and per-sheet Calculate invoked; state=' + $excelApp.CalculationState + '; independent copy saved for numerical audit')
        } finally { $book.Close($false) }
    }
} finally {
    if ($excelApp) { $excelApp.Quit(); [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excelApp) | Out-Null }
}
