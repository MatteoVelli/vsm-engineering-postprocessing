param([Parameter(Mandatory = $true)][string]$AuditPath)
$ErrorActionPreference = 'Stop'
$auditFile = (Resolve-Path -LiteralPath $AuditPath).Path
$destination = Join-Path (Split-Path -Parent $auditFile) 'native_excel'
New-Item -ItemType Directory -Path $destination -Force | Out-Null
$agroExcel = $null
$results = @()
try {
    $agroExcel = New-Object -ComObject Excel.Application
    $agroExcel.Visible = $false
    $agroExcel.DisplayAlerts = $false
    $agroExcel.AutomationSecurity = 3
    $agroExcel.AskToUpdateLinks = $false
    foreach ($entry in (Get-Content -LiteralPath $auditFile -Raw | ConvertFrom-Json)) {
        $inputPath = if ($entry.cache_cleared_workbook) { $entry.cache_cleared_workbook } else { $entry.workbook }
        $book = $agroExcel.Workbooks.Open($inputPath, 0, $true)
        try {
            $book.ForceFullCalculation = $false
            $agroExcel.CalculateFullRebuild()
            $checks = @()
            foreach ($chartCheck in $entry.chart_checks) {
                $series = $book.Worksheets.Item($chartCheck.sheet).ChartObjects($chartCheck.chart).Chart.SeriesCollection($chartCheck.series)
                $values = @($series.Values)
                if ($values.Count -ne $chartCheck.count) { throw ('Chart sample count mismatch: ' + $entry.case) }
                $limits = $values | Measure-Object -Minimum -Maximum
                foreach ($pair in @(@($limits.Minimum,$chartCheck.minimum),@($limits.Maximum,$chartCheck.maximum))) {
                    if ([Math]::Abs([double]$pair[0]-[double]$pair[1]) -gt (1e-10*[Math]::Max(1,[Math]::Abs([double]$pair[1])))) {
                        throw ('Chart sample values mismatch: ' + $entry.case)
                    }
                }
                Write-Output ($entry.case + ': native chart source samples verified')
            }
            $stages = @(@{name = 'initial'; metrics = $entry.metrics; edits = @()})
            if ($entry.mutations) { $stages += $entry.mutations }
            foreach ($stage in $stages) {
                foreach ($edit in $stage.edits) {
                    $editValue = $edit.value
                    $editCell = $book.Worksheets.Item($edit.sheet).Range($edit.cell)
                    if ($null -eq $editValue) { [void]$editCell.ClearContents() }
                    else {
                        if ($editValue -is [ValueType]) { $editValue = [double]$editValue }
                        $editCell.Value2 = $editValue
                    }
                }
                $agroExcel.CalculateFullRebuild()
                $book.Worksheets.Item('Statistics').Calculate()
                foreach ($metric in $stage.metrics.PSObject.Properties) {
                    # Existing unavailable kg/km cells are intentionally unchanged.
                    if ($metric.Value.type -ne 'f') { continue }
                    $cell = $book.Worksheets.Item('Statistics').Range($metric.Value.value_cell)
                    if (-not $cell.HasFormula) { throw ('Value is not a formula: ' + $cell.Address()) }
                    $actual = $cell.Value2
                    $expected = $metric.Value.cached_value
                    if ($expected -is [string]) {
                        if ($actual -ne $expected) { throw ('Text mismatch: ' + $entry.case + '/' + $metric.Name + ' = ' + $actual) }
                    } elseif ($actual -is [string] -or $null -eq $actual -or [Math]::Abs([double]$actual - [double]$expected) -gt (1e-11 * [Math]::Max(1, [Math]::Abs([double]$expected)))) {
                        throw ('Numerical mismatch: ' + $entry.case + '/' + $metric.Name + ' expected ' + $expected + ', got ' + $actual)
                    }
                    $checks += @{stage = $stage.name; metric = $metric.Name; cell = $metric.Value.value_cell;
                        has_formula = [bool]$cell.HasFormula; formula_bar = $cell.Formula; recalculated_value = $actual}
                }
                Write-Output ($entry.case + '/' + $stage.name + ': values verified')
            }
            $outputPath = Join-Path $destination ($entry.case + '.xlsx')
            $book.SaveCopyAs($outputPath)
            $results += @{case = $entry.case; input = $inputPath; saved_copy = $outputPath; checks = $checks}
            Write-Output ($entry.case + ': native Excel formula bar and recalculation PASS (' + $checks.Count + ' checks)')
        } finally { $book.Close($false) }
    }
    ConvertTo-Json -InputObject $results -Depth 10 | Set-Content -LiteralPath (Join-Path $destination 'results.json') -Encoding UTF8
} finally {
    if ($agroExcel) { $agroExcel.Quit(); [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($agroExcel) }
}
