param([Parameter(Mandatory = $true)][string]$AuditPath)
$ErrorActionPreference = 'Stop'
$auditFile = (Resolve-Path -LiteralPath $AuditPath).Path
$destination = Join-Path (Split-Path -Parent $auditFile) 'office'
if (Test-Path -LiteralPath $destination) { throw "Preserve existing Office evidence: $destination" }
New-Item -ItemType Directory -Path $destination | Out-Null
$excel = $null
$powerpoint = $null
$checks = @()
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3
    $powerpoint = New-Object -ComObject PowerPoint.Application
    $index = 0
    foreach ($entry in (Get-Content -LiteralPath $auditFile -Raw | ConvertFrom-Json)) {
        $label = 'case_' + $index
        $index++
        $inputBook = Join-Path $destination ($label + '_empty.xlsx')
        $calculatedBook = Join-Path $destination ($label + '_calculated.xlsx')
        $inputDeck = Join-Path $destination ($label + '.pptx')
        Copy-Item -LiteralPath $entry.cache_cleared_workbook -Destination $inputBook
        Copy-Item -LiteralPath $entry.powerpoint -Destination $inputDeck
        $book = $excel.Workbooks.Open($inputBook, 0, $true)
        $charts = @()
        try {
            $book.ForceFullCalculation = $false
            $excel.Calculation = -4105
            $excel.CalculateFullRebuild()
            foreach ($sheet in $book.Worksheets) { $sheet.Calculate() }
            $timer = [Diagnostics.Stopwatch]::StartNew()
            while ([int]$excel.CalculationState -eq 1 -and $timer.ElapsedMilliseconds -lt 30000) {
                Start-Sleep -Milliseconds 100
            }
            if ([int]$excel.CalculationState -eq 1) { throw 'Excel calculation did not finish' }
            $sheet = $book.Worksheets.Item(1)
            foreach ($object in $sheet.ChartObjects()) {
                $chart = $object.Chart
                if (-not $chart.HasTitle) { continue }
                $title = [string]$chart.ChartTitle.Text
                if ($title -notmatch 'Wheel Power|Wheel Torque|Power at Wheels') { continue }
                $image = Join-Path $destination ($label + '_chart_' + $charts.Count + '.png')
                $object.Activate()
                if (-not $chart.Export($image, 'PNG')) { throw "Chart export failed: $title" }
                $series = @()
                foreach ($line in $chart.SeriesCollection()) {
                    $series += @{name=[string]$line.Name; formula=[string]$line.Formula}
                }
                $charts += @{title=$title; series=$series; preview=$image}
            }
            $book.SaveCopyAs($calculatedBook)
        } finally { $book.Close($false) }
        $deck = $powerpoint.Presentations.Open($inputDeck, $true, $false, $false)
        try {
            $deck.SaveAs((Join-Path $destination ($label + '.pdf')), 32)
            $deck.Slides.Item(9).Export((Join-Path $destination ($label + '_traction.png')), 'PNG', 1600, 900)
            $checks += @{source=$entry.source; original_excel=$entry.excel; empty_cache_copy=$inputBook;
                         recalculated_excel=$calculatedBook; charts=$charts; slides=$deck.Slides.Count}
        } finally { $deck.Close() }
        $checks | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $destination 'checks.json') -Encoding UTF8
        Write-Output ('PASS desktop Office: ' + $entry.source)
    }
} finally {
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    if ($powerpoint) { $powerpoint.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($powerpoint) }
}
