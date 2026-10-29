param([string]$AuditRoot = 'outputs/torque_maps_2026_10_28', [string]$SourcePattern = '*')
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath $AuditRoot).Path
$preview = Join-Path $root 'office_preview'
New-Item -ItemType Directory -Path $preview -Force | Out-Null
$records = Get-Content -LiteralPath (Join-Path $root 'validation.json') -Raw | ConvertFrom-Json
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
    foreach ($property in $records.PSObject.Properties) {
        if ($property.Name -notlike $SourcePattern) { continue }
        $record = $property.Value
        $label = [IO.Path]::GetFileName((Split-Path -Parent (Split-Path -Parent $record.excel)))
        # Short disposable copies avoid native Office's path-length restrictions.
        $officeBook = Join-Path $preview ($label+'.xlsx')
        $officeDeck = Join-Path $preview ($label+'.pptx')
        Copy-Item -LiteralPath $record.excel -Destination $officeBook -Force
        Copy-Item -LiteralPath $record.powerpoint -Destination $officeDeck -Force
        $workbooks = $excel.Workbooks
        $book = $workbooks.GetType().InvokeMember('Open', [Reflection.BindingFlags]::InvokeMethod,
            $null, $workbooks, @([string]$officeBook, 0, $true))
        try {
            $sheet = $book.Worksheets.Item('Torque Maps')
            $sheet.Activate()
            $window = $book.Windows.Item(1)
            if ($window.FreezePanes -or $window.Split -or $window.SplitRow -ne 0 -or $window.SplitColumn -ne 0) {
                throw "Torque Maps has frozen or split panes: $label"
            }
            $window.ScrollRow = 40
            $window.ScrollColumn = 15
            if ($window.ScrollRow -ne 40 -or $window.ScrollColumn -ne 15) {
                throw "Torque Maps cannot scroll freely: $label"
            }
            $window.ScrollRow = 1
            $window.ScrollColumn = 1
            $charts = $sheet.ChartObjects()
            if ($charts.Count -ne 2) { throw "Expected two native Torque Maps charts: $label" }
            $seriesCounts = @()
            for ($index = 1; $index -le 2; $index++) {
                $object = $charts.Item($index)
                $chart = $object.Chart
                # Excel classifies an empty/marker-free scatter as subtype 75.
                # Populated operating maps must retain marker-only subtype -4169.
                $available = $record.excel_checks.series_counts[$index-1] -gt 0
                if (($available -and $chart.ChartType -ne -4169) -or
                    (-not $available -and $chart.ChartType -notin @(-4169,75))) {
                    throw "Incorrect native chart type: $label $index = $($chart.ChartType)"
                }
                $seriesCounts += $chart.SeriesCollection().Count
                $object.Activate()
                if (-not $chart.Export((Join-Path $preview ($label+'_excel_'+$index+'.png')), 'PNG')) {
                    throw "Excel chart export failed: $label $index"
                }
            }
            if ($seriesCounts[0] -gt 0) {
                $excel.Calculation = -4105
                $series = $charts.Item(1).Chart.SeriesCollection().Item(1)
                $sourceCell = [regex]::Match($series.Formula, '''Torque Maps''!(?<cell>\$[A-Z]+\$\d+)')
                if (-not $sourceCell.Success) { throw "Missing native RPM source reference: $label" }
                $cell = $sheet.Range($sourceCell.Groups['cell'].Value)
                $original = $cell.Value2
                $cell.Value2 = $original + 1
                $sheet.Calculate()
                $excel.CalculateFullRebuild()
                $charts.Item(1).Activate()
                $charts.Item(1).Chart.Refresh()
                [void]$charts.Item(1).Chart.Export((Join-Path $preview ($label+'_edit_check.png')), 'PNG')
                Start-Sleep -Milliseconds 500
                $series = $charts.Item(1).Chart.SeriesCollection().Item(1)
                $points = $series.XValues
                $actual = $points.GetValue($points.GetLowerBound(0))
                $editMethod = 'live'
                if ($actual -ne $original+1) {
                    # Hidden, read-only Excel can retain its display cache. Verify
                    # persistent editability independently in a disposable copy.
                    $editedPath = Join-Path $preview ($label+'_edited.xlsx')
                    $book.SaveCopyAs($editedPath)
                    $editedBook = $workbooks.Open([string]$editedPath, 0, $true)
                    try {
                        $editedSheet = $editedBook.Worksheets.Item('Torque Maps')
                        $editedSheet.Activate()
                        $editedSheet.ChartObjects().Item(1).Activate()
                        $editedPoints = $editedSheet.ChartObjects().Item(1).Chart.SeriesCollection().Item(1).XValues
                        $actual = $editedPoints.GetValue($editedPoints.GetLowerBound(0))
                    } finally { $editedBook.Close($false) }
                    $editMethod = 'save-copy/reopen'
                }
                if ($actual -ne $original+1) { throw "Edited source does not reach chart after reload: $label" }
                $cell.Value2 = $original
            }
            $sheet.PageSetup.PrintArea = '$A$1:$Z$31'
            $sheet.ExportAsFixedFormat(0, (Join-Path $preview ($label+'_excel.pdf')))
        } finally { $book.Close($false) }
        $deck = $powerpoint.Presentations.Open([string]$officeDeck, $true, $false, $false)
        try {
            $slide = $deck.Slides.Item($deck.Slides.Count-1)
            $expectedSubtitle = 'VSM simulation operating points showing the relationship between rotational speed and torque.'
            $subtitles = @($slide.Shapes | Where-Object { $_.HasTextFrame -and $_.TextFrame.TextRange.Text -ceq $expectedSubtitle })
            if ($subtitles.Count -ne 1) { throw "PowerPoint subtitle differs: $label" }
            $slide.Export((Join-Path $preview ($label+'_powerpoint.png')), 'PNG', 1800, 1013)
            $overflow = @()
            foreach ($shape in $slide.Shapes) {
                if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                    if ($shape.TextFrame.TextRange.BoundHeight -gt $shape.Height+1 -or
                        $shape.TextFrame.TextRange.BoundWidth -gt $shape.Width+1) {
                        $overflow += $shape.TextFrame.TextRange.Text
                    }
                }
            }
            if ($overflow.Count) { throw "Torque map text overflow: $label $overflow" }
            $checks += @{source=$property.Name; excel_native_charts=2; native_series_counts=$seriesCounts;
                         no_frozen_or_split_panes=$true; free_horizontal_vertical_scroll=$true; subtitle_exact=$true;
                         operating_series_counts=$record.excel_checks.series_counts;
                         source_edit_propagates=$true; source_edit_validation=$editMethod; slides=$deck.Slides.Count; map_slide=$slide.SlideIndex; text_overflows=$overflow.Count}
        } finally { $deck.Close() }
        $checks | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $preview 'checks.json') -Encoding UTF8
        Write-Output ('PASS native Office: '+$property.Name)
    }
} finally {
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    if ($powerpoint) { $powerpoint.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($powerpoint) }
}
