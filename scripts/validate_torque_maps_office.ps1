param([string]$AuditRoot = 'outputs/torque_maps_2026_10_28', [string]$SourcePattern = '*',
      [switch]$ExerciseMotorOverlays)
$ErrorActionPreference = 'Stop'

function Test-LiveChartEdits($book) {
    # Excel's COM Value2 writes can leave chart caches stale. Keyboard entry
    # exercises the user's normal edit path and verifies live X and Y updates.
    if (-not ('TorqueValidationWindow' -as [type])) {
        Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class TorqueValidationWindow {
 [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hwnd);
 [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
}
"@
    }
    $application = $book.Application
    $application.Visible = $true
    $maps = $book.Worksheets.Item('Torque Maps')
    $details = @()
    try {
        for ($index = 1; $index -le 2; $index++) {
            $chart = $maps.ChartObjects().Item($index).Chart
            $series = $chart.SeriesCollection().Item(1)
            $references = [regex]::Matches($series.Formula,
                "'(?<sheet>(?:[^']|'')+)'!(?<cell>\`$[A-Z]+\`$\d+):\`$[A-Z]+\`$\d+")
            if ($references.Count -ne 2) { throw 'Missing native X/Y source ranges' }
            for ($coordinate = 0; $coordinate -lt 2; $coordinate++) {
                $reference = $references[$coordinate]
                $source = $book.Worksheets.Item($reference.Groups['sheet'].Value.Replace("''", "'"))
                $cell = $source.Range($reference.Groups['cell'].Value)
                $original = [double]$cell.Value2
                $axis = $chart.Axes($coordinate+1)
                $changed = [Math]::Ceiling(3*[Math]::Max([Math]::Abs($axis.MinimumScale), [Math]::Abs($axis.MaximumScale))+1)
                foreach ($value in @($changed, $original)) {
                    [void]$source.Activate()
                    [void]$cell.Select()
                    [void][TorqueValidationWindow]::SetForegroundWindow([IntPtr]$application.Hwnd)
                    Start-Sleep -Milliseconds 250
                    if ([TorqueValidationWindow]::GetForegroundWindow() -ne [IntPtr]$application.Hwnd) {
                        throw 'Private Excel window lacks focus; keyboard entry cancelled'
                    }
                    $entry = $value.ToString('G15', [Globalization.CultureInfo]::InvariantCulture).Replace('.', [string]$application.International(3))
                    $application.SendKeys(($entry+'{ENTER}'), $true)
                    Start-Sleep -Milliseconds 300
                    [void]$maps.Activate()
                    [void]$maps.ChartObjects().Item($index).Activate()
                    $points = if ($coordinate -eq 0) { $series.XValues } else { $series.Values }
                    $actual = $points.GetValue($points.GetLowerBound(0))
                    if ($actual -ne $value) { throw "Normal Excel entry failed: chart=$index coordinate=$coordinate requested=$value cell=$($cell.Value2) plotted=$actual" }
                    if ($value -eq $changed) {
                        if ($axis.MaximumScale -lt $changed) { throw 'Live automatic axis excludes edited point' }
                        $details += @{chart=$index; coordinate=$(if ($coordinate -eq 0) {'X'} else {'Y'});
                                      source_sheet=$source.Name; source_cell=$cell.Address(); original=$original;
                                      edited=$actual; live_updated=$true; auto_axis_includes_edit=$true}
                    }
                }
            }
        }
        return $details
    } finally { $application.Visible = $false }
}

function Test-ManualMotorOverlays($book, $powerpoint, $preview, $label) {
    # Controlled validation points exist only in these disposable Office copies.
    $maps = $book.Worksheets.Item('Torque Maps')
    $inputs = $maps
    $firstRow = $maps.UsedRange.Row + $maps.UsedRange.Rows.Count + 3
    $originals = @()
    for ($index = 1; $index -le 2; $index++) {
        $chart = $maps.ChartObjects().Item($index).Chart
        $count = $chart.SeriesCollection().Count
        $formulas = @($chart.SeriesCollection() | ForEach-Object { [string]$_.Formula })
        foreach ($axis in @($chart.Axes(1), $chart.Axes(2))) {
            if (-not $axis.MinimumScaleIsAuto -or -not $axis.MaximumScaleIsAuto -or -not $axis.MajorUnitIsAuto) {
                throw 'Torque-map axes must be automatic before manual overlays'
            }
        }
        $xLimit = 1 + 2 * [Math]::Max([Math]::Abs($chart.Axes(1).MinimumScale), [Math]::Abs($chart.Axes(1).MaximumScale))
        $yLimit = 1 + 2 * [Math]::Max([Math]::Abs($chart.Axes(2).MinimumScale), [Math]::Abs($chart.Axes(2).MaximumScale))
        $column = if ($index -eq 1) { 1 } else { 4 }
        $xs = @((-$xLimit), 0.0, $xLimit)
        $ys = @($yLimit, ($yLimit / 2), (-$yLimit))
        for ($point = 0; $point -lt 3; $point++) {
            $inputs.Cells.Item($firstRow+$point, $column).Value2 = [double]$xs[$point]
            $inputs.Cells.Item($firstRow+$point, $column+1).Value2 = [double]$ys[$point]
        }
        $overlay = $chart.SeriesCollection().NewSeries()
        $overlay.Name = 'VALIDATION ONLY - manual curve ' + $index
        $overlay.XValues = $inputs.Range($inputs.Cells.Item($firstRow,$column), $inputs.Cells.Item($firstRow+2,$column))
        $overlay.Values = $inputs.Range($inputs.Cells.Item($firstRow,$column+1), $inputs.Cells.Item($firstRow+2,$column+1))
        $overlay.Format.Line.Visible = -1
        $overlay.Format.Line.ForeColor.RGB = 0
        $overlay.MarkerStyle = -4142
        if ($chart.SeriesCollection().Count -ne $count+1) { throw 'Manual series was not added' }
        for ($item = 1; $item -le $count; $item++) {
            if ($chart.SeriesCollection().Item($item).Formula -cne $formulas[$item-1]) {
                throw 'Adding a curve changed simulation series references'
            }
        }
        $originals += @{count=$count; x_limit=$xLimit; y_limit=$yLimit}
    }
    $book.Application.CalculateFullRebuild()
    $editedPath = Join-Path $preview ($label+'_manual_overlay.xlsx')
    $book.SaveCopyAs($editedPath)
    $edited = $book.Application.Workbooks.Open([string]$editedPath, 0, $true)
    $deck = $null
    try {
        $sheet = $edited.Worksheets.Item('Torque Maps')
        $sheet.Activate()
        $deck = $powerpoint.Presentations.Add($false)
        $deck.PageSetup.SlideWidth = 960
        $deck.PageSetup.SlideHeight = 540
        $slide = $deck.Slides.Add(1, 12)
        $details = @()
        for ($index = 1; $index -le 2; $index++) {
            $object = $sheet.ChartObjects().Item($index)
            $chart = $object.Chart
            $object.Activate()
            $chart.Refresh()
            $original = $originals[$index-1]
            if ($chart.SeriesCollection().Count -ne $original.count+1) { throw 'Manual series did not persist' }
            if ($chart.Axes(1).MinimumScale -gt -$original.x_limit -or $chart.Axes(1).MaximumScale -lt $original.x_limit -or
                $chart.Axes(2).MinimumScale -gt -$original.y_limit -or $chart.Axes(2).MaximumScale -lt $original.y_limit) {
                throw 'Automatic axes clip the added curve'
            }
            $image = Join-Path $preview ($label+'_manual_chart_'+$index+'.png')
            if (-not $chart.Export($image,'PNG')) { throw 'Manual chart export failed' }
            $chart.ChartArea.Copy()
            Start-Sleep -Milliseconds 300
            try { $shape = $slide.Shapes.PasteSpecial(10).Item(1) }
            catch { $shape = $slide.Shapes.PasteSpecial(0, 0, '', 0, '', 0).Item(1) }
            $editableType = $null
            $dataLinked = $false
            if ($shape.HasChart) {
                if ($shape.Chart.SeriesCollection().Count -ne $original.count+1) { throw 'Pasted chart lost series' }
                $dataLinked = [bool]$shape.Chart.ChartData.IsLinked
                $editableType = 'Native PowerPoint chart'
            } elseif ($shape.Type -eq 7 -and $shape.OLEFormat.ProgID -like 'Excel.*') {
                $editableType = [string]$shape.OLEFormat.ProgID
            } else {
                throw 'Pasted chart is not an editable embedded Excel object'
            }
            $shape.Left = 20 + ($index-1)*470
            $shape.Top = 90
            $shape.Width = 450
            $shape.Height = 260
            $details += @{series=$chart.SeriesCollection().Count; original_series=$original.count;
                          auto_axes_include_overlay=$true; editable_powerpoint_type=$editableType;
                          powerpoint_data_linked=$dataLinked}
        }
        $deck.SaveAs((Join-Path $preview ($label+'_manual_overlay.pptx')),24)
        $slide.Export((Join-Path $preview ($label+'_manual_powerpoint.png')),'PNG',1600,900)
        return $details
    } finally {
        if ($deck) { $deck.Close() }
        $edited.Close($false)
    }
}
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
            $null, $workbooks, @([string]$officeBook, 0, $false))
        $manualChecks = @()
        $liveChecks = @()
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
                if ($available) {
                    if ($chart.Axes(1).AxisTitle.Text -notlike '*rpm*' -or
                        $chart.Axes(2).AxisTitle.Text -notlike '*Nm*') { throw 'Incorrect numeric axis units' }
                    foreach ($simulation in $chart.SeriesCollection()) {
                        $references = [regex]::Matches($simulation.Formula,
                            "'(?<sheet>(?:[^']|'')+)'!(?<range>\`$[A-Z]+\`$\d+:\`$[A-Z]+\`$\d+)")
                        if ($references.Count -ne 2) { throw 'Expected independent native X/Y cell ranges' }
                        for ($coordinate = 0; $coordinate -lt 2; $coordinate++) {
                            $reference = $references[$coordinate]
                            $source = $book.Worksheets.Item($reference.Groups['sheet'].Value.Replace("''", "'"))
                            $expected = $source.Range($reference.Groups['range'].Value).Value2
                            $actualPoints = if ($coordinate -eq 0) { $simulation.XValues } else { $simulation.Values }
                            if ($actualPoints.Length -ne $expected.GetLength(0)) { throw 'Native series sample count differs' }
                            for ($point = 0; $point -lt $actualPoints.Length; $point++) {
                                if ($actualPoints.GetValue($actualPoints.GetLowerBound(0)+$point) -ne
                                    $expected.GetValue($expected.GetLowerBound(0)+$point, $expected.GetLowerBound(1))) {
                                    throw 'Native plotted coordinate differs from its numeric source'
                                }
                            }
                        }
                    }
                }
                $object.Activate()
                if (-not $chart.Export((Join-Path $preview ($label+'_excel_'+$index+'.png')), 'PNG')) {
                    throw "Excel chart export failed: $label $index"
                }
            }
            $editChecks = @()
            for ($editIndex = 1; $editIndex -le 2; $editIndex++) {
                if ($record.excel_checks.series_counts[$editIndex-1] -eq 0) { continue }
                $excel.Calculation = -4105
                $series = $charts.Item($editIndex).Chart.SeriesCollection().Item(1)
                $sourceCell = [regex]::Match($series.Formula, "'(?<sheet>(?:[^']|'')+)'!(?<cell>\`$[A-Z]+\`$\d+)")
                if (-not $sourceCell.Success) { throw "Missing native RPM source reference: $label" }
                $sourceSheet = $book.Worksheets.Item($sourceCell.Groups['sheet'].Value.Replace("''", "'"))
                $cell = $sourceSheet.Range($sourceCell.Groups['cell'].Value)
                $original = $cell.Value2
                $cell.Value2 = $original + 1
                $sheet.Calculate()
                $excel.CalculateFullRebuild()
                $charts.Item($editIndex).Activate()
                $charts.Item($editIndex).Chart.Refresh()
                [void]$charts.Item($editIndex).Chart.Export((Join-Path $preview ($label+'_edit_check_'+$editIndex+'.png')), 'PNG')
                Start-Sleep -Milliseconds 500
                $series = $charts.Item($editIndex).Chart.SeriesCollection().Item(1)
                $points = $series.XValues
                $actual = $points.GetValue($points.GetLowerBound(0))
                $editMethod = 'live'
                if ($actual -ne $original+1) {
                    # Hidden Excel can retain its display cache. Verify
                    # persistent editability independently in a disposable copy.
                    $editedPath = Join-Path $preview ($label+'_edited_'+$editIndex+'.xlsx')
                    $book.SaveCopyAs($editedPath)
                    $editedBook = $workbooks.Open([string]$editedPath, 0, $true)
                    try {
                        $editedSheet = $editedBook.Worksheets.Item('Torque Maps')
                        $editedSheet.Activate()
                        $editedSheet.ChartObjects().Item($editIndex).Activate()
                        $editedPoints = $editedSheet.ChartObjects().Item($editIndex).Chart.SeriesCollection().Item(1).XValues
                        $actual = $editedPoints.GetValue($editedPoints.GetLowerBound(0))
                    } finally { $editedBook.Close($false) }
                    $editMethod = 'save-copy/reopen'
                }
                if ($actual -ne $original+1) { throw "Edited source does not reach chart after reload: $label" }
                $cell.Value2 = $original
                $editChecks += @{chart=$editIndex; source_sheet=$sourceSheet.Name; source_cell=$cell.Address();
                                 before=$original; edited=$actual; validation=$editMethod}
            }
            $sheet.PageSetup.PrintArea = '$A$1:$Z$31'
            $sheet.ExportAsFixedFormat(0, (Join-Path $preview ($label+'_excel.pdf')))
            if ($ExerciseMotorOverlays) {
                $liveChecks = @(Test-LiveChartEdits $book | Where-Object { $null -ne $_ })
                if ($liveChecks.Count -ne 4) { throw 'Expected four live X/Y cell-edit checks' }
                $manualChecks = @(Test-ManualMotorOverlays $book $powerpoint $preview $label | Where-Object { $null -ne $_ })
                if ($manualChecks.Count -ne 2) { throw 'Expected two editable pasted-chart validation records' }
            }
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
                         operating_series_counts=$record.excel_checks.series_counts; all_native_points_match_source=$true;
                         source_edit_propagates=$true; source_edit_checks=$editChecks; slides=$deck.Slides.Count; map_slide=$slide.SlideIndex; text_overflows=$overflow.Count;
                         manual_curve_checks=$manualChecks; live_cell_edit_checks=$liveChecks}
        } finally { $deck.Close() }
        $checks | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $preview 'checks.json') -Encoding UTF8
        Write-Output ('PASS native Office: '+$property.Name)
    }
} finally {
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    if ($powerpoint) { $powerpoint.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($powerpoint) }
}
