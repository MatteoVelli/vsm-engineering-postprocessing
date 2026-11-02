param([string]$AuditRoot = 'outputs/bcapvalid')
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
    $case = 0
    foreach ($property in $records.PSObject.Properties) {
        $record = $property.Value
        $path = Join-Path $preview ('case_'+$case+'.xlsx')
        $pptPath = Join-Path $preview ('case_'+$case+'.pptx')
        Copy-Item -LiteralPath $record.excel -Destination $path
        Copy-Item -LiteralPath $record.powerpoint -Destination $pptPath
        $book = $excel.Workbooks.Open([string]$path, 0, $false)
        try {
            $excel.Calculation = -4105
            $excel.CalculateFullRebuild()
            $sheet = $book.Worksheets.Item('Statistics')
            $ids = @{}
            for ($row=2; $row -le $sheet.UsedRange.Rows.Count; $row++) {
                $id = $sheet.Cells.Item($row,1).Value2
                if ($id) { $ids[[string]$id] = $row }
            }
            $values = @{}
            foreach ($expected in $record.kpis.PSObject.Properties) {
                $actual = $sheet.Cells.Item($ids[$expected.Name],3).Value2
                $tolerance = [Math]::Max(1e-10, [Math]::Abs([double]$expected.Value)*2e-12)
                if ($actual -isnot [double] -or [Math]::Abs($actual-[double]$expected.Value) -gt $tolerance) {
                    throw "Excel recomputed KPI differs: $($expected.Name) actual=$actual expected=$($expected.Value)"
                }
                $values[$expected.Name] = $actual
            }
            $maps = $book.Worksheets.Item('Torque Maps')
            if ($maps.ChartObjects().Count -ne 2) { throw 'Missing native Torque Maps charts' }
            foreach ($chartObject in $maps.ChartObjects()) {
                if ($chartObject.Chart.SeriesCollection().Count -ne 4) { throw 'Missing Torque Maps simulation series' }
                foreach ($series in $chartObject.Chart.SeriesCollection()) {
                    if ($series.Formula -notlike "*'Torque Maps'!*") { throw 'Torque Maps must use local numerical tables' }
                }
            }
            $book.SaveCopyAs((Join-Path $preview ('case_'+$case+'_recalculated.xlsx')))
        } finally { $book.Close($false) }
        $deck = $powerpoint.Presentations.Open([string]$pptPath, $true, $false, $false)
        try {
            if ($deck.Slides.Count -ne 16) { throw 'Electric deck must retain 16 slides' }
            $allText = @()
            foreach ($slide in $deck.Slides) {
                foreach ($shape in $slide.Shapes) {
                    if ($shape.HasTextFrame -and $shape.TextFrame.HasText) { $allText += $shape.TextFrame.TextRange.Text }
                }
            }
            $formatted = ([double]$record.kpis.range_85_battery_km).ToString('F2',[Globalization.CultureInfo]::InvariantCulture)
            if (($allText -join ' ') -notlike ('*'+$formatted+'*')) { throw 'PowerPoint range text differs from Excel' }
            $deck.Slides.Item(7).Export((Join-Path $preview ('case_'+$case+'_battery_range.png')),'PNG',1800,1013)
            $deck.Slides.Item(15).Export((Join-Path $preview ('case_'+$case+'_torque_maps.png')),'PNG',1800,1013)
            $checks += @{source=$property.Name; reference_status=$record.capacity_interpretation.reference_status;
                         excel_recalculated_kpis=$values; all_kpis_match_python=$true;
                         powerpoint_range_text=$formatted; slides=16; native_local_torque_maps_preserved=$true}
        } finally { $deck.Close() }
        $checks | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $preview 'checks.json') -Encoding UTF8
        Write-Output ('PASS Excel recalculation / PowerPoint: '+$property.Name)
        $case++
    }
} finally {
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    if ($powerpoint) { $powerpoint.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($powerpoint) }
}
