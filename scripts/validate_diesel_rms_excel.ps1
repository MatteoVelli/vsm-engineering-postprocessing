param(
    [string]$OutputRoot = (Join-Path $PSScriptRoot '../outputs/diesel_rms_cleanup')
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
$taskRoot = (Resolve-Path -LiteralPath $OutputRoot).Path
$validation = Get-Content -LiteralPath (Join-Path $taskRoot 'final.json') -Raw | ConvertFrom-Json
$visualRoot = Join-Path $taskRoot 'visual'
New-Item -ItemType Directory -Force -Path $visualRoot | Out-Null
$excel = New-Object -ComObject Excel.Application
$book = $null
try {
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $book = $excel.Workbooks.Open($validation.diesel.xlsx, 0, $true)
    $sheet = $book.Worksheets.Item(1)
    $excel.CalculateFullRebuild()
    $results = @{}
    foreach ($metric in @('engine_torque', 'engine_power')) {
        $expected = $validation.diesel.headline_rms.$metric
        $cell = $sheet.Range($expected.cell)
        $actual = [double]$cell.Value2
        if ([Math]::Abs($actual - $expected.value) -gt 1e-11) {
            throw "Excel RMS mismatch for ${metric}: $actual vs $($expected.value)"
        }
        if ($cell.Formula -ne $expected.formula) { throw "Formula changed for $metric" }
        $results[$metric] = @{value = $actual; formula = $cell.Formula; cell = $expected.cell; unit = $expected.unit}
    }
    # Change a source sample and prove that Excel follows the visible formula.
    $source = $sheet.Range('F5')
    $original = [double]$source.Value2
    $source.Value2 = $original + 17.0
    $excel.CalculateFullRebuild()
    $changed = [double]$sheet.Range('F2').Value2
    $expectedChange = [Math]::Sqrt([Math]::Pow($results.engine_torque.value, 2) + (([Math]::Pow($original + 17.0, 2) - [Math]::Pow($original, 2)) / $validation.diesel.samples))
    if ([Math]::Abs($changed - $expectedChange) -gt 1e-11) { throw 'Live RMS edit failed' }
    $source.Value2 = $original
    $excel.CalculateFullRebuild()
    $book.SaveCopyAs((Join-Path $visualRoot 'diesel_excel_recalculated.xlsx'))
    $results['source_edit_recalculates'] = $true
    [void]$sheet.Activate()
    $excel.ActiveWindow.DisplayGridlines = $true
    $excel.ActiveWindow.Zoom = 100
    $areas = @(
        @{name = 'engine_torque_rms'; range = 'E1:H12'},
        @{name = 'engine_power_rms'; range = 'X1:AA12'},
        @{name = 'summary_and_charts'; range = 'AC1:AQ101'},
        @{name = 'bottom_statistics'; range = 'A225:AA232'}
    )
    foreach ($area in $areas) {
        $range = $sheet.Range($area.range)
        [void]$range.Select()
        [void]$range.CopyPicture(1, 2)
        $bitmap = [System.Windows.Forms.Clipboard]::GetImage()
        if ($null -eq $bitmap) { throw "No native Excel rendering for $($area.name)" }
        try {
            $bitmap.Save((Join-Path $visualRoot ($area.name + '.png')), [System.Drawing.Imaging.ImageFormat]::Png)
        } finally { $bitmap.Dispose() }
    }
    foreach ($chart in $sheet.ChartObjects()) {
        $index = $chart.Index
        $path = Join-Path $visualRoot ('chart_{0:00}.png' -f $index)
        if (-not $chart.Chart.Export($path, 'PNG')) { throw "Chart $index did not render" }
    }
    $sheet.PageSetup.PrintArea = 'AC1:AQ101'
    $sheet.PageSetup.Orientation = 2
    $sheet.PageSetup.Zoom = $false
    $sheet.PageSetup.FitToPagesWide = 1
    $sheet.PageSetup.FitToPagesTall = 1
    $sheet.ExportAsFixedFormat(0, (Join-Path $visualRoot 'diesel_summary_charts.pdf'))
    $results | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $visualRoot 'excel_recalculation.json') -Encoding UTF8
    $results | ConvertTo-Json -Depth 5
} finally {
    if ($null -ne $book) { $book.Close($false) }
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
}
