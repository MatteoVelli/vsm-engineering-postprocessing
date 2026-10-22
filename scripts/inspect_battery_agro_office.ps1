param([string[]]$Scenarios = @('electric_confirmed', 'hybrid_invalid_load', 'electric_zero_load', 'october_electric', 'october_hybrid', 'full_size_diesel'), [string]$AuditRoot = '')
$ErrorActionPreference = 'Stop'
if (-not $AuditRoot) { $AuditRoot = Join-Path $PSScriptRoot '..\outputs\battery_agro_validation' }
$auditRoot = (Resolve-Path -LiteralPath $auditRoot).Path
$previewRoot = Join-Path $auditRoot 'preview'
New-Item -ItemType Directory -Path $previewRoot -Force | Out-Null
$pptApp = $null
$excelApp = $null
$checks = @()
try {
    $pptApp = New-Object -ComObject PowerPoint.Application
    $excelApp = New-Object -ComObject Excel.Application
    $excelApp.Visible = $false
    $excelApp.DisplayAlerts = $false
    $excelApp.AskToUpdateLinks = $false
    $excelApp.AutomationSecurity = 3
    foreach ($scenario in $Scenarios) {
        $folder = Join-Path $auditRoot $scenario
        $manifestPath = Join-Path $folder 'excel\profile_excel_report_manifest.json'
        $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
        $deckPath = (Get-ChildItem -LiteralPath (Join-Path $folder 'powerpoint') -Filter '*.pptx' | Select-Object -First 1).FullName
        $deck = $pptApp.Presentations.Open($deckPath, $true, $false, $false)
        try {
            $page = if ($scenario -in @('full_size_diesel', 'diesel')) { 3 } else { 8 }
            $slide = $deck.Slides.Item($page)
            $slide.Export((Join-Path $previewRoot ($scenario + '_agro.png')), 'PNG', 1600, 900)
            $overflows = @()
            foreach ($shape in $slide.Shapes) {
                if ($shape.HasTextFrame -and $shape.TextFrame2.TextRange.Text.Trim() -and $shape.Top -ge 110 -and $shape.Top -lt 470) {
                    if ($shape.TextFrame2.TextRange.BoundHeight -gt ($shape.Height + 1)) {
                        $overflows += $shape.TextFrame2.TextRange.Text
                    }
                }
            }
            if ($scenario -notin @('full_size_diesel', 'diesel')) {
                $deck.Slides.Item(5).Export((Join-Path $previewRoot ($scenario + '_battery.png')), 'PNG', 1600, 900)
            }
            $allOverflows = @()
            foreach ($pageSlide in $deck.Slides) {
                foreach ($pageShape in $pageSlide.Shapes) {
                    if ($pageShape.HasTextFrame -and $pageShape.TextFrame2.TextRange.Text.Trim()) {
                        if ($pageShape.TextFrame2.TextRange.BoundHeight -gt ($pageShape.Height + 1)) {
                            $allOverflows += [pscustomobject]@{slide=$pageSlide.SlideIndex; text=$pageShape.TextFrame2.TextRange.Text; bounds_height=$pageShape.TextFrame2.TextRange.BoundHeight; frame_height=$pageShape.Height}
                        }
                    }
                }
            }
            $slides = $deck.Slides.Count
        } finally { $deck.Close() }
        $book = $excelApp.Workbooks.Open($manifest.report_file, 0, $true)
        try {
            $excelApp.CalculateFullRebuild()
            $main = $book.Worksheets.Item(1)
            $summary = @()
            $startColumn = $manifest.exported_report_channel_count + 2
            $lastColumn = $main.UsedRange.Columns.Count
            for ($col = $startColumn; $col -le $lastColumn; $col++) {
                $label = [string]$main.Cells.Item(3, $col).Value2
                if ($label -like '*Agro Capacity*' -or $label -like '*Agro Discharge Rate*' -or $label -like '*Battery Pack Mass*' -or $label -like '*Agro Mass Discharge Diagnostic*') {
                    if ([string]$main.Cells.Item(4, $col).Text -match '^#+$') { throw ('Clipped numeric reporting cell: ' + $label) }
                    $summary += [pscustomobject]@{label=$label; value=$main.Cells.Item(4, $col).Value2; formula=$main.Cells.Item(4, $col).Formula; display=$main.Cells.Item(4, $col).Text; number_format=$main.Cells.Item(4, $col).NumberFormat}
                }
            }
            $checks += [pscustomobject]@{scenario=$scenario; slides=$slides; charts=$main.ChartObjects().Count; new_summary=$summary; card_text_overflows=$overflows; all_slide_text_overflows=$allOverflows}
        } finally { $book.Close($false) }
        Write-Output ($scenario + ' inspected in PowerPoint and Excel')
    }
    $checks | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $auditRoot 'office_inspection.json') -Encoding UTF8
} finally {
    if ($null -ne $excelApp) { $excelApp.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excelApp) }
    if ($null -ne $pptApp) { $pptApp.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pptApp) }
}
