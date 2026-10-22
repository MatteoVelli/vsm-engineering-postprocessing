param([string]$AuditPath = 'outputs\october_08_04_validation\audit.json')
$ErrorActionPreference = 'Stop'
$auditFile = (Resolve-Path -LiteralPath $AuditPath).Path
$destination = Join-Path (Split-Path -Parent $auditFile) 'visual_checks'
New-Item -ItemType Directory -Path $destination -Force | Out-Null
$powerpoint = $null
$excel = $null
$checks = @()
try {
    $powerpoint = New-Object -ComObject PowerPoint.Application
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3
    foreach ($entry in (Get-Content -LiteralPath $auditFile -Raw | ConvertFrom-Json)) {
        $deck = $powerpoint.Presentations.Open($entry.presentation, $true, $false, $false)
        $overflow = @()
        try {
            $pages = if ($entry.profile -eq 'diesel') { @(3) } else { @(5,8) }
            foreach ($page in $pages) {
                $slide = $deck.Slides.Item($page)
                $slide.Export((Join-Path $destination ($entry.profile + '_slide_' + $page + '.png')), 'PNG', 1600, 900)
                foreach ($shape in $slide.Shapes) {
                    if ($shape.HasTextFrame -and $shape.TextFrame2.TextRange.Text.Trim() -and $shape.Top -ge 110 -and $shape.Top -lt 220) {
                        if ($shape.TextFrame2.TextRange.BoundHeight -gt ($shape.Height + 1) -or $shape.TextFrame2.TextRange.BoundWidth -gt ($shape.Width + 1)) {
                            $overflow += [pscustomobject]@{slide=$page; text=$shape.TextFrame2.TextRange.Text; bound_height=$shape.TextFrame2.TextRange.BoundHeight; frame_height=$shape.Height; bound_width=$shape.TextFrame2.TextRange.BoundWidth; frame_width=$shape.Width}
                        }
                    }
                }
            }
            $deck.SaveAs((Join-Path $destination ($entry.profile + '.pdf')), 32)
        } finally { $deck.Close() }
        $book = $excel.Workbooks.Open($entry.excel, 0, $true)
        try {
            $excel.CalculateFullRebuild()
            $statistics = $book.Worksheets.Item('Statistics')
            $cell = $statistics.Columns.Item(1).Find('agro_discharge_rate_kg_per_km')
            if ($null -eq $cell) { throw 'Missing discharge KPI' }
            $value = $statistics.Cells.Item($cell.Row, 3).Value2
            $formula = $statistics.Cells.Item($cell.Row, 3).Formula
            $checks += [pscustomobject]@{profile=$entry.profile; card_text_overflows=$overflow; recalculated_rate=$value; formula=$formula}
        } finally { $book.Close($false) }
        Write-Output ($entry.profile + ' rendered; card overflows: ' + $overflow.Count)
    }
    $checks | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $destination 'office_checks.json') -Encoding UTF8
} finally {
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    if ($powerpoint) { $powerpoint.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($powerpoint) }
}
