param([string]$AuditPath = 'outputs/sergio_final_2026_10_17/audit.json')
$ErrorActionPreference = 'Stop'
$auditFile = (Resolve-Path -LiteralPath $AuditPath).Path
$destination = Join-Path (Split-Path -Parent $auditFile) 'visual_checks'
New-Item -ItemType Directory -Path $destination -Force | Out-Null
$powerpoint = New-Object -ComObject PowerPoint.Application
$checks = @()
try {
    foreach ($entry in (Get-Content -LiteralPath $auditFile -Raw | ConvertFrom-Json)) {
        $cases = @(@{ label = $entry.profile; presentation = $entry.presentation })
        foreach ($variant in $entry.variants) {
            $cases += @{ label = $entry.profile + '_' + $variant.label; presentation = $variant.presentation }
        }
        foreach ($case in $cases) {
            $deck = $powerpoint.Presentations.Open($case.presentation, $true, $false, $false)
            try {
                $deck.Slides.Item(1).Export((Join-Path $destination ($case.label + '_cover.png')), 'PNG', 1600, 900)
                $source = $deck.Slides.Item(1).Shapes.Item('VSM source filename')
                if ($source.TextFrame.TextRange.BoundHeight -gt $source.Height + .5 -or
                    $source.TextFrame.TextRange.BoundWidth -gt $source.Width + .5) {
                    throw ('Source filename overflow: ' + $case.label + '; height ' + $source.TextFrame.TextRange.BoundHeight + '/' + $source.Height + '; width ' + $source.TextFrame.TextRange.BoundWidth + '/' + $source.Width)
                }
                $checks += @{ case = $case.label; source_text = $source.TextFrame.TextRange.Text;
                    text_height = $source.TextFrame.TextRange.BoundHeight; box_height = $source.Height;
                    text_width = $source.TextFrame.TextRange.BoundWidth; box_width = $source.Width }
                if ($case.label -eq $entry.profile) {
                    for ($i = 0; $i -lt $entry.slide_ids.Count; $i++) {
                        if ($entry.slide_ids[$i] -in @('tyre_total','soil_bulldozing','soil_compaction','diesel_engine','diesel_controls')) {
                            $slide = $deck.Slides.Item($i + 1)
                            $slide.Export((Join-Path $destination ($case.label + '_' + $entry.slide_ids[$i] + '.png')), 'PNG', 1600, 900)
                            foreach ($shape in $slide.Shapes) {
                                if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                                    if ($shape.TextFrame.TextRange.BoundHeight -gt $shape.Height + 1 -or
                                        $shape.TextFrame.TextRange.BoundWidth -gt $shape.Width + 1) {
                                        throw ('Text overflow: ' + $case.label + ' ' + $entry.slide_ids[$i] + ' ' + $shape.TextFrame.TextRange.Text)
                                    }
                                }
                            }
                        }
                    }
                    $deck.SaveAs((Join-Path $destination ($case.label + '.pdf')), 32)
                }
            } finally { $deck.Close() }
            Write-Output ($case.label + ': native render and text bounds PASS')
        }
    }
    $checks | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $destination 'native_text_bounds.json') -Encoding UTF8
} finally {
    $powerpoint.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($powerpoint)
}
