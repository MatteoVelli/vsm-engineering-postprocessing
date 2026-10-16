param([Parameter(Mandatory = $true)][string]$AuditPath)
$ErrorActionPreference = 'Stop'
$auditFile = (Resolve-Path -LiteralPath $AuditPath).Path
$destination = Join-Path (Split-Path -Parent $auditFile) 'visual_checks'
New-Item -ItemType Directory -Path $destination -Force | Out-Null
$powerpoint = $null
try {
    $powerpoint = New-Object -ComObject PowerPoint.Application
    foreach ($entry in (Get-Content -LiteralPath $auditFile -Raw | ConvertFrom-Json)) {
        $deck = $powerpoint.Presentations.Open($entry.presentation, $true, $false, $false)
        try {
            $deck.SaveAs((Join-Path $destination ($entry.profile + '.pdf')), 32)
            $deck.Slides.Item(9).Export((Join-Path $destination ($entry.profile + '_slide_9.png')), 'PNG', 1600, 900)
            Write-Output ($entry.profile + ': PDF and slide 9 rendered by desktop PowerPoint')
        } finally { $deck.Close() }
    }
} finally {
    if ($powerpoint) {
        $powerpoint.Quit()
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($powerpoint)
    }
}
