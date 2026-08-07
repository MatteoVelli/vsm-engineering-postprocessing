$ErrorActionPreference = "Stop"

$sourceFile = Join-Path $PSScriptRoot "..\reference_files\Sprayer_Caiman_SP_9300Kg_Hybrid_Gen80kW_30kph_74Ht_4000KgAQ_57-4pcSOC_5-80_1C2G_02.xlsx"
$configFile = Join-Path $PSScriptRoot "..\config\channel_selection_example.yaml"
$outputDirectory = Join-Path $PSScriptRoot "..\outputs\channel_selection"

if (-not (Test-Path $sourceFile)) {
    throw "Source workbook not found: $sourceFile"
}
if (-not (Test-Path $configFile)) {
    throw "Selection configuration not found: $configFile"
}

python -m vsm_postprocessing.channel_cli `
    $sourceFile `
    --config $configFile `
    --output-dir $outputDirectory
