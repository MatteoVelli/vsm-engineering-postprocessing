$ErrorActionPreference = "Stop"

$sourceFile = Join-Path $PSScriptRoot "..\reference_files\Sprayer_Caiman_SP_9300Kg_Hybrid_Gen80kW_30kph_74Ht_4000KgAQ_57-4pcSOC_5-80_1C2G_02.xlsx"
$outputDirectory = Join-Path $PSScriptRoot "..\outputs\source_inspection"

if (-not (Test-Path $sourceFile)) {
    throw "Source workbook not found: $sourceFile"
}

python -m vsm_postprocessing $sourceFile --output-dir $outputDirectory
