$ErrorActionPreference = "Stop"

$sourceFile = Join-Path $PSScriptRoot "..\reference_files\RoboSprayer_3500Kg_Electric_12kph_Batt_50kW_Mot_63RPM_Susp_Cool_Rough_Grad_Discharge.csv"
$outputDirectory = Join-Path $PSScriptRoot "..\outputs\source_inspection"

if (-not (Test-Path $sourceFile)) {
    throw "Source workbook not found: $sourceFile"
}

python -m vsm_postprocessing $sourceFile --output-dir $outputDirectory
