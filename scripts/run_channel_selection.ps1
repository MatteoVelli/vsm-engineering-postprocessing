$ErrorActionPreference = "Stop"

$sourceFile = Join-Path $PSScriptRoot "..\reference_files\RoboSprayer_3500Kg_Electric_12kph_Batt_50kW_Mot_63RPM_Susp_Cool_Rough_Grad_Discharge.csv"
$configFile = Join-Path $PSScriptRoot "..\config\channel_selection_example.yaml"
$outputDirectory = Join-Path $PSScriptRoot "..\outputs\channel_selection"

if (-not (Test-Path $sourceFile)) {
    throw "Source file not found: $sourceFile"
}
if (-not (Test-Path $configFile)) {
    throw "Selection configuration not found: $configFile"
}

python -m vsm_postprocessing.channel_cli `
    $sourceFile `
    --config $configFile `
    --output-dir $outputDirectory
