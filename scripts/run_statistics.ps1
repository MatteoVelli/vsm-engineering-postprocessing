$ErrorActionPreference = "Stop"

$source = ".\reference_files\RoboSprayer_3500Kg_Electric_12kph_Batt_50kW_Mot_63RPM_Susp_Cool_Rough_Grad_Discharge.csv"
$mathConfig = ".\config\math_channels_example.yaml"
$statisticsConfig = ".\config\statistics_example.yaml"
$output = ".\outputs\statistics"

vsm-stats $source `
    --config $statisticsConfig `
    --math-config $mathConfig `
    --output-dir $output
