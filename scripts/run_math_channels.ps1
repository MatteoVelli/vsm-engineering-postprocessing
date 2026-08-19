$ErrorActionPreference = "Stop"

$source = ".\reference_files\RoboSprayer_3500Kg_Electric_12kph_Batt_50kW_Mot_63RPM_Susp_Cool_Rough_Grad_Discharge.csv"
$config = ".\config\math_channels_example.yaml"
$output = ".\outputs\math_channels"

vsm-math $source --config $config --output-dir $output
