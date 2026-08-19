$ErrorActionPreference = "Stop"

$SourceFile = ".\reference_files\RoboSprayer_3500Kg_Electric_12kph_Batt_50kW_Mot_63RPM_Susp_Cool_Rough_Grad_Discharge.csv"
$PlotConfig = ".\config\plotting_example.yaml"
$MathConfig = ".\config\math_channels_example.yaml"
$OutputDir = ".\outputs\plots"

vsm-plot $SourceFile `
    --config $PlotConfig `
    --math-config $MathConfig `
    --output-dir $OutputDir
