$ErrorActionPreference = "Stop"

$SourceFile = ".\reference_files\Sprayer_Caiman_SP_9300Kg_Hybrid_Gen80kW_30kph_74Ht_4000KgAQ_57-4pcSOC_5-80_1C2G_02.xlsx"
$PlotConfig = ".\config\plotting_example.yaml"
$MathConfig = ".\config\math_channels_example.yaml"
$OutputDir = ".\outputs\plots"

vsm-plot $SourceFile `
    --config $PlotConfig `
    --math-config $MathConfig `
    --output-dir $OutputDir
