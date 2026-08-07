$ErrorActionPreference = "Stop"

$source = ".\reference_files\Sprayer_Caiman_SP_9300Kg_Hybrid_Gen80kW_30kph_74Ht_4000KgAQ_57-4pcSOC_5-80_1C2G_02.xlsx"
$config = ".\config\math_channels_example.yaml"
$output = ".\outputs\math_channels"

vsm-math $source --config $config --output-dir $output
