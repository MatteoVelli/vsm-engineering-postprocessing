$ErrorActionPreference = "Stop"

$source = "reference_files\Sprayer_Caiman_SP_9300Kg_Hybrid_Gen80kW_30kph_74Ht_4000KgAQ_57-4pcSOC_5-80_1C2G_02.xlsx"

vsm-excel $source `
  --config "config\excel_report_example.yaml" `
  --math-config "config\math_channels_example.yaml" `
  --statistics-config "config\statistics_excel_report.yaml" `
  --plotting-config "config\plotting_example.yaml" `
  --output-dir "outputs\excel_report"
