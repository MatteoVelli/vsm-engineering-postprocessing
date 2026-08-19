$ErrorActionPreference = "Stop"

$source = "reference_files\RoboSprayer_3500Kg_Electric_12kph_Batt_50kW_Mot_63RPM_Susp_Cool_Rough_Grad_Discharge.csv"

vsm-excel $source `
  --config "config\excel_report_example.yaml" `
  --math-config "config\math_channels_example.yaml" `
  --statistics-config "config\statistics_excel_report.yaml" `
  --plotting-config "config\plotting_example.yaml" `
  --output-dir "outputs\excel_report"
