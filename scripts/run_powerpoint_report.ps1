$ErrorActionPreference = "Stop"

if (-not (Test-Path ".\.venv\Scripts\python.exe")) {
    throw "Virtual environment not found. Run .\scripts\setup_windows.ps1 first."
}

& ".\.venv\Scripts\python.exe" -m vsm_postprocessing.powerpoint_report_cli `
    ".\reference_files\RoboSprayer_3500Kg_Electric_12kph_Batt_50kW_Mot_63RPM_Susp_Cool_Rough_Grad_Discharge.csv" `
    --config ".\config\powerpoint_report_example.yaml" `
    --math-config ".\config\math_channels_example.yaml" `
    --statistics-config ".\config\statistics_excel_report.yaml" `
    --plotting-config ".\config\plotting_example.yaml" `
    --output-dir ".\outputs\powerpoint_report"
