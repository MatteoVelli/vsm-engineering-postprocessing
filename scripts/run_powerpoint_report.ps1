$ErrorActionPreference = "Stop"

if (-not (Test-Path ".\.venv\Scripts\python.exe")) {
    throw "Virtual environment not found. Run .\scripts\setup_windows.ps1 first."
}

& ".\.venv\Scripts\python.exe" -m vsm_postprocessing.powerpoint_report_cli `
    ".\reference_files\Sprayer_Caiman_SP_9300Kg_Hybrid_Gen80kW_30kph_74Ht_4000KgAQ_57-4pcSOC_5-80_1C2G_02.xlsx" `
    --config ".\config\powerpoint_report_example.yaml" `
    --math-config ".\config\math_channels_example.yaml" `
    --statistics-config ".\config\statistics_excel_report.yaml" `
    --plotting-config ".\config\plotting_example.yaml" `
    --output-dir ".\outputs\powerpoint_report"
