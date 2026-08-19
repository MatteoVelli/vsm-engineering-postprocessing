$ErrorActionPreference = "Stop"

$ReportFile = ".\reference_files\Robo_Sprayer_Electrification_Tamplate_Electric_03.xlsx"
$PlotConfig = ".\config\plotting_reference_report.yaml"
$OutputDir = ".\outputs\plots_reference_report"

vsm-plot $ReportFile `
    --config $PlotConfig `
    --output-dir $OutputDir `
    --header-row 3 `
    --unit-row 4 `
    --data-start-row 5 `
    --data-end-row 17422 `
    --last-channel-column 70
