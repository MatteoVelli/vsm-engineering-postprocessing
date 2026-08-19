$ErrorActionPreference = "Stop"

$report = ".\reference_files\Robo_Sprayer_Electrification_Tamplate_Electric_03.xlsx"
$config = ".\config\statistics_reference_report.yaml"
$output = ".\outputs\statistics_reference_validation"

vsm-stats $report `
    --config $config `
    --output-dir $output `
    --header-row 3 `
    --unit-row 4 `
    --data-start-row 5 `
    --data-end-row 17422 `
    --last-channel-column 70
