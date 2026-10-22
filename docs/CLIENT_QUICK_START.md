# Client Quick Start

## First-Time Setup

1. Extract the complete ZIP into a writable Windows folder; do not run it inside the ZIP.
2. Open the extracted folder.
3. Double-click `SETUP_VSM_TOOL.bat`.
4. Follow the on-screen instructions.
5. When setup reports SUCCESS, double-click `START_VSM_TOOL.bat`.

Setup automatically checks Python, offers to install the supported official CPython 3.11.9 runtime if a validated Python 3.11.x runtime is missing, creates the package-local `.venv`, installs Streamlit and all required libraries, and validates the installation. No manual `pip`, `python -m venv`, Streamlit, PowerShell, YAML editing or PATH configuration is required for normal use.

The first setup requires internet access if Python or Python packages are not already available locally.

## Normal Use

1. Double-click `START_VSM_TOOL.bat`.
2. Select **Engineering Report**.
3. Upload the source VSM `.xlsx` or `.csv` results file.
4. Review the automatically detected **Diesel**, **Hybrid** or **Electric** profile and machine information; change the selection only if needed.
5. Click **STEP 3 - Validate**. Resolve any missing required channels or invalid data before continuing; the UI lists the affected channels and blocks report generation.
6. For Electric or Hybrid, enter **Battery pack mass (kg)** if known. Leave it blank to show N/A. Enter a finite positive value; no mass is inferred automatically.
7. Click **STEP 4 - Generate Engineering Report**.
8. Download the Excel report and, when needed, the PowerPoint report from **STEP 5 - Access Output**. The profile workflow generates both reports; using the PowerPoint is optional.

The launcher opens the local Streamlit interface in your browser. If the browser does not open automatically, use `http://localhost:8501`. Keep the launcher running while using the application; close it when finished.

## Workflows

**Engineering Report** generates the client Excel report and optional PowerPoint directly from the uploaded source data and selected reporting profile.

**Custom Analysis** analyzes the uploaded file exactly as supplied and keeps the manual channel/math/statistics/plot/report controls for engineering investigation.

## Where Runs Are Stored

UI runs are isolated under:

```text
outputs/ui_runs/<timestamp>/
```

Profile reports are saved inside the run's `profile_excel_report/` and `profile_powerpoint_report/` folders. You can also save copies anywhere using the browser download buttons. Custom Analysis pipeline runs include their manifests, summaries and logs.

## Equations in Excel

Select a mathematical data cell to inspect its live Excel formula when one is available. Every mathematical column also has an equation note on its main-sheet heading; hover over the heading or show its note in Excel. The mapping worksheet retains the full equation inventory.

The current October profiles use live formulas for supported arithmetic,
observed-interval signed trapezoidal energy, cumulative channels and statistics.
Every applicable configured KPI Value cell contains a genuine formula with
Python-validated cached results. Select a Value cell on the Statistics sheet to
inspect its equation. The hidden Formula Support sheet documents engineering
parameters and any additional formula/raw dependencies; unhide it to audit them.
Missing optional measurements disable dependent KPIs rather than appearing as zero.

Legacy profiles keep their original engineering conventions. Some visible legacy
MATH columns retain Python values; KPI dependencies use live helper calculations
where necessary. Regenerate the report after source changes to refresh all retained
values, charts and PowerPoint consistently. PowerPoint uses deterministic Python
results; editing Excel does not update an already generated presentation.

## October Reporting Update

Diesel reports include agrochemical mass, auxiliary power, wheel torque/power,
20 native Excel charts when their sources are available, a 24 pt first row and
five configured KPI formulas. The existing thermal slide shows separate oil and
water plots with four MIN/MAX cards, preserving ten slides for complete data.

Electric/Hybrid profiles use measured gearbox-output torque, signed wheel power,
bulldozing/compaction power and energy, and optional road height. The exact HV
auxiliary source `User_defined_Data_High_Voltage_Auxiliary_signal1_0` is interpreted
as kW only under its documented channel-specific unit declaration. Other unit
validation remains strict. Complete Electric inputs produce 25 Excel charts,
12 KPI formulas and 15 PowerPoint slides, including auxiliary energy on slide 7.

Current Electric and Hybrid reports automatically add four separate time-based
plots: Total Bulldozing Resistance Power, Tyre Bulldozing Resistance Energy
Accumulated, Total Compaction Resistance Power and Tyre Compaction Resistance
Energy Accumulated. Power is in kW, accumulated energy in kWh and time in seconds.
PowerPoint pages 12 and 13 show the two mechanisms before the final summary.
Each page now includes six Wheel Steering-style KPI cards: MAX POWER, MEAN
POWER, RMS POWER, TOTAL ENERGY, ENERGY PER KM and DISTANCE. Mean is time-weighted;
RMS uses the existing sample RMS. Energy is the accumulated channel's final minus
initial value, divided by the existing distance KPI for energy per kilometre.
Unavailable measurements show N/A; zero distance never produces a ratio.
The existing combined soil charts and all other plots remain available. Missing
or invalid wheel resistance measurements disable only their dependent plots;
PowerPoint identifies unavailable plot data. Explicit legacy profiles are unchanged.

The final 22 October delivery folder and ZIP are named
`Astauto_VSM_October_22102026`. Keep the two root launchers beside `Tool`.

Release checks use Diesel `_07`, Electric UoL `_04` and the available Caiman SP Hybrid reference. Caiman's negative inferred chemical contents are correctly flagged; a newer Hybrid export beyond this reference is not claimed as tested.

## Battery and Agrochemical Reporting

Electric and Hybrid have an optional **Battery pack mass (kg)** field. Enter the
pack mass for the current report if known. The seventh card on Battery & Electrical
Energy System displays that value in kg; blank input displays N/A. The value is
recorded as user-provided. Diesel does not request a battery pack mass.

**Agro Discharge Rate** is calculated automatically in kg/km from initial minus
final remaining chemical mass, divided by final cumulative distance. It uses the
current uploaded data and needs no chemical density. Invalid or missing source
measurements show N/A with an explanation. A nonzero starting distance or a
possible refill is reported in the diagnostics; the rate remains the specified
net endpoint calculation. The current templates are Electric 08, Hybrid 08 and
Diesel 04. The older volumetric diagnostics remain separate in Excel/metadata.

Agro Capacity is the maximum observed inferred chemical load in kg, not a
manufacturer's nominal tank capacity. Caiman's available reference has negative
inferred contents, so its mass-discharge rate correctly shows N/A. Existing battery
capacity consistency warnings and other engineering diagnostics remain valid.
Regenerate reports after changing source data or report inputs. Generated Excel
and PowerPoint filenames use the uploaded file's stem; machine-name overrides
change report content without renaming the source-based outputs.

## Supported Runtime

Application version remains 1.3.0. Use Python >=3.11,<3.12; automatic bootstrap
targets official CPython 3.11.9. The tested Streamlit version is 1.49.1 (declared
range >=1.49.1,<1.50). Setup installs NumPy >=1.26,<3, openpyxl >=3.1,<4,
PyYAML >=6,<7, matplotlib >=3.9,<4 and python-pptx >=1,<2.
MATLAB and Microsoft Office are not required for report generation.

## Setup Problems

If Windows blocks the downloaded archive, right-click the ZIP, select **Properties**, use **Unblock** if offered, and extract again. Follow your organisation's application policy; do not disable security software. Run setup from a writable folder and keep the package files together.

First-time setup needs access to the official Python download and Python package servers. If dependency installation fails, inspect `logs/setup.log`, check internet/proxy permissions, and rerun setup. If downloads are blocked by company policy, ask for an approved Python 3.11 installation or an offline dependency installation. See `docs/TROUBLESHOOTING.md` for details.

## Health Check

From PowerShell:

```powershell
.\.venv\Scripts\python.exe -m vsm_postprocessing.doctor_cli --project-root .
```

`PASS` means there are no blocking installation/configuration problems. `WARN` items are informational/non-blocking. `FAIL` must be corrected before production use.
