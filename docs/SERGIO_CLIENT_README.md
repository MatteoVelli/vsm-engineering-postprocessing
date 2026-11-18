# ASTAUTO VSM Engineering Post-Processing - Sergio delivery

Version 1.4.1 - 18 November 2026 (18.11.26)
Package: VSM_Engineering_PostProcessing_v1.4.1_Sergio_Ready_18.11.26.zip

## Replace your previous installation

This ZIP is the complete application. Downloading it does not update an existing
installation automatically.

1. Close the old application browser tab and its launcher window.
2. Back up simulation files, saved reports, custom configurations and motor
   inputs. Keep the old Tool/config and personal files for reference.
3. Once your data are safe, rename the old installation folder as a backup, or
   remove it yourself. Setup does not remove personal files.
4. Download and extract the complete ZIP into C:\VSM_PostProcessing\.
   Open C:\VSM_PostProcessing\VSM_Sergio_v1.4.1\. The two BAT launchers must
   remain beside the Tool folder. Do not run files from inside the ZIP.
5. Use a fresh local folder outside OneDrive, SharePoint or other sync folders.
   Desktop and Documents may be synced by OneDrive. A short writable local path
   avoids sync locks and Windows long-path problems. If C:\ is restricted, ask IT
   for a short writable local folder. Paths containing spaces are supported.
6. Do not copy the old .venv, .python, setup_cache, caches or application code into
   the new folder. Review custom configuration changes separately; do not
   overwrite the supplied defaults wholesale.

## Set up and start

1. Double-click SETUP_VSM_TOOL.bat. Allow internet access for first setup.
2. Setup detects 64-bit Python 3.11.x. If missing, choose Y to let the existing
   bootstrap install official CPython 3.11.9 locally for this tool. It creates a
   new environment in Tool/.venv, installs the libraries and checks installation.
   Administrator rights are normally unnecessary.
3. When setup completes, double-click START_VSM_TOOL.bat. Keep its window open.
   Streamlit opens in your browser; if necessary, open http://localhost:8501.
   Use START for subsequent sessions; SETUP is for installation or repair.

This is an online source distribution, not an offline executable. It needs no
OneDrive, MATLAB or Microsoft Office to generate reports, and no files from your
previous installation.

## Generate reports

1. Choose Engineering Report and upload your VSM CSV or Excel simulation file.
2. Review the detected Diesel, Electric or Hybrid profile and machine.
3. For Electric/Hybrid motor characteristics, open Motor Performance Curves
   (Optional) and enable Include one shared motor capability curve.
4. Download the blank Excel template if needed. Put Motor RPM, Peak Torque and
   Cont Torque in row 1, with at least two complete numeric rows in rpm and Nm.
   Use the applicable manufacturer's actual peak and continuous limits.
5. Select Upload Motor Torque Map (.xlsx). Review the import confirmation,
   preview and curves. Choose a worksheet if offered. Correct any displayed
   errors in the workbook, save it and upload again. No copy-and-paste is needed.
6. Enter the correct final-drive ratio and driveline efficiency from the vehicle
   specifications. No universal ratio, efficiency or manufacturer map is supplied.
   Enter optional battery mass if known; leave unknown optional inputs blank.
7. Choose STEP 3 - Validate and resolve required-input problems.
8. Choose STEP 4 - Generate Engineering Report, then download Excel and PowerPoint
   from STEP 5 - Access Output. Copies remain in Tool/outputs/ui_runs.

Diesel charts explicitly titled "Vs Time" or "Time Based" use Time [s]. Other
ordinary Diesel traces use zero-based Distance [km]; torque maps retain RPM.
The Diesel torque-map slide has a 0.25-inch gap between its two chart cards.
Engineering values and approved Electric/Hybrid layouts remain unchanged.
Editing Excel does not update PowerPoint; regenerate both from the tool.

## If setup or startup fails

- Confirm the ZIP was fully extracted into a writable local folder outside OneDrive.
- If Windows offers Properties > Unblock for the ZIP, use it and extract again,
  following company policy. Do not disable security software.
- For download failures, check internet/proxy permissions with IT and rerun SETUP.
  IT can provide approved 64-bit Python 3.11 if automatic bootstrap is blocked.
- For a missing environment or failed health check, rerun SETUP from the new folder.
  Share Tool/logs/setup.log and Tool/logs/doctor.log with support.
- If reports cannot be overwritten, close them in Excel/PowerPoint and retry.
- Missing optional measurements may show N/A. Density is required for mass-to-volume
  conversion. The older Caiman reference's negative inferred contents remain flagged.

See Tool/docs/CLIENT_QUICK_START.md, CONFIGURATION_GUIDE.md, KNOWN_LIMITATIONS.md
and TROUBLESHOOTING.md for details. Motor import preparation and validation are
explained in Tool/docs/MOTOR_TORQUE_EXCEL_IMPORT.md. The manual table fallback is
available when no motor workbook is uploaded.
