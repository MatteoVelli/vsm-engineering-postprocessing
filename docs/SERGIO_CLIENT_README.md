# VSM Engineering Post-Processing — Sergio final delivery

Version 1.4.1 · 4 November 2026

1. Extract the complete ZIP to a short writable Windows location, such as your Desktop. Open `VSM_Sergio_v1.4.1` and keep the two BAT files beside `Tool`. Avoid deeply nested extraction paths because long VSM filenames can exceed Windows path limits.
2. Run `SETUP_VSM_TOOL.bat` once and follow the prompts. First setup needs internet access when Python or packages are unavailable. The existing bootstrap installs official Python 3.11.9 when necessary and creates a local environment.
3. Run `START_VSM_TOOL.bat`. Keep its window open while using the local browser application.
4. Choose Engineering Report, upload a VSM CSV or Excel file, and review the detected Electric or Hybrid profile. Enter the battery mass if known.
5. For motor characteristics, enter the actual peak and continuous torque-versus-RPM curves in Motor Performance Curves, together with ratio and efficiency for the operating-map overlays. No manufacturer ratings are supplied automatically.
6. Validate, generate, then download Excel and PowerPoint. Results remain in `Tool/outputs/ui_runs`.

The real Electric 03 acceptance case produces 30 native Excel charts and 17
PowerPoint slides. The four motor curves use live shaft-power calculations;
ratio and efficiency do not change motor-shaft power. PowerPoint retains the
operating-map analysis and uses the corrected chart-container styling. Fill in
the editable MOTOR BRAND and MOTOR TYPE / MODEL fields as appropriate.

Missing motor inputs produce a visible omission notice. Density remains required
for mass-to-volume conversion; legitimate data-quality warnings are preserved.
Editing Excel does not update an existing PowerPoint: regenerate the reports.
Latest motor-reporting acceptance with real Hybrid data remains pending;
synthetic Hybrid tests do not establish real-data acceptance. Installation was
checked in an independent environment on the available Windows host, not every
customer PC. This is an online source distribution, not an offline executable.

See `Tool/docs/CLIENT_QUICK_START.md`, `CONFIGURATION_GUIDE.md`,
`KNOWN_LIMITATIONS.md` and `TROUBLESHOOTING.md` for details.
