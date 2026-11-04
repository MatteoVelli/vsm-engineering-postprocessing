# Sergio final release - 4 November 2026

Application version: **1.4.1**. Branch: `main`.
Packaged source commit: `21cc18cec9708fab3ca74bb96a4d68477a98d69b`.

## Included functionality

VSM CSV/XLSX import and validation, channel selection and mathematical channels,
RMS/MAX/MIN/LAST statistics, deterministic engineering calculations and automatic
Excel/PowerPoint reporting remain available. Electric/Hybrid include supplied
motor curves, continuous-limit analysis and Motor Duty Summary; the native
four-curve Motor Torque and Power chart uses live motor-shaft power formulas.
PowerPoint includes editable blank brand/model fields, the added characteristics
slide, retained operating maps, CONSUMPTION labels and approved plot cards.
No synthetic motor ratings are substituted for missing input.

## Verification

- Complete validated-runtime suite: **1,122 passed, 1 skipped, 0 failures/errors**,
  out of 1,123 tests, in 1732.788 seconds. All six formerly skipped
  ZIP tests ran. The remaining skip is the existing Diesel soil-applicability case.
- Focused Excel/PowerPoint tests: **50 passed**. Two expected sparse synthetic-curve
  warnings are retained in the full suite.
- Real Electric 03: **30 native Excel charts**, all **29 original chart definitions
  byte-identical**, four **23-point** motor curves and **46 correct power formulas**.
  Maximum peak/continuous shaft power is **55.000000141031 / 44.510100769599 kW**.
  Original numeric cells/formula caches are preserved; fresh provenance identifies
  the new source/run. All 69 shared Excel/PPT metric occurrences match. Additional
  plot-only PPT metrics retain the original reference values.
- **17 PowerPoint slides**: characteristics on 15, operating maps on 16, summary on
  17. Eleven plot containers have the approved geometry/shadow. Native Office
  checks passed; all 17 rendered slides match the previously approved fixed deck
  pixel-for-pixel. Existing KPI values, slide text and plots are preserved.
- Synthetic Hybrid motor reporting passes. The available older real Caiman
  workflow also generated 27 charts and 15 slides through the installed UI with
  no motor curves supplied; its motor-characteristics chart/slide is explicitly
  omitted and its negative-load warning remains visible. **Latest real Hybrid
  motor-characteristics and full gearbox acceptance remains pending.**
- Both profiles are selectable. Independent extraction, missing-Python bootstrap,
  fresh environments, pip checks, doctor, actual root SETUP/START launchers and
  real Electric UI generation passed. The final archive was tested independently
  from the developer environment. These are checks on the available Windows host,
  not every customer PC. Cached official CPython payloads were reused; the final
  environment reused only the validated standalone runtime, not a prior venv.
- Runtime policy and dependencies are unchanged: official CPython 3.11.9,
  Streamlit 1.49.1, numpy 2.4.6, matplotlib 3.11.2, openpyxl 3.1.5, python-pptx
  1.0.2 and PyYAML 6.0.3. Thirty-seven configuration/template/launcher/runtime
  files are byte-identical to the previous delivery.

## Package and replacement

- File: `dist/Sergio_Request/VSM_Engineering_PostProcessing_v1.4.1_Sergio_Requested_Final.zip`
- Size: **94,488,983 bytes**; **124 files**.
- SHA-256: `1f41414952ba9695779cd576cadc3d8b9f08864d350e6fe2ce47b1314862760d`
- Internal folder: `VSM_Sergio_v1.4.1`, retaining the root BAT launchers beside
  `Tool`. A long initial candidate folder caused a Windows path-limit failure;
  the shortened folder fixed it and passed repeated clean-install generation.
  Use a short extraction location. The requested outer ZIP filename is retained.
- ZIP integrity and manifest hashes pass. Two builds are byte-identical. The
  final application payload matches the independently validated candidate;
  only manifest commit/build provenance differs.
- Previous ZIP: `Astauto_VSM_v1.4.1_02112026.zip`, **94,452,529 bytes**,
  SHA-256 `402d77625aa715ca22de196bf1d8b6dd76523c7f3cff21a393ff63c2ec4e2335`.
- Archived copy: `dist/Archived_Client_Packages/2026-11-04_before_sergio_requested_final/Astauto_VSM_v1.4.1_02112026.zip`.
  Its checksum was verified before the redundant active copy was removed. The
  previous extracted delivery folder and checksum were archived too.
- Latest-package replacement succeeded. `dist/CURRENT_SERGIO_DELIVERY.json`
  and `dist/README_CURRENT_SERGIO.txt` identify the sole active delivery.

## Remaining diagnostics and reproduction

Density is still required for mass-to-volume conversion. Existing source-unit,
capacity, load-quality and other legitimate engineering diagnostics remain.
Doctor reports that the private example CSV is not bundled; UI uploads work.
No warnings were suppressed to obtain a passing status. First setup may need
internet access; this is not an offline executable distribution.

Rebuild at the packaged source commit using the existing validated environment:

```powershell
python scripts/build_sergio_release.py --output-dir outputs/new_release_build --release-date 2026-11-04
```

The builder refuses to overwrite an existing release. Tests and private validation
evidence are retained locally under `outputs/release_2026_11_04/`, including JUnit,
installed-browser, native-Office, numerical, preservation and pixel comparisons.
Generated reports, environments, caches and ZIPs are excluded from Git. The
accompanying `RELEASE_RECORD.json` records final synchronization and the separate
release-documentation commit without creating a self-referential archive hash.
