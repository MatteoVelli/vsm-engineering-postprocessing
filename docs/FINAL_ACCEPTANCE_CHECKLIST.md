# Final Acceptance Checklist - v1.2.1

## Engineering

- [ ] `pytest` passes on the maintainer machine.
- [ ] `scripts/run_doctor.ps1` reports zero blocking FAIL items.
- [ ] End-to-end pipeline completes on the supplied reference VSM workbook.
- [ ] Key Excel RMS/MAX/MIN/LAST/SUM values are spot-checked against validated reference values.
- [ ] Excel report opens without repair warnings.
- [ ] Optional PowerPoint opens without repair warnings and contains the intended plots/KPIs.
- [ ] Source-only PowerPoint identifies the 1,866-row input as a `Source 74 ha Field Cycle`, not the complete Sergio duty cycle.
- [ ] Excel top RMS pair is Battery Power RMS + Battery Heatflow RMS.

## Client workflow

- [ ] `START_VSM_TOOL.bat` starts from a clean extracted release folder.
- [ ] UI loads a `.xlsx` source file.
- [ ] UI completes every enabled stage with PASS.
- [ ] Excel download/open actions work.
- [ ] PowerPoint download/open actions work when enabled.

## Traceability

- [ ] `pipeline_manifest.json` is present.
- [ ] `pipeline_summary.txt` is present.
- [ ] `pipeline.log` is present.
- [ ] source/config hashes are recorded.
- [ ] software/Python/platform metadata are recorded.

## Release package

- [ ] version is `1.2.1` in both `pyproject.toml` and `vsm_postprocessing.__version__`.
- [ ] an environment package snapshot is retained for the validated release machine.
- [ ] client ZIP contains no reference/client workbooks.
- [ ] client ZIP contains no generated report/run outputs.
- [ ] `RELEASE_MANIFEST.json` is present in the client ZIP.
- [ ] external `.sha256` checksum matches the client ZIP.
- [ ] Git working tree is clean before the release tag is created.
