# Final Acceptance Checklist - v1.3.0

## Engineering

- [ ] `pytest` passes on the maintainer machine.
- [ ] `scripts/run_doctor.ps1` reports zero blocking FAIL items.
- [ ] Profile Excel reports complete for the latest Electric and Hybrid source CSVs.
- [ ] Profile PowerPoint reports open without repair warnings.
- [ ] Every generated PowerPoint slide includes the Astauto logo.
- [ ] Electric_03 and Hybrid_04 profile channel counts match the latest templates when optional source channels are present.
- [ ] Optional `Track_Height` / `Road Height` is exported when present and omitted without failure when absent.
- [ ] Agrochemical Discharge remains mapped from `HitchRear_Force_Z_VehicleCoordinates` using `-force / 9.81`.
- [ ] Hybrid PowerPoint KPI labels and values describe the same statistics.

## Client Workflow

- [ ] `START_VSM_TOOL.bat` starts from a clean extracted release folder.
- [ ] UI loads a `.csv` or `.xlsx` source file.
- [ ] UI completes the selected profile report with PASS.
- [ ] Excel download/open actions work.
- [ ] PowerPoint download/open actions work when enabled.

## Traceability

- [ ] `pipeline_manifest.json` is present for generic pipeline runs.
- [ ] `pipeline_summary.txt` is present for generic pipeline runs.
- [ ] `pipeline.log` is present for generic pipeline runs.
- [ ] source/config hashes are recorded.
- [ ] software/Python/platform metadata are recorded.

## Release Package

- [ ] version is `1.3.0` in both `pyproject.toml` and `vsm_postprocessing.__version__`.
- [ ] an environment package snapshot is retained for the validated release machine.
- [ ] client ZIP contains no reference/client workbooks.
- [ ] client ZIP contains no generated report/run outputs.
- [ ] client ZIP contains no retired scenario/provider assets.
- [ ] `RELEASE_MANIFEST.json` is present in the client ZIP.
- [ ] external `.sha256` checksum matches the client ZIP.
- [ ] Git working tree is reviewed before the release tag is created.
