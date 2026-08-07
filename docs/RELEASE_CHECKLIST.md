# Release / Client Acceptance Checklist

Before handing a release to another user:

- [ ] `scripts/run_doctor.ps1` reports no FAIL items.
- [ ] Regression suite passes (`pytest`).
- [ ] End-to-end pipeline completes on a known reference dataset.
- [ ] Excel report opens correctly and key engineering values are spot-checked.
- [ ] Optional PowerPoint opens correctly and contains the intended plots/KPIs.
- [ ] `pipeline_manifest.json`, `pipeline_summary.txt`, and `pipeline.log` are present.
- [ ] Source and configuration hashes are present in the manifest.
- [ ] Original client/reference files are not bundled unless explicitly authorised.
- [ ] The release version in `pyproject.toml` and `vsm_postprocessing.__version__` matches.
- [ ] `START_VSM_TOOL.bat` launches the UI on the target Windows machine.

## Final packaging

- [ ] `scripts/build_release.ps1` completes successfully.
- [ ] Client ZIP contains `RELEASE_MANIFEST.json`.
- [ ] `.sha256` file matches the generated client ZIP.
- [ ] No `.xlsx`, `.pptx`, client/reference files or generated run outputs are present in the client ZIP.
- [ ] `scripts/snapshot_environment.ps1` is retained with the validated release record.
