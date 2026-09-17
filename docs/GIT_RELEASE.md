# Git / GitHub Release Procedure - v1.2.5

Run these commands only after the local final acceptance checklist has passed.

```powershell
git status
git add .
git commit -m "Milestone 13A.1 Sergio fidelity specification and source-report fixes"
git tag -a v1.2.5 -m "VSM Engineering Post-Processing v1.2.5"
git push origin main
git push origin v1.2.5
```

Then create a GitHub Release from tag `v1.2.5` and attach:

- `VSM_Engineering_PostProcessing_v1.2.5_Client.zip`
- `VSM_Engineering_PostProcessing_v1.2.5_Client.sha256`

Do not attach Sergio's original Excel/PowerPoint reference files unless explicit permission has been given.

The GitHub Actions workflow runs the public regression suite. Acceptance tests that require private reference files are intentionally skipped in CI when those files are absent.


## Canonical build inputs and clean checkouts

The supported package builder is `src/vsm_postprocessing/release_builder.py`,
invoked by `scripts/build_release.ps1` or `python -m vsm_postprocessing.release_cli`.
Python bootstrap/download validation lives in `scripts/client_setup.ps1`; its client
instructions are in `docs/CLIENT_QUICK_START.md`. The historical one-off script
under `outputs/client_delivery/` is not a build or test dependency.

`outputs/` is disposable and may be absent before tests/builds. Packaging generates
its empty `outputs/.gitkeep` entry; doctor creates and verifies the writable runtime
output directory. Required branding/layout assets are listed in
`reference_files/README.md` and must be committed, including their `.gitignore`
exceptions. Original private input datasets and generated reports/ZIPs remain
excluded. Tests validate builds without existing outputs and continue to reject
missing required runtime assets.
