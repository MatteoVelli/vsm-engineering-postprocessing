# Git / GitHub Release Procedure - v1.4.1

The 4 November 2026 requested final release uses application version 1.4.1 and
`VSM_Engineering_PostProcessing_v1.4.1_Sergio_Requested_Final.zip`.
The reproducible portable builder is `scripts/build_sergio_release.py`; it uses
the existing release collector and deterministic ZIP writer. Run it in the
validated environment after committing the tested source:

```powershell
python scripts/build_sergio_release.py --output-dir outputs/release_staging --release-date 2026-11-04
```

Use a new output directory; the builder refuses to overwrite a release. Verify
the ZIP and its manifest, test a fresh extraction, archive the previous delivery,
then promote the validated ZIP and update `dist/CURRENT_SERGIO_DELIVERY.json`.
Final artifact checksums belong in the accompanying release record, avoiding a
self-referential package checksum. Real Hybrid motor acceptance stays pending
unless matching real inputs and motor curves have actually been tested.

The explicitly authorised workflow is:

1. Validate the full regression suite, the requested report workflows and reference
   hashes, clean package installation/bootstrap and shipped setup/start launchers.
   Record real-data limitations explicitly rather than treating synthetic fixtures
   as client engineering acceptance.
2. Review the complete source diff and stage only release source, configuration,
   tests and documentation. Exclude unrelated user files, outputs and ZIPs.
3. Commit the validated source and record the commit/tree in the portable manifest.
4. Reuse the established deterministic collector/writer and portable Tool layout.
   Validate the staged archive and a fresh extraction before replacing the ZIP
   identified by `dist/CURRENT_SERGIO_DELIVERY.json`. Preserve rollback in
   `dist/Archived_Client_Packages` and keep one active distribution.
5. Verify main and its expected upstream/remote, then push normally to origin/main.
   Stop the push if the branch/upstream is unexpected or the remote has diverged.
   No force push, tag or GitHub Release is required by this delivery request.

The previous 28 October prohibition on pushing applied to that historical task.
The current release request explicitly authorises the commit and normal push.

## Historical publication example

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
