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
