# Final Deterministic Release Notes - v1.2.0

## Purpose

Version 1.2.0 packages the complete deterministic VSM post-processing workflow as a reproducible client release. No new engineering calculations were introduced in this milestone.

## Validated workflow

```text
VSM CSV/XLSX
    -> inspection
    -> channel selection
    -> math channels
    -> statistics
    -> plots
    -> Excel report
    -> optional PowerPoint report
```

The same workflow is accessible through the local Streamlit user interface.

## Release hardening

- runtime dependencies remain bounded in `pyproject.toml`; the installed environment can be snapshotted for audit with `scripts/snapshot_environment.ps1`;
- Python 3.11 is declared in `.python-version`;
- client launcher creates a Python 3.11 environment and installs only the bounded runtime dependency set;
- developer setup installs the runtime plus development/test dependency set;
- deterministic client ZIP builder added;
- client ZIP contains an internal `RELEASE_MANIFEST.json` with per-file SHA-256 hashes;
- an external SHA-256 checksum is emitted for the release ZIP;
- GitHub CI configuration added for regression testing without private client files;
- version consistency is regression-tested;
- client/reference datasets and generated run outputs remain excluded from the release package.

## Client data policy

The Sergio reference Excel/PowerPoint files are validation inputs and are not bundled into the distributable release unless explicit authorisation is provided.

## Acceptance boundary

The supplied reference data validates the complete deterministic workflow. A future full-channel VSM dataset should be used for a separate performance/scale acceptance test because the supplied source workbook has 70 channels rather than the target 500-700-channel production range.
