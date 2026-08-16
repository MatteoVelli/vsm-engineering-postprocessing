from __future__ import annotations

from pathlib import Path

import pytest

from conftest import require_private_reference_file


def test_private_reference_helper_returns_existing_path(tmp_path: Path) -> None:
    reference = tmp_path / "reference.csv"
    reference.write_text("ok", encoding="utf-8")

    assert require_private_reference_file(reference, "synthetic private reference") == reference


def test_private_reference_helper_skips_missing_path(tmp_path: Path) -> None:
    missing = tmp_path / "missing.csv"

    with pytest.raises(pytest.skip.Exception, match="private client reference dataset not available"):
        require_private_reference_file(missing, "synthetic private reference")
