from __future__ import annotations

import shutil
import zipfile
from dataclasses import replace
from pathlib import Path

import numpy as np
import pytest
import yaml

from vsm_postprocessing.duty_cycle import (
    WorkbookRowProfileProvider,
    build_composition_plan,
    compose_duty_cycle,
    export_duty_cycle_composition,
    load_duty_cycle_config,
    load_profile_provider_config,
)
from vsm_postprocessing.errors import ConfigurationError, DataValidationError, VSMPostProcessingError
from vsm_postprocessing.importer import load_data_file
from vsm_postprocessing.utils import sha256_file


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SCENARIO_CONFIG = PROJECT_ROOT / "config" / "duty_cycle_sergio_reference.yaml"
PROFILE_CONFIG = PROJECT_ROOT / "config" / "duty_cycle_profiles_sergio_reference.yaml"
SOURCE_WORKBOOK = PROJECT_ROOT / "reference_files" / (
    "Sprayer_Caiman_SP_9300Kg_Hybrid_Gen80kW_30kph_74Ht_4000KgAQ_57-4pcSOC_5-80_1C2G_02.xlsx"
)
REFERENCE_WORKBOOK = PROJECT_ROOT / "reference_files" / "Sprayer_Caiman_SP_9300Kg_Electrification_03.xlsx"


def _copy_with_different_package_fingerprint(source: Path, destination: Path) -> Path:
    shutil.copyfile(source, destination)
    with zipfile.ZipFile(destination, mode="a") as workbook:
        workbook.comment = b"client-compatible-package-fingerprint"
    assert sha256_file(destination) != sha256_file(source)
    return destination


@pytest.fixture(scope="module")
def reference_assets():
    if not (SOURCE_WORKBOOK.exists() and REFERENCE_WORKBOOK.exists()):
        pytest.skip("Client source/reference workbooks are not present")
    scenario = load_duty_cycle_config(SCENARIO_CONFIG)
    source = load_data_file(SOURCE_WORKBOOK)
    provider = WorkbookRowProfileProvider(load_profile_provider_config(PROFILE_CONFIG), REFERENCE_WORKBOOK)
    return scenario, source, provider


def test_sergio_profile_provider_config_is_explicit_and_portable() -> None:
    config = load_profile_provider_config(PROFILE_CONFIG)

    assert config.provider_id == "sergio_reference_missing_phase_profiles"
    assert config.provider_type == "workbook_phase_rows"
    assert config.value_policy == "absolute_reference"
    assert config.channel_alignment == "column_position"
    assert config.phase_ids == ("P05", "P06", "P08", "P10")
    assert config.phase_rows == {
        "P05": (5364, 7081),
        "P06": (7082, 9617),
        "P08": (10518, 12080),
        "P10": (12981, 14656),
    }
    assert config.expected_filename == "Sprayer_Caiman_SP_9300Kg_Electrification_03.xlsx"
    assert config.expected_sha256 == "1ac1fd0cc4b73a584410e78edbbd65d966077e15499da89ec004070c8130cd80"
    assert config.data_start_row == 5
    assert config.data_end_row == 17422
    assert config.last_channel_column == 70


def test_profile_provider_config_rejects_unknown_keys(tmp_path: Path) -> None:
    raw = yaml.safe_load(PROFILE_CONFIG.read_text(encoding="utf-8"))
    raw["provider"]["mystery_setting"] = True
    path = tmp_path / "bad_provider.yaml"
    path.write_text(yaml.safe_dump(raw, sort_keys=False), encoding="utf-8")

    with pytest.raises(ConfigurationError, match="Unknown profile-provider configuration key"):
        load_profile_provider_config(path)


def test_column_position_alignment_requires_file_hash(tmp_path: Path) -> None:
    raw = yaml.safe_load(PROFILE_CONFIG.read_text(encoding="utf-8"))
    raw["provider"].pop("expected_sha256")
    path = tmp_path / "provider_without_hash.yaml"
    path.write_text(yaml.safe_dump(raw, sort_keys=False), encoding="utf-8")
    config = load_profile_provider_config(path)

    with pytest.raises(ConfigurationError, match="requires expected_sha256"):
        WorkbookRowProfileProvider(config, tmp_path / "missing.xlsx")


def test_sergio_reference_provider_validates_provenance_and_positional_channel_layout(reference_assets) -> None:
    scenario, source, provider = reference_assets
    config = load_profile_provider_config(PROFILE_CONFIG)

    validation = provider.validate(scenario, source)

    assert validation.provider_id == config.provider_id
    assert validation.source_file == REFERENCE_WORKBOOK.name
    assert validation.source_sha256 == config.expected_sha256
    assert validation.supported_phase_ids == ("P05", "P06", "P08", "P10")
    assert validation.sample_count == 17418
    assert validation.channel_count == 70
    assert validation.channel_alignment == "column_position"
    # Sergio renamed many report headers; exact channel IDs therefore differ,
    # while the fixed 70-column layout is locked by the reference SHA-256.
    assert not validation.channel_ids_match
    assert validation.channel_layout_compatible
    assert validation.nominal_time_step_s == 1.0


def test_profile_provider_compatible_mode_accepts_different_file_fingerprint(
    tmp_path: Path,
    reference_assets,
) -> None:
    scenario, source, _provider = reference_assets
    config = load_profile_provider_config(PROFILE_CONFIG)
    workbook = _copy_with_different_package_fingerprint(
        REFERENCE_WORKBOOK,
        tmp_path / REFERENCE_WORKBOOK.name,
    )

    provider = WorkbookRowProfileProvider(config, workbook, validation_mode="compatible")
    validation = provider.validate(scenario, source)

    assert validation.source_sha256 == sha256_file(workbook)
    assert validation.expected_sha256 == config.expected_sha256
    assert validation.reference_sha256_matches is False
    assert validation.validation_mode == "compatible"
    assert validation.channel_layout_compatible
    assert validation.supported_phase_ids == ("P05", "P06", "P08", "P10")


def test_profile_provider_accepts_prefixed_persisted_path_when_original_filename_matches(
    tmp_path: Path,
    reference_assets,
) -> None:
    scenario, source, _provider = reference_assets
    config = load_profile_provider_config(PROFILE_CONFIG)
    persisted = tmp_path / f"1ac1fd0cc4b7_{REFERENCE_WORKBOOK.name}"
    shutil.copyfile(REFERENCE_WORKBOOK, persisted)

    provider = WorkbookRowProfileProvider(
        config,
        persisted,
        validation_mode="compatible",
        original_filename=REFERENCE_WORKBOOK.name,
    )
    validation = provider.validate(scenario, source)

    assert provider.workbook_path == persisted.resolve()
    assert validation.source_file == REFERENCE_WORKBOOK.name
    assert validation.expected_filename == REFERENCE_WORKBOOK.name
    assert validation.reference_filename_matches is True
    assert validation.reference_sha256_matches is True
    assert validation.sample_count == 17418


def test_profile_provider_strict_mode_rejects_different_file_fingerprint(
    tmp_path: Path,
    reference_assets,
) -> None:
    _scenario, _source, _provider = reference_assets
    config = load_profile_provider_config(PROFILE_CONFIG)
    workbook = _copy_with_different_package_fingerprint(
        REFERENCE_WORKBOOK,
        tmp_path / REFERENCE_WORKBOOK.name,
    )

    with pytest.raises(DataValidationError, match="SHA-256 does not match"):
        WorkbookRowProfileProvider(config, workbook)


def test_profile_provider_compatible_mode_still_rejects_malformed_workbook(tmp_path: Path) -> None:
    config = load_profile_provider_config(PROFILE_CONFIG)
    workbook = tmp_path / REFERENCE_WORKBOOK.name
    workbook.write_text("not an Excel workbook", encoding="utf-8")

    with pytest.raises(VSMPostProcessingError):
        WorkbookRowProfileProvider(config, workbook, validation_mode="compatible")


def test_profile_provider_compatible_mode_still_rejects_missing_channel(reference_assets) -> None:
    scenario, source, _provider = reference_assets
    config = load_profile_provider_config(PROFILE_CONFIG)
    config = replace(config, last_channel_column=69)
    provider = WorkbookRowProfileProvider(config, REFERENCE_WORKBOOK, validation_mode="compatible")

    with pytest.raises(DataValidationError, match="channel layout"):
        provider.validate(scenario, source)


def test_profile_provider_compatible_mode_still_rejects_insufficient_phase_rows(reference_assets) -> None:
    scenario, source, _provider = reference_assets
    config = load_profile_provider_config(PROFILE_CONFIG)
    phase_rows = dict(config.phase_rows or {})
    phase_rows["P05"] = (5364, 5365)
    config = replace(config, phase_rows=phase_rows)
    provider = WorkbookRowProfileProvider(config, REFERENCE_WORKBOOK, validation_mode="compatible")

    with pytest.raises(DataValidationError, match="provider row count"):
        provider.validate(scenario, source)


def test_provider_resolves_exactly_the_four_missing_phases_in_the_plan(reference_assets) -> None:
    scenario, source, provider = reference_assets
    provider.validate(scenario, source)

    plan = build_composition_plan(scenario, profile_provider=provider)

    assert plan.unresolved_phase_ids == ()
    provider_rows = [row for row in plan.rows if row.profile_provider_id is not None]
    assert {row.phase_id for row in provider_rows} == {"P05", "P06", "P08", "P10"}
    assert len(provider_rows) == 1718 + 2536 + 1563 + 1676
    assert all(row.profile_provider_id == provider.provider_id for row in provider_rows)
    assert all(row.profile_provider_source == REFERENCE_WORKBOOK.name for row in provider_rows)
    assert all(row.generation_mode == f"external_profile:{provider.provider_id}" for row in provider_rows)


def test_full_composition_materialises_all_17418_samples_with_traceable_external_profiles(reference_assets) -> None:
    scenario, source, provider = reference_assets

    result = compose_duty_cycle(scenario, source, provider)
    indexes = {role: source.channel_index(channel_id) for role, channel_id in scenario.channel_roles.items()}

    assert result.sample_count == 17418
    assert result.completed_phase_ids == tuple(f"P{i:02d}" for i in range(1, 13))
    assert [item.phase_id for item in result.profile_provenance] == ["P05", "P06", "P08", "P10"]
    assert result.provenance[0].report_row == 5
    assert result.provenance[-1].report_row == 17422

    # Full-mission numerical targets established during 13A/13A.1.
    assert result.values[-1, indexes["track_time_s"]] == pytest.approx(17417.0)
    assert result.values[-1, indexes["time_minutes"]] == pytest.approx(290.28333333333336)
    assert result.values[-1, indexes["distance_m"]] == pytest.approx(114001.1)
    assert result.values[-1, indexes["distance_km"]] == pytest.approx(114.0011)
    assert result.values[-1, indexes["battery_soc_pct"]] == pytest.approx(23.9383)
    assert np.max(result.values[:, indexes["speed_kph"]]) == pytest.approx(62.6233)
    assert np.max(result.values[:, indexes["generator_total_power_kw"]]) == pytest.approx(80.02669061662199)

    # Canonical loading integration intentionally includes every one-second fuel
    # increment, leaving one 0.00406 kg step above Sergio's spreadsheet result.
    assert result.values[-1, indexes["fuel_consumption_kg"]] == pytest.approx(39.84212)


def test_external_profile_phases_are_exact_reference_replays_while_canonical_logic_remains_separate(reference_assets) -> None:
    scenario, source, provider = reference_assets
    result = compose_duty_cycle(scenario, source, provider)
    reference = provider.dataset.values

    phase_by_id = {phase.phase_id: phase for phase in scenario.phases}
    for phase_id in ("P05", "P06", "P08", "P10"):
        phase = phase_by_id[phase_id]
        assert phase.report_rows is not None
        start = phase.report_rows.start_row - 5
        stop = phase.report_rows.end_row - 5 + 1
        assert np.array_equal(result.values[start:stop, :], reference[start:stop, :])

    indexes = {role: source.channel_index(channel_id) for role, channel_id in scenario.channel_roles.items()}
    # Core electrical mission channels are equivalent throughout the whole
    # hybrid composition; known differences are confined to documented
    # canonical boundary/integration choices.
    for role in (
        "track_time_s",
        "time_minutes",
        "battery_power_kw",
        "battery_heatflow_kw",
        "battery_energy_kwh",
        "battery_soc_pct",
    ):
        column = indexes[role]
        assert np.max(np.abs(result.values[:, column] - reference[:, column])) < 1e-9

    fuel_difference = result.values[:, indexes["fuel_consumption_kg"]] - reference[:, indexes["fuel_consumption_kg"]]
    assert np.max(np.abs(fuel_difference)) == pytest.approx(0.00406)
    assert fuel_difference[-1] == pytest.approx(0.00406)


def test_full_composition_csv_contains_all_rows_and_provider_provenance(tmp_path: Path, reference_assets) -> None:
    scenario, source, provider = reference_assets
    result = compose_duty_cycle(scenario, source, provider)

    path = export_duty_cycle_composition(result, tmp_path / "full.csv")
    with path.open(encoding="utf-8") as handle:
        header = handle.readline().rstrip("\n").split(",")
        rows = sum(1 for _ in handle)

    assert rows == 17418
    assert header[:6] == [
        "output_sample_index",
        "phase_id",
        "phase_local_index",
        "generation_mode",
        "profile_provider_id",
        "profile_provider_source",
    ]
