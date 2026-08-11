from __future__ import annotations

import csv
from pathlib import Path

import pytest
import yaml

from vsm_postprocessing.duty_cycle import (
    build_composition_plan,
    export_composition_plan,
    load_duty_cycle_config,
    validate_source_dataset,
)
from vsm_postprocessing.errors import ConfigurationError
from vsm_postprocessing.importer import load_data_file


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CONFIG = PROJECT_ROOT / "config" / "duty_cycle_sergio_reference.yaml"
FORMAL_SPEC = PROJECT_ROOT / "docs" / "milestone_13A_1" / "proposed_duty_cycle_spec_v2.yaml"
REFERENCE_MAPPING = PROJECT_ROOT / "docs" / "milestone_13A" / "source_to_report_row_mapping.csv"


SOURCE_WORKBOOK = PROJECT_ROOT / "reference_files" / (
    "Sprayer_Caiman_SP_9300Kg_Hybrid_Gen80kW_30kph_74Ht_4000KgAQ_57-4pcSOC_5-80_1C2G_02.xlsx"
)


def test_sergio_duty_cycle_config_loads_and_freezes_12_phase_structure() -> None:
    scenario = load_duty_cycle_config(CONFIG)

    assert scenario.scenario_id == "sergio_reference_hybrid_1c2g_30_60kph"
    assert scenario.sample_rate_hz == 1.0
    assert scenario.expected_sample_count == 17418
    assert scenario.configured_sample_count == 17418
    assert scenario.first_timestamp_s == 0.0
    assert scenario.last_timestamp_s == 17417.0
    assert scenario.source_data_start_row == 5
    assert len(scenario.phases) == 12
    assert [phase.phase_id for phase in scenario.phases] == [f"P{index:02d}" for index in range(1, 13)]
    assert sum(phase.phase_type == "loading_opportunity_charge" for phase in scenario.phases) == 5


def test_formal_phase_13a1_spec_is_accepted_by_runtime_loader() -> None:
    scenario = load_duty_cycle_config(FORMAL_SPEC)
    assert scenario.expected_sample_count == 17418
    assert scenario.source_data_start_row == 5
    assert len(scenario.phases) == 12
    assert scenario.phases[-1].phase_id == "P12"


def test_composition_plan_has_exact_reference_row_and_time_boundaries() -> None:
    scenario = load_duty_cycle_config(CONFIG)
    plan = build_composition_plan(scenario)

    assert plan.sample_count == 17418
    assert plan.rows[0].output_sample_index == 0
    assert plan.rows[0].report_row == 5
    assert plan.rows[0].track_time_s == 0.0
    assert plan.rows[-1].output_sample_index == 17417
    assert plan.rows[-1].report_row == 17422
    assert plan.rows[-1].track_time_s == 17417.0

    p02 = plan.rows_for_phase("P02")
    assert len(p02) == 900
    assert p02[0].report_row == 1767
    assert p02[-1].report_row == 2666
    assert all(row.source_row_aligned is None for row in p02)

    p12 = plan.rows_for_phase("P12")
    assert len(p12) == 1866
    assert p12[0].source_row_aligned == 5
    assert p12[0].source_sample_index_aligned == 0
    assert p12[-1].source_row_aligned == 1870
    assert p12[-1].source_sample_index_aligned == 1865


def test_composition_plan_matches_phase_13a_structural_mapping_row_for_row() -> None:
    scenario = load_duty_cycle_config(CONFIG)
    plan = build_composition_plan(scenario)

    with REFERENCE_MAPPING.open(encoding="utf-8-sig", newline="") as handle:
        reference_rows = list(csv.DictReader(handle))

    assert len(reference_rows) == plan.sample_count
    for planned, reference in zip(plan.rows, reference_rows, strict=True):
        assert planned.output_sample_index == int(reference["sample_index"])
        assert planned.report_row == int(reference["report_row"])
        assert planned.track_time_s == float(reference["track_time_s"])
        assert planned.phase_id == reference["phase_id"]
        assert planned.phase_type == reference["phase_type"]
        assert planned.phase_local_index == int(reference["phase_local_index"])

        expected_source_row = int(reference["source_row_aligned"]) if reference["source_row_aligned"] else None
        expected_source_sample = (
            int(reference["source_sample_index_aligned"])
            if reference["source_sample_index_aligned"]
            else None
        )
        assert planned.source_row_aligned == expected_source_row
        assert planned.source_sample_index_aligned == expected_source_sample


def test_plan_explicitly_flags_reference_profiles_that_cannot_yet_be_materialised() -> None:
    scenario = load_duty_cycle_config(CONFIG)
    plan = build_composition_plan(scenario)

    assert plan.unresolved_phase_ids == ("P05", "P06", "P08", "P10")
    assert all(not row.profile_definition_resolved for row in plan.rows_for_phase("P06"))
    assert all("road/reference source" in row.notes for row in plan.rows_for_phase("P06"))
    assert all(row.profile_definition_resolved for row in plan.rows_for_phase("P12"))


def test_composition_plan_export_is_complete_and_traceable(tmp_path: Path) -> None:
    scenario = load_duty_cycle_config(CONFIG)
    plan = build_composition_plan(scenario)
    output = export_composition_plan(plan, tmp_path / "plan.csv")

    with output.open(encoding="utf-8", newline="") as handle:
        rows = list(csv.DictReader(handle))

    assert len(rows) == 17418
    assert rows[0]["phase_id"] == "P01"
    assert rows[-1]["phase_id"] == "P12"
    assert rows[-1]["track_time_s"] == "17417.0"
    assert rows[-1]["source_row_aligned"] == "1870"


def test_invalid_duty_cycle_config_rejects_duplicate_phase_ids(tmp_path: Path) -> None:
    raw = yaml.safe_load(CONFIG.read_text(encoding="utf-8"))
    raw["phases"][1]["id"] = raw["phases"][0]["id"]
    invalid = tmp_path / "invalid.yaml"
    invalid.write_text(yaml.safe_dump(raw, sort_keys=False), encoding="utf-8")

    with pytest.raises(ConfigurationError, match="Duplicate duty-cycle phase id"):
        load_duty_cycle_config(invalid)


def test_invalid_duty_cycle_config_rejects_sample_count_mismatch(tmp_path: Path) -> None:
    raw = yaml.safe_load(CONFIG.read_text(encoding="utf-8"))
    raw["scenario"]["expected_sample_count"] = 17419
    raw["timeline"]["last_timestamp_s"] = 17418
    invalid = tmp_path / "invalid.yaml"
    invalid.write_text(yaml.safe_dump(raw, sort_keys=False), encoding="utf-8")

    with pytest.raises(ConfigurationError, match="phases contain 17418 samples; expected 17419"):
        load_duty_cycle_config(invalid)


@pytest.mark.skipif(not SOURCE_WORKBOOK.exists(), reason="Client source workbook is not present")
def test_sergio_source_workbook_is_compatible_with_duty_cycle_source_alignments() -> None:
    scenario = load_duty_cycle_config(CONFIG)
    dataset = load_data_file(SOURCE_WORKBOOK)
    validation = validate_source_dataset(scenario, dataset)

    assert validation.source_sample_count == 1866
    assert validation.required_max_source_sample_index == 1865
    assert validation.source_data_start_row == 5
    assert validation.scenario_source_data_start_row == 5
    assert validation.source_nominal_time_step_s == 1.0
    assert validation.scenario_sample_period_s == 1.0


@pytest.mark.skipif(not SOURCE_WORKBOOK.exists(), reason="Client source workbook is not present")
def test_supported_prefix_materialises_p01_to_p04_and_stops_before_unresolved_p05() -> None:
    from vsm_postprocessing.duty_cycle import compose_supported_prefix

    scenario = load_duty_cycle_config(CONFIG)
    source = load_data_file(SOURCE_WORKBOOK)
    result = compose_supported_prefix(scenario, source)
    index = {role: source.channel_index(channel_id) for role, channel_id in scenario.channel_roles.items()}

    assert result.completed_phase_ids == ("P01", "P02", "P03", "P04")
    assert result.stopped_before_phase_id == "P05"
    assert result.sample_count == 5359
    assert len(result.provenance) == 5359
    assert result.provenance[-1].report_row == 5363

    values = result.values
    assert values[0, index["battery_soc_pct"]] == pytest.approx(95.0)
    assert values[1761, index["battery_soc_pct"]] == pytest.approx(61.8803)
    assert values[1762, index["battery_power_kw"]] == pytest.approx(60.0)
    assert values[2661, index["battery_soc_pct"]] == pytest.approx(76.8803)
    assert values[2661, index["fuel_consumption_kg"]] == pytest.approx(3.654)
    assert values[2661, index["generator_total_power_kw"]] == pytest.approx(80.00422042560322)
    assert values[2661, index["tank_mass_kg"]] == pytest.approx(3998.63394495)

    # P03 deliberately applies the explicit scenario reset rather than forcing
    # continuity from the preceding charging result.
    assert values[2662, index["battery_soc_pct"]] == pytest.approx(77.0)
    assert values[4458, index["battery_soc_pct"]] == pytest.approx(43.7648)

    # Canonical composer distance remains continuous across loading; it does not
    # reproduce Sergio's +1 m spreadsheet boundary artefact.
    assert values[1762, index["distance_m"]] == pytest.approx(values[1761, index["distance_m"]])
    assert values[2662, index["distance_m"]] == pytest.approx(values[2661, index["distance_m"]])

    # The deterministic loading action applies all 900 fuel increments in P04,
    # unlike the reference workbook's first-row fuel boundary artefact.
    assert values[-1, index["fuel_consumption_kg"]] == pytest.approx(7.308)
    assert values[-1, index["battery_soc_pct"]] == pytest.approx(58.7648)


@pytest.mark.skipif(
    not (SOURCE_WORKBOOK.exists() and (PROJECT_ROOT / "reference_files" / "Sprayer_Caiman_SP_9300Kg_Electrification_03.xlsx").exists()),
    reason="Client source/reference workbooks are not present",
)
def test_supported_prefix_matches_reference_on_resolved_core_logic_with_documented_canonical_differences() -> None:
    import numpy as np
    from vsm_postprocessing.duty_cycle import compose_supported_prefix
    from vsm_postprocessing.importer import ImportOptions

    report_workbook = PROJECT_ROOT / "reference_files" / "Sprayer_Caiman_SP_9300Kg_Electrification_03.xlsx"
    scenario = load_duty_cycle_config(CONFIG)
    source = load_data_file(SOURCE_WORKBOOK)
    reference = load_data_file(
        report_workbook,
        ImportOptions(
            header_row=3,
            unit_row=4,
            data_start_row=5,
            data_end_row=17422,
            last_channel_column=70,
            strict=True,
        ),
    )
    result = compose_supported_prefix(scenario, source)
    index = {role: source.channel_index(channel_id) for role, channel_id in scenario.channel_roles.items()}
    ref = reference.values[: result.sample_count]

    exact_or_near_roles = [
        "track_time_s",
        "time_minutes",
        "battery_power_kw",
        "battery_power_squared",
        "battery_heatflow_kw",
        "battery_heatflow_squared",
        "battery_energy_kwh",
        "battery_soc_pct",
        "energy_recuperated_kwh",
        "generator_2_power_kw",
        "tank_mass_kg",
    ]
    for role in exact_or_near_roles:
        column = index[role]
        assert np.max(np.abs(result.values[:, column] - ref[:, column])) < 1e-6

    # Canonical differences intentionally retained instead of spreadsheet quirks.
    distance_error = result.values[:, index["distance_m"]] - ref[:, index["distance_m"]]
    assert np.min(distance_error) == pytest.approx(-1.0)
    assert np.max(distance_error) == pytest.approx(0.0)
    assert result.values[-1, index["fuel_consumption_kg"]] - ref[-1, index["fuel_consumption_kg"]] == pytest.approx(0.00406)

    # P04's first reference row forces generator power to zero even though the
    # configured loading action is already active. Canonical output removes it.
    generator_error = result.values[:, index["generator_total_power_kw"]] - ref[:, index["generator_total_power_kw"]]
    nonzero = np.flatnonzero(np.abs(generator_error) > 1e-9)
    assert nonzero.tolist() == [4459]
    assert generator_error[4459] == pytest.approx(80.00422042560322)
