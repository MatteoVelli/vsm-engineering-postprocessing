from __future__ import annotations

import csv
from pathlib import Path

import yaml


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SPEC = PROJECT_ROOT / "docs" / "milestone_13A_1" / "proposed_duty_cycle_spec_v2.yaml"
CONFLICTS = PROJECT_ROOT / "docs" / "milestone_13A_1" / "reference_conflict_register.csv"
KPIS = PROJECT_ROOT / "docs" / "milestone_13A_1" / "kpi_definition_registry.csv"


def test_sergio_fidelity_spec_freezes_mission_structure_and_control_rules() -> None:
    spec = yaml.safe_load(SPEC.read_text(encoding="utf-8"))

    assert spec["specification_version"] == 2
    assert spec["scenario"]["expected_sample_count"] == 17418
    assert spec["timeline"]["first_timestamp_s"] == 0
    assert spec["timeline"]["last_timestamp_s"] == 17417
    assert spec["timeline"]["timestamp_span_s"] == 17417
    assert spec["battery"]["mission_nominal_capacity_kwh"] == 100.0
    assert spec["battery"]["physical_pack_reference_energy_kwh"] == 100.77

    phases = spec["phases"]
    assert len(phases) == 12
    assert sum(int(phase["output_samples"]) for phase in phases) == 17418
    loading = [phase for phase in phases if phase["type"] == "loading_opportunity_charge"]
    assert len(loading) == 5
    assert all(phase["output_samples"] == 900 for phase in loading)

    final_cycle = next(phase for phase in phases if phase["id"] == "P12")
    assert final_cycle["generator_start_threshold_pct"] == 5.0
    assert final_cycle["generator_stop_threshold_pct"] == 80.0

    controls = spec["controls"]["drive_range_extender"]
    assert controls["default"]["start_when_soc_below_pct"] == 40.0
    assert controls["default"]["stop_when_soc_at_or_above_pct"] == 80.0
    assert controls["state_transfer"]["P05_to_P06"] == "inherited_on"


def test_sergio_reference_conflicts_have_explicit_resolution_status() -> None:
    with CONFLICTS.open(encoding="utf-8-sig", newline="") as handle:
        rows = list(csv.DictReader(handle))

    assert len(rows) >= 15
    assert all(row["status"] for row in rows)
    assert any(row["topic"] == "Fuel energy and engine efficiency" for row in rows)
    assert any(row["status"].startswith("OPEN_") for row in rows)


def test_sergio_kpi_registry_distinguishes_battery_charge_and_discharge_semantics() -> None:
    with KPIS.open(encoding="utf-8-sig", newline="") as handle:
        rows = list(csv.DictReader(handle))

    by_id = {row["registry_id"]: row for row in rows}
    discharge = by_id["max_battery_discharge_power_magnitude_kw"]
    assert discharge["reference_location"] == "CT4"
    assert "-1*M17424" in discharge["reference_formula"]
    assert "negative_min_battery_power" in discharge["canonical_operation"]

    assert by_id["top_battery_heatflow_rms"]["current_or_next_tool_mapping"] == "report_battery_heatflow_rms"
    assert by_id["fuel_energy_consumed_kwh"]["decision_status"] == "REFERENCE_ONLY_UNRESOLVED"
