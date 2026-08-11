from __future__ import annotations

import argparse
from pathlib import Path

from .duty_cycle import (
    WorkbookRowProfileProvider,
    build_composition_plan,
    compose_duty_cycle,
    compose_supported_prefix,
    export_composition_plan,
    export_duty_cycle_composition,
    export_partial_composition,
    load_duty_cycle_config,
    load_profile_provider_config,
    validate_source_dataset,
)
from .importer import load_data_file


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Build, validate and materialise deterministic VSM duty-cycle scenarios."
    )
    parser.add_argument("config", type=Path, help="Duty-cycle YAML configuration")
    parser.add_argument(
        "--source",
        type=Path,
        default=None,
        help="Optional source VSM workbook/CSV to validate/materialise against the plan",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("outputs/duty_cycle_composition_plan.csv"),
        help="Output provenance CSV path",
    )
    parser.add_argument(
        "--materialize-supported-prefix",
        action="store_true",
        help="Materialise native source-backed/loading phases until the first unresolved phase",
    )
    parser.add_argument(
        "--prefix-output",
        type=Path,
        default=Path("outputs/duty_cycle_supported_prefix.csv"),
        help="Output CSV for --materialize-supported-prefix",
    )
    parser.add_argument(
        "--profile-config",
        type=Path,
        default=None,
        help="Optional external phase-profile provider YAML",
    )
    parser.add_argument(
        "--profile-workbook",
        type=Path,
        default=None,
        help="Workbook supplying externally resolved phase profiles",
    )
    parser.add_argument(
        "--materialize-full",
        action="store_true",
        help="Materialise all phases; requires --source, --profile-config and --profile-workbook",
    )
    parser.add_argument(
        "--full-output",
        type=Path,
        default=Path("outputs/duty_cycle_full_composition.csv"),
        help="Output CSV for --materialize-full",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    scenario = load_duty_cycle_config(args.config)

    source_validation = None
    source_dataset = None
    if args.source is not None:
        source_dataset = load_data_file(args.source)
        source_validation = validate_source_dataset(scenario, source_dataset)

    provider = None
    provider_validation = None
    if args.profile_config is not None or args.profile_workbook is not None:
        if args.profile_config is None or args.profile_workbook is None:
            parser.error("--profile-config and --profile-workbook must be supplied together")
        provider_config = load_profile_provider_config(args.profile_config)
        provider = WorkbookRowProfileProvider(provider_config, args.profile_workbook)
        if source_dataset is not None:
            provider_validation = provider.validate(scenario, source_dataset)

    plan = build_composition_plan(scenario, profile_provider=provider)
    output = export_composition_plan(plan, args.output)

    partial = None
    prefix_output = None
    if args.materialize_supported_prefix:
        if source_dataset is None:
            parser.error("--materialize-supported-prefix requires --source")
        partial = compose_supported_prefix(scenario, source_dataset)
        prefix_output = export_partial_composition(partial, args.prefix_output)

    full = None
    full_output = None
    if args.materialize_full:
        if source_dataset is None:
            parser.error("--materialize-full requires --source")
        if provider is None:
            parser.error("--materialize-full requires --profile-config and --profile-workbook")
        full = compose_duty_cycle(scenario, source_dataset, provider)
        full_output = export_duty_cycle_composition(full, args.full_output)

    print(f"Scenario: {scenario.scenario_id}")
    print(f"Samples: {plan.sample_count}")
    print(f"Phases: {len(scenario.phases)}")
    print(f"Final track time: {plan.rows[-1].track_time_s:.6f} s")
    if source_validation is not None:
        print(
            "Source validation: PASS "
            f"({source_validation.source_sample_count} samples, "
            f"required max index {source_validation.required_max_source_sample_index})"
        )
    if provider_validation is not None:
        print(
            "Profile provider validation: PASS "
            f"({provider_validation.provider_id}; phases "
            + ", ".join(provider_validation.supported_phase_ids)
            + ")"
        )
    if plan.unresolved_phase_ids:
        print("Profile-definition unresolved phases: " + ", ".join(plan.unresolved_phase_ids))
    else:
        print("Profile-definition unresolved phases: none")
    if partial is not None:
        print("Materialised prefix phases: " + ", ".join(partial.completed_phase_ids))
        print(f"Materialised prefix samples: {partial.sample_count}")
        print(f"Stopped before phase: {partial.stopped_before_phase_id or 'none'}")
        print(f"Materialised prefix CSV: {prefix_output}")
    if full is not None:
        print("Materialised full phases: " + ", ".join(full.completed_phase_ids))
        print(f"Materialised full samples: {full.sample_count}")
        print(
            "External profile phases: "
            + ", ".join(item.phase_id for item in full.profile_provenance)
        )
        print(f"Materialised full CSV: {full_output}")
    print(f"Provenance CSV: {output}")
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
