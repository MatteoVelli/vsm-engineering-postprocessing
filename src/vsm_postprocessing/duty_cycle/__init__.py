"""Configurable deterministic duty-cycle composition primitives."""

from .composer import (
    build_composition_plan,
    compose_duty_cycle,
    compose_supported_prefix,
    export_composition_plan,
    export_duty_cycle_composition,
    export_pipeline_dataset,
    export_partial_composition,
    validate_source_dataset,
)
from .config import load_duty_cycle_config, load_profile_provider_config
from .models import (
    DutyCycleComposition,
    DutyCycleCompositionPlan,
    DutyCyclePartialComposition,
    DutyCyclePhase,
    DutyCycleRowProvenance,
    DutyCycleScenario,
    OpportunityChargeConfig,
    DutyCycleSourceValidation,
    ReportRowRange,
    SourceRowRange,
)
from .profile_provider import (
    PhaseProfileProvenance,
    PhaseProfileProvider,
    ProfileProviderValidation,
    WorkbookProfileProviderConfig,
    WorkbookRowProfileProvider,
)

__all__ = [
    "DutyCycleComposition",
    "DutyCycleCompositionPlan",
    "DutyCyclePartialComposition",
    "DutyCyclePhase",
    "DutyCycleRowProvenance",
    "DutyCycleScenario",
    "OpportunityChargeConfig",
    "DutyCycleSourceValidation",
    "ReportRowRange",
    "SourceRowRange",
    "PhaseProfileProvenance",
    "PhaseProfileProvider",
    "ProfileProviderValidation",
    "WorkbookProfileProviderConfig",
    "WorkbookRowProfileProvider",
    "build_composition_plan",
    "compose_duty_cycle",
    "compose_supported_prefix",
    "export_composition_plan",
    "export_duty_cycle_composition",
    "export_pipeline_dataset",
    "export_partial_composition",
    "load_duty_cycle_config",
    "load_profile_provider_config",
    "validate_source_dataset",
]
