"""Explicit, traceable reporting metadata; never used by vehicle calculations."""
from dataclasses import dataclass, fields
import math
from typing import Mapping, Any

from .errors import ConfigurationError


@dataclass(frozen=True)
class ReportingInputs:
    machine_configuration_id: str | None = None
    battery_mass_kg: float | None = None
    battery_mass_status: str = 'not_provided'
    battery_mass_provenance: str | None = None
    density_kg_per_L: float | None = None
    density_provenance: str | None = None
    hitch_force_is_remaining_agro_mass: bool = False
    agro_mass_provenance: str | None = None
    agro_volume_source: str | None = None
    agro_volume_provenance: str | None = None
    agro_flow_source: str | None = None
    agro_flow_provenance: str | None = None

    def __post_init__(self):
        for key in ('battery_mass_kg', 'density_kg_per_L'):
            value = getattr(self, key)
            if value is not None and (isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value) or value <= 0):
                raise ConfigurationError(f'{key} must be strictly positive and finite')
        if self.battery_mass_status not in {'confirmed', 'estimated', 'user_provided', 'not_provided'}:
            raise ConfigurationError('battery_mass_status must be confirmed, estimated, user_provided or not_provided')
        if self.battery_mass_kg is not None and (self.battery_mass_status == 'not_provided' or not self.battery_mass_provenance):
            raise ConfigurationError('Battery mass requires status and provenance')
        if self.battery_mass_kg is None and self.battery_mass_status != 'not_provided':
            raise ConfigurationError('Battery mass status requires a mass value')
        if self.density_kg_per_L is not None and not self.density_provenance:
            raise ConfigurationError('Agrochemical density requires provenance in kg/L')
        if not isinstance(self.hitch_force_is_remaining_agro_mass, bool):
            raise ConfigurationError('hitch_force_is_remaining_agro_mass must be boolean')
        for item in fields(self):
            if item.name in {'battery_mass_kg', 'density_kg_per_L', 'hitch_force_is_remaining_agro_mass'}:
                continue
            value = getattr(self, item.name)
            if value is not None and (not isinstance(value, str) or not value.strip()):
                raise ConfigurationError(f'{item.name} must be nonempty text')

    @classmethod
    def from_mapping(cls, raw: Mapping[str, Any] | None):
        if raw is None:
            return cls()
        if not isinstance(raw, Mapping):
            raise ConfigurationError('reporting_inputs must be a mapping')
        raw = dict(raw)
        if raw.get('battery_mass_kg') is not None:
            raw.setdefault('battery_mass_status', 'user_provided')
            raw.setdefault('battery_mass_provenance', 'Manual report input')
        # Public UI/profile name; keep the shorter calculator argument compatible.
        if 'agro_density_kg_per_L' in raw:
            if 'density_kg_per_L' in raw:
                raise ConfigurationError('Specify density only once')
            raw['density_kg_per_L'] = raw.pop('agro_density_kg_per_L')
        unknown = set(raw) - {f.name for f in fields(cls)}
        if unknown:
            raise ConfigurationError(f'Unknown reporting inputs: {sorted(unknown)}')
        if raw.get('density_kg_per_L') is not None and not raw.get('density_provenance'):
            raise ConfigurationError('Agrochemical density requires provenance in kg/L')
        return cls(**raw)

    @property
    def battery_display(self):
        if self.battery_mass_kg is None:
            return 'Not specified'
        prefix = '~' if self.battery_mass_status == 'estimated' else ''
        return f'{prefix}{self.battery_mass_kg:,.6g} kg ({self.battery_mass_status})'

    @property
    def battery_card_display(self):
        return 'N/A' if self.battery_mass_kg is None else f'{self.battery_mass_kg:,.6g} kg'
