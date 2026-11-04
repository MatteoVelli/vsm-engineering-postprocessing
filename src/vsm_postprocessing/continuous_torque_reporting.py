"""Optional reporting inputs/results; never part of the VSM engineering KPIs."""
from dataclasses import asdict, dataclass
import math

from .continuous_torque import ContinuousTorqueCurve, ContinuousTorqueAnalysis, analyse_dataset
from .errors import ConfigurationError
from .motor_performance import _number


@dataclass(frozen=True)
class ContinuousTorqueInputs:
    curve: ContinuousTorqueCurve | None = None
    manual_duration_s: float | None = None
    manual_provenance: str | None = None
    final_drive_ratio: float | None = None
    driveline_efficiency: float | None = None

    def __post_init__(self):
        if self.curve is not None and not isinstance(self.curve, ContinuousTorqueCurve):
            raise ConfigurationError('Continuous torque requires a validated continuous curve.')
        if self.manual_duration_s is not None:
            duration = _number(self.manual_duration_s, 'Manual maximum segment duration')
            if duration < 0 or not isinstance(self.manual_provenance, str) or not self.manual_provenance.strip():
                raise ConfigurationError('Manual duration must be nonnegative and have an external calculation source.')
            object.__setattr__(self, 'manual_duration_s', duration)
        for name in ('final_drive_ratio', 'driveline_efficiency'):
            value = getattr(self, name)
            if value is not None:
                value = _number(value, name)
                if value <= 0 or name == 'driveline_efficiency' and value > 1:
                    raise ConfigurationError('Continuous wheel overlays require positive ratio and efficiency in (0, 1].')
                object.__setattr__(self, name, value)
        if self.curve and self.final_drive_ratio is not None and self.driveline_efficiency is not None:
            r, e = self.final_drive_ratio, self.driveline_efficiency
            if any(not math.isfinite(v) for rpm, torque in zip(self.curve.motor_speed_rpm, self.curve.continuous_torque_nm)
                   for v in (rpm/r, torque*r*e, -torque*r/e)):
                raise ConfigurationError('Continuous wheel capability conversion overflows; check ratio and efficiency.')


@dataclass(frozen=True)
class ContinuousTorqueAssessment:
    inputs: ContinuousTorqueInputs
    analysis: ContinuousTorqueAnalysis | None

    @property
    def duration_s(self):
        return self.analysis.maximum_segment_s if self.analysis and self.analysis.available else self.inputs.manual_duration_s

    @property
    def origin(self):
        if self.analysis and self.analysis.available:
            return 'automatic'
        return 'manual_fallback' if self.inputs.manual_duration_s is not None else None

    def summary(self):
        return {'inputs': asdict(self.inputs), 'analysis': self.analysis.summary() if self.analysis else None,
                'duration_s': self.duration_s, 'origin': self.origin,
                'manual_provenance': self.inputs.manual_provenance if self.origin == 'manual_fallback' else None,
                'excel_editing_note': 'Editing an Excel limit updates its chart. Re-run the application to recalculate episode metrics and the PowerPoint KPI.'}


def assess_continuous_torque(dataset, profile, inputs, resolution=None):
    if not isinstance(inputs, ContinuousTorqueInputs):
        raise ConfigurationError('Continuous analysis requires validated reporting inputs.')
    analysis = analyse_dataset(dataset, profile, inputs.curve, resolution) if inputs.curve else None
    return ContinuousTorqueAssessment(inputs, analysis)
