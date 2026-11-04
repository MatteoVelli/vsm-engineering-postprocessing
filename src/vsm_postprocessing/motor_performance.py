"""Optional shared motor capability inputs, independent of VSM measurements."""
from dataclasses import dataclass
import math
from numbers import Real

from .errors import ConfigurationError


SPEED = 'Motor speed [rpm]'
TORQUE = 'Motor peak torque [Nm]'
GENERATOR = 'Generator peak torque [Nm]'
WHEEL_SPEED = 'Wheel speed [rpm]'
POSITIVE = 'Positive wheel torque [Nm]'
NEGATIVE = 'Negative wheel torque [Nm]'
ASSUMPTION = ('One user-provided curve shared by all four motors. Generator peak torque '
              'is assumed to be the negative of motor peak torque, as in Sergio\'s example; '
              'this symmetry is not a universal motor characteristic. Capability overlays '
              'are reference inputs, not VSM measurements or engineering statistics.')


def _number(value, name):
    if isinstance(value, bool) or not isinstance(value, Real) or not math.isfinite(value):
        raise ConfigurationError(name + ' must be a finite number.')
    return float(value)


@dataclass(frozen=True)
class MotorPerformanceCurve:
    motor_speed_rpm: tuple[float, ...]
    motor_peak_torque_nm: tuple[float, ...]
    final_drive_ratio: float
    driveline_efficiency: float

    def __post_init__(self):
        rpm = tuple(_number(v, 'Motor speed') for v in self.motor_speed_rpm)
        torque = tuple(_number(v, 'Motor peak torque') for v in self.motor_peak_torque_nm)
        ratio = _number(self.final_drive_ratio, 'Final-drive ratio')
        efficiency = _number(self.driveline_efficiency, 'Driveline efficiency')
        if len(rpm) != len(torque) or len(rpm) < 2:
            raise ConfigurationError('Enter at least two complete RPM/torque pairs.')
        if len(rpm) > 1048572:
            raise ConfigurationError('Motor curve exceeds the available Excel table rows.')
        if any(t < 0 for t in torque):
            raise ConfigurationError('Motor peak torque must be a nonnegative magnitude; generator torque is derived separately.')
        if ratio <= 0:
            raise ConfigurationError('Final-drive ratio must be greater than zero (dimensionless, not percent).')
        if not 0 < efficiency <= 1:
            raise ConfigurationError('Driveline efficiency must be greater than zero and no greater than one.')
        for name, value in (('motor_speed_rpm', rpm), ('motor_peak_torque_nm', torque),
                            ('final_drive_ratio', ratio), ('driveline_efficiency', efficiency)):
            object.__setattr__(self, name, value)
        if any(not math.isfinite(v) for row in self.calculated_rows() for v in row.values()):
            raise ConfigurationError('Motor capability calculations overflow; check RPM, torque, ratio and efficiency.')

    @classmethod
    def from_rows(cls, rows, ratio, efficiency):
        pairs = []
        for index, row in enumerate(rows, 1):
            x, y = row.get(SPEED), row.get(TORQUE)
            # The editor's unused insertion rows are not motor data.
            blank = lambda v: v is None or isinstance(v, str) and not v.strip() or isinstance(v, Real) and math.isnan(v)
            if blank(x) and blank(y):
                continue
            if blank(x) or blank(y):
                raise ConfigurationError(f'Motor curve row {index}: enter both RPM and peak torque, or remove the row.')
            pairs.append((_number(x, f'Row {index} RPM'), _number(y, f'Row {index} torque')))
        return cls(tuple(p[0] for p in pairs), tuple(p[1] for p in pairs), ratio, efficiency)

    def calculated_rows(self):
        return [{GENERATOR: -torque, WHEEL_SPEED: rpm/self.final_drive_ratio,
                 POSITIVE: torque*self.final_drive_ratio*self.driveline_efficiency,
                 NEGATIVE: -torque*self.final_drive_ratio/self.driveline_efficiency}
                for rpm, torque in zip(self.motor_speed_rpm, self.motor_peak_torque_nm)]

    def provenance(self):
        return {'input_origin': 'User-provided shared motor capability curve', 'assumption': ASSUMPTION,
                'motor_speed_rpm': self.motor_speed_rpm, 'motor_peak_torque_nm': self.motor_peak_torque_nm,
                'final_drive_ratio': self.final_drive_ratio, 'driveline_efficiency': self.driveline_efficiency,
                'formulas': {GENERATOR: '-motor_peak_torque', WHEEL_SPEED: 'motor_rpm / final_drive_ratio',
                             POSITIVE: 'motor_peak_torque * final_drive_ratio * driveline_efficiency',
                             NEGATIVE: '-motor_peak_torque * final_drive_ratio / driveline_efficiency'}}
