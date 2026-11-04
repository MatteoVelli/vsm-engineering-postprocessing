"""Deterministic motor-side continuous-duty assessment, independent of reports/UI."""
from dataclasses import asdict, dataclass
import math

import numpy as np

from .errors import ConfigurationError
from .motor_performance import SPEED, _number


CONTINUOUS = 'Motor continuous torque [Nm]'
MOTORS = ('EM1 / FL', 'EM2 / FR', 'EM3 / RL', 'EM4 / RR')
POLICY = ('Linear interpolation at absolute motor RPM; symmetric positive/negative continuous '
          'torque magnitude supplied by the user, not a universal motor property. Strict '
          'exceedance: abs(torque) > limit. Sample i holds on [t[i], t[i+1]); the final '
          'sample has zero duration. No speed extrapolation, terminal time extension or '
          'threshold-crossing interpolation. Invalid/unsupported samples make that motor '
          'unavailable; a whole-vehicle maximum requires all four motors. Exceedance alone '
          'does not establish thermal failure.')


def blank(value):
    return value is None or isinstance(value, str) and not value.strip() or isinstance(value, (float, np.floating)) and math.isnan(value)


@dataclass(frozen=True)
class ContinuousTorqueCurve:
    motor_speed_rpm: tuple[float, ...]
    continuous_torque_nm: tuple[float, ...]

    def __post_init__(self):
        speed = tuple(_number(v, 'Continuous curve RPM') for v in self.motor_speed_rpm)
        torque = tuple(_number(v, 'Continuous torque') for v in self.continuous_torque_nm)
        if len(speed) != len(torque) or not 2 <= len(speed) <= 1048572:
            raise ConfigurationError('Enter at least two complete continuous RPM/torque pairs within Excel row capacity.')
        if any(v < 0 for v in speed) or any(v <= 0 for v in torque):
            raise ConfigurationError('Continuous curve RPM must be nonnegative and continuous torque must be positive.')
        pairs = sorted(zip(speed, torque))
        if any(a[0] == b[0] for a, b in zip(pairs, pairs[1:])):
            raise ConfigurationError('Continuous curve RPM points must be unique.')
        object.__setattr__(self, 'motor_speed_rpm', tuple(p[0] for p in pairs))
        object.__setattr__(self, 'continuous_torque_nm', tuple(p[1] for p in pairs))

    @classmethod
    def from_rows(cls, rows):
        pairs = []
        for index, row in enumerate(rows, 1):
            torque = row.get(CONTINUOUS)
            if blank(torque):  # Peak-only points belong to the independent peak curve.
                continue
            speed = row.get(SPEED)
            if blank(speed):
                raise ConfigurationError(f'Continuous curve row {index}: enter RPM for the continuous torque.')
            pairs.append((speed, torque))
        return cls(tuple(p[0] for p in pairs), tuple(p[1] for p in pairs))


@dataclass(frozen=True)
class ExceedanceEpisode:
    first_sample: int
    last_sample: int
    start_s: float
    last_flagged_sample_s: float
    interval_end_s: float
    duration_s: float


@dataclass(frozen=True)
class MotorContinuousResult:
    motor: str
    status: str
    reason: str | None = None
    sample_count: int = 0
    flagged_samples: int | None = None
    maximum_segment_s: float | None = None
    total_time_s: float | None = None
    episode_count: int | None = None
    maximum_utilisation: float | None = None
    unsupported_samples: int = 0
    episodes: tuple[ExceedanceEpisode, ...] = ()
    flags: tuple[bool | None, ...] = ()
    limits_nm: tuple[float | None, ...] = ()

    def summary(self):
        return {k: v for k, v in asdict(self).items() if k not in {'flags', 'limits_nm'}}


@dataclass(frozen=True)
class ContinuousTorqueAnalysis:
    motors: tuple[MotorContinuousResult, ...]
    channel_provenance: dict

    @property
    def available(self):
        return len(self.motors) == 4 and all(m.status == 'available' for m in self.motors)

    @property
    def maximum_segment_s(self):
        return max(m.maximum_segment_s for m in self.motors) if self.available else None

    @property
    def responsible_motors(self):
        maximum = self.maximum_segment_s
        return tuple(m.motor for m in self.motors if maximum is not None and maximum > 0 and m.maximum_segment_s == maximum)

    def summary(self):
        return {'status': 'available' if self.available else 'unavailable', 'policy': POLICY,
                'maximum_segment_s': self.maximum_segment_s, 'responsible_motors': self.responsible_motors,
                'motors': [m.summary() for m in self.motors], 'channel_provenance': self.channel_provenance}


def analyse_motor(time_s, speed_rpm, torque_nm, curve, motor='Motor'):
    """Use observed forward intervals; never substitute a sample count for seconds."""
    try:
        time, speed, torque = (np.asarray(v, dtype=float) for v in (time_s, speed_rpm, torque_nm))
    except (TypeError, ValueError):
        return MotorContinuousResult(motor, 'unavailable', 'Numeric paired time, RPM and torque samples are required.')
    unavailable = lambda reason, **kw: MotorContinuousResult(motor, 'unavailable', reason, time.size if time.ndim == 1 else 0, **kw)
    if any(v.ndim != 1 for v in (time, speed, torque)) or len(time) < 2 or len(speed) != len(time) or len(torque) != len(time):
        return unavailable('At least two paired time, RPM and torque samples are required.')
    with np.errstate(over='ignore', invalid='ignore'):
        intervals = np.diff(time)
    if not np.all(np.isfinite(time)) or not np.all(np.isfinite(intervals)) or np.any(intervals <= 0):
        return unavailable('Timestamps must be finite and strictly increasing.')
    magnitude = np.abs(speed)
    supported = np.isfinite(speed) & np.isfinite(torque) & (magnitude >= curve.motor_speed_rpm[0]) & (magnitude <= curve.motor_speed_rpm[-1])
    if not np.all(supported):
        return unavailable('Nonfinite motor samples or RPM outside the supplied continuous curve; no extrapolation.',
                           unsupported_samples=int(np.count_nonzero(~supported)))
    limits = np.interp(magnitude, curve.motor_speed_rpm, curve.continuous_torque_nm)
    if not np.all(np.isfinite(limits)) or np.any(limits <= 0):
        return unavailable('Interpolated continuous limits are numerically unsupported.')
    flags = np.abs(torque) > limits
    with np.errstate(over='ignore', divide='ignore'):
        utilisation = np.abs(torque) / limits
    if not np.all(np.isfinite(utilisation)):
        return unavailable('Continuous torque utilisation overflow; check continuous ratings.')
    widths = np.r_[intervals, 0.0]
    starts = np.flatnonzero(flags & ~np.r_[False, flags[:-1]])
    ends = np.flatnonzero(flags & ~np.r_[flags[1:], False])
    episodes = tuple(ExceedanceEpisode(int(a), int(b), float(time[a]), float(time[b]),
                        float(time[min(b+1, len(time)-1)]), float(np.sum(widths[a:b+1]))) for a, b in zip(starts, ends))
    return MotorContinuousResult(motor, 'available', sample_count=len(time), flagged_samples=int(flags.sum()),
        maximum_segment_s=max((e.duration_s for e in episodes), default=0.0),
        total_time_s=float(np.sum(widths[flags])), episode_count=len(episodes),
        maximum_utilisation=float(utilisation.max()), episodes=episodes,
        flags=tuple(bool(v) for v in flags), limits_nm=tuple(float(v) for v in limits))


def analyse_dataset(dataset, profile, curve, resolution=None):
    """Bind actual raw channels using the selected profile's semantic mappings."""
    from .report_profile import resolve_profile
    from .utils import normalize_display_unit
    resolution = resolution or resolve_profile(dataset, profile)
    provenance = {'source': str(dataset.source_path), 'source_sha256': dataset.quality.source_sha256,
                  'profile': profile.profile_id, 'curve': asdict(curve), 'channels': {}}

    def values(name, kind):
        resolved = resolution.resolved.get(name)
        if resolved is None:
            raise ValueError(f'Actual source channel unavailable: {name}')
        channel = resolved.channel
        unit = (normalize_display_unit(channel.unit) or '').lower()
        factors = {'time': {'s': 1, 'ms': .001, 'min': 60, 'h': 3600},
                   'speed': {'rpm': 1, '1/min': 1, 'rad/s': 60/(2*math.pi), '1/s': 60, 'hz': 60},
                   'torque': {'nm': 1}}
        if unit not in factors[kind]:
            raise ValueError(f'Unsupported {kind} unit for {name}: {channel.unit}')
        factor = factors[kind][unit]
        provenance['channels'][name] = {'source_name': channel.source_name, 'unit': channel.unit,
                                        'conversion_factor': factor, 'channel_id': channel.channel_id}
        return np.asarray(dataset.values[:, dataset.channel_index(channel.channel_id)], dtype=float) * factor

    try:
        time = values('track_time', 'time')
    except ValueError as exc:
        return ContinuousTorqueAnalysis(tuple(MotorContinuousResult(m, 'unavailable', str(exc)) for m in MOTORS), provenance)
    results = []
    for index, label in enumerate(MOTORS, 1):
        try:
            speed = values(f'electricsystem_em{index}_speed', 'speed')
            torque = values(f'electricsystem_em{index}_torque', 'torque')
            results.append(analyse_motor(time, speed, torque, curve, label))
        except ValueError as exc:
            results.append(MotorContinuousResult(label, 'unavailable', str(exc), len(time)))
    return ContinuousTorqueAnalysis(tuple(results), provenance)
