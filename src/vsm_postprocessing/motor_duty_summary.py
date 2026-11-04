"""Manufacturer discussion evidence derived from an existing automatic assessment.

Continuous flags and episodes belong exclusively to continuous_torque. This
module adds sample statistics; it never changes those results or engineering KPIs.
"""
from dataclasses import asdict
import math

import numpy as np


def _rms(values):
    scale = float(np.max(np.abs(values)))
    return scale * math.sqrt(float(np.mean((values / scale) ** 2))) if scale else 0.0


def _peak_check(speed, torque, time, peak):
    unavailable = lambda reason: dict(status='unavailable', reason=reason)
    if peak is None:
        return unavailable('No user-provided peak curve.')
    pairs = sorted(zip(peak.motor_speed_rpm, peak.motor_peak_torque_nm))
    x, y = np.asarray(pairs, dtype=float).T
    if np.any(x < 0) or np.any(np.diff(x) <= 0) or np.any(y <= 0):
        return unavailable('Peak checks require unique nonnegative RPM and positive torque limits.')
    if np.any(speed < x[0]) or np.any(speed > x[-1]):
        return unavailable('Motor RPM outside supplied peak curve; no extrapolation or partial-cycle pass.')
    limits = np.interp(speed, x, y)
    with np.errstate(over='ignore', divide='ignore'):
        utilisation = np.abs(torque) / limits
    if not np.all(np.isfinite(utilisation)):
        return unavailable('Peak utilisation is numerically unsupported.')
    residual = np.abs(torque) - limits
    # Retain strict residuals without treating display precision as a physical
    # tolerance. Even an above-resolution result is not a thermal/suitability verdict.
    numerical = residual > 0
    maximum = float(utilisation.max())
    return dict(status='available', reason=None, maximum_utilisation=maximum,
                strict_flagged_samples=int(numerical.sum()),
                strict_elapsed_s=float(np.sum(np.r_[np.diff(time), 0.0][numerical])),
                maximum_positive_residual_nm=max(0.0, float(residual.max())),
                above_display_precision=round(maximum * 100, 1) > 100.0,
                precision_note='One decimal percentage is a reporting screen, not a manufacturer acceptance tolerance.')


def _band_edges(maximum):
    if maximum <= 0:
        return np.array([0., 1.])
    target = maximum / 5
    magnitude = 10. ** math.floor(math.log10(target)) if target else maximum
    step = next((v * magnitude for v in (1, 2, 2.5, 5, 10) if v * magnitude >= target), maximum)
    count = max(1, int(math.ceil(maximum / step)))
    if not math.isfinite(step * count):
        return np.linspace(0, maximum, 6)
    return np.arange(count + 1, dtype=float) * step


def build_motor_duty_summary(dataset, assessment, peak=None):
    """Return evidence only when all four automatic motor assessments are valid."""
    if assessment is None or assessment.origin != 'automatic' or not assessment.analysis.available:
        return None
    analysis = assessment.analysis
    provenance = analysis.channel_provenance
    channels = provenance['channels']

    def values(name):
        binding = channels[name]
        return np.asarray(dataset.values[:, dataset.channel_index(binding['channel_id'])], dtype=float) * binding['conversion_factor']

    time = values('track_time')
    duration = float(time[-1] - time[0])
    if not math.isfinite(duration) or duration <= 0:
        return None
    raw = [(np.abs(values(f'electricsystem_em{i}_speed')), values(f'electricsystem_em{i}_torque')) for i in range(1, 5)]
    edges = _band_edges(max(float(speed.max()) for speed, _ in raw))
    bands = [dict(lower_rpm=float(a), upper_rpm=float(b), centre_rpm=float(a / 2 + b / 2),
                  rms_nm=[], sample_counts=[]) for a, b in zip(edges[:-1], edges[1:])]
    motors, events, diagnostics = [], [], []
    curve = assessment.inputs.curve
    for index, (result, (speed, torque)) in enumerate(zip(analysis.motors, raw), 1):
        mean_speed = float(speed.mean())
        limit = float(np.interp(mean_speed, curve.motor_speed_rpm, curve.continuous_torque_nm))
        rms = _rms(torque)
        positive, negative = torque[torque > 0], torque[torque < 0]
        peak_result = _peak_check(speed, torque, time, peak)
        motor = dict(motor=result.motor, maximum_segment_s=result.maximum_segment_s,
                     total_time_s=result.total_time_s, fraction_above=result.total_time_s / duration,
                     episode_count=result.episode_count, maximum_utilisation=result.maximum_utilisation,
                     flagged_samples=result.flagged_samples, rms_nm=rms,
                     maximum_speed_rpm=float(speed.max()), mean_speed_rpm=mean_speed,
                     maximum_driving_nm=float(positive.max()) if positive.size else 0.,
                     maximum_regeneration_nm=float(negative.min()) if negative.size else 0.,
                     mean_driving_nm=float(positive.mean()) if positive.size else None,
                     continuous_at_mean_speed_nm=limit, rms_continuous_ratio=rms / limit,
                     episodes_per_hour=result.episode_count * 3600 / duration, peak=peak_result)
        if positive.size == 0:
            diagnostics.append(f'{result.motor}: mean positive driving torque N/A; no positive-torque samples.')
        if peak_result['status'] == 'unavailable':
            diagnostics.append(f"{result.motor}: peak check N/A; {peak_result['reason']}")
        motors.append(motor)
        for band_index, band in enumerate(bands):
            mask = (speed >= band['lower_rpm']) & ((speed <= band['upper_rpm']) if band_index == len(bands) - 1 else (speed < band['upper_rpm']))
            band['sample_counts'].append(int(mask.sum()))
            band['rms_nm'].append(_rms(torque[mask]) if mask.any() else None)
            if not mask.any():
                diagnostics.append(f"{result.motor}: RMS N/A in empty speed band {band['lower_rpm']:g} to {band['upper_rpm']:g} rpm.")
        for number, event in enumerate(result.episodes, 1):
            selection = slice(event.first_sample, event.last_sample + 1)
            t = torque[selection]
            utilisation = np.abs(t) / np.asarray(result.limits_nm[selection])
            mode = 'Mixed' if np.any(t > 0) and np.any(t < 0) else 'Driving' if np.any(t > 0) else 'Regeneration' if np.any(t < 0) else 'Zero torque'
            events.append(dict(motor=result.motor, episode=number, **asdict(event),
                               flagged_samples=event.last_sample - event.first_sample + 1,
                               mean_speed_rpm=float(speed[selection].mean()), mean_torque_nm=float(t.mean()),
                               maximum_torque_nm=float(t.max()), maximum_absolute_torque_nm=float(np.abs(t).max()),
                               mean_utilisation=float(utilisation.mean()), maximum_utilisation=float(utilisation.max()), mode=mode))
        if np.any(values(f'electricsystem_em{index}_speed') < 0):
            diagnostics.append(f'{result.motor}: reverse RPM observed; driving/regeneration labels follow torque sign and are indicative.')

    def maximum(key):
        value = max(m[key] for m in motors)
        # Exact source maxima distinguish close-but-different utilisation results.
        return dict(value=value, motors=[m['motor'] for m in motors if m[key] == value])

    kpis = {key: maximum(key) for key in ('maximum_segment_s', 'total_time_s', 'fraction_above', 'maximum_utilisation')}
    longest = ', '.join(kpis['maximum_segment_s']['motors'])
    exceeded = any(m['flagged_samples'] > 0 for m in motors)
    interpretation = [
        'Supplied continuous rating exceeded during this cycle.' if exceeded else 'No observed torque exceedance of the supplied continuous rating during this cycle.',
        f"Longest exceedance: {kpis['maximum_segment_s']['value']:g} s ({longest})." if exceeded else 'Consecutive and accumulated exceedance durations are zero for all motors.',
        f"Maximum instantaneous continuous utilisation: {kpis['maximum_utilisation']['value']:.1%} ({', '.join(kpis['maximum_utilisation']['motors'])}).",
    ]
    available = [m for m in motors if m['peak']['status'] == 'available']
    above = [m for m in available if m['peak']['above_display_precision']]
    residuals = [m for m in available if m['peak']['strict_flagged_samples']]
    if above:
        interpretation.append('Peak curve exceeded beyond displayed precision: ' + ', '.join(m['motor'] for m in above) + '. Manufacturer transient criteria require confirmation.')
    elif residuals:
        interpretation.append('Small positive peak-curve residuals round to 100.0% utilisation; physical significance is unconfirmed. See numerical evidence below.')
    elif available:
        interpretation.append('No numerical peak-curve exceedances observed in supported checks.')
    else:
        interpretation.append('Peak-curve analysis unavailable: ' + motors[0]['peak']['reason'])
    if len(available) not in (0, 4):
        interpretation.append('Some peak checks unavailable because of input or RPM coverage; see diagnostics.')
    interpretation.append('Thermal acceptability requires manufacturer confirmation. Sample RMS and the indicative ratio are not a validated thermal model.')
    return dict(source=str(dataset.source_path), source_sha256=dataset.quality.source_sha256,
                duration_s=duration, sample_count=len(time), motors=motors, kpis=kpis,
                events=events, bands=bands, diagnostics=diagnostics, interpretation=interpretation,
                continuous_policy=analysis.summary()['policy'], channel_provenance=provenance,
                definitions=dict(rms='sqrt(mean(T^2)); all samples, including the final sample; not time-weighted or a thermal equivalent.',
                                 means='Arithmetic sample means; speed uses absolute RPM. Driving/regeneration labels use source torque sign.',
                                 time='Forward observed intervals [t_i,t_(i+1)); final sample has zero duration. Percent = elapsed exceedance / observed cycle duration.',
                                 ratio='Sample RMS divided by continuous limit interpolated at arithmetic mean absolute RPM; indicative only.',
                                 peak='Strict absolute torque comparison with linearly interpolated user peak magnitude; symmetric regeneration assumed; no extrapolation. Report display precision is not a physical tolerance.',
                                 bands='Shared, dynamically generated absolute-RPM bands; lower bound inclusive, upper exclusive except the last band; empty bands are gaps, not zero RMS.'))
