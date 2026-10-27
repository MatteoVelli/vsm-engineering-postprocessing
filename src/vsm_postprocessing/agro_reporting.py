"""Conservative agro reporting independent of the validated legacy KPI arrays."""
from dataclasses import asdict, dataclass, field
import math
from typing import Any

import numpy as np

from .reporting_inputs import ReportingInputs
from .agro_mass import resolve_mass_values
from .diagnostics import Diagnostic


@dataclass(frozen=True)
class ReportingMetric:
    metric_id: str
    label: str
    value: float | None
    unit: str
    reason: str = ''
    source: str = ''
    definition: str = ''
    status: str = 'unavailable'
    interpretation_status: str = 'unconfirmed'
    quality_status: str = 'unavailable'
    provenance: str = ''


@dataclass(frozen=True)
class AgroReport:
    capacity: ReportingMetric
    rate: ReportingMetric
    trace: dict[str, Any]
    net_mass_rate: ReportingMetric | None = None
    mass_values: np.ndarray | None = field(default=None, repr=False, compare=False)

    @property
    def mass_rate(self):
        return ReportingMetric('agro_mass_discharge_rate_kg_per_m', 'Agro Mass Discharge Diagnostic',
                               self.trace.get('mass_rate_kg_per_m'), 'kg/m',
                               reason='' if self.trace.get('mass_rate_kg_per_m') is not None else (
                                   self.capacity.reason or self.rate.reason or 'Selected mass/window does not support a monotonic mass-loss diagnostic; inspect quality and refill warnings'),
                               source=self.capacity.source,
                               interpretation_status=self.capacity.interpretation_status,
                               quality_status=self.capacity.quality_status, provenance=self.capacity.provenance,
                               definition='Initial minus final selected remaining mass / full-window travelled distance; not volumetric discharge',
                               status=self.capacity.status if self.trace.get('mass_rate_kg_per_m') is not None else 'unavailable')

    def to_dict(self):
        return {'capacity': asdict(self.capacity), 'rate': asdict(self.rate), 'trace': self.trace,
                'net_mass_rate': asdict(self.net_mass_rate) if self.net_mass_rate else None, 'mass_rate': asdict(self.mass_rate)}


def calculate_net_mass_discharge_rate(mass_kg, mass_unit, distance, distance_unit):
    """Sergio's net start-minus-end mass / final cumulative distance KPI.

    Inputs contain simulation samples only (the importer excludes summary rows).
    Mass is already in kg; profile math owns the force conversion and gravity.
    """
    definition = '(initial_agro_mass_kg - final_agro_mass_kg) / final_distance_km'
    metric = ReportingMetric('agro_discharge_rate_kg_per_km', 'Agro Discharge Rate',
                             None, 'kg/km', definition=definition)
    trace = {'definition': definition, 'warnings': []}

    def unavailable(reason):
        trace['warnings'].append('Agro discharge rate unavailable: ' + reason)
        return ReportingMetric(**{**asdict(metric), 'reason': reason}), trace

    if mass_kg is None:
        return unavailable('remaining agrochemical mass channel missing or invalid; check its force source and units')
    if distance is None:
        return unavailable('cumulative distance channel missing or invalid')
    if str(mass_unit).casefold() != 'kg' or distance_unit not in {'m', 'km'}:
        return unavailable('incompatible units: remaining mass must be kg and distance m or km')
    mass, d = np.asarray(mass_kg, dtype=float), np.asarray(distance, dtype=float)
    if mass.ndim != 1 or d.ndim != 1 or len(mass) != len(d) or len(d) < 2:
        return unavailable('at least two aligned simulation samples required')
    if not np.isfinite(mass).all() or not np.isfinite(d).all():
        return unavailable('nonfinite mass or distance samples; check endpoints and source data')
    if np.any(mass < 0):
        return unavailable('negative remaining agrochemical mass; verify source sign and load interpretation')
    if np.any(d < 0) or np.any(np.diff(d) < 0):
        return unavailable('negative or nonmonotonic cumulative distance')
    final_km = float(d[-1]) / (1000 if distance_unit == 'm' else 1)
    if final_km <= 0 or not math.isfinite(final_km):
        return unavailable('positive finite final cumulative distance required')
    if not np.isclose(d[0], 0, rtol=0, atol=1e-9):
        trace['warnings'].append('Cumulative distance does not start at zero; rate uses final cumulative distance as specified, not end-minus-start distance')
    if np.any(np.diff(mass) > 0):
        trace['warnings'].append('Remaining mass increases (possible refill); rate uses net start-minus-end mass, not cumulative discharge')
    discharged = float(mass[0]) - float(mass[-1])
    if discharged < 0:
        return unavailable('final mass exceeds initial mass; net discharge is negative')
    rate = discharged / final_km
    if not math.isfinite(rate):
        return unavailable('nonfinite net mass discharge rate')
    trace.update(initial_mass_kg=float(mass[0]), final_mass_kg=float(mass[-1]),
                 discharged_mass_kg=discharged, final_distance_km=final_km, sample_count=len(d))
    return ReportingMetric(**{**asdict(metric), 'value': rate, 'status': 'calculated',
                              'source': 'agrochemical_discharge [kg]; track_distance'}), trace


def calculate_agro_kpis(distance, distance_unit, *, inputs: ReportingInputs,
                        force_N=None, force_unit=None, volume_L=None, volume_unit=None,
                        flow=None, flow_unit=None, time_s=None, time_unit='s',
                         mass_kg=None, mass_unit=None, source_name=None):
    trace = {'analysis_window': 'entire report sample window', 'sample_count': len(distance) if distance is not None else 0,
             'density_kg_per_L': inputs.density_kg_per_L, 'density_provenance': inputs.density_provenance,
             'method': None, 'status': 'unavailable', 'warnings': [], 'gravity_m_per_s2': 9.81,
             'nominal_tank_capacity': 'not specified; not inferred from the observed maximum'}
    trace['physical_interpretation_provenance'] = inputs.agro_mass_provenance
    capacity = ReportingMetric('agro_capacity_observed_kg', 'Agro Capacity', None, 'kg',
                               'remaining-load source unavailable', definition='Observed maximum inferred remaining agrochemical mass; not nominal tank capacity')
    rate = ReportingMetric('agro_discharge_rate_L_per_m', 'Agro Discharge Rate', None, 'L/m',
                           definition='Discharged volume [L] / distance travelled [m], entire report window')

    def unavailable(reason):
        if reason not in trace['warnings']:
            trace['warnings'].append(reason)
        return AgroReport(capacity, ReportingMetric(**{**asdict(rate), 'reason': reason,
            'source': trace.get('source') or capacity.source, 'provenance': trace.get('source_provenance') or capacity.provenance,
            'interpretation_status': capacity.interpretation_status}), trace, mass_values=mass)

    # Resolve one explicitly mapped source, irrespective of machine identity.
    mass, mass_trace = resolve_mass_values(inputs, force=force_N, force_unit=force_unit,
                                           mass=mass_kg, mass_unit=mass_unit, source_name=source_name)
    trace['mass_source'] = mass_trace
    trace['gravity_m_per_s2'] = mass_trace['gravity_m_per_s2']
    trace['warnings'].extend(mass_trace['warnings'])
    trace['capacity_method'] = mass_trace['formula']
    source = inputs.agro_mass_source or 'Explicit manual rear-hitch force mapping'
    trace['mass_source_name'] = source
    mass_problem = mass_trace.get('reason')
    if mass is not None:
        trace.update(initial_remaining_mass_kg=float(mass[0]), final_remaining_mass_kg=float(mass[-1]))
        kind = 'inferred' if mass_trace['policy'] == 'force_inference' else 'simulated'
        capacity = ReportingMetric(**{**asdict(capacity), 'value': float(np.max(mass)), 'reason': '',
            'label': 'Agro Capacity (' + kind + ')', 'status': kind, 'source': source,
            'interpretation_status': mass_trace['interpretation_status'],
            'quality_status': mass_trace['quality_status'], 'provenance': mass_trace['provenance'] or '',
            'definition': 'Observed maximum ' + kind + ' remaining mass; not nominal tank capacity; ' + mass_trace['assumption']})
        if np.any(mass < 0):
            mass_problem = 'nonphysical negative inferred contents; verify source model and operating phases'
    else:
        capacity = ReportingMetric(**{**asdict(capacity), 'reason': mass_problem or 'remaining mass unavailable',
            'source': source, 'interpretation_status': mass_trace['interpretation_status'],
            'provenance': mass_trace['provenance'] or ''})

    if distance_unit not in {'m', 'km'}:
        return unavailable('distance unit mismatch (m or km required)')
    d = np.asarray(distance, dtype=float) if distance is not None else np.array([])
    if len(d) < 2 or not np.isfinite(d).all():
        return unavailable('nonfinite or insufficient distance data')
    with np.errstate(over='ignore', invalid='ignore'):
        d = d * (1000 if distance_unit == 'km' else 1)
    if not np.isfinite(d).all():
        return unavailable('nonfinite distance after conversion to meters')
    if np.any(np.diff(d) < 0) or np.any(d < 0):
        return unavailable('reversed/nonmonotonic or negative distance trace')
    travelled = float(d[-1]) - float(d[0])
    trace.update(distance_start_m=float(d[0]), distance_end_m=float(d[-1]), distance_travelled_m=travelled)
    if not np.isfinite(travelled) or travelled <= 0:
        return unavailable('positive finite travelled distance required')

    # Reject isolated jumps over 20x the median increment (>=5 nonzero steps).
    # This is a conservative quality screen, not a validated segmentation method.
    def jump(values):
        delta = np.abs(np.diff(values))
        active = delta[delta > 0]
        return len(active) >= 5 and bool(np.max(active) > 20 * np.median(active))

    if jump(d):
        return unavailable('distance discontinuity: verify operating window')
    if (mass is not None and not mass_problem and len(mass) == len(d)
            and not np.any(np.diff(mass) > 0) and not jump(mass)):
        discharged_mass = float(mass[0] - mass[-1])
        mass_rate = discharged_mass / travelled
        trace.update(initial_remaining_mass_kg=float(mass[0]), final_remaining_mass_kg=float(mass[-1]),
                     discharged_mass_kg=discharged_mass, mass_rate_kg_per_m=mass_rate if np.isfinite(mass_rate) else None)
    volume = None
    if inputs.agro_volume_source:
        trace['method'] = 'direct cumulative volume'
        trace['source'] = inputs.agro_volume_source
        trace['source_provenance'] = inputs.agro_volume_provenance
        if not inputs.agro_volume_provenance:
            return unavailable('direct volume source requires trusted provenance')
        if volume_L is None or volume_unit not in {'L', 'l'}:
            return unavailable('direct volume source missing or unit mismatch (L required)')
        v = np.asarray(volume_L, dtype=float)
        if len(v) != len(d) or not np.isfinite(v).all() or np.any(v < 0):
            return unavailable('nonfinite, negative or insufficient direct volume data')
        if np.any(np.diff(v) < 0) or jump(v):
            return unavailable('cumulative volume reset/nonmonotonic trace or discontinuity; validated segmentation required')
        volume = float(v[-1] - v[0])
    elif inputs.agro_flow_source:
        trace.update(method='direct volumetric flow integration', source=inputs.agro_flow_source,
                     source_provenance=inputs.agro_flow_provenance)
        if not inputs.agro_flow_provenance or flow is None or flow_unit not in {'L/s', 'l/s', 'L/min', 'l/min', 'L/h', 'l/h'}:
            return unavailable('flow source requires trusted provenance and confirmed volumetric flow units')
        f = np.asarray(flow, dtype=float)
        t = np.asarray(time_s, dtype=float) if time_s is not None else np.array([])
        if time_unit != 's' or len(t) != len(d) or len(f) != len(d) or not np.isfinite(t).all() or not np.isfinite(f).all() or np.any(f < 0) or np.any(np.diff(t) <= 0):
            return unavailable('nonfinite/negative flow or invalid time window (strictly increasing seconds required)')
        if jump(f):
            return unavailable('flow discontinuity: verify operating window')
        divisor = 60 if flow_unit.lower().endswith('/min') else 3600 if flow_unit.lower().endswith('/h') else 1
        volume = float(np.sum((f[1:] + f[:-1]) / 2 * np.diff(t)) / divisor)
    else:
        trace.update(method=mass_trace['policy'], source=source,
                     formula=mass_trace['formula'] + '; volume_L = (mass_first - mass_last) / density_kg_per_L')
        if mass_problem:
            return unavailable(mass_problem)
        if mass is None or len(mass) != len(d):
            return unavailable('remaining-mass channel missing or insufficient data')
        if np.any(np.diff(mass) > 0):
            return unavailable('refill/loading event or increasing load: validated per-cycle segmentation required')
        if jump(mass):
            return unavailable('remaining-load discontinuity: verify operating window')
        discharged_mass = float(mass[0] - mass[-1])
        mass_rate = discharged_mass / travelled
        trace.update(discharged_mass_kg=discharged_mass, mass_rate_kg_per_m=mass_rate if np.isfinite(mass_rate) else None)
        if inputs.density_kg_per_L is None:
            return unavailable('density required (kg/L) for mass-to-volume conversion')
        volume = discharged_mass / inputs.density_kg_per_L
    if not np.isfinite(volume) or volume < 0 or not np.isfinite(volume / travelled):
        return unavailable('nonfinite or negative discharge volume/rate')
    trace.update(discharged_volume_L=volume, status='available')
    rate = ReportingMetric(**{**asdict(rate), 'value': volume / travelled, 'source': trace['source'],
                              'status': capacity.status if trace['method'] in {'force_inference', 'direct_mass'} else 'confirmed',
                              'interpretation_status': capacity.interpretation_status if trace['method'] in {'force_inference', 'direct_mass'} else 'configured_volume_source',
                              'quality_status': 'valid', 'provenance': trace.get('source_provenance') or capacity.provenance})
    return AgroReport(capacity, rate, trace, mass_values=mass)


def calculate_reporting_kpis(statistics_result, metadata):
    """Resolve exact raw channels; do not mistake the legacy mass channel for liters."""
    dataset = statistics_result.dataset
    resolved = statistics_result.resolution.resolved

    def semantic(name):
        item = resolved.get(name)
        if item is None:
            return None, None
        return dataset.values[:, dataset.channel_index(item.channel.channel_id)], item.effective_unit

    def direct(source):
        if 'fuel' in source.casefold():
            return None, None  # Fuel consumption is never an agrochemical volume source.
        matches = [c for c in dataset.channels if c.source_name == source]
        if len(matches) != 1:
            return None, None
        c = matches[0]
        return dataset.values[:, dataset.channel_index(c.channel_id)], c.unit

    inputs = metadata.reporting_inputs
    d, du = semantic('track_distance')
    t, tu = semantic('track_time')
    source = inputs.agro_mass_source
    source_item = resolved.get(source) if source else None
    if source_item is not None:
        selected, selected_unit = semantic(source)
    else:
        selected, selected_unit = direct(source) if source else (None, None)
    inferred = inputs.agro_mass_source_policy == 'force_inference'
    v, vu = direct(inputs.agro_volume_source) if inputs.agro_volume_source else (None, None)
    f, f_unit = direct(inputs.agro_flow_source) if inputs.agro_flow_source else (None, None)
    report = calculate_agro_kpis(d, du, inputs=inputs,
        force_N=selected if inferred else None, force_unit=selected_unit if inferred else None,
        mass_kg=None if inferred else selected, mass_unit=None if inferred else selected_unit,
        volume_L=v, volume_unit=vu, flow=f, flow_unit=f_unit, time_s=t, time_unit=tu,
        source_name=source_item.channel.source_name if source_item else source)
    # The endpoint calculation itself is unchanged; only its source is unified.
    metric, net_trace = calculate_net_mass_discharge_rate(report.mass_values, 'kg', d, du)
    mass_trace = report.trace['mass_source']
    metric = ReportingMetric(**{**asdict(metric), 'source': report.capacity.source,
        'status': report.capacity.status if metric.value is not None else 'unavailable',
        'interpretation_status': report.capacity.interpretation_status,
        'quality_status': report.capacity.quality_status if metric.value is not None else 'unavailable',
        'provenance': report.capacity.provenance,
        'definition': metric.definition + '; net mass loss, not cumulative sprayed mass; ' + mass_trace['assumption']})
    if metric.value is None and report.mass_values is None:
        metric = ReportingMetric(**{**asdict(metric), 'reason': report.capacity.reason})
    net_trace.update(mass_expression=mass_trace['formula'], source=report.capacity.source,
                     interpretation_status=mass_trace['interpretation_status'], provenance=mass_trace['provenance'])
    unconfirmed = inferred and mass_trace['interpretation_status'] != 'user_confirmed'
    resolved_source_name = source_item.channel.source_name if source_item else source
    net_trace['warnings'].append(Diagnostic(mass_trace['assumption']
        + f'; source: {resolved_source_name} [{selected_unit}]; formula: {mass_trace["formula"]}',
        code='force_interpretation_unconfirmed' if unconfirmed else 'simulation_mass_interpretation',
        severity='warning' if unconfirmed else 'info', user_visible=unconfirmed))
    report = AgroReport(report.capacity, report.rate, report.trace, metric, report.mass_values)
    report.trace['net_mass_discharge'] = net_trace
    report.trace['warnings'].extend(net_trace['warnings'])
    matches = ([source_item.channel] if source_item else [c for c in dataset.channels if c.source_name == source])
    report.trace['mass_source'].update(
        raw_header=matches[0].source_name if len(matches) == 1 else source,
        raw_unit=matches[0].unit if len(matches) == 1 else None,
        effective_unit=selected_unit,
        raw_channel_id=matches[0].channel_id if len(matches) == 1 else None,
        resolution='resolved' if selected is not None else 'missing, ambiguous or unit-incompatible')
    report.trace['analysis_window'] = {'first_sample': 0, 'last_sample': dataset.quality.sample_count - 1,
                                      'time_start': dataset.quality.time_start, 'time_end': dataset.quality.time_end,
                                      'time_unit': dataset.quality.time_unit, 'selection': 'entire report source window'}
    report.trace['distance_source'] = resolved['track_distance'].channel.source_name if 'track_distance' in resolved else None
    legacy = statistics_result.profile.math_by_semantic_name().get('agrochemical_discharge')
    if legacy:
        report.trace['preserved_legacy_mass_expression'] = legacy.expression
    force_resolved = resolved.get('agrochemical_discharge_force')
    report.trace['force_source'] = ({
        'raw_header': force_resolved.channel.source_name,
        'raw_unit': force_resolved.channel.unit,
        'effective_unit': force_resolved.effective_unit,
        'semantic_mapping': force_resolved.definition.semantic_name,
        'physical_interpretation_supported': mass_trace['interpretation_status'] == 'user_confirmed',
        'interpretation_provenance': inputs.agro_mass_provenance,
    } if force_resolved else None)
    for key, source in (('direct_volume', inputs.agro_volume_source), ('direct_flow', inputs.agro_flow_source)):
        if source:
            matches = [c for c in dataset.channels if c.source_name == source]
            report.trace[key + '_resolution'] = {
                'raw_header': source, 'matches': len(matches),
                'raw_unit': matches[0].unit if len(matches) == 1 else None,
                'status': 'resolved' if len(matches) == 1 else 'ambiguous' if matches else 'missing',
            }
    return report


def reporting_metric_rows(result, metadata):
    """One display registry for both renderers; legacy numeric metrics stay intact."""
    rows = []
    inputs = metadata.reporting_inputs
    if result.profile.metadata.powertrain in {'electric', 'hybrid'} or inputs.battery_mass_kg is not None:
        rows.append({'metric_id': 'battery_mass_kg', 'label': 'Battery Pack Mass', 'ppt_label': 'BATTERY PACK MASS',
                     'value': inputs.battery_mass_kg, 'unit': 'kg', 'display': inputs.battery_card_display,
                     'status': inputs.battery_mass_status, 'reason': 'Battery pack mass not specified' if inputs.battery_mass_kg is None else '',
                     'source': inputs.battery_mass_provenance or 'Not specified', 'definition': 'Reporting metadata only; no vehicle/energy recalculation'})
    if result.reporting_agro is not None:
        metrics = (result.reporting_agro.capacity, result.reporting_agro.rate, result.reporting_agro.mass_rate,
                   result.reporting_agro.net_mass_rate)
        for metric in metrics:
            if metric is None:
                continue
            display = 'N/A'
            if metric.value is not None:
                display = (f'{metric.value:,.2f} {metric.unit}' if metric is result.reporting_agro.capacity
                           else f'{metric.value:.2f} {metric.unit}' if metric is result.reporting_agro.net_mass_rate
                           else f'{metric.value:.6g} {metric.unit}')
            rows.append({**asdict(metric), 'ppt_label': metric.label.upper(),
                         'display': display,
                         'number_format': ('#,##0.00' if metric is result.reporting_agro.capacity or metric is result.reporting_agro.net_mass_rate else _rate_number_format(metric.value)),
                         'status': metric.status})
    return rows


def _rate_number_format(value):
    """Excel precision corresponding to the shared six-significant-digit display."""
    if value is None:
        return 'General'
    if 'e' in f'{value:.6g}':
        return '0.00000E+00'
    places = max(0, 5 - math.floor(math.log10(abs(value)))) if value else 5
    return '0' + ('.' + '0' * places if places else '')
