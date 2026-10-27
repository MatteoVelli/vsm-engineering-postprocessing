"""Reporting mass policy: numerical support is independent of machine identity."""
import numpy as np
from .diagnostics import Diagnostic


def resolve_mass_values(inputs, *, force=None, force_unit=None, mass=None, mass_unit=None, source_name=None):
    policy = inputs.agro_mass_source_policy
    provenance = inputs.agro_mass_source_provenance
    inferred = policy == 'force_inference'
    scope = inputs.agro_mass_confirmation_source
    confirmed = bool(inputs.hitch_force_is_remaining_agro_mass and inputs.agro_mass_provenance
                     and (scope is None or scope == source_name))
    trace = {'policy': policy, 'mapping': inputs.agro_mass_source,
             'provenance': provenance, 'status': 'inferred' if inferred else 'simulated' if policy == 'direct_mass' else 'unavailable',
             'interpretation_status': ('user_confirmed' if confirmed else 'assumed') if inferred else (
                                           'configured_simulation_source' if policy == 'direct_mass' else 'unconfigured'),
             'physical_validation_provenance': inputs.agro_mass_provenance,
             'confirmation_source': scope,
             'confirmation_applied': confirmed if inferred else False,
             'formula': '-source_N / 9.81' if inferred else 'source_kg',
             'gravity_m_per_s2': 9.81 if inferred else None,
             'quality_status': 'unavailable', 'warnings': [],
             'assumption': (('User-confirmed simulation tank-load interpretation; equivalent mass is inferred, not a direct tank-mass measurement or independently validated sprayed quantity'
                             if confirmed else 'Configured force is treated as remaining agrochemical load; channel name/trend is not physical proof; not physically validated discharge')
                            if inferred else 'Configured simulation mass; not asserted to be a measured mass or nominal tank capacity')}

    def invalid(reason):
        trace['reason'] = reason
        return None, trace

    if policy == 'unconfigured':
        return invalid('agro mass source policy not configured; select direct_mass or force_inference with provenance')
    raw, unit = (force, force_unit) if inferred else (mass, mass_unit)
    expected = 'N' if inferred else 'kg'
    if raw is None:
        return invalid('configured agro mass source missing, ambiguous or unresolved; no implicit source fallback')
    if (unit != 'N' if inferred else str(unit).casefold() != 'kg'):
        return invalid('agro mass source unit mismatch (' + expected + ' required)')
    values = np.asarray(raw, dtype=float)
    if values.ndim != 1 or len(values) < 2 or not np.isfinite(values).all():
        return invalid('nonfinite or insufficient remaining-mass data; at least two finite aligned samples required')
    values = -values / 9.81 if inferred else values.copy()
    if not np.isfinite(values).all():
        return invalid('nonfinite mass after force conversion')
    if inferred and np.all(values == 0):
        return invalid('all-zero hitch force: no inferable payload (empty tank and inactive/unmapped force cannot be distinguished)')
    if np.max(values) < 0 or (inferred and np.max(values) == 0 and np.any(values < 0)):
        return invalid('nonphysical negative inferred contents; verify source sign and load interpretation')
    trace.update(quality_status='valid', sample_count=len(values),
                 initial_remaining_mass_kg=float(values[0]), final_remaining_mass_kg=float(values[-1]),
                 observed_maximum_mass_kg=float(values.max()), minimum_mass_kg=float(values.min()))
    if np.any(values < 0):
        trace['quality_status'] = 'questionable'
        trace['warnings'].append('nonphysical negative remaining contents; positive observed maximum retained with a quality warning; discharge unavailable')
        if inferred and np.any(values > 0):
            trace['warnings'].append(Diagnostic(
                'Hitch-force sign reversal produces negative equivalent remaining mass; loading direction does not establish valid remaining contents. Verify the simulation phases.',
                code='force_sign_change'))
    if np.any(np.diff(values) > 0):
        trace['warnings'].append('Remaining mass increases: possible refill or dynamic force reactions; a decreasing trace is not physical validation either')
    if np.all(np.diff(values) == 0):
        trace['warnings'].append('Constant mass/force trace; may be static load or inactive discharge, not evidence of spraying')
    active = np.abs(np.diff(values))
    active = active[active > 0]
    if len(active) >= 5 and np.max(active) > 20 * np.median(active):
        trace['quality_status'] = 'questionable'
        trace['warnings'].append('Remaining-load discontinuity: net endpoint loss is not validated spraying or per-cycle discharge')
    return values, trace
