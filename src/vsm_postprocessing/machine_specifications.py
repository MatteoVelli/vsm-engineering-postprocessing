"""Machine specifications for reporting only; never a numerical model input.

Only a matched configuration can supply specifications. Filename parameters are
claims requiring confirmation, and template text is never an authority.
"""
import math
import re
from collections.abc import Mapping

from .errors import ConfigurationError


SPECIFICATIONS = {
    'gvw_kg': ('GVW', 'kg'),
    'target_speed_kph': ('Target speed', 'kph'),
    'battery': ('Battery specification', ''),
    'motor': ('Motor specification', ''),
    'generator': ('Generator specification', ''),
    'engine': ('Engine specification', ''),
}


def validate_specifications(raw):
    if not isinstance(raw, Mapping) or set(raw) - SPECIFICATIONS.keys():
        raise ConfigurationError('specifications must contain only supported machine specification keys')
    for key, item in raw.items():
        if not isinstance(item, Mapping) or set(item) != {'value', 'status', 'provenance'}:
            raise ConfigurationError('Each specification requires value, status and provenance')
        if item['status'] not in {'confirmed', 'user_provided', 'estimated'}:
            raise ConfigurationError('Specification status must be confirmed, user_provided or estimated')
        if not isinstance(item['provenance'], str) or not item['provenance'].strip():
            raise ConfigurationError('Machine specifications require provenance')
        value = item['value']
        if key in {'gvw_kg', 'target_speed_kph'}:
            if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value) or value <= 0:
                raise ConfigurationError('Mass and target speed specifications must be positive and finite')
        elif not isinstance(value, str) or not value.strip():
            raise ConfigurationError('Component specifications must be nonempty text with explicit units where applicable')


def resolve_machine_specifications(filename, configuration=None):
    records = {key: {'label': label, 'value': None, 'unit': unit,
                     'status': 'unavailable', 'provenance': None,
                     'reason': 'Not specified; requires engineering confirmation'}
               for key, (label, unit) in SPECIFICATIONS.items()}
    # A filename weight may mean vehicle mass, GVW or payload. Do not decide
    # that physical meaning here, or mistake a later chemical-load token for GVW.
    prefix = re.split(r'[_\s-](?:electric|hybrid|diesel)(?:[_\s.-]|$)', filename, flags=re.I)[0]
    claims = {
        'gvw_kg': re.findall(r'(?:^|[_\s])([0-9]+(?:[.,][0-9]+)?)kg(?=[_\s.]|$)', prefix, re.I),
        'target_speed_kph': re.findall(r'(?:^|[_\s])([0-9]+(?:[.,][0-9]+)?)kph(?=[_\s.]|$)', filename, re.I),
    }
    for key, values in claims.items():
        if len(values) == 1:
            records[key].update(value=float(values[0].replace(',', '.')),
                                status='filename_claim', provenance=filename,
                                reason='Filename claim; physical meaning and specification require confirmation')
    configured = (configuration or {}).get('specifications', {})
    validate_specifications(configured)
    for key, item in configured.items():
        records[key].update(item, reason='')
    return records


def specification_display(metadata, key, *, compact=False):
    item = metadata.machine_specifications.get(key)
    if not item or item['value'] is None:
        return 'Not specified'
    value = item['value']
    text = f'{value:,.6g}' if isinstance(value, (int, float)) else value
    text += (' ' + item['unit']) if item['unit'] else ''
    status = item['status']
    if status == 'filename_claim':
        text += ' (TBC)' if compact else ' (filename claim; requires confirmation)'
    elif status != 'confirmed':
        text += f' ({status})'
    return text


def machine_metadata_pills(metadata):
    mass = metadata.machine_specifications.get('gvw_kg', {})
    mass_label = 'Mass' if mass.get('status') == 'filename_claim' else 'GVW'
    components = ('engine', 'generator') if metadata.powertrain_name == 'Diesel' else ('battery', 'motor')
    return (f'{mass_label} {specification_display(metadata, "gvw_kg", compact=True)}',
            f'{metadata.powertrain_name} powertrain',
            f'Target {specification_display(metadata, "target_speed_kph", compact=True)}',
            *(key.title() + ' ' + specification_display(metadata, key, compact=True) for key in components))


def machine_overview_lines(metadata):
    if metadata.powertrain_name == 'Diesel':
        return (f'Machine: {metadata.machine_name}; powertrain: Diesel.',
                'Mass / GVW: ' + specification_display(metadata, 'gvw_kg', compact=True)
                + '; target speed: ' + specification_display(metadata, 'target_speed_kph', compact=True),
                'Engine: ' + specification_display(metadata, 'engine', compact=True),
                'Generator: ' + specification_display(metadata, 'generator', compact=True))
    last = ('Generator: ' + specification_display(metadata, 'generator', compact=True)
            + '; engine: ' + specification_display(metadata, 'engine', compact=True)
            if metadata.powertrain_name == 'Hybrid' else 'Missing specifications require engineering confirmation.')
    return (f'Machine: {metadata.machine_name}; powertrain: {metadata.powertrain_name}.',
            'Mass / GVW: ' + specification_display(metadata, 'gvw_kg', compact=True)
            + '; target speed: ' + specification_display(metadata, 'target_speed_kph', compact=True),
            'Battery: ' + specification_display(metadata, 'battery', compact=True)
            + '; motor: ' + specification_display(metadata, 'motor', compact=True),
            last)
