"""Presentation aggregation and compact native cards for existing analysis."""
from pptx.util import Inches


def continuous_torque_cards(assessment):
    if assessment is None or assessment.duration_s is None:
        return []
    definitions = [('duration', 'MAX CONSECUTIVE\nEXCEEDANCE', 'maximum_segment_s', 's'),
                   ('accumulated', 'TOTAL TIME ABOVE LIMIT\nPER MOTOR', 'total_time_s', 's'),
                   ('episodes', 'EXCEEDANCE EPISODES\nPER MOTOR', 'episode_count', ''),
                   ('utilisation', 'MAX CONTINUOUS\nUTILISATION', 'maximum_utilisation', '%')]
    if assessment.origin != 'automatic':
        return [dict(key='duration', label=definitions[0][1], value=assessment.duration_s,
                     display=f'{assessment.duration_s:.6g} s', origin=assessment.origin,
                     responsible_motors=[], provenance=assessment.inputs.manual_provenance)]
    cards = []
    for key, label, attribute, unit in definitions:
        motors = assessment.analysis.motors
        maximum = max(getattr(m, attribute) for m in motors)
        value = maximum * 100 if unit == '%' else maximum
        cards.append(dict(key=key, label=label, value=value,
                          display=f'{value:.1f}%' if unit == '%' else f'{value:.6g}' + (f' {unit}' if unit else ''),
                          origin='automatic', aggregation=f'maximum per motor of {attribute}; never a sum across motors',
                          responsible_motors=[m.motor for m in motors if getattr(m, attribute) == maximum]))
    return cards


def add_continuous_torque_kpi(slide, assessment):
    from .profile_powerpoint_report_engine import (
        _add_custom_kpi_cards, _add_text_box, _ASTAUTO_MUTED,
        _V7_KPI_CARD_LEFT, _V7_KPI_CARD_WIDTH, _V7_KPI_CARD_GAP,
    )
    cards = continuous_torque_cards(assessment)
    if cards:
        span = 6 * _V7_KPI_CARD_WIDTH + 5 * _V7_KPI_CARD_GAP
        width = (span - 3 * _V7_KPI_CARD_GAP) / 4
        left = _V7_KPI_CARD_LEFT + ((span - width) / 2 if len(cards) == 1 else 0)
        shapes = _add_custom_kpi_cards(
            slide, {}, tuple((c['label'], c['key'], '') for c in cards),
            card_width=width, card_left=left, display_values={c['key']: c['display'] for c in cards},
        )
        for record, group in zip(cards, shapes):
            for part, shape in zip(('card', 'indicator', 'label', 'value'), group):
                shape.name = f'Continuous torque {record["key"]} {part}'
        caption = ('Automatic analysis | Per-motor maxima above continuous torque; motor identities and ties in notes.'
                   if assessment.origin == 'automatic' else 'Manual duration | External calculation; provenance recorded in slide notes.')
        source = _add_text_box(slide, caption, Inches(.6), Inches(3.08), Inches(12.13), Inches(.25),
                               font_size=9, color=_ASTAUTO_MUTED, font_name='Calibri')
        source.name = 'Continuous torque duration provenance'
    return cards
