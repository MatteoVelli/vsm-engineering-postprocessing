"""Profile identifiers must survive Streamlit's formatted-label round trip."""
from dataclasses import asdict, replace
from pathlib import Path

import pytest
from streamlit.testing.v1 import AppTest

from vsm_postprocessing import ui_app
from vsm_postprocessing.errors import ConfigurationError
from vsm_postprocessing.report_profile import load_reporting_profile
from vsm_postprocessing.report_selection import ProfileReportSelection, select_profile
from vsm_postprocessing.ui_config import get_reporting_profile_definition


@pytest.mark.parametrize('kind', ['electric', 'hybrid', 'diesel'])
def test_profile_widgets_preserve_identifiers_across_reruns_and_edits(kind, monkeypatch):
    name = 'full_size_sprayer_diesel' if kind == 'diesel' else 'robosprayer_' + kind
    definition = get_reporting_profile_definition(ui_app.PROJECT_ROOT, name)
    definition = replace(definition, profile_path=definition.profile_path.parent / 'sergio_2026_10' / definition.profile_path.name)
    profile = load_reporting_profile(definition.profile_path)
    # Plot titles, like channel report names, are presentation text and need
    # not be unique in a valid custom reporting profile.
    profile = replace(profile, plots=tuple(replace(p, title='Shared title') for p in profile.plots))
    monkeypatch.setattr(ui_app, 'load_reporting_profile', lambda _: profile)
    monkeypatch.setattr(ui_app, '_selection_test_definition', definition, raising=False)
    code = '''
import streamlit as st
from dataclasses import asdict
from types import SimpleNamespace
from vsm_postprocessing import ui_app
inspection = SimpleNamespace(channels=[], quality=SimpleNamespace(sample_count=4))
_, selection, _ = ui_app._render_profile_report_options(st, ui_app._selection_test_definition, inspection)
st.session_state['selected'] = asdict(selection)
'''
    at = AppTest.from_string(code).run()
    expected = asdict(ProfileReportSelection(**at.session_state['selected']))
    for _ in range(2):
        at.run()
        assert not at.exception
        assert at.session_state['selected'] == expected
    select_profile(profile, ProfileReportSelection(**at.session_state['selected']))
    channels = next(w for w in at.multiselect if w.label == 'Channels exported to Excel')
    # Keep the FR channel while removing its identically named FL sibling.
    removed = 'driveshaft_torque_fl' if kind == 'diesel' else 'suspension_damper_force_fl'
    retained = 'driveshaft_torque_fr' if kind == 'diesel' else 'suspension_damper_force_fr'
    chosen = [n for n in expected['export_channels'] if n != removed]
    channels.set_value(chosen).run()
    at.run()
    assert not at.exception
    assert at.session_state['selected']['export_channels'] == tuple(chosen)
    assert retained in at.session_state['selected']['export_channels']
    # Empty choices are intentional and must stay empty, not reset to defaults.
    next(w for w in at.multiselect if w.label == 'Channels exported to Excel').set_value([]).run()
    at.run()
    assert at.session_state['selected']['export_channels'] == ()


@pytest.mark.parametrize('requested,detail', [
    (('track_time', 'track_time'), "duplicate identifiers: ['track_time']"),
    (('unknown_channel',), "unknown identifiers: ['unknown_channel']"),
    (('unknown_channel', 'unknown_channel'), "duplicate identifiers: ['unknown_channel']"),
])
def test_invalid_selection_reports_responsible_identifiers(requested, detail):
    profile = load_reporting_profile(Path('config/report_profiles/sergio_2026_10/robosprayer_electric.yaml'))
    with pytest.raises(ConfigurationError) as error:
        select_profile(profile, ProfileReportSelection(export_channels=requested))
    assert 'Invalid export channels selection' in str(error.value)
    assert detail in str(error.value)
