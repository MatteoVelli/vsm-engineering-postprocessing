"""Automatic specifications, physical confidence and latest reference regressions."""
from pathlib import Path
import numpy as np
import pytest
from vsm_postprocessing.report_profile import load_reporting_profile
from vsm_postprocessing.report_metadata import resolve_report_metadata
from vsm_postprocessing.importer import load_data_file, ImportOptions
from vsm_postprocessing.profile_statistics import calculate_profile_statistics
from vsm_postprocessing.agro_reporting import calculate_reporting_kpis, calculate_agro_kpis, reporting_metric_rows
from vsm_postprocessing.reporting_inputs import ReportingInputs
from vsm_postprocessing.errors import ConfigurationError

BASE = Path('config/report_profiles')

def profile(kind):
    return load_reporting_profile(BASE / ('full_size_sprayer_diesel.yaml' if kind == 'diesel' else 'robosprayer_' + kind + '.yaml'))

@pytest.mark.parametrize('filename,kind,mass,status', [
    ('RoboSprayer_3500Kg_Electric_UoL_04.csv','electric',None,'not_provided'),
    ('RoboSprayer_V5_3500Kg_Electric_99.csv','electric',None,'not_provided'),
    ('Unrelated_3500Kg_Electric.csv','electric',None,'not_provided'),
    ('RoboSprayer_V4_3500Kg_Electric.csv','electric',None,'not_provided'),
    ('RoboSprayer_8500Kg_Electric.csv','electric',None,'not_provided'),
    ('Full_Size_Sprayer_8500Kg_Electric.csv','electric',None,'not_provided'),
    ('Full_Size_Sprayer_8500Kg_Hybrid.csv','hybrid',None,'not_provided'),
    ('Full_Size_Sprayer_8500Kg_Diesel.csv','diesel',None,'not_provided'),
    ('Full_Size_Sprayer_8500Kg_Diesel.csv','electric',None,'not_provided'),
    ('Sprayer_Caiman_SP_9300Kg_Hybrid.csv','hybrid',None,'not_provided'),
])
def test_automatic_specifications_require_model_configuration_and_powertrain(filename,kind,mass,status):
    metadata = resolve_report_metadata(filename,profile(kind))
    assert metadata.reporting_inputs.battery_mass_kg == mass
    assert metadata.reporting_inputs.battery_mass_status == status
    if mass:
        assert metadata.reporting_inputs.battery_mass_provenance
        assert metadata.configuration_detection == 'profile model/configuration/powertrain mapping'


def test_diesel_cannot_select_electric_configuration_or_manual_pack():
    with pytest.raises(ConfigurationError):
        resolve_report_metadata('Full_Size_8500Kg_Diesel.csv',profile('diesel'),reporting_inputs={'machine_configuration_id':'full_size_sprayer_8500kg'})
    with pytest.raises(ConfigurationError):
        resolve_report_metadata('Full_Size_8500Kg_Diesel.csv',profile('diesel'),reporting_inputs={'battery_mass_kg':1700,'battery_mass_status':'estimated','battery_mass_provenance':'hypothetical'})

@pytest.mark.parametrize('kind,pattern,samples,expected_rate', [
    ('electric','*3500Kg_Electric_UoL*csv',3954,0.04533818805785114),
    ('diesel','*8500Kg_Diesel_1000m*csv',232,0.4998508170252672),
    ('hybrid','*Caiman_SP_9300Kg_Hybrid_12x1Km*csv',1770,None),
])
def test_latest_sources_without_manual_confirmation(kind,pattern,samples,expected_rate):
    files = list(Path('reference_files').glob(pattern))
    if not files:
        pytest.skip('Private simulation reference not installed')
    source = files[0]
    selected = profile(kind)
    result = calculate_profile_statistics(load_data_file(source,ImportOptions(strict=False)),selected)
    assert result.dataset.quality.sample_count == samples
    metadata = resolve_report_metadata(source,selected)
    assert metadata.reporting_inputs.hitch_force_is_remaining_agro_mass
    assert '27 October 2026' in metadata.reporting_inputs.agro_mass_provenance
    assert metadata.reporting_inputs.agro_mass_source_policy == 'force_inference'
    agro = calculate_reporting_kpis(result,metadata)
    assert agro.capacity.value > 0 and agro.capacity.status == 'inferred'
    assert 'not nominal' in agro.capacity.definition
    assert agro.rate.value is None and agro.rate.unit == 'L/m'
    if expected_rate is None:
        assert agro.mass_rate.value is None and 'negative' in agro.rate.reason
    else:
        assert agro.mass_rate.value == pytest.approx(expected_rate,rel=1e-12)
        assert agro.mass_rate.status == 'inferred' and agro.mass_rate.unit == 'kg/m'
        assert 'density required' in agro.rate.reason
    if kind == 'electric':
        assert agro.capacity.value == next(s.value for s in result.statistics if s.definition.statistic_id == 'agrochemical_discharge_max')


def test_multiple_spraying_phases_use_full_distance_not_only_active_samples():
    result = calculate_agro_kpis([100,200,300,400,500], 'm',force_N=-9.81*np.array([500,450,450,400,400]),force_unit='N',inputs=ReportingInputs(agro_mass_source_policy='force_inference', agro_mass_source='Test force', agro_mass_source_provenance='Explicit synthetic mapping',density_kg_per_L=1.25,density_provenance='explicit hypothetical test'))
    assert result.rate.value == .2
    assert result.mass_rate.value == .25
    assert result.capacity.status == result.rate.status == 'inferred'

@pytest.mark.parametrize('distance,mass', [
    ([0,1,2,3,4,5,200],[500,499,498,497,496,495,494]),
    ([0,1,2,3,4,5,6],[500,499,498,497,496,495,100]),
])
def test_discontinuities_reject_mass_and_volume_rate(distance,mass):
    result = calculate_agro_kpis(distance,'m',force_N=-9.81*np.array(mass),force_unit='N',inputs=ReportingInputs(agro_mass_source_policy='force_inference', agro_mass_source='Test force', agro_mass_source_provenance='Explicit synthetic mapping'))
    assert result.mass_rate.value is result.rate.value is None
    assert 'discontinuity' in result.rate.reason


def test_flow_integration_uses_observed_time_and_unit_conversion():
    result = calculate_agro_kpis([0,100,200],'m',time_s=[10,70,130],flow=[60,120,60],flow_unit='L/min',inputs=ReportingInputs(agro_flow_source='Chemical flow',agro_flow_provenance='hypothetical calibrated volume flow'))
    assert result.rate.value == .9
    assert result.rate.status == 'confirmed'


def test_display_name_override_does_not_erase_verified_source_identity():
    metadata = resolve_report_metadata('RoboSprayer_3500Kg_Electric.csv',profile('electric'),machine_name_override='Sergio demonstration')
    assert metadata.machine_name == 'Sergio demonstration'
    assert metadata.reporting_inputs.battery_mass_kg is None


def test_profile_configuration_validation(tmp_path):
    import yaml
    for invalid in ({'configuration_id': 'missing_fields'}, {'configuration_id':'x','machine_names':['X'],'identity_tokens':['3500kg'],'powertrains':['electric'],'reporting_inputs':{'battery_mass_kg':-1}}):
        path = tmp_path / 'invalid.yaml'
        path.write_text(yaml.safe_dump({'version':1,'profile':{'profile_id':'test','name':'Test','machine_configurations':[invalid]},'channels':{'raw':[]}}))
        with pytest.raises(ConfigurationError):
            load_reporting_profile(path)


@pytest.mark.parametrize('value,expected', [(0.04533818805785114,'0.0000000'),(0.4998508170252672,'0.000000'),(0,'0.00000'),(None,'General'),(1e-8,'0.00000E+00'),(1e7,'0.00000E+00')])
def test_excel_rate_precision_matches_shared_significant_digits(value,expected):
    from vsm_postprocessing.agro_reporting import _rate_number_format
    assert _rate_number_format(value) == expected
