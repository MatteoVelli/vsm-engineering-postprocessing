import numpy as np
import pytest

from vsm_postprocessing.continuous_torque import (
    CONTINUOUS, ContinuousTorqueCurve, analyse_motor, analyse_dataset,
)
from vsm_postprocessing.errors import ConfigurationError
from vsm_postprocessing.motor_performance import SPEED


def constant(limit=10):
    return ContinuousTorqueCurve((0, 1000), (limit, limit))


def test_interpolation_sign_symmetry_strict_comparison_and_independent_peak_data():
    curve = ContinuousTorqueCurve.from_rows([{SPEED:100, CONTINUOUS:5}, {SPEED:0,CONTINUOUS:10},
                                            {SPEED:200,CONTINUOUS:None}])
    result = analyse_motor([0,.2,.8,2], [50,-50,0,100], [8,-8,10,5], curve)
    assert result.limits_nm == (7.5,7.5,10,5)
    assert result.flags == (True,True,False,False)
    assert result.maximum_segment_s == pytest.approx(.8)
    assert result.total_time_s == pytest.approx(.8)
    assert result.episode_count == 1


@pytest.mark.parametrize('time,flags,duration,total,count', [
    ([0,1,2,3],[True,True,False,True],2,2,2),
    ([0,.1,.4,2],[True,True,False,True],.4,.4,2),
    ([10,12,15,21],[False,True,True,True],9,9,1),
    ([0,3,7],[True,True,True],7,7,1),
    ([0,3,7],[False,False,True],0,0,1),
    ([0,3,7],[False,False,False],0,0,0),
])
def test_observed_intervals_and_final_sample_are_independent_of_sample_frequency(time,flags,duration,total,count):
    r=analyse_motor(time,[100]*len(time),[20 if f else 0 for f in flags],constant())
    assert r.status == 'available'
    assert r.maximum_segment_s == pytest.approx(duration)
    assert r.total_time_s == pytest.approx(total)
    assert r.episode_count == count
    assert r.flagged_samples == sum(flags)


@pytest.mark.parametrize('time', [[0,0,1],[0,2,1],[0,float('nan'),2],[0,float('inf'),2],[0]])
def test_invalid_timestamps_are_unavailable_never_fabricated_zero(time):
    r=analyse_motor(time,[0]*len(time),[0]*len(time),constant())
    assert r.status == 'unavailable' and r.maximum_segment_s is None
    assert r.total_time_s is None and r.episode_count is None


@pytest.mark.parametrize('time,speed,torque', [(0,0,0),(['bad',1],[0,0],[0,0]),([0,1],[0],[0,0])])
def test_malformed_vectors_are_unavailable(time,speed,torque):
    assert analyse_motor(time,speed,torque,constant()).status == 'unavailable'


@pytest.mark.parametrize('speed,torque', [([0,1001],[0,0]), ([float('nan'),1],[0,0]),
                                        ([0,1],[0,float('inf')])])
def test_missing_or_uncovered_samples_cannot_be_reported_as_zero(speed,torque):
    r=analyse_motor([0,1],speed,torque,constant())
    assert r.status == 'unavailable' and r.unsupported_samples == 1
    assert r.maximum_segment_s is None


@pytest.mark.parametrize('speed,torque', [((0,0),(1,2)), ((-1,2),(1,2)), ((0,1),(1,0)),
                                        ((0,1),(1,-1)), ((0,1),(1,float('nan'))), ((0,),(1,))])
def test_invalid_continuous_curves(speed,torque):
    with pytest.raises(ConfigurationError): ContinuousTorqueCurve(speed,torque)


def test_higher_limits_never_create_more_flagged_samples():
    time=[0,.4,1,3,4,8]
    torque=[0,5,15,-20,40,100]
    results=[analyse_motor(time,[100]*6,torque,constant(v)) for v in (1,10,1000)]
    assert [r.flagged_samples for r in results] == [5,4,0]
    assert results[-1].maximum_segment_s == results[-1].total_time_s == results[-1].episode_count == 0


def test_dataset_adapter_uses_semantic_sources_and_four_independent_motors(tmp_path):
    from test_sergio_october_15 import dataset_for, csv_for, profile_path
    from vsm_postprocessing.importer import load_data_file, ImportOptions
    from vsm_postprocessing.report_profile import load_reporting_profile, resolve_profile
    data=load_data_file(csv_for(dataset_for('electric'),tmp_path/'different.csv'),ImportOptions(strict=False))
    profile=load_reporting_profile(profile_path('electric'))
    resolution=resolve_profile(data,profile)
    for index,torque in enumerate((0,20,30,40),1):
        binding=resolution.resolved[f'electricsystem_em{index}_torque'].channel.channel_id
        data.values[:,data.channel_index(binding)]=torque
    r=analyse_dataset(data,profile,ContinuousTorqueCurve((0,1e9),(10,10)),resolution)
    assert r.available
    assert r.motors[0].episode_count == 0 and all(m.episode_count == 1 for m in r.motors[1:])
    assert r.responsible_motors == ('EM2 / FR','EM3 / RL','EM4 / RR')
    resolution.resolved.pop('electricsystem_em4_torque')
    missing=analyse_dataset(data,profile,constant(),resolution)
    assert not missing.available and missing.maximum_segment_s is None
    assert missing.motors[3].status == 'unavailable'
