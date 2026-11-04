"""Optional shared peak/continuous motor specifications and live assessment UI."""
from dataclasses import dataclass
import hashlib

from .continuous_torque import CONTINUOUS, ContinuousTorqueCurve, POLICY, blank
from .continuous_torque_reporting import ContinuousTorqueInputs, assess_continuous_torque
from .errors import ConfigurationError, VSMPostProcessingError
from .motor_performance import MotorPerformanceCurve, SPEED, TORQUE, ASSUMPTION


@dataclass(frozen=True)
class MotorReportingInputs:
    peak: MotorPerformanceCurve | None = None
    continuous: ContinuousTorqueInputs | None = None
    valid: bool = True


def render_motor_performance_inputs(st, source_scope):
    prefix = 'motor_performance_' + hashlib.sha256(source_scope.encode()).hexdigest()[:16] + '_'
    revision = st.session_state.get(prefix+'revision', 0)
    key = lambda name: prefix+str(revision)+'_'+name
    seed = st.session_state.get(prefix+'seed', {})
    with st.expander('Motor Performance Curves (Optional)'):
        previous = st.session_state.get('motor_specifications_for_explicit_reuse')
        if previous and previous['scope'] != source_scope:
            st.caption('This source/profile has no inherited motor specification. Reuse is an explicit action; manual durations are never copied.')
            if st.button('Reuse previous motor specifications for this source', key=key('reuse')):
                st.session_state[prefix+'seed'] = {k:v for k,v in previous.items() if k != 'scope'}
                st.session_state[prefix+'revision'] = revision+1
                st.rerun()
        enabled = st.checkbox('Include one shared motor capability curve', value=bool(seed), key=key('enabled'))
        st.caption('One reference curve shared by all four motors. Peak overlays and continuous-duty analysis are independent reporting features.')
        st.caption('Supply both curves to include Motor Torque and Power in Excel and PowerPoint. Shaft power uses only motor RPM and torque; each curve needs a nonnegative, strictly increasing speed grid.')
        if enabled and st.button('Clear motor curve', key=key('clear')):
            st.session_state.pop(prefix+'seed',None)
            st.session_state[prefix+'revision'] = revision+1
            st.rerun()
        st.caption(':blue[Editable: motor RPM, peak torque, continuous torque, ratio and efficiency.] Calculated outputs are read-only.')
        ratio = st.number_input('Final-drive ratio [dimensionless]', min_value=0.0, value=seed.get('ratio'),
            format='%.6f', key=key('ratio'), disabled=not enabled,
            help='Needed for peak/wheel overlays; continuous motor-side assessment needs no gearbox conversion.')
        efficiency = st.number_input('Driveline efficiency [fraction]', min_value=0.0, max_value=1.0,
            value=seed.get('efficiency'), format='%.6f', key=key('efficiency'), disabled=not enabled)
        initial = seed.get('rows', [{SPEED:float('nan'), TORQUE:float('nan'), CONTINUOUS:float('nan')} for _ in range(2)])
        initial = [{name:float('nan') if blank(row.get(name)) else row[name] for name in (SPEED,TORQUE,CONTINUOUS)} for row in initial]
        rows = st.data_editor(initial,num_rows='dynamic',hide_index=True,key=key('table'),disabled=not enabled,
            column_config={SPEED:st.column_config.NumberColumn(SPEED,format='%.6f'),
                           TORQUE:st.column_config.NumberColumn(TORQUE,min_value=0.0,format='%.6f'),
                           CONTINUOUS:st.column_config.NumberColumn(CONTINUOUS,min_value=0.0,format='%.6f')},width='stretch')
        peak = None; curve = None
        if enabled:
            st.caption(ASSUMPTION)
            peak_rows = [r for r in rows if not blank(r.get(TORQUE))]
            if peak_rows:
                try:
                    peak = MotorPerformanceCurve.from_rows(peak_rows,ratio,efficiency)
                    st.dataframe(peak.calculated_rows(),hide_index=True,width='stretch')
                except ConfigurationError as exc:
                    st.error('Peak overlays unavailable: '+str(exc))
            if any(not blank(r.get(CONTINUOUS)) for r in rows):
                try:
                    curve = ContinuousTorqueCurve.from_rows(rows)
                    st.caption('Continuous analysis uses motor-side measurements; ratio and efficiency do not affect its result. Continuous RPM points are sorted; duplicate RPM is rejected. Blank rating cells belong only to the other curve.')
                except ConfigurationError as exc:
                    st.error('Continuous analysis unavailable: '+str(exc))
        fallback = st.checkbox('Enter externally calculated maximum duration (optional fallback)', key=key('manual_enabled'))
        manual = st.number_input('Maximum Segment Time Above Continuous Torque [s] — manual fallback',
            min_value=0.0,value=None,key=key('manual_duration'),disabled=not fallback,
            help='Empty by default. Automatic assessment takes precedence. This value is never copied to another source.')
        provenance = st.text_input('External calculation source / provenance',value='',key=key('manual_provenance'),disabled=not fallback)
        manual_inputs = None
        if fallback and manual is not None:
            try:
                manual_inputs = ContinuousTorqueInputs(manual_duration_s=manual,manual_provenance=provenance)
            except ConfigurationError as exc:
                st.error(str(exc))
        continuous = None
        if curve is not None or manual_inputs is not None:
            wheel = ratio is not None and ratio > 0 and efficiency is not None and 0 < efficiency <= 1
            try:
                continuous = ContinuousTorqueInputs(curve,manual_inputs.manual_duration_s if manual_inputs else None,
                    manual_inputs.manual_provenance if manual_inputs else None,
                    ratio if wheel else None,efficiency if wheel else None)
            except ConfigurationError as exc:
                st.error('Continuous wheel overlays unavailable: '+str(exc))
                continuous = ContinuousTorqueInputs(curve,manual_inputs.manual_duration_s if manual_inputs else None,
                    manual_inputs.manual_provenance if manual_inputs else None)
        if enabled and (peak is not None or curve is not None):
            st.session_state['motor_specifications_for_explicit_reuse'] = {'scope':source_scope,'rows':rows,'ratio':ratio,'efficiency':efficiency}
        valid = not (enabled or fallback) or peak is not None or continuous is not None
        if enabled and not peak and not continuous:
            st.caption('Enter at least two RPM/torque pairs for either curve, or clear/disable the optional specification. Continuous-only analysis needs neither peak torque nor drivetrain parameters.')
        return MotorReportingInputs(peak,continuous,valid)


def render_continuous_torque_results(st, source_path, profile_path, inputs, source_sha256, import_options):
    """Cache imported samples by source content and recalculate for current specs/profile."""
    if inputs is None:
        return None
    from .importer import load_data_file
    from .report_profile import load_reporting_profile
    profile = load_reporting_profile(profile_path)
    data_key = (str(source_path.resolve()),source_sha256,repr(import_options))
    analysis_key = (data_key,repr(profile),repr(inputs))
    try:
        if st.session_state.get('continuous_assessment_key') != analysis_key:
            if inputs.curve is not None:
                if st.session_state.get('continuous_dataset_key') != data_key:
                    st.session_state['continuous_dataset'] = load_data_file(source_path,import_options)
                    st.session_state['continuous_dataset_key'] = data_key
                data = st.session_state['continuous_dataset']
            else:
                data = None
            st.session_state['continuous_assessment'] = assess_continuous_torque(data,profile,inputs)
            st.session_state['continuous_assessment_key'] = analysis_key
        assessment = st.session_state['continuous_assessment']
    except VSMPostProcessingError as exc:
        st.info('Continuous torque analysis unavailable: '+str(exc))
        return None
    st.subheader('Continuous Torque Analysis')
    if assessment.duration_s is not None:
        st.metric('Maximum Segment Time Above Continuous Torque [s]',f'{assessment.duration_s:.6g}')
        if assessment.origin == 'automatic':
            st.caption('Automatically calculated from the current VSM source and continuous motor curve. Responsible motor(s): '+(', '.join(assessment.analysis.responsible_motors) or 'None — no positive-duration exceedance.'))
            if inputs.manual_duration_s is not None:
                st.caption('The external manual duration is not used because automatic analysis is available.')
        else:
            st.caption('Manual fallback — external calculation source: '+inputs.manual_provenance)
    else:
        st.info('Continuous torque assessment unavailable; no duration will be exported.')
    if assessment.analysis:
        rows=[]
        for motor in assessment.analysis.motors:
            rows.append({'Motor':motor.motor,'Status':motor.status,
                'Maximum Segment Time Above Continuous Torque [s]':motor.maximum_segment_s,
                'Total Time Above Continuous Torque [s]':motor.total_time_s,
                'Number of Exceedance Episodes':motor.episode_count,
                'Maximum Continuous Torque Utilisation':motor.maximum_utilisation,
                'Flagged samples':motor.flagged_samples,'Unavailable reason':motor.reason or ''})
        st.dataframe(rows,hide_index=True,width='stretch')
        st.caption('Times and episode counts are per motor. A vehicle maximum requires all four motors to be available.')
        with st.expander('Continuous torque calculation policy and provenance'):
            st.write(POLICY)
            st.json(assessment.summary())
    return assessment
