"""Validační chyba nesmí odpojit widgety ani změnit zadaný investiční scénář."""
import json
import pytest
from test_investment_navigation_reference import audited_app, set_case, REFERENCE, CASES

KEYS = ('inv_machine_1', 'inv_machine_2', 'inv_scenario', 'inv_fixed_cost',
        'inv_expected_days', 'inv_n', 'inv_budget')


def values(at):
    widgets = {w.key: w.value for kind in (at.selectbox, at.number_input, at.slider) for w in kind}
    return {key: widgets[key] for key in KEYS}


@pytest.mark.parametrize('case', sorted(CASES))
def test_all_seven_inputs_remain_mounted_during_validation_and_survive_repair(tmp_path, case):
    import navigace
    at = audited_app(tmp_path)
    expected = set_case(at, case)
    assert expected == json.loads(REFERENCE.read_text())[case]
    original = values(at)
    at.selectbox(key='inv_machine_2').set_value(original['inv_machine_1']).run()
    assert not at.exception
    assert values(at) == {**original, 'inv_machine_2': original['inv_machine_1']}
    assert any('různé' in w.value for w in at.warning)
    assert not at.get('plotly_chart')
    assert not at.get('download_button')
    assert not any('class="de-hero ' in m.value for m in at.main.markdown)
    for slug in navigace.INVESTICE_VIEWS:
        at.query_params['pohled'] = slug; at.run()
        assert not at.exception
        assert values(at) == {**original, 'inv_machine_2': original['inv_machine_1']}
        assert not at.get('plotly_chart') and not at.get('download_button')
    at.selectbox(key='inv_machine_2').set_value(original['inv_machine_2']).run()
    assert not at.exception
    assert values(at) == original
    assert at.session_state['_investment_audit'] == expected
    for slug in navigace.INVESTICE_VIEWS:
        at.query_params['pohled'] = slug; at.run()
        assert not at.exception
        assert values(at) == original
        assert at.session_state['_investment_audit'] == expected
