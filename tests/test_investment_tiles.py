"""Stále vykreslený editor, skutečné hodnoty souhrnu a dlaždice zachovávají scénář."""
import pytest
import investice as inv
import navigace
from formatting import fmt_czk
from test_investice import machines, _fingerprint
from test_investment_navigation_reference import audited_app, set_case
from test_investment_validation import values


def editor(at):
    return next(e for e in at.main.expander if e.label == 'Upravit investiční scénář')


def test_editor_open_close_and_tiles_preserve_every_input_and_reference(tmp_path):
    at = audited_app(tmp_path)
    expected = set_case(at, 'optimisticky')
    original = values(at)
    assert not editor(at).proto.expanded
    for slug in navigace.INVESTICE_VIEWS:
        at.button(key='inv_view_'+slug).click().run()
        assert at.query_params['pohled'] == slug
        for opened in [True, False, True, False]:
            at.session_state['de_inv_editor'] = opened; at.run()
            assert not at.exception
            assert editor(at).proto.expanded == opened
            assert values(at) == original
            assert at.session_state['_investment_audit'] == expected
            assert at.query_params['pohled'] == slug


@pytest.mark.parametrize('slug', list(navigace.INVESTICE_VIEWS))
def test_edit_from_every_view_recalculates_without_navigation_change(tmp_path, slug):
    at = audited_app(tmp_path); set_case(at, 'optimisticky')
    at.button(key='inv_view_'+slug).click().run()
    at.session_state['de_inv_editor'] = True; at.run()
    at.number_input(key='inv_fixed_cost').set_value(65000).run()
    assert not at.exception
    assert at.query_params['pohled'] == slug
    assert at.number_input(key='inv_fixed_cost').value == 65000
    data = machines().set_index('machine_type')
    lo, hi = inv.scenario_days('Optimistický')
    expected = inv.evaluate(data.loc['Kolový bagr'], data.loc['Silniční válec'], 'Kolový bagr',
                            'Silniční válec', 65000, lo, hi, 25, 2000, seed=42)
    assert at.session_state['_investment_audit']['simulation'] == _fingerprint(expected)
    summary = next(m.value for m in at.main.markdown if 'class="de-inv-summary"' in m.value)
    for value in ['Kolový bagr', 'Silniční válec', 'Optimistický', fmt_czk(65000), fmt_czk(5000000), '25', '2\u00a0000']:
        assert value in summary
    at.session_state['de_inv_editor'] = False; at.run()
    assert at.query_params['pohled'] == slug
    assert at.number_input(key='inv_fixed_cost').value == 65000
