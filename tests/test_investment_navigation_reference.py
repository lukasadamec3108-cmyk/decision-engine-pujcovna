"""Přesná reference scénářů z běžící aplikace před změnou investiční navigace.

Původní matematické a DOCX reference zůstávají nedotčené. Test navíc hlídá předání
vstupů v app.py, celá simulační pole, doporučení, rizika a nákupní portfolio.
"""
import json
import os
import secrets
from pathlib import Path

import pytest
from streamlit.testing.v1 import AppTest

from test_investice import CASES, ROOT

REFERENCE = ROOT / 'tests/fixtures/investment_navigation_reference.json'
BUDGETS = {'vychozi': 3000000, 'pesimisticky': 2500000, 'optimisticky': 5000000, 'oba_ztratove': 100000}
AUDIT = '''
if PAGE == "investice":
    from test_investice import _fingerprint
    from test_report_docx import parts
    st.session_state["_investment_audit"] = dict(
        inputs=[machine_1, machine_2, scenario, fixed_cost, expected_days, n, budget],
        simulation=_fingerprint(_inv),
        recommendation=[better_machine, recommendation, both_loss, scorecard_mini, scorecard_dumper],
        portfolio=_fingerprint(purchase_best),
        frontier=_fingerprint(purchase_frontier),
        docx=parts(build_docx_report()),
    )
'''


def audited_app(tmp_path, source=None):
    """Pouze testovací obal; produkční app.py se neinstrumentuje ani nemění."""
    os.chdir(ROOT)
    wrapper = tmp_path / 'audit_app.py'
    source = ((ROOT / 'app.py').read_text() if source is None else source) + AUDIT
    wrapper.write_text('exec(compile(' + repr(source) + ', ' + repr(str(ROOT / 'app.py')) + ', "exec"))')
    at = AppTest.from_file(str(wrapper), default_timeout=300)
    password = secrets.token_urlsafe(16)
    at.secrets['password'] = password
    at.run(); at.text_input[0].set_value(password).run()
    at.switch_page(str(ROOT / 'stranky/stranka_investice.py')).run()
    assert not at.exception, [e.value for e in at.exception]
    return at


def set_case(at, case):
    m1, m2, scenario, fixed, days, n = CASES[case]
    for key, value in [('inv_machine_1',m1),('inv_machine_2',m2),('inv_scenario',scenario)]:
        at.selectbox(key=key).set_value(value)
    at.number_input(key='inv_fixed_cost').set_value(fixed)
    at.number_input(key='inv_budget').set_value(BUDGETS[case])
    at.slider(key='inv_expected_days').set_value(days)
    at.slider(key='inv_n').set_value(n)
    at.run()
    assert not at.exception, [e.value for e in at.exception]
    return at.session_state['_investment_audit']


@pytest.mark.parametrize('case',sorted(CASES))
def test_app_scenario_matches_pre_navigation_reference(tmp_path, case):
    at = audited_app(tmp_path)
    expected = json.loads(REFERENCE.read_text())[case]
    assert set_case(at, case) == expected
    import navigace
    for slug in navigace.INVESTICE_VIEWS:
        at.query_params["pohled"] = slug
        at.run()
        assert not at.exception, [e.value for e in at.exception]
        assert at.session_state['_investment_audit'] == expected


def test_every_scenario_input_recalculates_immediately_and_survives_all_views(tmp_path):
    import investice as inv
    import portfolio as pf
    import report_docx as rd
    import navigace
    from test_investice import _fingerprint, machines
    from test_report_docx import inputs, parts

    at = audited_app(tmp_path)
    values = dict(inv_machine_1="Mini bagr", inv_machine_2="Pásový dumper", inv_scenario="Realistický",
                  inv_fixed_cost=100000, inv_expected_days=18, inv_n=1000, inv_budget=3000000)
    changes = [("inv_machine_1", "Kolový bagr"), ("inv_machine_2", "Silniční válec"),
               ("inv_scenario", "Optimistický"), ("inv_budget", 2500000), ("inv_fixed_cost", 60000),
               ("inv_expected_days", 25), ("inv_n", 2000)]
    data = machines()
    options = pf.all_portfolios(data)
    for key, value in changes:
        values[key] = value
        widget = at.selectbox if key in ("inv_machine_1", "inv_machine_2", "inv_scenario") else (
            at.slider if key in ("inv_expected_days", "inv_n") else at.number_input)
        widget(key=key).set_value(value).run()  # žádné potvrzovací tlačítko ani formulář
        assert not at.exception, [e.value for e in at.exception]
        m1, m2 = values["inv_machine_1"], values["inv_machine_2"]
        row1 = data[data.machine_type == m1].iloc[0]
        row2 = data[data.machine_type == m2].iloc[0]
        lo, hi = inv.scenario_days(values["inv_scenario"])
        expected = inv.evaluate(row1, row2, m1, m2, values["inv_fixed_cost"], lo, hi,
                                values["inv_expected_days"], values["inv_n"], seed=inv.SIMULATION_SEED)
        audit = at.session_state["_investment_audit"]
        assert audit["simulation"] == _fingerprint(expected), key
        assert audit["recommendation"][0] == expected["better_machine"], key
        assert audit["portfolio"] == _fingerprint(pf.best_portfolio(options, values["inv_budget"])), key
        report = inputs("vychozi")
        report.update({k: expected[k] for k in report if k in expected})
        report.update(machine_1=m1, machine_2=m2, scenario=values["inv_scenario"],
                      fixed_cost=values["inv_fixed_cost"], expected_days=values["inv_expected_days"],
                      n=values["inv_n"], min_days=lo, max_days=hi,
                      avg_daily_mini=float(row1.daily_price), avg_daily_dumper=float(row2.daily_price))
        assert audit["docx"] == parts(rd.build_report(**report)), key
        for kept_key, kept_value in values.items():
            assert at.session_state["_keep_" + kept_key] == kept_value
    final = at.session_state["_investment_audit"]
    for slug in navigace.INVESTICE_VIEWS:
        at.query_params["pohled"] = slug
        at.run()
        assert not at.exception
        assert at.session_state["_investment_audit"] == final
