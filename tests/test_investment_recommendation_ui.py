"""Prezentace Doporučení; původní reference výpočtů a reportu zůstávají beze změny."""
import json
import sys

import pytest
from test_investment_navigation_reference import audited_app, set_case, REFERENCE, ROOT, CASES


def recommendation(at):
    area = [n for n in at.main.children.values() if getattr(n, 'key', None) == 'de_inv_obsah'][0]
    return [n for n in area.children.values() if getattr(n, 'key', None) == 'de_inv_obsah_1_on'][0]


def hero(at):
    return [m.value for m in recommendation(at).markdown if 'class="de-hero ' in m.value][0]


@pytest.mark.parametrize('case', sorted(CASES))
def test_recommendation_preserves_reference_states_and_three_kpis(tmp_path, case):
    at = audited_app(tmp_path)
    audit = set_case(at, case)
    assert audit == json.loads(REFERENCE.read_text())[case]
    main = recommendation(at)
    [cards] = [m.value for m in main.markdown if 'class="de-kpis"' in m.value]
    assert cards.count('class="de-kpi"') == 3
    assert all(label in cards for label in ['Očekávaný měsíční zisk stroje', 'Šance ztrátového měsíce', 'Bod zvratu'])
    if audit['recommendation'][2]:
        assert 'Žádný stroj není ziskový' in hero(at) and 'Relativně lépe:' in hero(at)
        assert 'class="de-hero bad"' in hero(at)
    else:
        assert 'Doporučený stroj' in hero(at)
    direct = ' '.join(n.value for n in main.children.values() if getattr(n, 'type', None) == 'markdown')
    assert 'Zisk stroje nezahrnuje společné náklady' in direct and 'ukázkové' in direct
    assert 'Co udělat dál' in direct and 'Srovnání strojů' in direct
    [table] = [m.value for m in main.markdown if 'class="de-cmp-scroll"' in m.value and 'Očekávaný měsíční zisk' in m.value]
    assert 'tabindex="0"' in table and 'role="region"' in table
    assert 'Ukazatel' in table and 'Pořizovací cena' in table
    assert len(main.get('download_button')) == 1


def test_equal_selected_machines_keep_validation_and_editable_scenario(tmp_path):
    at = audited_app(tmp_path)
    at.selectbox(key='inv_machine_2').set_value(at.selectbox(key='inv_machine_1').value).run()
    assert not at.exception
    assert any('různé' in w.value for w in at.warning)
    assert at.selectbox(key='inv_machine_1') and at.selectbox(key='inv_machine_2')
    assert not any('class="de-hero ' in m.value for m in at.main.markdown)


@pytest.mark.parametrize('state', ['equal', 'higher_risk'])
def test_synthetic_render_states_do_not_invent_a_recommendation_rule(tmp_path, state):
    # Jen testovací obal předá rendereru krajní stav. Produkční výpočty, data a seed se nemění.
    source = (ROOT / 'app.py').read_text()
    patch = ('_inv.update(expected_profit_mini=50000.0, expected_profit_dumper=50000.0, '
             'better_machine=machine_1, profit_diff=0.0)\n' if state == 'equal' else
             '_inv.update(expected_profit_mini=50000.0, expected_profit_dumper=60000.0, '
             'better_machine=machine_2, prob_loss_mini=0.1, prob_loss_dumper=0.2, '
             'risk_loss_mini=10.0, risk_loss_dumper=20.0)\n')
    source = source.replace('# Simulace a základní metriky\n', patch + '# Simulace a základní metriky\n', 1)
    at = audited_app(tmp_path, source=source)
    if state == 'equal':
        assert 'Stroje vycházejí stejně' in hero(at)
        tables = [m.value for m in recommendation(at).markdown if 'class="de-cmp-scroll"' in m.value]
        assert tables and all('<span class="tag">' not in t for t in tables)
    else:
        assert 'Doporučený stroj' in hero(at)
        assert 'Vyšší očekávaný zisk, ale také vyšší riziko ztráty.' in hero(at)


def test_photo_styles_apply_only_to_recommendation_and_panel_is_shared(tmp_path):
    import navigace
    import ui_theme as ui
    at = audited_app(tmp_path)
    for i, slug in enumerate(navigace.INVESTICE_VIEWS):
        at.query_params['pohled'] = slug; at.run()
        assert not at.exception
        [header] = [m.value for m in at.main.markdown if 'class="de-header' in m.value]
        assert ('zahlavi_dashboard.webp' in header) == True
        styles = [m.value for m in at.main.markdown if m.value.strip() == ui.INVESTMENT_RECOMMENDATION_CSS.strip()]
        assert bool(styles) == (i == 0)
        assert len([n for n in at.main if getattr(n, 'key', None) == 'de_inv_inputs']) == 1


def test_regression_extractor_ignores_only_scroll_focus_metadata():
    sys.path.insert(0, str(ROOT / 'tools/regression'))
    import extract
    assert extract.numbers('<div class="de-cmp-scroll" tabindex="0">0 Kč 123 Kč</div>') == ['0', '123']
