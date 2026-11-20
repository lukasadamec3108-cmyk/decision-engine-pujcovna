"""Riziko: přesná data, původní kategorie a čitelná prezentace krajních stavů."""
import json
import pytest
from formatting import fmt_czk, fmt_pct
from investice import risk_light
from test_investice import run, machines, ROOT
from test_investment_navigation_reference import audited_app, set_case, CASES, REFERENCE


def view(at):
    area = next(n for n in at.main.children.values() if getattr(n, 'key', None) == 'de_inv_obsah')
    return next(n for n in area.children.values() if getattr(n, 'key', None) == 'de_inv_obsah_3_on')


def risk_view(at):
    at.query_params['pohled'] = 'riziko'; at.run()
    assert not at.exception
    return view(at)


@pytest.mark.parametrize('case', sorted(CASES))
def test_risk_values_categories_and_graph_match_reference(tmp_path, case):
    at = audited_app(tmp_path)
    audit = set_case(at, case)
    assert audit == json.loads(REFERENCE.read_text())[case]
    node = risk_view(at)
    assert at.session_state['_investment_audit'] == audit
    r = run(case)
    cards = next(m.value for m in node.markdown if 'class="de-res-grid"' in m.value)
    chunks = cards.split('<div class="de-res-card ')[1:]
    assert len(chunks) == 2 and cards.count('class="de-res-row"') == 8
    for text, suffix, name in zip(chunks, ('mini', 'dumper'), CASES[case][:2]):
        assert risk_light(r['prob_loss_'+suffix])[2:] in text
        assert fmt_pct(r['prob_loss_'+suffix]*100) in text
        assert fmt_czk(r['var95_'+suffix]) in text
        if r['var95_'+suffix] < 0:
            assert f'class="v neg">{fmt_czk(r["var95_"+suffix])}' in text
        failure = machines().set_index('machine_type').loc[name, 'failure_prob']
        assert fmt_pct(failure*100) in text
    raw = ' '.join(m.value for m in node.markdown)
    assert 'Skutečný výsledek může být i horší' in raw
    assert ('Oba stroje mají záporný očekávaný měsíční výsledek' in raw) == audit['recommendation'][2]
    assert ('při těchto předpokladech <b>vysoké</b>' in raw) == (
        r['prob_loss_mini'] >= .15 and r['prob_loss_dumper'] >= .15)
    [fig] = [json.loads(c.proto.spec) for c in node.get('plotly_chart')]
    for i in range(2):
        assert fig['data'][i]['x'] == [r['rr_x'+str(i+1)]]
        assert fig['data'][i]['y'] == [r['rr_y'+str(i+1)]]
        assert fig['data'][i]['name'] == CASES[case][i]
    assert fig['data'][0]['marker']['symbol'] != fig['data'][1]['marker']['symbol']
    assert fig['layout']['showlegend']
    assert fig['layout']['legend']['y'] < 0
    assert fig['layout']['xaxis']['range'][1] <= 110
    assert fig['layout']['yaxis']['range'][0] < min(0,r['rr_y1'],r['rr_y2'])
    assert fig['layout']['yaxis']['range'][1] > max(0,r['rr_y1'],r['rr_y2'])
    assert fig['layout']['shapes'][0]['y0'] == 0
    assert any(e.label == 'Jak číst ukazatele rizika' for e in node.expander)


@pytest.mark.parametrize('prob,category,tone', [(0,'Nízké riziko','good'),(.05,'Střední riziko','warn'),
                                             (.15,'Vysoké riziko','bad'),(1,'Vysoké riziko','bad')])
def test_equal_results_and_category_boundaries_are_presented_without_false_tradeoff(tmp_path, prob, category, tone):
    # Pouze renderer dostává syntetický krajní stav; model, seed a produkční data se nemění.
    source = (ROOT/'app.py').read_text()
    patch = (f'_inv.update(prob_loss_mini={prob}, prob_loss_dumper={prob}, '
             f'expected_profit_mini=-1000., expected_profit_dumper=-1000., '
             f'rr_x1={prob*100}, rr_x2={prob*100}, rr_y1=-1000., rr_y2=-1000., both_loss=True)\n')
    source = source.replace('# Simulace a základní metriky\n', patch+'# Simulace a základní metriky\n',1)
    at = audited_app(tmp_path,source=source)
    node = risk_view(at)
    raw = ' '.join(m.value for m in node.markdown)
    assert 'shodnou šanci ztrátového měsíce i očekávaný výsledek' in raw
    assert 'jeden má vyšší' not in raw
    assert f'class="de-status {tone}">{category}' in raw
    assert 'neznamená ziskovost ani doporučení k nákupu' in raw


def test_risk_css_and_photo_are_isolated_and_panel_stays_single(tmp_path):
    import navigace
    import ui_theme as ui
    at = audited_app(tmp_path)
    for i, slug in enumerate(navigace.INVESTICE_VIEWS):
        at.query_params['pohled'] = slug; at.run()
        assert not at.exception
        header = next(m.value for m in at.main.markdown if 'class="de-header' in m.value)
        assert ('zahlavi_dashboard.webp' in header) == True
        assert any(m.value.strip() == ui.INVESTMENT_RISK_CSS.strip() for m in at.main.markdown) == (i == 2)
        assert len([n for n in at.main if getattr(n,'key',None)=='de_inv_inputs']) == 1
