"""Citlivost: původní výsledky, data všech grafů a zachované rozdíly metodik."""
import json
import numpy as np
import pytest
from formatting import fmt_czk, fmt_days
from test_investice import run, ROOT, machines
from test_investment_navigation_reference import audited_app, set_case, CASES, REFERENCE
from test_investment_results_ui import array


def view(at):
    area = next(n for n in at.main.children.values() if getattr(n, 'key', None) == 'de_inv_obsah')
    return next(n for n in area.children.values() if getattr(n, 'key', None) == 'de_inv_obsah_4_on')


def open_view(at):
    at.query_params['pohled'] = 'vliv-na-zisk'; at.run()
    assert not at.exception
    return view(at)


@pytest.mark.parametrize('case',sorted(CASES))
def test_sensitivity_prices_curves_and_profitability_map_match_reference(tmp_path,case):
    import ui_theme as ui
    at = audited_app(tmp_path)
    audit = set_case(at,case)
    assert audit == json.loads(REFERENCE.read_text())[case]
    node = open_view(at)
    assert at.session_state['_investment_audit'] == audit
    r = run(case)
    raw = ' '.join(m.value for m in node.markdown)
    assert 'ne simulací' in raw and 'Scénář trhu tuto část neovlivňuje' in raw
    assert 'Graf neukazuje zisk' in raw and '<b>bez servisních nákladů a poruch</b>' in raw
    assert 'Servisní náklady nezahrnují' in raw
    cards = next(m.value for m in node.markdown if 'class="de-res-grid"' in m.value)
    assert cards.count('class="de-res-row"') == 8
    for text,i in zip(cards.split('<div class="de-res-card ')[1:],(1,2)):
        actual = float(machines().set_index('machine_type').loc[CASES[case][i-1],'daily_price'])
        required = r['required_price_'+str(i)]
        assert fmt_czk(required) in text and fmt_czk(actual) in text
        assert ('Cena dostačuje' in text) == (actual>=required)
        assert ('Cena je pod minimem' in text) == (actual<required)
    figures=[json.loads(c.proto.spec) for c in node.get('plotly_chart')]
    assert len(figures)==3
    tornado,be,heat=figures
    short={'Denní cena':'Denní cena','Vytížení (dny/měsíc)':'Dny pronájmu',
           'Fixní náklad':'Fixní náklad','Servisní náklad':'Servisní náklad'}
    effects={}
    for label,value in r['sensitivity_data']:
        driver,name=label.split(' – ',1);effects[short[driver],name]=value
    for i,name in enumerate(CASES[case][:2]):
        trace=tornado['data'][i]
        assert trace['name']==name and trace['marker']['color']==ui.SERIES[i]
        expected=np.array([effects[d,name] for d in trace['y']])
        assert np.array_equal(array(trace['x']),expected)
        assert trace['text']==[fmt_czk(x) for x in expected]
        assert trace['customdata']==[fmt_czk(x/2) for x in expected]
        curve=be['data'][i]
        assert curve['line']['color']==ui.SERIES[i]
        assert np.array_equal(array(curve['x']),r['days_grid'])
        assert np.array_equal(array(curve['y']),r['profit_curve_mini' if i==0 else 'profit_curve_dumper'])
    for i,suffix in enumerate(('mini','dumper')):
        point=be['data'][i+2]
        assert point['x']==[r['break_even_'+suffix]] and point['y']==[0]
        assert fmt_days(r['break_even_'+suffix],1) in point['text'][0]
    assert be['layout']['yaxis']['title']['text']=='Tržby − fixní náklad (Kč/měsíc)'
    assert be['layout']['legend']['y']<0
    for trace in heat['data'][:2]:
        assert np.array_equal(array(trace['x']),r['days_range'])
        assert np.array_equal(array(trace['y']),r['price_range'])
        z=array(trace['z']).reshape(r['profit_matrix'].shape)
        assert np.array_equal(z,r['profit_matrix'])
    assert heat['data'][0]['zmid']==0
    assert heat['data'][0]['colorscale']==[[0,ui.BAD],[.5,'#F1F3F5'],[1,ui.BLUE]]
    assert {e.label for e in node.expander}=={'Mapa ziskovosti – kombinace ceny a počtu dní','Jak se počítají údaje na této záložce'}
    if r['profit_matrix'].max()<=0:
        assert 'tržby nikde nepokryjí ani fixní náklad' in raw


@pytest.mark.parametrize('days,fixed',[(0,100000),(30,0)])
def test_zero_demand_and_zero_cost_keep_original_edge_behaviour(tmp_path,days,fixed):
    at=audited_app(tmp_path)
    at.slider(key='inv_expected_days').set_value(days)
    at.number_input(key='inv_fixed_cost').set_value(fixed).run()
    node=open_view(at)
    raw=' '.join(m.value for m in node.markdown)
    assert ('nelze náklady pokrýt žádnou cenou' in raw)==(days==0)
    if fixed==0:assert 'pokrývá fixní náklad už od prvního dne' in raw
    assert len(node.get('plotly_chart'))==3


def test_sensitivity_styles_photo_and_shared_panel_are_isolated(tmp_path):
    import navigace
    import ui_theme as ui
    at=audited_app(tmp_path)
    for i,slug in enumerate(navigace.INVESTICE_VIEWS):
        at.query_params['pohled']=slug;at.run()
        assert not at.exception
        header=next(m.value for m in at.main.markdown if 'class="de-header' in m.value)
        assert ('zahlavi_dashboard.webp' in header)==True
        assert any(m.value.strip()==ui.INVESTMENT_SENSITIVITY_CSS.strip() for m in at.main.markdown)==(i==3)
        assert len([n for n in at.main.children.values() if getattr(n,'key',None)=='de_inv_inputs'])==1
