"""Nákup: přesné portfolio a rozpočtová křivka, původní stavy bez ziskového nákupu."""
import json
import numpy as np
import pytest
import portfolio as pf
from formatting import fmt_czk, fmt_num
from test_investice import machines, ROOT
from test_investment_navigation_reference import audited_app, set_case, CASES, REFERENCE, BUDGETS
from test_investment_results_ui import array


def view(at):
    area=next(n for n in at.main.children.values() if getattr(n,'key',None)=='de_inv_obsah')
    return next(n for n in area.children.values() if getattr(n,'key',None)=='de_inv_obsah_5_on')


def open_view(at):
    at.query_params['pohled']='nakup-stroju';at.run()
    assert not at.exception
    return view(at)


def raw(node):return ' '.join(m.value for m in node.markdown)


@pytest.mark.parametrize('case',sorted(CASES))
def test_exact_portfolio_kpis_machine_details_and_frontier(tmp_path,case):
    at=audited_app(tmp_path);audit=set_case(at,case)
    assert audit==json.loads(REFERENCE.read_text())[case]
    node=open_view(at)
    assert at.session_state['_investment_audit']==audit
    options=pf.all_portfolios(machines());budget=BUDGETS[case];pb=pf.best_portfolio(options,budget)
    text=raw(node)
    assert 'portfolio.py' in text and 'průměrnou poptávkou bez simulace' in text
    assert 'nikoli skutečný čistý zisk celé firmy' in text
    if pb['units']:
        hero=next(m.value for m in node.markdown if 'class="de-hero ' in m.value)
        for name in pb['counts']:assert name in hero
        kpis=next(m.value for m in node.markdown if 'class="de-kpis"' in m.value)
        assert kpis.count('class="de-kpi"')==3
        assert fmt_czk(pb['investment']) in kpis and fmt_czk(pb['profit']) in kpis
        assert fmt_czk(budget-pb['investment']) in kpis
        assert fmt_num(pf.payback_months(pb),1)+' měsíce' in kpis
        assert 'Modelovaný měsíční provozní zisk' in kpis
        cards=next(m.value for m in node.markdown if 'class="de-res-grid"' in m.value)
        chunks=cards.split('<div class="de-res-card ')[1:]
        assert len(chunks)==len(pb['lines'])
        for (name,line),chunk in zip(pb['lines'],chunks):
            assert f"{name} – {line['units']} ks" in chunk
            for key in ('investment','revenue','costs','profit'):assert fmt_czk(line[key]) in chunk
            assert fmt_num(line['days'],1) in chunk
        assert fmt_czk(pf.profit_with_demand(machines(),pb['counts'],.8)) in text
        assert 'úroků a daní' in text
    else:
        assert 'Zatím nekupovat' in text
        assert not any('class="de-kpis"' in m.value for m in node.markdown)
    [fig]=[json.loads(c.proto.spec) for c in node.get('plotly_chart')]
    xmax=max(3000000,budget*1.25)
    steps=[(x,p) for x,p in pf.budget_frontier(options) if x<=xmax]
    assert np.array_equal(array(fig['data'][0]['x']),[x for x,_ in steps]+[xmax])
    assert np.array_equal(array(fig['data'][0]['y']),[p['profit'] for _,p in steps]+[steps[-1][1]['profit']])
    assert fig['data'][0]['line']['shape']=='hv'
    assert len(fig['data'][0]['customdata'])==len(steps)+1
    for label,(_,p) in zip(fig['data'][0]['customdata'],steps):
        assert all(name in label for name in p['counts'])
    assert fig['data'][1]['x']==[budget] and fig['data'][1]['y']==[pb['profit']]
    assert fig['layout']['showlegend'] and fig['layout']['legend']['y']<0
    assert fig['layout']['yaxis']['title']['text']=='Modelovaný měsíční provozní zisk (Kč)'
    assert any(e.label=='Předpoklady a metodika' for e in node.expander)
    assert 'Ekonomika jednoho kusu podle dat' in text


@pytest.mark.parametrize('budget',[100000,100000000])
def test_budget_extremes_keep_existing_selection_and_no_invented_utilization(tmp_path,budget):
    at=audited_app(tmp_path);at.number_input(key='inv_budget').set_value(budget).run()
    node=open_view(at);text=raw(node)
    if budget==100000:
        assert 'Zatím nekupovat' in text and 'Rozpočet nestačí' in text
        assert not any('class="de-kpis"' in m.value for m in node.markdown)
    else:assert 'Doporučená sestava' in text
    [fig]=[json.loads(c.proto.spec) for c in node.get('plotly_chart')]
    expected=pf.best_portfolio(pf.all_portfolios(machines()),budget)
    assert fig['data'][1]['y']==[expected['profit']]
    assert fig['data'][1]['x']==[budget]


def test_no_profitable_portfolio_does_not_display_fabricated_kpis(tmp_path):
    # Pouze testovací obal; produkční data ani algoritmus se nemění.
    source=(ROOT/'app.py').read_text().replace('purchase_options = pf.all_portfolios(df)',
        '_test_df = df.copy()\n_test_df["monthly_fixed_cost"] = 1e9\npurchase_options = pf.all_portfolios(_test_df)',1)
    at=audited_app(tmp_path,source=source);node=open_view(at);text=raw(node)
    assert 'Nekupovat žádný stroj' in text and 'Žádná sestava strojů nemá' in text
    assert 'portfolio.py' in text and 'nikoli skutečný čistý zisk celé firmy' in text
    assert not any('class="de-kpis"' in m.value for m in node.markdown)
    [fig]=[json.loads(c.proto.spec) for c in node.get('plotly_chart')]
    assert np.array_equal(array(fig['data'][0]['y']),[0,0])
    assert fig['data'][1]['y']==[0]


def test_purchase_styles_are_isolated_and_shared_panel_stays_single(tmp_path):
    import navigace
    import ui_theme as ui
    at=audited_app(tmp_path)
    for i,slug in enumerate(navigace.INVESTICE_VIEWS):
        at.query_params['pohled']=slug;at.run()
        assert not at.exception
        header=next(m.value for m in at.main.markdown if 'class="de-header' in m.value)
        assert 'zahlavi_dashboard.webp' in header
        assert any(m.value.strip()==ui.INVESTMENT_PURCHASE_CSS.strip() for m in at.main.markdown)==(i==4)
        assert len([n for n in at.main.children.values() if getattr(n,'key',None)=='de_inv_inputs'])==1
