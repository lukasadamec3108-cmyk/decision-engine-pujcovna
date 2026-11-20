"""Možné výsledky: přesná data simulace a správné rozlišení percentilů a pravděpodobností."""
import base64
import json
import numpy as np
import pytest
from test_investment_navigation_reference import audited_app, set_case, REFERENCE, CASES
from test_investice import run
from formatting import fmt_czk, fmt_pct


def view(at):
    area = [n for n in at.main.children.values() if getattr(n, 'key', None) == 'de_inv_obsah'][0]
    return [n for n in area.children.values() if getattr(n, 'key', None) == 'de_inv_obsah_2_on'][0]


def array(value):
    if isinstance(value, dict) and 'bdata' in value:
        return np.frombuffer(base64.b64decode(value['bdata']), dtype=value['dtype'])
    return np.asarray(value)


@pytest.mark.parametrize('case', sorted(CASES))
def test_results_show_exact_reference_percentiles_means_loss_rates_and_chart_data(tmp_path, case):
    at = audited_app(tmp_path)
    audit = set_case(at, case)
    assert audit == json.loads(REFERENCE.read_text())[case]
    at.query_params['pohled'] = 'vysledky'; at.run()
    assert not at.exception
    assert at.session_state['_investment_audit'] == audit
    r = run(case)
    node = view(at)
    [cards] = [m.value for m in node.markdown if 'class="de-res-grid"' in m.value]
    chunks = cards.split('<div class="de-res-card ')[1:]
    assert len(chunks) == 2 and cards.count('class="de-res-row"') == 10
    for text, suffix in zip(chunks, ('mini', 'dumper')):
        for label in ['Slabší měsíc (P10)', 'Typický měsíc (P50 · medián)', 'Silnější měsíc (P90)',
                      'Průměr všech měsíců', 'Šance ztrátového měsíce']:
            assert label in text
        values = [r['p10_'+suffix], r['p50_'+suffix], r['p90_'+suffix], r['expected_profit_'+suffix]]
        for number in values:
            assert fmt_czk(number) in text
            if number < 0:
                assert f'class="v neg">{fmt_czk(number)}' in text
        assert fmt_pct(r['prob_loss_'+suffix] * 100) in text
    raw = ' '.join(m.value for m in node.markdown)
    assert 'Nejde o pravděpodobnosti tří tržních scénářů ani procenta využití či obsazenosti' in raw
    if audit['recommendation'][2]:
        assert 'Oba stroje mají v tomto scénáři záporný průměrný měsíční výsledek' in raw
        assert cards.count('Relativně lepší') == 1
    specs = [json.loads(c.proto.spec) for c in node.get('plotly_chart')]
    histogram = next(s for s in specs if s['data'][0]['type'] == 'histogram')
    cdf = next(s for s in specs if s['data'][0]['type'] == 'scatter')
    for i, suffix in enumerate(('mini', 'dumper')):
        samples = r['profit_'+suffix]
        assert np.array_equal(array(histogram['data'][i]['x']), samples)
        assert histogram['data'][i]['nbinsx'] == 25
        assert np.array_equal(array(cdf['data'][i]['x']), np.sort(samples))
        assert np.array_equal(array(cdf['data'][i]['y']), np.arange(1,len(samples)+1)/len(samples))
        assert histogram['data'][i]['marker']['color'] == cdf['data'][i]['line']['color']
    assert histogram['layout']['barmode'] == 'overlay'
    assert histogram['layout']['xaxis']['title']['text'] == 'Měsíční zisk / ztráta (Kč)'
    assert histogram['layout']['yaxis']['title']['text'] == 'Počet simulovaných měsíců'
    assert histogram['layout']['legend']['y'] < 0
    assert cdf['layout']['yaxis']['tickformat'] == '.0%'
    assert [shape['x0'] for shape in histogram['layout']['shapes']] == [r['p50_mini'],r['p50_dumper'],0]
    assert any('Podrobnější analýza výsledků' == e.label for e in node.expander)
    for suffix in ('mini','dumper'):
        assert f"{fmt_czk(r['ci_'+suffix][0])} až {fmt_czk(r['ci_'+suffix][1])}" in raw


def test_results_styles_and_photo_are_isolated_from_other_views(tmp_path):
    import navigace
    import ui_theme as ui
    at = audited_app(tmp_path)
    for i, slug in enumerate(navigace.INVESTICE_VIEWS):
        at.query_params['pohled'] = slug; at.run()
        assert not at.exception
        [header] = [m.value for m in at.main.markdown if 'class="de-header' in m.value]
        assert ('zahlavi_dashboard.webp' in header) == True
        shown = any(m.value.strip() == ui.INVESTMENT_RESULTS_CSS.strip() for m in at.main.markdown)
        assert shown == (i == 1)
