"""Stránka Strojový park (modernizace UI): stejné údaje jako dřív, nový vzhled.

Počty kusů jsou z dat flotily, ekonomika jednoho kusu z portfolio.type_month – stejná tabulka jako
Investice → Nákup strojů. Test ověřuje shodu tabulky s Investicemi, grafu s tabulkou a KPI s daty
a to, že se nezobrazuje zavádějící souhrn tržeb/zisku za celou flotilu. Heslo je náhodné, jen pro test.
"""

import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))
import datamodel as dm  # noqa: E402
import portfolio as pf  # noqa: E402
from app_pages import open_page  # noqa: E402
from test_dashboard import dash_kpi, plain  # noqa: E402
from test_navigation import app, ok  # noqa: E402

COLUMNS = ["Stroj", "Cena (Kč)", "Poptávka (dny)", "Tržby (Kč/měs.)", "Náklady (Kč/měs.)", "Provozní zisk (Kč/měs.)"]
_STATE = {}


def fleet_app():
    if "at" not in _STATE:
        at = app()
        open_page(at, "strojovy_park")
        ok(at)
        _STATE["at"] = at
    return _STATE["at"]


def num(text: str) -> int:
    return int(text.replace(" ", "").replace(" ", "").replace("Kč", "").replace("−", "-"))


def table(at) -> list:
    html_all = [m.value for m in at.main.markdown if 'class="dd-cmp"' in m.value][0]
    head = [plain(h) for h in re.findall(r"<th>(.*?)</th>", html_all)]
    rows = [[plain(c) for c in re.findall(r"<td[^>]*>(.*?)</td>", tr)]
            for tr in re.findall(r"<tr>(.*?)</tr>", html_all.split("<tbody>")[1])]
    return head, rows


def test_unit_and_type_counts_match_fleet_data():
    at = fleet_app()
    cap = dm.fleet_capacity(dm.load_model())
    assert dash_kpi(at, "Počet strojů ve flotile") == f"{sum(cap.values())}kusů" == "9kusů"
    assert dash_kpi(at, "Počet typů strojů") == f"{len(cap)}typů" == "8typů"
    grid = [m.value for m in at.main.markdown if 'class="dd-units"' in m.value][0]
    shown = {plain(n): int(plain(c).removesuffix("ks"))
             for n, c in re.findall(r'class="n">(.*?)</span><span class="c">(.*?)</span>', grid)}
    assert shown == cap and "Vibrační deska" in shown and shown["Vibrační deska"] == 2


def test_economics_table_identical_to_investice_and_model():
    at = fleet_app()
    head, rows = table(at)
    assert head == COLUMNS
    inv = app()
    open_page(inv, "investice")
    inv_df = [d.value for d in inv.main.dataframe if "Provozní zisk (Kč/měs.)" in d.value.columns][0]
    assert list(inv_df.columns) == COLUMNS
    assert [[plain(str(v)) for v in r] for r in inv_df.values.tolist()] == rows        # stejné texty jako dřív
    df = dm.legacy_machines(dm.load_model())
    for (_, row), shown in zip(df.iterrows(), rows):                                   # a stejný výpočet
        one = pf.type_month(row, 1)
        assert shown[0] == row["machine_type"]
        assert [num(x) for x in shown[3:]] == [round(one["revenue"]), round(one["costs"]), round(one["profit"])]


def test_chart_uses_exactly_table_profits_sorted():
    at = fleet_app()
    _, rows = table(at)
    charts = [json.loads(c.proto.spec) for c in at.get("plotly_chart")]
    assert len(charts) == 1
    bar = charts[0]["data"][0]
    expected = sorted(((r[0], num(r[5])) for r in rows), key=lambda x: x[1], reverse=True)
    assert list(zip(bar["y"], [round(x) for x in bar["x"]])) == expected
    assert [num(t) for t in bar["text"]] == [p for _, p in expected]
    assert "jeden kus" in charts[0]["layout"]["xaxis"]["title"]["text"]


def test_kpis_are_single_unit_ranges_not_fleet_totals():
    at = fleet_app()
    _, rows = table(at)
    revenue = [num(r[3]) for r in rows]
    profit = [num(r[5]) for r in rows]
    assert dash_kpi(at, "Měsíční tržby jednoho kusu") == f"{min(revenue):,}–{max(revenue):,} Kč".replace(",", " ")
    assert dash_kpi(at, "Měsíční provozní zisk jednoho kusu") == f"{min(profit):,}–{max(profit):,} Kč".replace(",", " ")
    shown = plain(" ".join(m.value for m in at.main.markdown if "<style>" not in m.value))
    assert "model · 1 kus" in shown and "všech 8 typů v zisku" in shown
    assert "Souhrn tržeb a zisku za celou flotilu se záměrně nezobrazuje" in shown
    cap = dm.fleet_capacity(dm.load_model())
    df = dm.legacy_machines(dm.load_model())
    for total in (sum(revenue), sum(profit),                                         # žádný součet za flotilu
                  sum(pf.type_month(r, cap[r["machine_type"]])["profit"] for _, r in df.iterrows())):
        assert f"{round(total):,}".replace(",", " ") not in shown


def test_methodology_data_info_and_links_kept():
    at = fleet_app()
    exp = [e for e in at.main.expander if e.label == "Metodika a data"][0]
    text = " ".join(m.value for m in exp.markdown)
    for item in ("Zdroj", "Tržby", "Náklady", "Provozní zisk", "Proč ne souhrn za flotilu", "Počty kusů"):
        assert f'<div class="t">{item}</div>' in text
    shown = " ".join(m.value for m in at.main.markdown)
    assert "stejně jako Investice → Nákup strojů" in shown and "Načtená data:" in shown
    assert "potvrzen" in shown and "předběžn" in shown and "servisní odstávky" in shown
    links = [p.proto.label for p in at.get("page_link")]
    assert links == ["Obsazenost strojů po dnech – Plánování kapacity", "Nákup strojů – Investice"]
    import navigace
    urls = {p[0]: p[5] for p in navigace.PAGES}
    assert [p.proto.page for p in at.get("page_link")] == [urls["kapacita"], urls["investice"]]
    for key in ("kapacita", "investice"):                                    # cílové stránky se otevřou
        target = app()
        open_page(target, key)
        ok(target)
