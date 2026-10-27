"""Stránka Pronájem strojů (modernizace UI): stejné vstupy, validace a výsledek plánovače, nový vzhled.

KPI karty a karty variant jen zobrazují výsledek planner_view.evaluate – test ověřuje, že odpovídají
kartám plánovače a výsledku v Rozhodovacím centru. Heslo je náhodné, jen pro test.
"""

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))
import planner_view as pv  # noqa: E402
import rental_view as rv  # noqa: E402
from app_pages import open_page  # noqa: E402
from test_navigation import app, ok  # noqa: E402


def plain(value: str) -> str:
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", value).replace(" ", " ")).strip()


def kpis(at) -> dict:
    """{popisek: (hodnota, štítek, vysvětlení)} KPI karet výsledku."""
    html_all = " ".join(m.value for m in at.main.markdown if "dd-kpis" in m.value and "<style>" not in m.value)
    out = {}
    for card in html_all.split('<div class="dd-kpi">')[1:]:
        label = plain(card.split('class="dd-lbl">')[1].split("</span>")[0])
        value = plain(card.split('<div class="dd-val')[1].split(">", 1)[1].split("</div>")[0])
        chip = re.search(r'class="dd-chip [a-z]+">(.*?)</span>', card)
        sub = plain(card.split('class="dd-sub">')[1].split("</div>")[0])
        out[label] = (value, plain(chip.group(1)) if chip else "", sub)
    return out


def submit(at, label, type_id, start, end, qty=1, price=None, check=True):
    at.selectbox(key="dp_r_type").set_value(type_id)
    at.number_input(key="dp_r_qty").set_value(qty)
    at.number_input(key="dp_r_from").set_value(start)
    at.number_input(key="dp_r_to").set_value(end)
    at.number_input(key="dp_r_price").set_value(price)
    [b for b in at.main.button if b.label == label][0].click().run()
    if check:
        ok(at)


def rental_app():
    at = app()
    open_page(at, "pronajem")
    ok(at)
    return at


def test_form_inputs_defaults_and_buttons_unchanged():
    at = rental_app()
    assert at.selectbox(key="dp_r_type").label == "Stroj"
    assert at.number_input(key="dp_r_qty").value == 1 and at.number_input(key="dp_r_qty").label == "Počet kusů"
    assert (at.number_input(key="dp_r_qty").min, at.number_input(key="dp_r_qty").max) == (1, 10)
    assert at.number_input(key="dp_r_from").value == 21 and at.number_input(key="dp_r_to").value == 24
    assert (at.number_input(key="dp_r_from").min, at.number_input(key="dp_r_from").max) == (1, 365)
    assert at.number_input(key="dp_r_price").value is None                       # nepovinná cena
    assert [b.label for b in at.main.button if b.label in ("Vyhodnotit požadavek", "Přidat k porovnání")] == [
        "Vyhodnotit požadavek", "Přidat k porovnání"]
    assert not [r for r in at.radio if r.key == "dp_kind"]                       # jen pronájem
    assert any(e.label == "Jak plánovač rozhoduje" for e in at.main.expander)
    assert any("Ukázková data" in m.value for m in at.main.markdown)


def test_validation_rejects_invalid_period_and_adds_nothing():
    at = rental_app()
    submit(at, "Vyhodnotit požadavek", "PASOVY_DUMPER", 24, 22, check=False)
    assert not at.exception
    expected = pv.form_problems(pv.rental_request("X", "PASOVY_DUMPER", 1, 24, 22))
    assert expected and [e.value for e in at.error] == [" ".join(expected)]
    assert at.session_state["dp_requests"] == []
    assert not kpis(at)


def test_single_evaluation_kpis_match_planner_card():
    at = rental_app()
    submit(at, "Vyhodnotit požadavek", "PASOVY_DUMPER", 22, 23)
    k = kpis(at)
    assert list(k) == ["Dostupná kapacita", "Ekonomický přínos", "Kapacitní konflikt", "Výsledné doporučení"]
    assert k["Dostupná kapacita"][0] == "Vejde se" and "vejde se do zbývající kapacity" in k["Dostupná kapacita"][2]
    assert k["Ekonomický přínos"][0] == "nejvýše 16 000 Kč"                        # horní mez z plánovače
    assert "nezadaném údaji" in k["Ekonomický přínos"][2] and "variabilní náklad pronájmu" in k["Ekonomický přínos"][2]
    assert k["Kapacitní konflikt"][0] == "Bez konfliktu" and k["Kapacitní konflikt"][1] == "✓ kapacita stačí"
    assert k["Výsledné doporučení"][0] == "Méně výhodné"
    shown = " ".join(m.value for m in at.main.markdown)
    assert 'class="dp-card"' in shown and "Ekonomické porovnání" in shown       # podrobnosti zůstávají
    assert "Porovnání variant" in shown                                            # soupeří s FIKT-P06 z dat
    assert len(at.session_state["dp_requests"]) == 1


def test_not_feasible_request_explains_why_values_cannot_be_determined():
    at = rental_app()
    submit(at, "Vyhodnotit požadavek", "SILNICNI_VALEC", 15, 18)
    k = kpis(at)
    assert k["Dostupná kapacita"][0] == "Nevejde se" and "kapacitu drží závazky" in k["Dostupná kapacita"][2]
    assert k["Ekonomický přínos"][0] == "Nelze určit" and "nevejde do kapacity" in k["Ekonomický přínos"][2]
    assert k["Kapacitní konflikt"][0] == "4 dny" and "15" in k["Kapacitní konflikt"][2]
    assert "konflikt potvrzených závazků" in k["Kapacitní konflikt"][2]
    assert k["Výsledné doporučení"][0] == "Neproveditelné"
    assert any(e.label.startswith("Co blokuje NP-") for e in at.main.expander)     # podrobná tabulka zůstává


def test_comparison_of_variants_and_same_result_as_decision_centre():
    at = rental_app()
    submit(at, "Přidat k porovnání", "MINI_BAGR", 40, 44, price=6100)
    submit(at, "Přidat k porovnání", "MINI_BAGR", 42, 43, price=9000)
    assert [r["ref"] for r in at.session_state["dp_requests"]] == ["NP-1", "NP-2"]
    rent = " ".join(m.value for m in at.main.markdown)
    variants = rent.split('<div class="dd-vars">')[1]
    assert variants.count('<div class="dd-var">') == 2
    assert "⇄ Záleží na nákladech" in plain(variants) and "nejvýše 30 500 Kč" in plain(variants)
    assert k_value(at, "Výsledné doporučení") == "⇄ Záleží na nákladech"           # poslední přidaný NP-2
    assert any(str(n.key).startswith(pv.SIM_KEY_PREFIX) for n in at.main.number_input)   # „Co kdyby“ zůstává
    rent_cards = [m.value for m in at.main.markdown if 'class="dp-card"' in m.value]
    rent_panels = [m.value for m in at.main.markdown if "dp-panel" in m.value and "<style>" not in m.value]
    open_page(at, "rozhodovaci_centrum")
    ok(at)
    assert [m.value for m in at.main.markdown if 'class="dp-card"' in m.value] == rent_cards
    assert [m.value for m in at.main.markdown if "dp-panel" in m.value and "<style>" not in m.value] == rent_panels


def k_value(at, label):
    return kpis(at)[label][0]


def test_economy_value_branches_without_inventing_numbers():
    base = dict(econ_badge=None, econ_why="Ekonomicky se nehodnotí – požadavek se nevejde do kapacity.")
    assert rv.dp_economy_value(dict(base, economy=None))[0] == "Nelze určit"
    exact = dict(base, economy=dict(status=pv.EXACT, value=12500.0, inputs=[]))
    assert rv.dp_economy_value(exact)[0] == "12 500 Kč"
    unknown = dict(base, economy=dict(status=pv.WITH_UNKNOWN, hi=None, lo=None, inputs=[
        dict(label="variabilní náklad", state=pv.NEZADANO, applies=True)]))
    value, why = rv.dp_economy_value(unknown)
    assert value == "Nelze určit" and "variabilní náklad" in why
    invalid = dict(base, economy=dict(status=pv.INVALID, invalid=["záporná cena"], inputs=[]))
    value, why = rv.dp_economy_value(invalid)
    assert value == "Nelze určit" and "záporná cena" in why


def test_explanation_groups_keep_every_original_item():
    import ast
    tree = ast.parse((ROOT / "app.py").read_text(encoding="utf-8"))
    names = {t.id for n in ast.walk(tree) if isinstance(n, ast.Assign) for t in n.targets if isinstance(t, ast.Name)}
    assert {"DP_EXPLAIN", "DP_EXPLAIN_GROUPS"} <= names
    at = rental_app()
    exp = [e for e in at.main.expander if e.label == "Jak plánovač rozhoduje"][0]
    text = " ".join(m.value for m in exp.markdown)
    for group in ("Provoz a kapacita", "Ekonomika", "Význam doporučení", "Konflikty a alternativy"):
        assert f'<div class="dd-group">{group}</div>' in text
    for item in ("Zdroj", "Provoz", "Ekonomika", "Štítky ekonomiky", "Doporučeno", "Nelze rozhodnout", "Konflikty",
                 "Alternativy", "Nic se nemění", "Předběžné rezervace"):
        assert f'<div class="t">{item}</div>' in text, item
