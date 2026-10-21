"""Testy simulace „Co kdyby“ v záložce 8 (planner_view, E2.5c krok 4D-B). Jen fiktivní testovací data.

Simulace jen dosazuje hypotetickou hodnotu jedné nezadané sazby do vzorců plánovače; tady se ověřuje
proti economics.py (kopie dat s doplněnou hodnotou) i proti plánovači spuštěnému nad touto kopií.
Spuštění: python -m pytest tests   nebo   python tests/test_planner_view_sim.py
"""

import copy
import random
import re
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))
import planner as pl  # noqa: E402
import planner_view as pv  # noqa: E402
from test_econ_formula import filled, files_hash  # noqa: E402
from test_planner import BAGR, model  # noqa: E402
from test_planner_view import DEMO, STARTS, app, click  # noqa: E402

SMYK = "SMYKOVY_NAKLADAC"
NBSP = " "


def rent(ref, s, e, price=None, t=BAGR):
    return pv.rental_request(ref, t, 1, s, e, price)


def job(ref, s, dur, revenue=150000, direct=60000):
    return pv.job_request(ref, ref, [(BAGR, 1)], s, dur, None, None, revenue, direct, None)


def panel_for(res, ref):
    return next(p for p in res["panels"] if ref in p["requests"])


def p07():
    return panel_for(pv.evaluate(DEMO, STARTS, [rent("NP-3", 20, 23, 2000, SMYK)]), "NP-3")


# ---------------------------------------------------------------------------
# Kontrolní scénář FIKT-P07 × NP-3 a hraniční hodnoty
# ---------------------------------------------------------------------------

def test_fikt_p07_values_winner_and_position():
    spec = p07()["sim"]
    assert spec["available"] and spec["key"] == ["variable_cost_rental_day", SMYK]
    r = pv.simulate(spec, 6000)
    assert [(x["label"], x["amount"], x["tag"]) for x in r["rows"]] == [
        ("FIKT-P07", "7 500 Kč", "vychází nejlépe"), ("nepřijmout nic", "0 Kč", ""), ("NP-3", "−16 000 Kč", "")]
    assert r["verdict"] == "Při tomto předpokladu vychází nejlépe FIKT-P07."
    assert r["position"]["text"] == "Hodnota leží v úseku „pod 7 500 Kč/den“ – k hranici 1 (7 500 Kč/den) zbývá 1 500 Kč/den."
    r = pv.simulate(spec, 9000)
    assert r["winners"] == ["nepřijmout nic"] and "nad hranicí 1 (7 500 Kč/den) o 1 500 Kč/den" in r["position"]["text"]


def test_exact_tie_on_boundary():
    r = pv.simulate(p07()["sim"], 7500)
    assert sorted(r["winners"]) == ["FIKT-P07", "nepřijmout nic"] and not r["near"]
    assert r["verdict"] == "Při tomto předpokladu vycházejí přesně stejně: FIKT-P07 a nepřijmout nic."
    assert r["position"] == dict(kind="boundary", n=1, text="Hodnota leží přesně na hranici 1 (7 500 Kč/den).")
    assert all(x["tag"] == "vychází stejně – nejlépe" for x in r["rows"] if x["label"] != "NP-3")


def test_zero_is_valid_hypothesis_and_empty_is_no_simulation():
    spec = p07()["sim"]
    r = pv.simulate(spec, 0)
    assert r["x_text"] == "0 Kč/den (hypotetická nula)" and r["winners"] == ["FIKT-P07"]
    assert r["rows"][0]["amount"] == "37 500 Kč"
    for bad in (None, -1):
        with pytest.raises(ValueError):
            pv.simulate(spec, bad)


def test_decimal_boundary_near_tie_in_halere():
    # NP-4 = 30 500 − 5x, NP-1 = 18 000 − 2x → hranice 12 500 / 3 = 4 166,666… Kč/den
    spec = pv.evaluate(model(), {}, [rent("NP-4", 1, 5, 6100), rent("NP-1", 3, 4, 9000)])["panels"][0]["sim"]
    r = pv.simulate(spec, 4166.66)
    assert r["winners"] == ["NP-4"] and r["near"] and abs(r["gap"] - 0.02) < 1e-6
    assert [x["amount"] for x in r["rows"][:2]] == ["9 666,70 Kč", "9 666,68 Kč"]
    assert r["verdict"] == "Prakticky stejné: NP-4 vychází lépe než NP-1 jen o 0,02 Kč."
    r = pv.simulate(spec, 4166.67)
    assert r["winners"] == ["NP-1"] and r["near"]
    assert "nad hranicí 1 (4 166,67 Kč/den) o méně než 0,01 Kč/den" in r["position"]["text"]
    # hodnota se počítá na haléře (jak ji pole ukazuje): 12 500/3 i 4 166,666667 = 4 166,67 → NP-1 o 0,01 Kč
    for typed in (12500 / 3, 4166.666667):
        r = pv.simulate(spec, typed)
        assert r["x"] == 4166.67 and r["winners"] == ["NP-1"] and r["position"]["kind"] == "segment"
        assert r["verdict"] == "Prakticky stejné: NP-1 vychází lépe než NP-4 jen o 0,01 Kč."
    # daleko od hranice: celé koruny
    assert pv.simulate(spec, 1000)["rows"][0]["amount"] == "25 500 Kč"


def test_audit_position_always_agrees_with_calculation():
    # audit PR #36: hodnota rovná zaokrouhlené hranici nesmí hlásit „přesně na hranici“ s jediným vítězem
    # ani rozdíl „0,00 Kč“; poloha i rovnost se řídí výpočtem
    rnd = random.Random(9)
    checked = 0
    for _ in range(60):
        reqs = []
        for i in range(rnd.randint(2, 4)):
            s = rnd.randint(1, 6)
            reqs.append(rent(f"NP-{i + 1}", s, s + rnd.randint(0, 4), rnd.choice([None, 2000, 4500, 6100, 9000])))
        for pn in pv.evaluate(model(units=rnd.randint(1, 2)), {}, reqs)["panels"]:
            spec = pn["sim"]
            if not spec or not spec["available"]:
                continue
            xs = [b["at"] for b in spec["boundaries"]] + [b["at"] + 1e-7 for b in spec["boundaries"]] \
                + [round(b["at"], 2) for b in spec["boundaries"]] + [rnd.uniform(0, 20000) for _ in range(5)]
            for x in xs:
                r = pv.simulate(spec, x)
                assert (r["position"]["kind"] == "boundary") == (len(r["winners"]) > 1)
                if r["position"]["kind"] == "segment":
                    assert set(r["winners"]) <= set(spec["segments"][r["position"]["index"]]["winners"])
                assert "0,00 Kč" not in r["verdict"]
                assert r["x"] == round(x, 2)
                checked += 1
    assert checked > 100


def test_multiple_boundaries_and_ties():
    spec = panel_for(pv.evaluate(model(), {}, [rent("NP-1", 1, 6, 6000), rent("NP-2", 1, 4, 7500),
                                               rent("NP-3", 1, 2, 11000), rent("NP-4", 1, 1, 15000)]), "NP-1")["sim"]
    expected = {1000: ["NP-1"], 3000: ["NP-1", "NP-2"], 3500: ["NP-2"], 4000: ["NP-2", "NP-3"], 5000: ["NP-3"],
                7000: ["NP-3", "NP-4"], 10000: ["NP-4"], 15000: ["NP-4", "nepřijmout nic"], 20000: ["nepřijmout nic"]}
    for x, w in expected.items():
        r = pv.simulate(spec, x)
        assert sorted(r["winners"]) == sorted(w), x
        assert (r["position"]["kind"] == "boundary") == (len(w) > 1), x


# ---------------------------------------------------------------------------
# Kdy se simulace nenabízí
# ---------------------------------------------------------------------------

def test_not_offered_for_certain_results():
    for m, reqs in ((model(var_rental=1500, var_job=2000, includes="ne"), [job("NP-1", 1, 10), rent("NP-2", 1, 5, 6500)]),
                    (model(var_rental=7000), [rent("NP-1", 1, 3, 6000), rent("NP-2", 2, 4, 6500)]),
                    (model(var_rental=500), [rent("NP-1", 1, 3, 500), rent("NP-2", 2, 4, 0)])):
        for pn in pv.evaluate(m, {}, reqs)["panels"]:
            assert pn["sim"] is None and "Co kdyby" not in pv.panel_html(pn)


def test_not_offered_with_several_unknowns_or_unclear_setting():
    # chybí nastavení „stroje v přímých nákladech“ (dva scénáře) → nenabízet
    pn = pv.evaluate(model(), {}, [job("NP-1", 1, 10), rent("NP-2", 1, 5, 6500), rent("NP-3", 6, 10, 7000)])["panels"][0]
    assert pn["sim"] == dict(available=False, reason=pv.SIM_SETTING)
    assert "není zadáno, zda přímé náklady zakázek obsahují vlastní stroje" in pv.panel_html(pn)
    # nastavení zadané „ne“, ale nezadané obě sazby (pronájem i zakázka) → víc neznámých → nenabízet
    pn = pv.evaluate(model(includes="ne"), {}, [job("NP-1", 1, 10), rent("NP-2", 1, 5, 6500)])["panels"][0]
    assert pn["kind"] == "conditional" and pn["sim"] == dict(available=False, reason=pv.SIM_MULTI)
    # nastavení „ne“ a známá sazba pronájmu → jediná neznámá (sazba na zakázce) → nabídnout
    pn = pv.evaluate(model(includes="ne", var_rental=1500), {}, [job("NP-1", 1, 10), rent("NP-2", 1, 5, 6500),
                                                                  rent("NP-3", 6, 10, 7000)])["panels"][0]
    assert pn["sim"]["available"] and pn["sim"]["key"] == ["variable_cost_job_day", BAGR]
    r = pv.simulate(pn["sim"], 3000)                       # J = 90 000 − 10 × 3 000 = 60 000 > R1 + R2 = 52 500
    assert r["winners"] == ["NP-1"] and r["rows"][0]["amount"] == "60 000 Kč"


# ---------------------------------------------------------------------------
# Shoda s economics.py i s plánovačem nad kopií dat (náhodně)
# ---------------------------------------------------------------------------

def test_simulation_matches_economics_and_planner_random():
    rnd = random.Random(1717)
    checked = 0
    for _ in range(35):
        reqs = []
        for i in range(rnd.randint(2, 4)):
            s = rnd.randint(1, 6)
            reqs.append(rent(f"NP-{i + 1}", s, s + rnd.randint(0, 4), rnd.choice([None, 2000, 4500, 6100, 9000])))
        m = model(units=rnd.randint(1, 2))
        res = pv.evaluate(m, {}, reqs)
        p = pl.plan(m, {}, [pv.to_request(r) for r in reqs])
        by = {r["ref"]: pv.to_request(r) for r in reqs}
        for pn in res["panels"]:
            spec = pn["sim"]
            if not spec or not spec["available"]:
                continue
            g = next(g for g in p.groups if set(g["requests"]) == set(pn["requests"]))
            for x in [0, 1234.5] + [b["at"] for b in spec["boundaries"]] + [rnd.uniform(0, 20000)]:
                x = round(x, 2)
                r = pv.simulate(spec, x)
                mf = filled(m, rental={BAGR: x})
                truth = {pv._vlabel(v): sum(pl.contribution(mf, by[ref]).lo for ref in v) for v in g["candidates"]}
                for row in r["rows"]:
                    assert abs(row["value"] - truth[row["label"]]) < 1e-6           # stejné jako economics.py
                top = max(truth.values())
                assert set(r["winners"]) == {k for k, v in truth.items() if abs(v - top) < 1e-6}
                # plánovač nad kopií dat s doplněnou hodnotou dojde ke stejnému vítězi
                q = pl.plan(mf, {}, [pv.to_request(rr) for rr in reqs])
                gq = next(gg for gg in q.groups if set(gg["requests"]) == set(pn["requests"]))
                if gq["verdict"] in (pl.RECOMMEND, pl.NONE_PAYS):
                    assert r["winners"] == [pv._vlabel(gq["choice"])]
                else:
                    assert gq["verdict"] == pl.EQUAL and set(r["winners"]) == {pv._vlabel(tuple(b)) for b in gq["best"]}
                checked += 1
    assert checked > 30


# ---------------------------------------------------------------------------
# Nic se nemění, společná hodnota, oddělení od doporučení, graf
# ---------------------------------------------------------------------------

def test_simulation_changes_nothing():
    before = files_hash()
    res = pv.evaluate(DEMO, STARTS, [rent("NP-3", 20, 23, 2000, SMYK)])
    snap = copy.deepcopy(res)
    pn = panel_for(res, "NP-3")
    for x in (0, 7500, 12000):
        pv.simulate(pn["sim"], x)
        pv.sim_html(pn["sim"], pv.simulate(pn["sim"], x))
        pv.panel_html(pn, pv.simulate(pn["sim"], x))
    assert res == snap and files_hash() == before                         # výsledek plánovače i data beze změny


def test_shared_value_for_same_cost_parameter():
    # dva oddělené konflikty na stejném stroji = stejná sazba → stejný klíč; jiný stroj → jiný klíč
    res = pv.evaluate(model(), {}, [rent("NP-1", 1, 3, 6000), rent("NP-2", 2, 4, 7000),
                                    rent("NP-3", 20, 22, 6500), rent("NP-4", 21, 23, 9000)])
    sims = [p["sim"] for p in res["panels"]]
    assert len(sims) == 2 and all(s["available"] for s in sims)
    assert pv.sim_key(sims[0], "abc") == pv.sim_key(sims[1], "abc")
    assert pv.sim_key(sims[0], "abc") != pv.sim_key(sims[0], "xyz")      # jiný otisk požadavků → nová hodnota
    other = p07()["sim"]
    assert pv.sim_key(other, "abc") != pv.sim_key(sims[0], "abc")


def test_sim_box_is_separated_from_recommendation():
    spec = p07()["sim"]
    out = pv.sim_html(spec, pv.simulate(spec, 6000)).replace(NBSP, " ")
    assert "Co kdyby – hypotetický výpočet &#40;není doporučení&#41;" in out and "Skutečné doporučení: ⇄ Záleží na nákladech" in out
    assert "#1E7B4A" not in out and "★" not in out and "ec-rec" not in out
    css = "".join(line for line in pv.CSS.splitlines() if line.startswith((".dp-sim", ".dp-seg.sim")))
    assert "#1E7B4A" not in css and "dashed" in css


def test_chart_marker_only_labels_position():
    pn = p07()
    ch = pn["conditions"][0]["chart"]
    plain = pv.chart_html(ch)
    seg = pv.chart_html(ch, pv.simulate(pn["sim"], 6000)).replace(NBSP, " ")
    bnd = pv.chart_html(ch, pv.simulate(pn["sim"], 7500)).replace(NBSP, " ")
    assert "▼ předpoklad 6 000 Kč/den" in seg and 'class="dp-seg c0 sim"' in seg
    assert "▼ předpoklad 7 500 Kč/den" in bnd and 'class="dp-bnd sim"' in bnd
    strip = lambda h: re.sub(r'<div class="dp-sim-tag">.*?</div>', "", h).replace(' sim"', '"')
    for x in (6000, 7500):                                                 # úseky, hranice a texty beze změny
        assert strip(pv.chart_html(ch, pv.simulate(pn["sim"], x))) == plain
    assert pv.panel_html(pn) == pv.panel_html(pn, None)


def test_variant_names_do_not_break_at_hyphen():
    spec = p07()["sim"]
    out = pv.sim_html(spec, pv.simulate(spec, 6000))
    assert "vychází nejlépe FIKT\u2011P07." in out and "<b>FIKT\u2011P07</b>" in out and "<b>NP\u20113</b>" in out
    assert pv.simulate(spec, 6000)["verdict"] == "Při tomto předpokladu vychází nejlépe FIKT-P07."   # data beze změny


def test_sim_html_escapes():
    spec = dict(p07()["sim"], unknown_label="<b>x</b>")
    r = pv.simulate(spec, 1000)
    r["rows"][0]["label"] = "[a](https://example.com)<i>"
    out = pv.sim_html(spec, r)
    assert "<b>x" not in out and "<i>" not in out and "](https://example.com)" not in out


# ---------------------------------------------------------------------------
# Záložka 8 v aplikaci: zadání, výsledek, vymazání po změně požadavků
# ---------------------------------------------------------------------------

def _sim_inputs(at):
    return [n for n in at.tabs[7].number_input if str(n.key).startswith(pv.SIM_KEY_PREFIX)]


def test_tab_simulation_flow_and_reset():
    before = files_hash()
    at = app()
    at.selectbox(key="dp_r_type").set_value("PASOVY_DUMPER")
    at.number_input(key="dp_r_from").set_value(22); at.number_input(key="dp_r_to").set_value(23)
    click(at, "Přidat k porovnání")
    (box,) = _sim_inputs(at)
    assert box.value is None                                               # výchozí stav prázdný
    shown = " ".join(m.value for m in at.tabs[7].markdown)
    assert "hypotetický výpočet" not in shown
    box.set_value(6000.0).run()
    shown = " ".join(m.value for m in at.tabs[7].markdown).replace(NBSP, " ").replace("\u2011", "-")
    assert "Co kdyby – hypotetický výpočet &#40;není doporučení&#41;" in shown
    assert "Při tomto předpokladu vychází nejlépe FIKT-P06." in shown and "▼ předpoklad 6 000 Kč/den" in shown
    assert "Skutečné doporučení: ⇄ Záleží na nákladech" in shown
    # skutečné štítky karet se nezměnily
    assert '<small>Ekonomika:</small>⇄ Záleží na nákladech' in shown and 'class="dp-badge ec-rec"' not in shown
    old_key = box.key
    # změna požadavků → nový klíč, hodnota vymazaná
    at.number_input(key="dp_r_from").set_value(30); at.number_input(key="dp_r_to").set_value(31)
    click(at, "Přidat k porovnání")
    (box2,) = _sim_inputs(at)
    assert box2.key != old_key and box2.value is None and old_key not in at.session_state
    assert "hypotetický výpočet" not in " ".join(m.value for m in at.tabs[7].markdown)
    assert files_hash() == before


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK", name)
