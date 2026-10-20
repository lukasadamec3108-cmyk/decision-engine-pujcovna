"""Testy panelu „Ekonomické porovnání“ v záložce 8 (planner_view, E2.5c krok 4C). Jen fiktivní data.

Spuštění: python -m pytest tests   nebo   python tests/test_planner_view_panel.py
"""

import pickle
import random
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))
import econ_formula as ef  # noqa: E402
import planner as pl  # noqa: E402
import planner_view as pv  # noqa: E402
from test_econ_formula import files_hash  # noqa: E402
from test_planner import BAGR, model  # noqa: E402
from test_planner_view import DEMO, STARTS  # noqa: E402

NBSP = " "


def rent(ref, s, e, price=None, t=BAGR):
    return pv.rental_request(ref, t, 1, s, e, price)


def job(ref, s, dur, revenue=150000, direct=60000, earliest=None, deadline=None, name=None):
    return pv.job_request(ref, name or ref, [(BAGR, 1)], s, dur, earliest, deadline, revenue, direct, None)


def run(m, reqs, starts=None):
    return pv.evaluate(m, starts or {}, reqs)


def one_panel(res):
    assert len(res["panels"]) == 1
    return res["panels"][0]


def text(panel):
    return pv.panel_html(panel).replace(NBSP, " ")


# ---------------------------------------------------------------------------
# Pět výsledků: odpověď, Proč, varianty
# ---------------------------------------------------------------------------

def test_recommend_answer_and_why_with_amounts():
    pn = one_panel(run(model(var_rental=1500, var_job=2000, includes="ne"),
                       [job("NP-1", 1, 10), rent("NP-2", 1, 5, 6500), rent("NP-3", 6, 10, 7000)]))
    assert pn["badge"] == pv.EC_REC and pn["headline"] == "Přijmout NP-1."
    assert pn["why"][0] == ("NP-1 přinese 70 000 Kč, nejlepší jiná varianta (NP-2 + NP-3) 52 500 Kč "
                            "– o 17 500 Kč méně.")
    labels = [v["label"] for v in pn["variants"]]
    assert labels[0] == "NP-1" and "NP-2 + NP-3" in labels and labels[-1] == "nepřijmout nic"
    assert pn["variants"][0]["tag"] == "doporučeno"
    assert pn["missing"] == [] and pn["conditions"] == []


def test_recommend_with_incomplete_data_gives_no_amount_difference():
    # J = 90 000 přesně, pronájmy s nezadaným nákladem → doporučení je jisté, ale rozdíl v Kč neznámý
    pn = one_panel(run(model(includes="ano"), [job("NP-1", 1, 10), rent("NP-2", 1, 5, 6500), rent("NP-3", 6, 10, 7000)]))
    assert pn["badge"] == pv.EC_REC
    assert not any(" méně" in x for x in pn["why"])
    assert any("Rozdíly v Kč se bez úplných údajů neuvádějí" in x for x in pn["why"])
    assert "NP-1 přinese přesně 90 000 Kč." in pn["why"]


def test_none_pays_panel():
    pn = one_panel(run(model(var_rental=7000), [rent("NP-1", 1, 3, 6000), rent("NP-2", 2, 4, 6500)]))
    assert pn["badge"] == pv.EC_NO and pn["headline"].startswith("Nepřijmout žádný z těchto požadavků")
    assert "NP-1: ztráta 3 000 Kč." in pn["why"] and "NP-2: ztráta 1 500 Kč." in pn["why"]
    nothing = next(v for v in pn["variants"] if v["nothing"])
    assert nothing["tag"] == "doporučeno" and nothing["amount"] == "0 Kč"
    assert "nezamítá" in pn["note"] and "žádná rezervace se neruší" in pn["note"]


def test_none_pays_with_incomplete_data_says_at_least():
    pn = one_panel(run(model(), [job("NP-1", 1, 3, revenue=10000, direct=20000),
                                 job("NP-2", 2, 3, revenue=5000, direct=9000)]))
    assert pn["kind"] == "none_pays"
    assert any("ztráta nejméně 10 000 Kč (nezadané náklady ji mohou jen zvětšit)" in x for x in pn["why"])


def test_equal_panel():
    pn = one_panel(run(model(var_rental=500), [rent("NP-1", 1, 3, 500), rent("NP-2", 2, 4, 0)]))
    assert pn["badge"] == pv.EC_EQ and "nepřijmout nic" in pn["headline"] and "NP-1" in pn["headline"]
    assert "Každá z nich přinese přesně 0 Kč." in pn["why"]
    assert {v["label"]: v["tag"] for v in pn["variants"]}["NP-1"] == "shodně nejlepší"


def test_conditional_single_unknown_panel():
    pn = one_panel(run(DEMO, [rent("NP-1", 22, 23, None, "PASOVY_DUMPER")], STARTS))
    assert pn["badge"] == pv.EC_COND
    assert pn["headline"] == "Záleží na nezadaném nákladu: Pásový dumper: variabilní náklad pronájmu (Kč/den)."
    (block,) = pn["conditions"]
    assert block["lines"][:3] == ["pod 8 000 Kč/den: FIKT-P06",
                                  "přesně 8 000 Kč/den: FIKT-P06, NP-1 a nepřijmout nic vycházejí stejně (rovnocenné)",
                                  "nad 8 000 Kč/den: nepřijmout nic"]
    assert "NP-1: nevychází nejlépe při žádné výši nákladu (stejně jako nejlepší jen přesně při 8 000 Kč/den)" \
        in block["lines"]
    assert block["break_even"] and "Hranice zlomu je výše nákladu" in text(pn)
    assert pn["missing"] == ["Pásový dumper: variabilní náklad pronájmu (Kč/den)"]      # ne smykový nakladač
    assert "nepočítá jako 0 Kč" in pn["missing_note"]


def test_conditional_multiple_unknowns_names_both_variants_and_units():
    pn = one_panel(run(model(), [job("NP-1", 1, 10), rent("NP-2", 1, 5, 6500), rent("NP-3", 6, 10, 7000)]))
    assert pn["badge"] == pv.EC_COND and pn["headline"].startswith("Záleží na nezadaných údajích")
    blocks = {b["title"]: b for b in pn["conditions"]}
    assert set(blocks) == {"Pokud přímé náklady zakázek obsahují stroje",
                           "Pokud přímé náklady zakázek neobsahují stroje"}
    ne = blocks["Pokud přímé náklady zakázek neobsahují stroje"]
    assert ne["lead"] == "Rozhoduje víc nezadaných nákladů najednou:" and not ne["break_even"]
    for sentence in ne["lines"] + ne["more"]:
        assert " je lepší než " in sentence or " je vždy lepší než " in sentence        # obě varianty
        assert "Kč" in sentence
    assert ("NP-1 je lepší než NP-2 + NP-3, pokud je Mini bagr: variabilní náklad na zakázce (Kč/den) vyšší než "
            "Mini bagr: variabilní náklad pronájmu (Kč/den) o méně než 2 250 Kč/den (nebo nižší).") in ne["lines"]
    assert ("NP-1 je lepší než nepřijmout nic, pokud je Mini bagr: variabilní náklad na zakázce (Kč/den) nižší "
            "než 9 000 Kč/den.") in ne["lines"]
    assert set(pn["missing"]) == {"Mini bagr: variabilní náklad na zakázce (Kč/den)",
                                  "Mini bagr: variabilní náklad pronájmu (Kč/den)",
                                  "zda přímé náklady zakázek obsahují vlastní stroje (ano/ne)"}


def test_undecided_panel():
    req = pv.job_request("NP-1", "J", [(BAGR, 1)], 1, 3, None, None, None, 1000, None)
    pn = one_panel(run(model(var_rental=0, includes="ano"), [req, rent("NP-2", 2, 4, 6500)]))
    assert pn["badge"] == pv.EC_UND and pn["headline"].startswith("Nelze rozhodnout – chybí nebo je neplatné")
    assert "NP-2 přinese přesně 19 500 Kč." in pn["why"]
    assert pn["missing"] == ["NP-1 (J): tržba (Kč)"]


def test_standalone_request_has_no_panel():
    res = run(model(), [rent("NP-1", 1, 3, 6000)])
    assert res["panels"] == [] and res["cards"][0]["group_refs"] is None


# ---------------------------------------------------------------------------
# Podmínky slovy – přesnost zjednodušení
# ---------------------------------------------------------------------------

def _diff(c, xs):
    return c["known"] + sum(t["coef"] * xs[tuple(t["key"])] for t in c["terms"])


def test_condition_sentence_simplifications_are_exact():
    a, b = ("variable_cost_job_day", BAGR), ("variable_cost_rental_day", BAGR)
    term = lambda key, coef: dict(key=list(key), label=f"{key[0]}/den", unit=pv.KC_DEN, kind=ef.RATE,
                                  coef=coef, coef_unit="strojodnů")
    base = dict(a_label="A", b_label="B", status=ef.CONDITIONAL, winner=None, strict=False, unit=pv.KC)
    # jedna neznámá: A − B = 22 500 − 10x > 0 ⇔ x < 2 250
    c = dict(base, known=22500.0, terms=[term(a, -10.0)])
    assert pv.condition_sentence(c).endswith("nižší než 2 250 Kč/den.")
    # dvě neznámé, stejný počet dní: 22 500 − 10x + 10y > 0 ⇔ x − y < 2 250
    c = dict(base, known=22500.0, terms=[term(a, -10.0), term(b, 10.0)])
    assert "o méně než 2 250 Kč/den" in pv.condition_sentence(c)
    rnd = random.Random(4)
    for _ in range(200):
        xs = {a: rnd.uniform(0, 9000), b: rnd.uniform(0, 9000)}
        assert (_diff(c, xs) > 0) == (xs[a] - xs[b] < 2250)
    # záporný rozdíl: −5 000 − 10x + 10y > 0 ⇔ x nižší než y o více než 500
    c = dict(base, known=-5000.0, terms=[term(a, -10.0), term(b, 10.0)])
    assert "nižší než" in pv.condition_sentence(c) and "o více než 500 Kč/den" in pv.condition_sentence(c)
    # různý počet dní → úplná nerovnice (žádné falešné zjednodušení)
    c = dict(base, known=55000.0, terms=[term(a, -10.0), term(b, 5.0)])
    s = pv.condition_sentence(c)
    assert "10 strojodnů ×" in s and "< 55 000 Kč + 5 strojodnů ×" in s
    # jistý výsledek ve scénáři
    assert pv.condition_sentence(dict(base, status=ef.CERTAIN, winner=ef.A, strict=False, known=0.0, terms=[])) == \
        "A je vždy lepší než B (při nulovém nákladu stejně)."


# ---------------------------------------------------------------------------
# Neznámé ≠ nula, „neuplatňuje se“, upozornění se neztratí
# ---------------------------------------------------------------------------

def test_three_states_on_cards():
    res = run(model(var_rental=0, includes="ano"),
              [pv.job_request("NP-1", "J", [(BAGR, 1)], 1, 3, None, None, None, 1000, None), rent("NP-2", 2, 4, 6500)])
    c1 = next(c for c in res["cards"] if c["ref"] == "NP-1")["economy_lines"]
    c2 = next(c for c in res["cards"] if c["ref"] == "NP-2")["economy_lines"]
    assert "nezadáno" in c1[0]                                                # nezadáno
    assert any(x.startswith("Nezadáno, ale neuplatňuje se") and "penále (zakázka končí v termínu)" in x for x in c1)
    assert any("Zadaná nula" in x for x in c2)                                # skutečná nula
    pn = one_panel(res)
    assert not any("penále" in m for m in pn["missing"])                      # neuplatňuje se → nedoplňovat
    assert not any("penále" in m for m in res["missing_relevant"])


def test_unknown_never_shown_as_zero_in_panel():
    pn = one_panel(run(model(), [rent("NP-1", 1, 5, 6500), rent("NP-2", 1, 3, 9000)]))
    for v in pn["variants"]:
        if not v["nothing"]:
            assert "nezadáno" in v["amount"] and v["amount"] != "0 Kč"


def test_important_warnings_survive_shortening():
    cases = [(model(var_rental=1500, var_job=2000, includes="ne"), [job("NP-1", 1, 10), rent("NP-2", 1, 5, 6500)]),
             (model(var_rental=7000), [rent("NP-1", 1, 3, 6000), rent("NP-2", 2, 4, 6500)]),
             (model(var_rental=500), [rent("NP-1", 1, 3, 500), rent("NP-2", 2, 4, 0)]),
             (model(), [rent("NP-1", 1, 5, 6500), rent("NP-2", 1, 3, 9000)]),
             (model(var_rental=0, includes="ano"),
              [pv.job_request("NP-1", "J", [(BAGR, 1)], 1, 3, None, None, None, 1000, None), rent("NP-2", 2, 4, 6500)])]
    for m, reqs in cases:
        res = run(m, reqs)
        pn = one_panel(res)
        assert "nezamítá" in pn["note"]                                      # jednou za skupinu
        if pn["kind"] in ("conditional", "undecided"):
            assert "nepočítá jako 0 Kč" in pn["missing_note"] and pn["missing"]
        for c in res["cards"]:
            assert "Ekonomické porovnání" in c["econ_why"]                   # karta odkazuje na panel
            for line in c["economy_lines"]:
                assert "nejvýše" not in line or c["economy"]["status"] != pv.EXACT


# ---------------------------------------------------------------------------
# Vazba panel ↔ karty, zelená jen z doporučení, HTML, mezipaměť, data
# ---------------------------------------------------------------------------

def test_panel_and_cards_are_linked_random():
    rnd = random.Random(44)
    for _ in range(30):
        reqs = []
        for i in range(rnd.randint(2, 5)):
            s = rnd.randint(1, 8)
            reqs.append(rent(f"NP-{i + 1}", s, s + rnd.randint(0, 3), rnd.choice([None, 0, 3000, 9000])))
        m = model(units=rnd.randint(1, 2), var_rental=rnd.choice([None, 0, 1500, 8000]))
        res = run(m, reqs)
        p = pl.plan(m, {}, [pv.to_request(r) for r in reqs])
        competing = [g for g in p.groups if g["verdict"] != pl.ALL_FIT]
        assert sorted(sorted(x["requests"]) for x in res["panels"]) == sorted(sorted(g["requests"]) for g in competing)
        for c in res["cards"]:
            if c["group_refs"] is None:
                assert c["econ_badge"] in (pv.EC_NONE, None)
                continue
            mine = [x for x in res["panels"] if set(x["requests"]) == set(c["group_refs"])]
            assert len(mine) == 1 and c["ref"] in mine[0]["requests"]
            g = next(g for g in competing if c["ref"] in g["requests"])
            assert mine[0]["badge"] == pv.PANEL_BADGE[pv.VERDICT_KIND[g["verdict"]]]
            if c["econ_badge"] == pv.EC_REC:
                assert g["verdict"] == pl.RECOMMEND and c["ref"] in g["choice"]
            if mine[0]["badge"] == pv.EC_REC:
                assert g["verdict"] == pl.RECOMMEND
            if any(" méně." in x for x in mine[0]["why"]):
                ge = next(x["economy"] for x in res["groups"] if set(x["requests"]) == set(g["requests"]))
                assert ge["all_exact"]                                        # rozdíly jen z úplných dat


def test_panel_html_escapes_user_text_and_has_no_tables():
    req = pv.job_request("NP-1", "[Klik](https://example.com)<b>x</b>", [(BAGR, 1)], 1, 3, None, None, None, 1000,
                         None)
    out = pv.panel_html(one_panel(run(model(var_rental=0, includes="ano"), [req, rent("NP-2", 2, 4, 6500)])))
    assert "<b>x" not in out and "](https://example.com)" not in out and "&lt;b&gt;x" in out
    assert "<table" not in out and "Ekonomické porovnání" in out


def test_panel_css_is_mobile_safe():
    css = pv.CSS
    assert ".dp-panel{" in css and "overflow-wrap:anywhere" in css.split(".dp-panel{")[1].split("}")[0]
    assert ".dp-var{display:flex;flex-wrap:wrap" in css


def test_panels_are_plain_and_pickle_safe():
    res = run(model(), [job("NP-1", 1, 10), rent("NP-2", 1, 5, 6500), rent("NP-3", 6, 10, 7000)])
    for pn in res["panels"]:
        assert pickle.loads(pickle.dumps(pn)) == pn
        stack = [pn]
        while stack:
            x = stack.pop()
            assert isinstance(x, (dict, list, str, int, float, bool, type(None))), type(x)
            stack += list(x.values()) if isinstance(x, dict) else x if isinstance(x, list) else []


def test_hidden_variants_counted():
    reqs = [rent(f"NP-{i}", 1 + i, 3 + i, 5000 + 100 * i) for i in range(1, 6)]
    pn = one_panel(run(model(var_rental=0), reqs))
    shown_rivals = [v for v in pn["variants"] if v["role"] == "horší" and not v["nothing"]]
    assert len(shown_rivals) <= 2 and pn["hidden_total"] >= len(pn["hidden"])
    assert "Další varianty" in text(pn) if pn["hidden_total"] else True


def test_data_unchanged():
    before = files_hash()
    run(DEMO, [rent("NP-1", 22, 23, None, "PASOVY_DUMPER")], STARTS)
    assert files_hash() == before



# ---------------------------------------------------------------------------
# Regresní testy z auditu PR #33
# ---------------------------------------------------------------------------

def test_audit_invalid_value_stays_in_missing_list():
    # samostatný požadavek s neplatnou sazbou: „Chybějící údaje“ ji musí dál uvádět (jako před 4C)
    m = model()
    m.machine_types.loc[m.machine_types["type_id"] == BAGR, "variable_cost_rental_day"] = "abc"
    res = run(m, [rent("NP-1", 1, 3, 6000)])
    assert any("abc" in x for x in res["missing_relevant"])
    assert res["cards"][0]["economy_lines"][0].startswith("Příspěvek nelze spočítat")


def test_audit_group_over_planner_limit_has_no_empty_variant_list():
    old = pl.MAX_GROUP
    pl.MAX_GROUP = 2
    try:
        pn = one_panel(run(model(), [rent(f"NP-{i}", 1, 3, 5000 + i) for i in range(1, 4)]))
    finally:
        pl.MAX_GROUP = old
    out = text(pn)
    assert pn["kind"] == "undecided" and pn["variants"] == [] and "limit" in pn["headline"]
    assert "Porovnávané varianty" not in out and "neporovnávaly" in out and "vždy i „nepřijmout nic“" not in out



def test_audit_panel_claims_hold_for_any_value_of_unknown_costs():
    # „Doporučeno“, „Rovnocenné“, „Nevyplatí se“ i rozdíly v Kč musí platit pro každé dosazení
    # nezadaných sazeb (dosazuje se jen to, co v modelu chybí)
    import re
    from test_econ_formula import filled
    num = lambda t: float(t.replace(" ", "").replace(NBSP, "").replace("−", "-"))
    rnd = random.Random(91)
    checked = 0
    for _ in range(40):
        reqs = []
        for i in range(rnd.randint(2, 4)):
            s = rnd.randint(1, 7)
            if rnd.random() < 0.4:
                reqs.append(job(f"NP-{i + 1}", s, rnd.randint(1, 4), rnd.choice([20000, 60000, 120000]),
                                rnd.choice([0, 10000, 50000])))
            else:
                reqs.append(rent(f"NP-{i + 1}", s, s + rnd.randint(0, 3), rnd.choice([None, 0, 4000, 9000])))
        vr, vj, inc0 = rnd.choice([None, 0, 1500]), rnd.choice([None, 0, 2500]), rnd.choice([None, "ano", "ne"])
        m = model(units=rnd.randint(1, 2), var_rental=vr, var_job=vj, includes=inc0)
        res = run(m, reqs)
        p = pl.plan(m, {}, [pv.to_request(r) for r in reqs])
        by = {r["ref"]: pv.to_request(r) for r in reqs}
        for pn in res["panels"]:
            g = next(g for g in p.groups if set(g["requests"]) == set(pn["requests"]))
            for xr in ((0, 700, 15000) if vr is None else (None,)):
                for xj in ((0, 3000, 20000) if vj is None else (None,)):
                    for inc in (["ano", "ne"] if inc0 is None else [None]):
                        mf = filled(m, rental=None if xr is None else {BAGR: xr},
                                    job_rate=None if xj is None else {BAGR: xj}, includes=inc)
                        vals = {v: sum(pl.contribution(mf, by[x]).lo for x in v) for v in g["candidates"]}
                        top = max(vals.values())
                        if pn["kind"] == "recommend":
                            ch = tuple(g["choice"])
                            assert all(vals[ch] > x for v, x in vals.items() if v != ch)
                            for w in pn["why"]:
                                mm = re.search(r"přinese (.+?) Kč, nejlepší jiná varianta \((.+)\) (.+?) Kč – o (.+?) Kč",
                                               w)
                                if mm:
                                    second = max(x for v, x in vals.items() if v != ch and v)
                                    assert abs(num(mm.group(1)) - vals[ch]) < 1 and abs(num(mm.group(3)) - second) < 1
                                    assert abs(num(mm.group(4)) - (vals[ch] - second)) < 1
                                    checked += 1
                        elif pn["kind"] == "none_pays":
                            assert all(x < 0 for v, x in vals.items() if v)
                        elif pn["kind"] == "equal":
                            assert all(abs(vals[tuple(b)] - top) < 1e-6 for b in g["best"])
    assert checked > 10



# ---------------------------------------------------------------------------
# Hranice zlomu: v hranici rovnost, ne vítězství (kontrola po 4C)
# ---------------------------------------------------------------------------

def test_boundary_fikt_p07_vs_np3_states_equality_at_threshold():
    # FIKT-P07 = 37 500 − 5x, NP-3 = 8 000 − 4x, nic = 0 → při x = 7 500 je FIKT-P07 přesně 0 Kč jako „nic“
    from test_econ_formula import filled
    res = run(DEMO, [rent("NP-3", 20, 23, 2000, "SMYKOVY_NAKLADAC")], STARTS)
    pn = next(x for x in res["panels"] if "NP-3" in x["requests"])
    lines = pn["conditions"][0]["lines"]
    assert lines == ["pod 7 500 Kč/den: FIKT-P07",
                     "přesně 7 500 Kč/den: FIKT-P07 a nepřijmout nic vycházejí stejně (rovnocenné)",
                     "nad 7 500 Kč/den: nepřijmout nic",
                     "NP-3: nevychází nejlépe při žádné výši nákladu"]       # NP-3 v hranici neremizuje
    cards = {c["ref"]: c for c in res["cards"]}
    assert "pod 7 500 Kč/den (přesně při 7 500 Kč/den stejně jako nepřijmout nic)" in cards["FIKT-P07"]["econ_why"]
    assert "stejně jako nejlepší" not in cards["NP-3"]["econ_why"]
    p = pl.plan(DEMO, STARTS, [pv.to_request(rent("NP-3", 20, 23, 2000, "SMYKOVY_NAKLADAC"))])
    p07 = p.requests["FIKT-P07"]["request"]
    for x, expected in ((7499, 5.0), (7500, 0.0), (7501, -5.0)):
        assert pl.contribution(filled(DEMO, rental={"SMYKOVY_NAKLADAC": x}), p07).lo == expected


def test_non_integer_threshold_is_not_rounded_across_the_boundary():
    # NP-1 = 42 700 − 7x, NP-2 = 36 000 − 4x → hranice 6 700 / 3 = 2 233,33 Kč/den
    pn = one_panel(run(model(), [rent("NP-1", 1, 7, 6100), rent("NP-2", 1, 4, 9000)]))
    lines = pn["conditions"][0]["lines"]
    assert lines[0] == "pod 2 233,33 Kč/den: NP-1"
    assert "přesně 2 233,33 Kč/den: NP-1 a NP-2 vycházejí stejně (rovnocenné)" in lines
    assert ef.rate_text(2233.3333) == "2 233,33 Kč" and ef.rate_text(7500.0) == "7 500 Kč"


def test_boundaries_match_true_values_random():
    from test_econ_formula import filled
    rnd = random.Random(77)
    checked = 0
    for _ in range(40):
        reqs = []
        for i in range(rnd.randint(2, 4)):
            s = rnd.randint(1, 6)
            reqs.append(rent(f"NP-{i + 1}", s, s + rnd.randint(0, 4), rnd.choice([None, 2000, 4500, 6100, 9000])))
        m = model(units=rnd.randint(1, 2))                                  # jediná neznámá: sazba pronájmu
        res = run(m, reqs)
        p = pl.plan(m, {}, [pv.to_request(r) for r in reqs])
        by = {r["ref"]: pv.to_request(r) for r in reqs}
        for x in res["groups"]:
            ge = x["economy"]
            if ge["kind"] != "conditional" or not ge["profiles"]:
                continue
            g = next(g for g in p.groups if set(g["requests"]) == set(ge["requests"]))
            label = {pv._vlabel(v): v for v in g["candidates"]}

            def best_at(rate):
                vals = {v: sum(pl.contribution(filled(m, rental={BAGR: rate}), by[r]).lo for r in v)
                        for v in g["candidates"]}
                top = max(vals.values())
                return {pv._vlabel(v) for v, val in vals.items() if abs(val - top) < 1e-6}
            prof = ge["profiles"][0]
            for b in prof["boundaries"]:                                     # přesně v hranici: rovnost
                assert set(b["winners"]) == best_at(b["at"]) and len(b["winners"]) >= 2
                checked += 1
            for seg in prof["segments"]:                                     # uvnitř úseku: vítěz úseku
                probe = seg["start"] + 0.5 if seg["end"] is None else (seg["start"] + seg["end"]) / 2
                assert set(seg["winners"]) == best_at(probe)
            for line in pn_lines(res, ge):
                assert not line.startswith("do ")                            # žádné nejednoznačné „do“
            assert all(w in label for seg in prof["segments"] for w in seg["winners"])
    assert checked > 10


def pn_lines(res, ge):
    pn = next(x for x in res["panels"] if set(x["requests"]) == set(ge["requests"]))
    return [line for block in pn["conditions"] for line in block["lines"]]


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK", name)
