"""Testy štítků a karet požadavků záložky 8 (planner_view, E2.5c krok 4B). Jen fiktivní data.

Spuštění: python -m pytest tests   nebo   python tests/test_planner_view_cards.py
"""

import random
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))
import planner as pl  # noqa: E402
import planner_view as pv  # noqa: E402
from test_econ_formula import files_hash  # noqa: E402
from test_planner import BAGR, model  # noqa: E402

GREEN = "#1E7B4A"


def rent(ref, s, e, price=None):
    return pv.rental_request(ref, BAGR, 1, s, e, price)


def job(ref, s, dur, revenue=150000, direct=60000, earliest=None, deadline=None, penalty=None):
    return pv.job_request(ref, ref, [(BAGR, 1)], s, dur, earliest, deadline, revenue, direct, penalty)


def cards(m, reqs):
    res = pv.evaluate(m, {}, reqs)
    return {c["ref"]: c for c in res["cards"]}


def badges(cs):
    return {ref: (c["op_badge"], c["econ_badge"]) for ref, c in cs.items()}


# ---------------------------------------------------------------------------
# Pět ekonomických výsledků + samostatný požadavek
# ---------------------------------------------------------------------------

def test_recommended_and_others():
    cs = cards(model(var_rental=1500, var_job=2000, includes="ne"),
               [job("NP-1", 1, 10), rent("NP-2", 1, 5, 6500), rent("NP-3", 6, 10, 7000)])
    assert badges(cs) == {"NP-1": (pv.OP_FITS, pv.EC_REC), "NP-2": (pv.OP_FITS, pv.EC_NO),
                          "NP-3": (pv.OP_FITS, pv.EC_NO)}
    assert "prokazatelně ekonomicky nejlepší" in cs["NP-1"]["econ_why"]
    assert "„NP-1“" in cs["NP-2"]["econ_why"] and "nezamítá" in cs["NP-2"]["econ_why"]
    assert cs["NP-1"]["economy_lines"] == ["Příspěvek: 70 000 Kč."]


def test_none_pays_explains_it_is_not_a_rejection():
    cs = cards(model(var_rental=7000), [rent("NP-1", 1, 3, 6000), rent("NP-2", 2, 4, 6500)])
    assert {b[1] for b in badges(cs).values()} == {pv.EC_NO}
    for c in cs.values():
        assert "nepřijmout žádný" in c["econ_why"]
        assert "nezamítá" in c["econ_why"] and "žádná rezervace se neruší" in c["econ_why"]
    assert cs["NP-1"]["economy_lines"] == ["Příspěvek: −3 000 Kč – ztráta."]


def test_equal_with_nothing():
    cs = cards(model(var_rental=500), [rent("NP-1", 1, 3, 500), rent("NP-2", 2, 4, 0)])
    assert badges(cs) == {"NP-1": (pv.OP_FITS, pv.EC_EQ), "NP-2": (pv.OP_FITS, pv.EC_NO)}
    assert "„nepřijmout nic“ (0 Kč)" in cs["NP-1"]["econ_why"]
    assert cs["NP-1"]["economy_lines"][0] == "Příspěvek: přesně 0 Kč (vychází ze zadaných údajů)."


def test_conditional_with_break_even_in_kc_per_day():
    cs = cards(model(), [rent("NP-1", 1, 5, 6500), rent("NP-2", 1, 3, 9000)])
    assert badges(cs) == {"NP-1": (pv.OP_FITS, pv.EC_COND), "NP-2": (pv.OP_FITS, pv.EC_COND)}
    assert "do 2 750 Kč/den" in cs["NP-1"]["econ_why"]
    assert "2 750–9 000 Kč/den" in cs["NP-2"]["econ_why"]
    assert "nepočítá jako 0 Kč" in cs["NP-1"]["econ_why"]


def test_conditional_never_winner_is_not_shown_as_conditional():
    # NP-1 (2 dny á 8 000 = 16 000 − 2x) vs NP-2 (5 dní á 8 000 = 40 000 − 5x): všechny se rovnají při 8 000 Kč/den;
    # NP-1 nevyhrává v žádném úseku → „Nevyplatí se“ (nanejvýš shodně na hranici)
    cs = cards(model(), [rent("NP-1", 2, 3, 8000), rent("NP-2", 1, 5, 8000)])
    assert badges(cs) == {"NP-1": (pv.OP_FITS, pv.EC_NO), "NP-2": (pv.OP_FITS, pv.EC_COND)}
    assert "nanejvýš shodně" in cs["NP-1"]["econ_why"]


def test_conditional_with_several_unknowns_does_not_invent_one_threshold():
    cs = cards(model(), [job("NP-1", 1, 10), rent("NP-2", 1, 5, 6500), rent("NP-3", 6, 10, 7000)])
    assert cs["NP-1"]["econ_badge"] == pv.EC_COND
    why = cs["NP-1"]["econ_why"]
    assert "rozhoduje víc neznámých najednou" in why and "hranice" not in why.split("neobsahují stroje")[1][:40]


def test_undecided():
    req = pv.job_request("NP-1", "J", [(BAGR, 1)], 1, 3, None, None, None, 1000, None)
    cs = cards(model(var_rental=0, includes="ano"), [req, rent("NP-2", 2, 4, 6500)])
    assert badges(cs) == {"NP-1": (pv.OP_FITS, pv.EC_UND), "NP-2": (pv.OP_FITS, pv.EC_UND)}
    assert "tržba" in cs["NP-1"]["econ_why"]
    assert cs["NP-1"]["economy_lines"][0] == "Příspěvek: −1 000 Kč + NP-1 (J): tržba (Kč, nezadáno)."


def test_standalone_request_has_no_comparison():
    cs = cards(model(), [rent("NP-1", 1, 3, 6000)])
    assert badges(cs) == {"NP-1": (pv.OP_FITS, pv.EC_NONE)}
    assert "se doporučení nevydává" in cs["NP-1"]["econ_why"]


def test_operational_states_are_separate():
    cs = cards(model(bookings=[("R1", 1, 5, 1, "confirmed")]), [rent("NP-1", 3, 4)])
    assert badges(cs) == {"NP-1": (pv.OP_NO, None)} and cs["NP-1"]["economy_lines"] == []
    bad = [(f"{BAGR}-09", BAGR, "owned", "x", "")]
    cs = cards(model(extra_units=bad), [rent("NP-1", 3, 4)])
    assert badges(cs) == {"NP-1": (pv.OP_UNKNOWN, None)}


# ---------------------------------------------------------------------------
# Skutečná nula × neznámý × neplatný údaj, scénáře, penále
# ---------------------------------------------------------------------------

def test_unknown_cost_is_never_shown_as_zero():
    lines = cards(model(), [rent("NP-1", 1, 3, 0)])["NP-1"]["economy_lines"]
    text = " ".join(lines)
    assert "nezadáno" in text and "přesně 0 Kč" not in text
    assert lines[0].startswith("Příspěvek: 0 Kč − 3 strojodny × Mini bagr: variabilní náklad pronájmu (Kč/den")
    assert "Zadaná nula" in text and "cena pronájmu 0 Kč/den" in text           # cena 0 je skutečně zadaná


def test_true_zero_cost_is_labelled_as_entered():
    lines = cards(model(var_rental=0), [rent("NP-1", 1, 3, 6000)])["NP-1"]["economy_lines"]
    assert lines[0] == "Příspěvek: 18 000 Kč."
    assert "Zadaná nula (skutečná hodnota, ne chybějící údaj): Mini bagr: variabilní náklad pronájmu 0 Kč/den." in lines


def test_invalid_value_explains_why():
    m = model()
    m.machine_types.loc[m.machine_types["type_id"] == BAGR, "variable_cost_rental_day"] = "abc"
    lines = cards(m, [rent("NP-1", 1, 3, 6000)])["NP-1"]["economy_lines"]
    assert lines[0].startswith("Příspěvek nelze spočítat – neplatný údaj:") and "abc" in lines[0]


def test_exact_amount_never_says_at_most():
    for m, reqs in ((model(var_rental=1500, var_job=2000, includes="ne"), [job("NP-1", 1, 10)]),
                    (model(var_rental=7000), [rent("NP-1", 1, 3, 6000)]),
                    (model(var_rental=500), [rent("NP-1", 1, 3, 500)])):
        for c in cards(m, reqs).values():
            assert c["economy"]["status"] == pv.EXACT
            assert not any("nejvýše" in x for x in c["economy_lines"])


def test_both_scenarios_of_direct_cost_includes_machines():
    lines = cards(model(), [job("NP-1", 1, 10)])["NP-1"]["economy_lines"]
    assert lines[0].startswith("Příspěvek závisí na tom, zda přímé náklady zakázek obsahují vlastní stroje")
    assert "– pokud obsahují stroje: 90 000 Kč." in lines
    assert ("– pokud neobsahují stroje: 90 000 Kč − 10 strojodnů × Mini bagr: variabilní náklad na zakázce "
            "(Kč/den, nezadáno).") in lines
    # zadané „ano“: náklad stroje se neodečte podruhé, „ne“ + sazba: přesně
    assert cards(model(var_job=2000, includes="ano"), [job("NP-1", 1, 10)])["NP-1"]["economy_lines"][0] == \
        "Příspěvek: 90 000 Kč."
    assert cards(model(var_job=2000, includes="ne"), [job("NP-1", 1, 10)])["NP-1"]["economy_lines"][0] == \
        "Příspěvek: 70 000 Kč."


def test_late_job_penalty_explained_on_contribution_line():
    cs = cards(model(var_job=0, includes="ne"), [job("NP-1", 1, 6, earliest=1, deadline=4, penalty=1000)])
    assert "2 dny po termínu – penále 2 000 Kč" in cs["NP-1"]["economy_lines"][-1]
    cs = cards(model(var_job=0, includes="ne"), [job("NP-1", 1, 6, earliest=1, deadline=4)])
    assert "2 dny po termínu × NP-1 (NP-1): penále (Kč/den, nezadáno)" in cs["NP-1"]["economy_lines"][0]
    assert "sazba penále chybí" in cs["NP-1"]["economy_lines"][-1]


# ---------------------------------------------------------------------------
# Barvy, text štítků, HTML
# ---------------------------------------------------------------------------

def _css_rule(cls):
    return re.search(r"\.dp-badge\." + cls + r"\{([^}]*)\}", pv.CSS).group(1)


def test_green_only_for_recommended():
    assert GREEN in _css_rule("ec-rec")
    for cls in ("op-ok", "op-no", "op-und", "ec-no", "ec-eq", "ec-cond", "ec-und", "ec-none"):
        assert GREEN not in _css_rule(cls), cls
    assert "#7C3AED" in _css_rule("ec-cond") and "#2563EB" in _css_rule("ec-eq")
    assert "#B42318" in _css_rule("op-no") and "#B7791F" in _css_rule("ec-und") and "#B7791F" in _css_rule("op-und")
    assert "#64748B" in _css_rule("op-ok") and "#4B5563" in _css_rule("ec-no")


def test_card_html_shows_both_labelled_badges_and_no_legacy_green():
    cs = cards(model(var_rental=1500, var_job=2000, includes="ne"), [job("NP-1", 1, 10), rent("NP-2", 1, 5, 6500)])
    rec = pv.card_html(cs["NP-1"], {BAGR: "Mini bagr"})
    other = pv.card_html(cs["NP-2"], {BAGR: "Mini bagr"})
    assert "<small>Provoz:</small>✓ Vejde se" in rec and "<small>Ekonomika:</small>★ Doporučeno" in rec
    assert 'class="dp-badge ec-rec"' in rec and "ec-rec" not in other
    for out in (rec, other):
        assert 'dp-badge ok"' not in out and 'dp-badge rec"' not in out          # dřívější zelené třídy
        assert "<b>Provoz:</b>" in out and "<b>Ekonomika:</b>" in out
    assert "Příspěvek: 70\u00a0000\u00a0Kč." in rec                           # částka se nezalomí


def test_every_badge_has_text():
    for label in (pv.OP_FITS, pv.OP_NO, pv.OP_UNKNOWN, pv.EC_REC, pv.EC_NO, pv.EC_EQ, pv.EC_COND, pv.EC_UND,
                  pv.EC_NONE):
        assert re.search(r"[A-Za-zÁ-ž]{3}", label), label


# ---------------------------------------------------------------------------
# Štítky odvozené z výsledku plánovače, předběžné rezervace, data
# ---------------------------------------------------------------------------

def test_econ_badges_follow_planner_verdict_random():
    rnd = random.Random(48)
    for _ in range(25):
        reqs = []
        for i in range(rnd.randint(2, 4)):
            s = rnd.randint(1, 6)
            reqs.append(rent(f"NP-{i + 1}", s, s + rnd.randint(0, 3), rnd.choice([0, 2000, 6500, 9000])))
        m = model(var_rental=rnd.choice([None, 0, 1500, 7000]))
        p = pl.plan(m, {}, [pv.to_request(r) for r in reqs])
        cs = cards(m, reqs)
        for g in p.groups:
            for ref in g["requests"]:
                lab = cs[ref]["econ_badge"]
                if g["verdict"] == pl.ALL_FIT:
                    assert lab == pv.EC_NONE
                elif g["verdict"] == pl.RECOMMEND:
                    assert (lab == pv.EC_REC) == (ref in g["choice"])
                elif g["verdict"] == pl.NONE_PAYS:
                    assert lab == pv.EC_NO
                elif g["verdict"] == pl.EQUAL:
                    assert lab in (pv.EC_EQ, pv.EC_NO)
                elif g["verdict"] == pl.CONDITIONAL:
                    assert lab in (pv.EC_COND, pv.EC_NO)
                    if lab == pv.EC_COND:
                        assert any(ref in c for c in g["contenders"])
                else:
                    assert lab == pv.EC_UND
                assert lab != pv.EC_REC or g["verdict"] == pl.RECOMMEND        # zelená jen z doporučení


def test_tentative_booking_card_gets_badges_and_keeps_note():
    before = files_hash()
    m = model(var_rental=1000, bookings=[("FIKT-T", 1, 5, 1, pl.TENTATIVE)])
    cs = cards(m, [rent("NP-1", 3, 4, 9000)])
    assert badges(cs) == {"NP-1": (pv.OP_FITS, pv.EC_NO), "FIKT-T": (pv.OP_FITS, pv.EC_REC)}
    assert "nic se automaticky neruší" in cs["FIKT-T"]["note"]
    assert files_hash() == before


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK", name)
