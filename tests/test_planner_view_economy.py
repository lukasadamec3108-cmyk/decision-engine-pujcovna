"""Testy strukturovaného ekonomického výstupu záložky 8 (planner_view, E2.5c krok 4A).

Jen fiktivní testovací data; modely se skládají v paměti, datové soubory se nemění.
Spuštění: python -m pytest tests   nebo   python tests/test_planner_view_economy.py
"""

import pickle
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))
import datamodel as dm  # noqa: E402
import planner as pl  # noqa: E402
import planner_view as pv  # noqa: E402
from test_econ_formula import files_hash  # noqa: E402
from test_planner import BAGR, model  # noqa: E402

PLAIN = (dict, list, str, int, float, bool, type(None))


def plain(x, path="res", strict=False):
    """Jen obyčejné typy (žádné objekty plánovače ani dataclass) – bezpečné pro mezipaměť.
    Nová část „economy“ (strict) smí obsahovat jen dict/list/str/čísla/None; dřívější klíče
    (např. period, needs) mají z kompatibility i tuple."""
    strict = strict or path.endswith(".economy")
    allowed = PLAIN if strict else PLAIN + (tuple,)
    assert isinstance(x, allowed), f"{path}: {type(x).__name__}"
    if isinstance(x, dict):
        for k, v in x.items():
            assert isinstance(k, str), f"{path}: klíč {k!r}"
            plain(v, f"{path}.{k}", strict)
    elif isinstance(x, (list, tuple)):
        for i, v in enumerate(x):
            plain(v, f"{path}[{i}]", strict)


def rent(ref, s, e, price=None):
    return pv.rental_request(ref, BAGR, 1, s, e, price)


def job(ref, s, dur, revenue=150000, direct=60000):
    return pv.job_request(ref, ref, [(BAGR, 1)], s, dur, None, None, revenue, direct, None)


def econ(res):
    gs = [g for g in res["groups"] if g["verdict"] != pl.ALL_FIT]
    assert len(gs) == 1
    return gs[0]["economy"]


def card(res, ref):
    return next(c for c in res["cards"] if c["ref"] == ref)


# ---------------------------------------------------------------------------
# Pět ekonomických výsledků
# ---------------------------------------------------------------------------

def test_recommend_structure():
    res = pv.evaluate(model(var_rental=1500, var_job=2000, includes="ne"), {},
                      [job("NP-1", 1, 10), rent("NP-2", 1, 5, 6500), rent("NP-3", 6, 10, 7000)])
    e = econ(res)
    assert e["kind"] == "recommend" and e["choice"] == ["NP-1"] and e["best"] == [["NP-1"]]
    top = e["variants"][0]
    assert top["refs"] == ["NP-1"] and top["role"] == "doporučeno"
    assert top["contribution"]["status"] == pv.EXACT and top["contribution"]["value"] == 70000.0
    nothing = next(v for v in e["variants"] if v["nothing"])
    assert nothing["label"] == "nepřijmout nic" and nothing["contribution"]["value"] == 0.0
    assert next(v for v in e["variants"] if v["refs"] == ["NP-2", "NP-3"])["contribution"]["value"] == 52500.0
    assert e["missing"] == [] and e["profiles"] == [] and e["conditions"] == []


def test_none_pays_structure():
    e = econ(pv.evaluate(model(var_rental=7000), {}, [rent("NP-1", 1, 3, 6000), rent("NP-2", 2, 4, 6500)]))
    assert e["kind"] == "none_pays" and e["choice"] == [] and e["best"] == [[]]
    assert e["variants"][0]["nothing"] and e["variants"][0]["role"] == "doporučeno"
    losses = {tuple(v["refs"]): v["contribution"]["value"] for v in e["variants"] if not v["nothing"]}
    assert losses == {("NP-1",): -3000.0, ("NP-2",): -1500.0}


def test_equal_structure_zero_ties_with_nothing():
    e = econ(pv.evaluate(model(var_rental=500), {}, [rent("NP-1", 1, 3, 500), rent("NP-2", 2, 4, 0)]))
    assert e["kind"] == "equal" and e["choice"] is None
    assert sorted(e["best"]) == [[], ["NP-1"]]
    roles = {v["label"]: v["role"] for v in e["variants"]}
    assert roles["NP-1"] == "nejlepší" and roles["nepřijmout nic"] == "nejlepší" and roles["NP-2"] == "horší"
    zero = next(v for v in e["variants"] if v["refs"] == ["NP-1"])["contribution"]
    assert zero["status"] == pv.EXACT and zero["value"] == 0.0                  # skutečná nula, ne neznámá


def test_conditional_single_unknown_profile_in_kc_per_day():
    e = econ(pv.evaluate(model(), {}, [rent("NP-1", 1, 5, 6500), rent("NP-2", 1, 3, 9000)]))
    assert e["kind"] == "conditional" and e["choice"] is None and e["best"] == []
    (prof,) = e["profiles"]
    assert prof["unit"] == pv.KC_DEN and prof["unknown"]["unit"] == pv.KC_DEN
    assert prof["unknown"]["label"] == "Mini bagr: variabilní náklad pronájmu/den"
    assert [(s["start"], s["end"], s["winners"]) for s in prof["segments"]] == [
        (0.0, 2750.0, ["NP-1"]), (2750.0, 9000.0, ["NP-2"]), (9000.0, None, ["nepřijmout nic"])]
    assert [m["key"] for m in e["missing"]] == [["variable_cost_rental_day", BAGR]]
    a = next(v for v in e["variants"] if v["refs"] == ["NP-1"])["contribution"]
    assert a["status"] == pv.WITH_UNKNOWN and a["value"] is None and a["hi"] == 32500.0 and a["lo"] is None
    (sc,) = a["scenarios"]
    assert sc["known"] == 32500.0 and sc["unit"] == pv.KC
    assert sc["terms"] == [dict(key=["variable_cost_rental_day", BAGR], label="Mini bagr: variabilní náklad pronájmu/den",
                                kind="sazba", unit=pv.KC_DEN, coef=-5.0, coef_unit="strojodnů")]


def test_conditional_multiple_unknowns_gives_structured_pairs_not_single_threshold():
    res = pv.evaluate(model(), {}, [job("NP-1", 1, 10), rent("NP-2", 1, 5, 6500), rent("NP-3", 6, 10, 7000)])
    e = econ(res)
    assert e["kind"] == "conditional" and e["scenarios"] == ["ano", "ne"]
    # ve scénáři „ne“ víc neznámých → podmínky dvojic, žádné jediné pásmo
    assert [p["scenario"] for p in e["profiles"]] == ["ano"]
    pairs = [c for c in e["conditions"] if c["scenario"] == "ne"]
    assert pairs and all(c["a"] == ["NP-1"] for c in pairs)
    vs_rentals = next(c for c in pairs if c["b"] == ["NP-2", "NP-3"])
    assert vs_rentals["known"] == 22500.0 and vs_rentals["unit"] == pv.KC
    assert {(t["key"][0], t["coef"], t["unit"]) for t in vs_rentals["terms"]} == {
        ("variable_cost_job_day", -10.0, pv.KC_DEN), ("variable_cost_rental_day", 10.0, pv.KC_DEN)}
    # rozdíl odpovídá vzorcům variant ve stejném scénáři
    var = {tuple(v["refs"]): v["contribution"] for v in e["variants"]}
    known = lambda refs: next(s["known"] for s in var[refs]["scenarios"] if s["scenario"] in ("ne", ""))
    assert vs_rentals["known"] == known(("NP-1",)) - known(("NP-2", "NP-3"))
    labels = [m["label"] for m in e["missing"]]
    assert "zda přímé náklady zakázek obsahují vlastní stroje" in labels and len(labels) == len(set(labels))


def test_undecided_structure_names_blocking_data():
    req = pv.job_request("NP-1", "J", [(BAGR, 1)], 1, 3, None, None, None, 1000, None)       # chybí tržba
    e = econ(pv.evaluate(model(var_rental=0, includes="ano"), {}, [req, rent("NP-2", 2, 4, 6500)]))
    assert e["kind"] == "undecided" and e["choice"] is None
    assert any("tržba" in b for b in e["blocking"])
    rev = next(m for m in e["missing"] if m["key"] == ["revenue", "NP-1"])
    assert rev["unit"] == pv.KC and rev["kind"] == "údaj"


# ---------------------------------------------------------------------------
# Skutečná nula × nezadáno × neplatné, jednotky
# ---------------------------------------------------------------------------

def test_inputs_distinguish_zero_missing_invalid():
    def inputs(m):
        return {i["field"]: i for i in card(pv.evaluate(m, {}, [rent("NP-1", 1, 3, 6000)]), "NP-1")["economy"]["inputs"]}
    zero = inputs(model(var_rental=0))["variable_cost_rental_day"]
    assert zero["state"] == pv.ZADANO and zero["value"] == 0.0 and zero["unit"] == pv.KC_DEN
    missing = inputs(model())["variable_cost_rental_day"]
    assert missing["state"] == pv.NEZADANO and missing["value"] is None
    bad = model()
    bad.machine_types.loc[bad.machine_types["type_id"] == BAGR, "variable_cost_rental_day"] = "abc"
    res = pv.evaluate(bad, {}, [rent("NP-1", 1, 3, 6000)])
    inv = {i["field"]: i for i in card(res, "NP-1")["economy"]["inputs"]}["variable_cost_rental_day"]
    assert inv["state"] == pv.NEPLATNE and inv["value"] is None and "abc" in inv["reason"]
    c = card(res, "NP-1")["economy"]
    assert c["status"] == pv.INVALID and c["value"] is None and c["invalid"]


def test_card_contribution_zero_vs_unknown():
    c0 = card(pv.evaluate(model(var_rental=0), {}, [rent("NP-1", 1, 3, 0)]), "NP-1")["economy"]
    assert c0["status"] == pv.EXACT and c0["value"] == 0.0
    cu = card(pv.evaluate(model(), {}, [rent("NP-1", 1, 3, 0)]), "NP-1")["economy"]
    assert cu["status"] == pv.WITH_UNKNOWN and cu["value"] is None                # nikdy 0 Kč
    assert cu["scenarios"][0]["known"] == 0.0 and cu["scenarios"][0]["terms"]


def test_job_inputs_and_scenarios():
    c = card(pv.evaluate(model(), {}, [job("NP-1", 1, 4)]), "NP-1")["economy"]
    fields = {i["field"]: i for i in c["inputs"]}
    assert fields["revenue"]["value"] == 150000.0 and fields["revenue"]["unit"] == pv.KC
    assert fields["penalty_per_day"]["state"] == pv.NEZADANO and fields["penalty_per_day"]["unit"] == pv.KC_DEN
    assert fields["direct_cost_includes_machines"]["state"] == pv.NEZADANO
    assert {s["scenario"] for s in c["scenarios"]} == {"ano", "ne"}
    c = card(pv.evaluate(model(includes="ano"), {}, [job("NP-1", 1, 4)]), "NP-1")["economy"]
    assert {i["field"]: i for i in c["inputs"]}["direct_cost_includes_machines"]["value"] == "ano"
    assert c["status"] == pv.EXACT and c["value"] == 90000.0


def test_not_feasible_card_has_no_economy():
    m = model(bookings=[("R1", 1, 5, 1, "confirmed")])
    assert card(pv.evaluate(m, {}, [rent("NP-1", 3, 4)]), "NP-1")["economy"] is None


# ---------------------------------------------------------------------------
# Mezipaměť, kompatibilita, data, předběžné rezervace
# ---------------------------------------------------------------------------

def test_result_is_plain_and_pickle_safe_for_all_verdicts():
    cases = [(model(var_rental=1500, var_job=2000, includes="ne"), [job("NP-1", 1, 10), rent("NP-2", 1, 5, 6500)]),
             (model(var_rental=7000), [rent("NP-1", 1, 3, 6000), rent("NP-2", 2, 4, 6500)]),
             (model(var_rental=500), [rent("NP-1", 1, 3, 500), rent("NP-2", 2, 4, 0)]),
             (model(), [job("NP-1", 1, 10), rent("NP-2", 1, 5, 6500), rent("NP-3", 6, 10, 7000)]),
             (model(), [rent("NP-1", 1, 3, 6000)])]
    for m, reqs in cases:
        res = pv.evaluate(m, {}, reqs)
        res = {k: v for k, v in res.items() if k != "explanation"}
        plain(res)
        assert pickle.loads(pickle.dumps(res)) == res


def test_existing_keys_unchanged():
    res = pv.evaluate(model(), {}, [rent("NP-1", 1, 5, 6500), rent("NP-2", 1, 3, 9000)])
    assert set(res) == {"cards", "groups", "manual", "missing", "notes", "explanation", "tentative"}
    g = res["groups"][0]
    assert set(g) == {"requests", "verdict", "text", "options", "choice", "economy"}
    c = card(res, "NP-1")
    assert {"ref", "title", "badge", "why", "operational", "economic", "blocking", "alternatives", "note",
            "economy"} <= set(c)


def test_variant_formulas_agree_with_planner():
    m = model(var_rental=1200, var_job=900, includes="ne")
    reqs = [rent("NP-1", 1, 3, 6000), rent("NP-2", 2, 4, 7000), job("NP-3", 3, 4, 40000, 10000)]
    res = pv.evaluate(m, {}, reqs)
    p = pl.plan(m, {}, [pv.to_request(r) for r in reqs])
    g = next(x for x in p.groups if x["verdict"] != pl.ALL_FIT)
    e = econ(res)
    for v in e["variants"]:
        val = g["values"].get(tuple(v["refs"]))
        if val is not None:
            assert v["contribution"]["value"] == val.lo


def test_tentative_booking_in_group_and_data_unchanged():
    before = files_hash()
    m = model(var_rental=1000, bookings=[("FIKT-T", 1, 5, 1, pl.TENTATIVE)])
    e = econ(pv.evaluate(m, {}, [rent("NP-1", 3, 4, 9000)]))
    assert e["kind"] == "recommend" and e["choice"] == ["FIKT-T"] and "FIKT-T" in e["requests"]
    assert (m.rental_bookings["status"] == pl.TENTATIVE).all() and files_hash() == before
    demo = dm.load_model()
    assert files_hash() == before and len(demo.rental_bookings) == 7


def test_others_are_capped_but_counted():
    reqs = [rent(f"NP-{i}", 1 + i, 3 + i, 5000 + 100 * i) for i in range(5)]
    e = econ(pv.evaluate(model(var_rental=0), {}, reqs))
    shown = [v for v in e["variants"] if v["role"] == "horší"]
    assert len(shown) <= pv.MAX_OTHER_VARIANTS + 1                          # + „nic“ (vždy předané)
    assert e["others_total"] >= len([v for v in shown if not v["nothing"]])


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK", name)
