"""Testy příspěvku jako vzorce s neznámými údaji (econ_formula.py, E2.5c krok 2).

Jen fiktivní testovací data; modely se skládají v paměti, datové soubory se nemění.
Spuštění: python -m pytest tests   nebo   python tests/test_econ_formula.py
"""

import hashlib
import random
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))
import datamodel as dm  # noqa: E402
import econ_formula as ef  # noqa: E402
import planner as pl  # noqa: E402
from planner import Request  # noqa: E402
from test_planner import BAGR, BASE, job, model, rent  # noqa: E402

DUMPER = "PASOVY_DUMPER"
RATE_R = ("variable_cost_rental_day", BAGR)
RATE_J = ("variable_cost_job_day", BAGR)


def files_hash():
    return {f.name: hashlib.sha256(f.read_bytes()).hexdigest() for f in sorted((ROOT / "data" / "model").glob("*.csv"))}


def two_types(var_rental=None, var_job=None, includes=None):
    """Model s jedním mini bagrem a jedním pásovým dumperem (bez zakázek a rezervací)."""
    m = model(var_rental=var_rental, var_job=var_job, includes=includes)
    t = {n: getattr(m, n).copy() for n in dm.SCHEMA}
    t["fleet_units"] = pd.concat([t["fleet_units"], pd.DataFrame(
        [(f"{DUMPER}-01", DUMPER, "owned", "", "")], columns=list(BASE.fleet_units.columns))], ignore_index=True)
    return dm.DataModel(**t)


def filled(m, rental=None, job_rate=None, includes=None):
    """Kopie modelu s doplněnými sazbami (pro kontrolu vzorce proti economics.py)."""
    t = {n: getattr(m, n).copy() for n in dm.SCHEMA}
    if rental is not None:
        for typ, v in rental.items():
            t["machine_types"].loc[t["machine_types"]["type_id"] == typ, "variable_cost_rental_day"] = str(v)
    if job_rate is not None:
        for typ, v in job_rate.items():
            t["machine_types"].loc[t["machine_types"]["type_id"] == typ, "variable_cost_job_day"] = str(v)
    if includes is not None:
        s = t["settings"]
        t["settings"] = pd.concat([s[s["key"] != "direct_cost_includes_machines"], pd.DataFrame(
            [dict(key="direct_cost_includes_machines", value=includes, unit="", note="")])], ignore_index=True)
    return dm.DataModel(**t)


def total(m, reqs):
    return sum(pl.contribution(m, r).lo for r in reqs)


# ---------------------------------------------------------------------------
# Tvar vzorce, skutečná nula × chybějící údaj
# ---------------------------------------------------------------------------

def test_formula_is_known_amount_minus_usage_times_unknown_rate():
    c = ef.request_formula(model(), rent("A", 1, 5, q=1, price=6500))
    f = c.scenarios[ef.ONLY]
    assert f.known == 32500 and f.terms == ((RATE_R, -5.0),)      # 5 strojodnů × neznámá sazba
    assert ef.describe(c) == ["32 500 Kč − 5 strojodnů × [Mini bagr: variabilní náklad pronájmu/den]"]
    # 2 kusy → 10 strojodnů
    f2 = ef.request_formula(model(units=2), rent("A", 1, 5, q=2, price=6500)).scenarios[ef.ONLY]
    assert f2.known == 65000 and f2.terms == ((RATE_R, -10.0),)


def test_true_zero_is_not_an_unknown_but_missing_value_is():
    zero = ef.request_formula(model(var_rental=0), rent("A", 1, 5, price=6500))
    assert zero.scenarios[ef.ONLY] == ef.Formula(32500.0) and zero.scenarios[ef.ONLY].exact
    missing = ef.request_formula(model(), rent("A", 1, 5, price=6500))
    assert not missing.scenarios[ef.ONLY].exact and RATE_R in missing.unknowns
    # skutečná nulová cena je známá částka 0, ne neznámá
    free = ef.request_formula(model(var_rental=0), rent("A", 1, 5, price=0))
    assert free.scenarios[ef.ONLY] == ef.Formula(0.0)
    # chybějící údaj se nikdy nenahradí nulou: vzorec s neznámou ≠ přesná hodnota
    assert missing.scenarios[ef.ONLY] != ef.Formula(32500.0)
    assert ef.bounds(missing) == (None, 32500.0) and ef.bounds(zero) == (32500.0, 32500.0)


def test_invalid_value_cannot_build_formula():
    m = model()
    m.machine_types.loc[m.machine_types["type_id"] == BAGR, "variable_cost_rental_day"] = "-5"
    c = ef.request_formula(m, rent("A", 1, 5, price=6500))
    assert not c.valid and c.invalid
    r = ef.compare(c, ef.NOTHING)
    assert r["status"] == ef.UNDECIDABLE and "neplatné" in r["text"]


# ---------------------------------------------------------------------------
# Vyrušení společné neznámé, různé využití, více sazeb
# ---------------------------------------------------------------------------

def test_same_unknown_rate_in_both_variants_cancels():
    m = model()
    a = ef.variant_formula(m, [rent("A", 1, 5, price=6500)])                     # 32 500 − 5x
    b = ef.variant_formula(m, [rent("B", 1, 3, price=7000), rent("C", 4, 5, price=6000)])   # 33 000 − 5x
    r = ef.compare(a, b)
    assert r["status"] == ef.CERTAIN and r["winner"] == ef.B and r["strict"]
    assert r["scenarios"][ef.ONLY]["diff"] == ef.Formula(-500.0)               # x se vyrušilo
    # plánovač (E2.5c krok 3) využije vyrušení: A je vyloučená, zbývá jen B + C proti „nic“ (podmíněně)
    p = pl.plan(m, new_requests=[rent("A", 1, 5, price=6500), rent("B", 1, 3, price=7000),
                                 rent("C", 4, 5, price=6000)])
    g = p.groups[0]
    assert g["verdict"] == pl.CONDITIONAL and ("A",) not in g["contenders"]
    # B + C = 33 000 − 5x, B = 21 000 − 3x, C = 12 000 − 2x: nad 6 000 Kč je C ztrátový, nad 7 000 Kč i B
    assert [(lo, hi, w) for lo, hi, w in g["profiles"][0]["profile"]["segments"]] == [
        (0.0, 6000.0, ("B + C",)), (6000.0, 7000.0, ("B",)), (7000.0, None, ("nepřijmout nic",))]
    assert "do 6 000 Kč: B + C; 6 000 Kč–7 000 Kč: B; nad 7 000 Kč: nepřijmout nic" in g["text"]
    # přesná shoda: stejné strojodny i tržba
    r = ef.compare(ef.variant_formula(m, [rent("A", 1, 4, price=5000)]),
                   ef.variant_formula(m, [rent("B", 1, 2, price=5000), rent("C", 3, 4, price=5000)]))
    assert r["status"] == ef.CERTAIN and r["winner"] == ef.EQUAL


def test_different_usage_of_same_machine_gives_break_even():
    m = model()
    a = ef.variant_formula(m, [rent("A", 1, 5, price=6500)])     # 32 500 − 5x
    b = ef.variant_formula(m, [rent("B", 1, 3, price=9000)])     # 27 000 − 3x
    r = ef.compare(a, b)
    assert r["status"] == ef.CONDITIONAL and r["winner"] is None
    assert r["scenarios"][ef.ONLY]["diff"] == ef.Formula(5500.0, ((RATE_R, -2.0),))
    assert "< 2 750 Kč (hranice zlomu)" in r["text"]
    # kontrola proti economics.py s doplněnou sazbou na obou stranách hranice
    for x, better in ((2000, "A"), (3000, "B")):
        mf = filled(m, rental={BAGR: x})
        ta, tb = pl.contribution(mf, rent("A", 1, 5, price=6500)).lo, pl.contribution(mf, rent("B", 1, 3, price=9000)).lo
        assert (ta > tb) == (better == "A")
    # méně strojodnů a vyšší známá částka → jistě lepší (rozdíl roste se sazbou)
    c = ef.variant_formula(m, [rent("C", 1, 3, price=12000)])    # 36 000 − 3x
    r = ef.compare(c, a)
    assert r["status"] == ef.CERTAIN and r["winner"] == ef.A


def test_multiple_unknown_rates():
    m = two_types()
    a = ef.variant_formula(m, [Request("A", "rental", pl.NEW, ((BAGR, 1), (DUMPER, 1)), 1, 4, price_day=7000)])
    b = ef.variant_formula(m, [rent("B", 1, 4, price=6500)])
    r = ef.compare(a, b)
    # A: 2 stroje × 4 dny × 7 000 = 56 000 − 4·xBagr − 4·xDumper; B: 26 000 − 4·xBagr → rozdíl 30 000 − 4·xDumper
    diff = r["scenarios"][ef.ONLY]["diff"]
    assert diff == ef.Formula(30000.0, ((("variable_cost_rental_day", DUMPER), -4.0),))
    assert r["status"] == ef.CONDITIONAL and "7 500 Kč" in r["text"]
    # dvě různé sazby, které se nevyruší → podmínka se dvěma neznámými
    j = ef.variant_formula(m, [job("J", 1, 4, revenue=60000, direct=10000)])
    r = ef.compare(j, b)
    assert r["status"] == ef.CONDITIONAL
    text = r["text"]
    assert "variabilní náklad na zakázce/den" in text and "variabilní náklad pronájmu/den" in text


def test_formula_matches_economics_for_random_rates_and_scenarios():
    rnd = random.Random(52)
    m = two_types()
    reqs = [rent("R1", 1, 3, price=6000), Request("R2", "rental", pl.NEW, ((DUMPER, 1),), 2, 6, price_day=7000),
            Request("J1", "job", pl.NEW, ((BAGR, 1), (DUMPER, 1)), 1, 8, revenue=120000, direct_cost=30000,
                    earliest=1, deadline=6, penalty_per_day=2000)]
    for _ in range(25):
        rates = {t: rnd.choice([0, 300, 1500, 4000]) for t in (BAGR, DUMPER)}
        jrates = {t: rnd.choice([0, 800, 2500]) for t in (BAGR, DUMPER)}
        for sc in ("ano", "ne"):
            mf = filled(m, rental=rates, job_rate=jrates, includes=sc)
            values = {("variable_cost_rental_day", t): v for t, v in rates.items()}
            values.update({("variable_cost_job_day", t): v for t, v in jrates.items()})
            for r in reqs:
                c = ef.request_formula(m, r)
                f = c.scenarios.get(sc, c.scenarios.get(ef.ONLY))
                assert abs(f.value(values) - pl.contribution(mf, r).lo) < 1e-6
            v = ef.variant_formula(m, reqs).scenarios[sc]
            assert abs(v.value(values) - total(mf, reqs)) < 1e-6


# ---------------------------------------------------------------------------
# Scénáře direct_cost_includes_machines
# ---------------------------------------------------------------------------

def test_both_scenarios_when_setting_missing_and_no_double_counting():
    j = job("J", 1, 5, q=2, revenue=100000, direct=40000)            # 2 bagry × 5 dní = 10 strojodnů
    c = ef.request_formula(model(units=2), j)
    assert set(c.scenarios) == {"ano", "ne"}
    assert c.scenarios["ano"] == ef.Formula(60000.0)                     # stroje už v přímých nákladech
    assert c.scenarios["ne"] == ef.Formula(60000.0, ((RATE_J, -10.0),))  # odečtou se jednou
    # zadané nastavení → jediný scénář shodný s příslušnou větví
    assert ef.request_formula(model(units=2, includes="ano"), j).scenarios == {ef.ONLY: ef.Formula(60000.0)}
    assert ef.request_formula(model(units=2, includes="ne"), j).scenarios == {
        ef.ONLY: ef.Formula(60000.0, ((RATE_J, -10.0),))}
    # pronájem na nastavení nezávisí – v kombinaci platí pro oba scénáře stejný vzorec
    v = ef.variant_formula(model(units=2), [j, rent("R", 6, 8, price=5000)])
    assert v.scenarios["ano"] == ef.Formula(75000.0, ((RATE_R, -3.0),))
    assert v.scenarios["ne"] == ef.Formula.of(75000.0, {RATE_R: -3.0, RATE_J: -10.0})


def test_scenarios_compared_consistently():
    m = model()
    j = ef.variant_formula(m, [job("J", 1, 10, revenue=150000, direct=60000)])
    r = ef.variant_formula(m, [rent("R1", 1, 5, price=6500), rent("R2", 6, 10, price=7000)])
    res = ef.compare(j, r)
    assert res["status"] == ef.CONDITIONAL
    # „ano“: 90 000 − (67 500 − 10·xR) = 22 500 + 10·xR > 0 pro každé xR ≥ 0 → jistě A
    assert res["scenarios"]["ano"]["status"] == ef.CERTAIN and res["scenarios"]["ano"]["winner"] == ef.A
    # „ne“: 22 500 − 10·xJ + 10·xR → záleží na obou sazbách
    assert res["scenarios"]["ne"]["status"] == ef.CONDITIONAL
    # se zadaným nastavením „ano“ je výsledek jistý
    res_ano = ef.compare(ef.variant_formula(model(includes="ano"), [job("J", 1, 10, revenue=150000, direct=60000)]),
                         ef.variant_formula(model(includes="ano"), [rent("R1", 1, 5, price=6500),
                                                                     rent("R2", 6, 10, price=7000)]))
    assert res_ano["status"] == ef.CERTAIN and res_ano["winner"] == ef.A
    # stejný požadavek ve scénáři „ano“ i „ne“ se proti „nic“ chová podle scénáře
    vs_nothing = ef.compare(j, ef.NOTHING)
    assert vs_nothing["scenarios"]["ano"]["winner"] == ef.A
    assert vs_nothing["scenarios"]["ne"]["status"] == ef.CONDITIONAL
    assert "9 000 Kč (hranice zlomu)" in vs_nothing["text"]


# ---------------------------------------------------------------------------
# „Nepřijmout nic“, ztrátové varianty, chybějící základní údaje
# ---------------------------------------------------------------------------

def test_nothing_variant_is_exact_zero_and_in_group_candidates():
    assert ef.NOTHING.scenarios == {ef.ONLY: ef.ZERO} and ef.bounds(ef.NOTHING) == (0.0, 0.0)
    assert ef.compare(ef.NOTHING, ef.NOTHING)["winner"] == ef.EQUAL
    m = model()
    p = pl.plan(m, new_requests=[rent("A", 1, 3, price=6000), rent("B", 2, 4, price=7000)])
    g = p.groups[0]
    assert g["candidates"][0] == pl.NOTHING and ("A",) in g["candidates"] and ("B",) in g["candidates"]
    forms = ef.group_formulas(m, p, g)
    assert forms[()].scenarios == {ef.ONLY: ef.ZERO}


def test_all_variants_loss_making_nothing_is_certainly_better():
    # ztráta je jistá i s neznámými náklady: tržba − známé náklady < 0 a neznámé ji jen zvětší
    m = model()
    p = pl.plan(m, new_requests=[job("J1", 1, 3, revenue=10000, direct=20000),
                                 job("J2", 2, 4, revenue=5000, direct=9000)])
    g = p.groups[0]
    forms = ef.group_formulas(m, p, g)
    for v, c in forms.items():
        if v:
            r = ef.compare(c, forms[()])
            assert r["status"] == ef.CERTAIN and r["winner"] == ef.B, v
    # úplná data: ztrátové pronájmy jistě horší než „nic“
    m = model(var_rental=7000)
    for r in (rent("A", 1, 3, price=6000), rent("B", 2, 4, price=6500)):
        res = ef.compare(ef.variant_formula(m, [r]), ef.NOTHING)
        assert res["status"] == ef.CERTAIN and res["winner"] == ef.B and res["strict"]


def test_profit_vs_nothing_depends_on_unknown_rate_not_false_decision():
    # tržba 32 500, neznámý náklad 5 strojodnů → proti „nic“ jen podmíněně (hranice 6 500 Kč/den)
    r = ef.compare(ef.variant_formula(model(), [rent("A", 1, 5, price=6500)]), ef.NOTHING)
    assert r["status"] == ef.CONDITIONAL and "6 500 Kč (hranice zlomu)" in r["text"]


def test_missing_basic_data_is_undecidable_unless_it_cancels():
    m = model(var_rental=0)
    unknown_rev = Request("J", "job", pl.NEW, ((BAGR, 1),), 1, 3, name="J", revenue=None, direct_cost=1000)
    c = ef.request_formula(model(var_rental=0, includes="ano"), unknown_rev)
    assert c.scenarios[ef.ONLY].terms == ((("revenue", "J"), 1.0),)
    r = ef.compare(c, ef.variant_formula(m, [rent("B", 1, 3, price=6500)]))
    assert r["status"] == ef.UNDECIDABLE and "tržba" in r["text"]
    # stejná zakázka v obou variantách → neznámá tržba se vyruší
    m2 = model(var_rental=0, includes="ano")
    a = ef.variant_formula(m2, [unknown_rev, rent("B", 4, 5, price=6500)])
    b = ef.variant_formula(m2, [unknown_rev])
    r = ef.compare(a, b)
    assert r["status"] == ef.CERTAIN and r["winner"] == ef.A
    # chybějící sazba penále u zakázky v termínu nehraje roli (0 dní po termínu)
    on_time = Request("K", "job", pl.NEW, ((BAGR, 1),), 1, 3, revenue=9000, direct_cost=1000, deadline=5)
    assert ef.request_formula(model(var_job=0, includes="ne"), on_time).scenarios == {ef.ONLY: ef.Formula(8000.0)}
    late = Request("L", "job", pl.NEW, ((BAGR, 1),), 1, 6, revenue=9000, direct_cost=1000, deadline=4)
    f = ef.request_formula(model(var_job=0, includes="ne"), late).scenarios[ef.ONLY]
    assert f == ef.Formula(8000.0, ((("penalty_per_day", "L"), -2.0),))        # 2 dny po termínu


# ---------------------------------------------------------------------------
# Shoda s dosavadními výsledky, neměnnost dat
# ---------------------------------------------------------------------------

def _demo_requests(m):
    req = m.job_requirements[m.job_requirements["resource_kind"] == "machine"]
    out = []
    for j in m.jobs.itertuples():
        needs = tuple((str(r.resource_id), int(r.quantity)) for r in req[req["job_id"] == j.job_id].itertuples())
        s = int(j.earliest_start_day) + 2
        out.append(Request(str(j.job_id), "job", pl.NEW, needs, s, s + int(j.duration_days) - 1,
                           earliest=int(j.earliest_start_day), deadline=int(j.deadline_day), in_data=True))
    for b in m.rental_bookings.itertuples():
        out.append(Request(str(b.booking_id), "rental", pl.NEW, ((str(b.type_id), int(b.units)),),
                           int(b.start_day), int(b.end_day), price_day=float(b.price_day), in_data=True))
    return out


def test_complete_data_gives_exact_formula_equal_to_current_results():
    before = files_hash()
    rates = {t: 1000 + 100 * i for i, t in enumerate(BASE.machine_types["type_id"])}
    for includes in ("ano", "ne"):
        mf = filled(BASE, rental=rates, job_rate=rates, includes=includes)
        for r in _demo_requests(mf):
            c = ef.request_formula(mf, r)
            v = pl.contribution(mf, r)
            assert v.exact, r.ref
            assert c.scenarios == {ef.ONLY: ef.Formula(v.lo)}, r.ref
    # skupina z plánovače: vzorce variant = dosavadní hodnoty planner.plan
    m = model(var_rental=1200, var_job=900, includes="ne")
    p = pl.plan(m, new_requests=[rent("A", 1, 3, price=6000), rent("B", 2, 4, price=7000),
                                 job("J", 3, 6, revenue=40000, direct=10000)])
    g = p.groups[0]
    forms = ef.group_formulas(m, p, g)
    assert set(forms) == set(g["candidates"])
    for v, val in g["values"].items():
        assert forms[v].scenarios == {ef.ONLY: ef.Formula(val.lo)}
    assert files_hash() == before


def test_bounds_agree_with_current_value_intervals():
    for m, r in ((model(), rent("A", 1, 5, price=6500)),
                 (model(), job("J", 1, 5, revenue=50000, direct=20000)),
                 (model(var_rental=0), rent("A", 1, 5, price=6500)),
                 (model(includes="ne"), job("J", 1, 5, revenue=50000, direct=20000))):
        v = pl.contribution(m, r)
        lo, hi = ef.bounds(ef.request_formula(m, r))
        assert hi == v.hi
        assert v.lo is None or lo == v.lo


def test_formula_does_not_change_model_or_data():
    before = files_hash()
    m = model()
    snap = {n: getattr(m, n).copy() for n in dm.SCHEMA}
    ef.variant_formula(m, [rent("A", 1, 5), job("J", 2, 6), rent("B", 3, 4, price=0)])
    assert all(snap[n].equals(getattr(m, n)) for n in dm.SCHEMA)
    assert files_hash() == before


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK", name)
