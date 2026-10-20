"""Testy robustního ekonomického verdiktu plánovače (E2.5c krok 3). Jen fiktivní testovací data.

Modely se skládají v paměti; datové soubory se nemění.
Spuštění: python -m pytest tests   nebo   python tests/test_planner_robust.py
"""

import random
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))
import capacity as cp  # noqa: E402
import datamodel as dm  # noqa: E402
import planner as pl  # noqa: E402
from planner import Request  # noqa: E402
from test_econ_formula import filled, files_hash  # noqa: E402
from test_planner import BAGR, job, model, rent  # noqa: E402

LONG = dict(revenue=150000, direct=60000)          # zakázka 10 dní na bagru


def group(p):
    gs = [g for g in p.groups if g["verdict"] != pl.ALL_FIT]
    assert len(gs) == 1
    return gs[0]


def long_vs_short(m, **job_kw):
    """Jedna dlouhá zakázka (dny 1–10) proti dvěma kratším pronájmům (1–5 á 6 500, 6–10 á 7 000)."""
    kw = dict(LONG, **job_kw)
    return pl.plan(m, new_requests=[job("J", 1, 10, revenue=kw["revenue"], direct=kw["direct"]),
                                    rent("R1", 1, 5, price=6500), rent("R2", 6, 10, price=7000)])


# ---------------------------------------------------------------------------
# Dva kratší pronájmy × jedna dlouhá zakázka, vyšší tržba ≠ vyšší příspěvek
# ---------------------------------------------------------------------------

def test_long_job_beats_two_short_rentals_with_complete_data():
    # J = 150 000 − 60 000 − 10 × 2 000 = 70 000; R1 + R2 = 67 500 − 10 × 1 500 = 52 500
    p = long_vs_short(model(var_rental=1500, var_job=2000, includes="ne"))
    g = group(p)
    assert g["verdict"] == pl.RECOMMEND and g["choice"] == ("J",)
    assert g["formulas"][("J",)] == ["70 000 Kč"]
    assert "prokazatelně" in g["text"]


def test_two_short_rentals_beat_long_job():
    # J = 120 000 − 60 000 − 20 000 = 40 000 < R1 + R2 = 52 500
    p = long_vs_short(model(var_rental=1500, var_job=2000, includes="ne"), revenue=120000)
    g = group(p)
    assert g["verdict"] == pl.RECOMMEND and g["choice"] == ("R1", "R2")
    assert ("R1", "R2") in g["candidates"] and pl.NOTHING in g["candidates"]


def test_higher_revenue_but_lower_contribution_loses():
    # tržba zakázky 150 000 > tržba pronájmů 67 500, ale příspěvek 150 000 − 110 000 − 20 000 = 20 000 < 52 500
    m = model(var_rental=1500, var_job=2000, includes="ne")
    p = long_vs_short(m, direct=110000)
    g = group(p)
    assert 150000 > 6500 * 5 + 7000 * 5
    assert g["verdict"] == pl.RECOMMEND and g["choice"] == ("R1", "R2")
    assert ("J",) not in g["contenders"]


# ---------------------------------------------------------------------------
# Ztrátové, shodné a nulové varianty
# ---------------------------------------------------------------------------

def test_all_loss_making_recommends_nothing():
    p = pl.plan(model(var_rental=7000), new_requests=[rent("A", 1, 3, price=6000), rent("B", 2, 4, price=6500)])
    g = group(p)
    assert g["verdict"] == pl.NONE_PAYS and g["choice"] == pl.NOTHING and g["contenders"] == [pl.NOTHING]
    # i s neznámými náklady, je-li ztráta jistá už ze známých částek
    p = pl.plan(model(), new_requests=[job("J1", 1, 3, revenue=10000, direct=20000),
                                       job("J2", 2, 4, revenue=5000, direct=9000)])
    assert group(p)["verdict"] == pl.NONE_PAYS


def test_equal_results_are_equal_not_recommended():
    p = pl.plan(model(var_rental=1000), new_requests=[rent("A", 1, 2, price=6000), rent("B", 2, 3, price=6000)])
    g = group(p)
    assert g["verdict"] == pl.EQUAL and g["choice"] is None
    assert set(g["best"]) == {("A",), ("B",)} and pl.NOTHING not in g["best"]


def test_zero_contribution_is_not_better_than_nothing():
    # A: příspěvek přesně 0 Kč (cena = variabilní náklad), B: ztráta → A je rovnocenné s „nic“
    p = pl.plan(model(var_rental=500), new_requests=[rent("A", 1, 3, price=500), rent("B", 2, 4, price=0)])
    g = group(p)
    assert g["verdict"] == pl.EQUAL and set(g["best"]) == {("A",), pl.NOTHING}
    assert "nepřijmout nic" in g["text"]


# ---------------------------------------------------------------------------
# Neznámé sazby: hranice zlomu, vyrušení, skutečná nula, oba scénáře
# ---------------------------------------------------------------------------

def test_decision_depends_on_unknown_rate_gives_break_even():
    # A = 32 500 − 5x, B = 27 000 − 3x, „nic“ = 0 → do 2 750 Kč A, 2 750–9 000 Kč B, nad 9 000 Kč nic
    p = pl.plan(model(), new_requests=[rent("A", 1, 5, price=6500), rent("B", 1, 3, price=9000)])
    g = group(p)
    assert g["verdict"] == pl.CONDITIONAL and g["choice"] is None
    segs = g["profiles"][0]["profile"]["segments"]
    assert [(lo, hi, w) for lo, hi, w in segs] == [(0.0, 2750.0, ("A",)), (2750.0, 9000.0, ("B",)),
                                                   (9000.0, None, ("nepřijmout nic",))]
    assert "do 2 750 Kč: A; 2 750 Kč–9 000 Kč: B; nad 9 000 Kč: nepřijmout nic" in g["text"]
    assert any("variabilní náklad pronájmu" in x for x in p.missing_data)


def test_cancelling_unknown_rate_allows_certain_recommendation():
    # náklady strojů jsou v přímých nákladech zakázky (ano) → J = 90 000 přesně;
    # R1 + R2 = 67 500 − 10x je horší pro každé x ≥ 0 → doporučení i bez sazby pronájmu
    p = long_vs_short(model(includes="ano"))
    g = group(p)
    assert g["verdict"] == pl.RECOMMEND and g["choice"] == ("J",) and g["formulas"][("J",)] == ["90 000 Kč"]
    # stejná neznámá sazba v obou variantách se vyruší: B + C jistě lepší než A
    p = pl.plan(model(), new_requests=[rent("A", 1, 5, price=6500), rent("B", 1, 3, price=7000),
                                       rent("C", 4, 5, price=6000)])
    assert ("A",) not in group(p)["contenders"]


def test_true_zero_versus_missing_rate():
    reqs = [rent("A", 1, 5, price=6500), rent("B", 1, 3, price=9000)]
    zero = group(pl.plan(model(var_rental=0), new_requests=reqs))
    assert zero["verdict"] == pl.RECOMMEND and zero["choice"] == ("A",)       # 32 500 > 27 000 > 0
    missing = group(pl.plan(model(), new_requests=reqs))
    assert missing["verdict"] == pl.CONDITIONAL and missing["choice"] is None  # chybí → nikdy jako nula


def test_both_scenarios_of_direct_cost_includes_machines():
    # sazba pronájmu známá (1 500), sazba stroje na zakázce chybí, nastavení chybí:
    # „ano“: J = 90 000 > 52 500 → J;  „ne“: J = 90 000 − 10y → J jen pro y < 3 750 Kč
    p = long_vs_short(model(var_rental=1500))
    g = group(p)
    assert g["verdict"] == pl.CONDITIONAL and g["choice"] is None
    by_sc = {x["scenario"]: x["profile"] for x in g["profiles"]}
    assert set(by_sc) == {"ano", "ne"}
    assert [w for _, _, w in by_sc["ano"]["segments"]] == [("J",)]
    # R1 a R2 samostatně i „nic“ jsou prokazatelně horší než R1 + R2 (52 500 Kč přesně) → vyřazené
    assert set(g["contenders"]) == {("J",), ("R1", "R2")}
    assert [(lo, hi, w) for lo, hi, w in by_sc["ne"]["segments"]] == [(0.0, 3750.0, ("J",)),
                                                                      (3750.0, None, ("R1 + R2",))]
    assert "do 3 750 Kč: J; nad 3 750 Kč: R1 + R2" in g["text"]
    assert "Pokud přímé náklady zakázek neobsahují stroje" in g["text"]
    # nastavení zadané: „ano“ + známá sazba stroje → náklad stroje se neodečte podruhé
    p = long_vs_short(model(var_rental=1500, var_job=2000, includes="ano"))
    g = group(p)
    assert g["verdict"] == pl.RECOMMEND and g["choice"] == ("J",) and g["formulas"][("J",)] == ["90 000 Kč"]


def test_weak_dominance_is_not_an_unconditional_recommendation():
    # A = 19 500 − 3x, B = 19 500 − 4x: při x = 0 shodné, při x > 0 je A lepší → žádné jisté doporučení
    p = pl.plan(model(), new_requests=[rent("A", 1, 3, price=6500), rent("B", 1, 4, price=4875)])
    g = group(p)
    assert g["verdict"] == pl.CONDITIONAL and g["choice"] is None
    prof = g["profiles"][0]["profile"]
    assert set(prof["at_zero"]) == {"A", "B"} and prof["segments"][0][2] == ("A",)
    assert "při nulové hodnotě shodně i B" in g["text"]
    # totéž s úplnými daty a kladnou sazbou → A jistě lepší
    g = group(pl.plan(model(var_rental=100), new_requests=[rent("A", 1, 3, price=6500),
                                                           rent("B", 1, 4, price=4875)]))
    assert g["verdict"] == pl.RECOMMEND and g["choice"] == ("A",)


# ---------------------------------------------------------------------------
# Chybějící základní údaje, neplatné údaje
# ---------------------------------------------------------------------------

def test_missing_basic_data_cannot_decide():
    j = Request("J", "job", pl.NEW, ((BAGR, 1),), 1, 3, name="J", revenue=None, direct_cost=1000)
    g = group(pl.plan(model(var_rental=0, includes="ano"), new_requests=[j, rent("B", 2, 4, price=6500)]))
    assert g["verdict"] == pl.NO_DECISION and g["choice"] is None and "tržba" in g["text"]


def test_invalid_value_cannot_decide():
    m = model()
    m.machine_types.loc[m.machine_types["type_id"] == BAGR, "variable_cost_rental_day"] = "abc"
    g = group(pl.plan(m, new_requests=[rent("A", 1, 3, price=6000), rent("B", 2, 4, price=7000)]))
    assert g["verdict"] == pl.NO_DECISION and g["choice"] is None


# ---------------------------------------------------------------------------
# Předběžné rezervace, kapacita, potvrzené závazky
# ---------------------------------------------------------------------------

def test_tentative_booking_competes_but_is_never_cancelled():
    before = files_hash()
    m = model(var_rental=1000, bookings=[("FIKT-T", 1, 5, 1, pl.TENTATIVE)])
    snap = m.rental_bookings.copy()
    p = pl.plan(m, new_requests=[rent("NP-1", 3, 4, price=9000)])
    g = group(p)
    # FIKT-T: 5 × (6 500 − 1 000) = 27 500 > NP-1: 2 × (9 000 − 1 000) = 16 000
    assert g["verdict"] == pl.RECOMMEND and g["choice"] == ("FIKT-T",)
    assert "nic se automaticky neruší" in p.requests["FIKT-T"]["note"]
    assert snap.equals(m.rental_bookings) and files_hash() == before


def test_recommended_variant_respects_capacity_and_confirmed_commitments():
    # 2 bagry; potvrzený pronájem drží 1 kus ve dnech 1–5 → nové požadavky soupeří o zbylý kus
    m = model(units=2, var_rental=1000, bookings=[("POTVRZ", 1, 5, 1, "confirmed")])
    reqs = [rent("A", 1, 3, price=7000), rent("B", 2, 4, price=8000), rent("C", 4, 5, price=6000)]
    p = pl.plan(m, new_requests=reqs)
    g = group(p)
    assert g["verdict"] == pl.RECOMMEND
    chosen = [r for r in reqs if r.ref in g["choice"]]
    cal = cp.from_model(m)
    added = [cp.Assignment("rental", r.ref, BAGR, r.start, r.end, 1) for r in chosen]
    test = cp.build_calendar(cal.units, cal.assignments + added, cal.open_jobs)
    assert (cp.daily_capacity(test, 1, 5, [BAGR])["free"] >= 0).all()
    # požadavek přes servis/potvrzený závazek bez volného kusu je neproveditelný a ekonomicky se neporovnává
    m = model(units=1, var_rental=1000, bookings=[("POTVRZ", 1, 5, 1, "confirmed")])
    p = pl.plan(m, new_requests=[rent("X", 2, 3, price=9000), rent("Y", 7, 8, price=5000)])
    assert p.requests["X"]["operational"] == pl.NOT_FEASIBLE and not p.groups


def test_service_blocks_capacity_before_economics():
    m = model(units=1, var_rental=0, maintenance=[(f"{BAGR}-01", 3, 2)])
    p = pl.plan(m, new_requests=[rent("A", 1, 5, price=9000), rent("B", 6, 7, price=5000)])
    assert p.requests["A"]["operational"] == pl.NOT_FEASIBLE
    assert p.requests["B"]["operational"] == pl.OK_ALONE


# ---------------------------------------------------------------------------
# Vlastnosti: shoda s hrubou silou (úplná data) a bezpečnost (neúplná data)
# ---------------------------------------------------------------------------

def _random_requests(rnd):
    out = []
    for i in range(rnd.randint(2, 5)):
        s = rnd.randint(1, 6)
        if rnd.random() < 0.3:
            out.append(job(f"J{i}", s, s + rnd.randint(1, 4), revenue=rnd.choice([10000, 30000, 60000]),
                           direct=rnd.choice([5000, 20000])))
        else:
            out.append(rent(f"R{i}", s, s + rnd.randint(0, 3), price=rnd.choice([0, 2000, 6500, 9000])))
    return out


def _true_values(m_filled, g, reqs):
    by_ref = {r.ref: r for r in reqs}
    return {v: sum(pl.contribution(m_filled, by_ref[x]).lo for x in v) for v in g["candidates"]}


def test_property_complete_data_matches_brute_force():
    rnd = random.Random(31)
    seen = set()
    for _ in range(40):
        reqs = _random_requests(rnd)
        m = model(units=rnd.randint(1, 2), var_rental=1500, var_job=1000, includes="ne")
        for g in pl.plan(m, new_requests=reqs).groups:
            if g["verdict"] == pl.ALL_FIT:
                continue
            vals = _true_values(m, g, reqs)
            top = max(vals.values())
            best = {v for v, x in vals.items() if abs(x - top) < 1e-6}
            seen.add(g["verdict"])
            if len(best) == 1:
                (w,) = best
                assert g["verdict"] == (pl.NONE_PAYS if w == pl.NOTHING else pl.RECOMMEND)
                assert g["choice"] == w
            else:
                assert g["verdict"] == pl.EQUAL and set(g["best"]) == best
    assert {pl.RECOMMEND, pl.NONE_PAYS} <= seen


def test_property_missing_rates_never_give_false_recommendation():
    rnd = random.Random(32)
    seen = set()
    for _ in range(30):
        reqs = _random_requests(rnd)
        m = model(units=rnd.randint(1, 2))                       # chybí sazby i nastavení
        p = pl.plan(m, new_requests=reqs)
        for g in p.groups:
            if g["verdict"] == pl.ALL_FIT:
                continue
            seen.add(g["verdict"])
            for xr in (0, 800, 3000, 12000):
                for xj in (0, 1500, 9000):
                    for inc in ("ano", "ne"):
                        vals = _true_values(filled(m, rental={BAGR: xr}, job_rate={BAGR: xj}, includes=inc), g, reqs)
                        top = max(vals.values())
                        if g["verdict"] in (pl.RECOMMEND, pl.NONE_PAYS):
                            # prokazatelně lepší: ostře lepší při jakýchkoli hodnotách neznámých
                            assert all(vals[g["choice"]] > x for v, x in vals.items() if v != g["choice"])
                        elif g["verdict"] == pl.EQUAL:
                            assert all(abs(vals[v] - top) < 1e-6 for v in g["best"])
                        elif g["verdict"] == pl.CONDITIONAL:
                            # skutečně nejlepší varianta je vždy mezi neodmítnutými
                            assert any(abs(vals[v] - top) < 1e-6 for v in g["contenders"])
    assert pl.CONDITIONAL in seen


def test_plan_does_not_change_data():
    before = files_hash()
    snap = {n: getattr(dm.load_model(), n).copy() for n in dm.SCHEMA}
    long_vs_short(model())
    assert files_hash() == before
    assert all(snap[n].equals(getattr(dm.load_model(), n)) for n in dm.SCHEMA)


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK", name)
