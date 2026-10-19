"""Testy plánovacího a rozhodovacího modelu (planner.py, E2.5a). Jen fiktivní testovací data.

Modely se skládají v paměti z kopie ukázkových dat – datové soubory se nemění.
Spuštění: python -m pytest tests   nebo   python tests/test_planner.py
"""

import hashlib
import random
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import capacity as cp  # noqa: E402
import datamodel as dm  # noqa: E402
import economics as ec  # noqa: E402
import planner as pl  # noqa: E402
from planner import Request  # noqa: E402

BASE = dm.load_model()
BAGR = "MINI_BAGR"           # ceník 6 500 Kč/den (fiktivní)


def model(units=1, extra_units=(), bookings=(), maintenance=(), var_rental=None, var_job=None, includes=None):
    """Malý model: jen mini bagry (units kusů), bez zakázek; volitelně úplná ekonomická data."""
    t = {n: getattr(BASE, n).copy() for n in dm.SCHEMA}
    fleet = [(f"{BAGR}-{i:02d}", BAGR, "owned", "", "") for i in range(1, units + 1)] + list(extra_units)
    t["fleet_units"] = pd.DataFrame(fleet, columns=list(BASE.fleet_units.columns))
    t["rental_bookings"] = pd.DataFrame(
        [(b, BAGR, str(s), str(e), str(q), "6500", "", st) for b, s, e, q, st in bookings],
        columns=list(BASE.rental_bookings.columns))
    t["maintenance"] = pd.DataFrame([(u, str(s), str(d), "fiktivní") for u, s, d in maintenance],
                                    columns=list(BASE.maintenance.columns))
    t["jobs"] = BASE.jobs.iloc[0:0].copy()
    t["job_requirements"] = BASE.job_requirements.iloc[0:0].copy()
    mt = t["machine_types"]
    if var_rental is not None:
        mt["variable_cost_rental_day"] = str(var_rental)
    if var_job is not None:
        mt["variable_cost_job_day"] = str(var_job)
    if includes is not None:
        t["settings"] = pd.concat([t["settings"], pd.DataFrame(
            [dict(key="direct_cost_includes_machines", value=includes, unit="", note="")])], ignore_index=True)
    return dm.DataModel(**t)


def rent(ref, s, e, q=1, price=None, status=pl.NEW):
    return Request(ref, "rental", status, ((BAGR, q),), s, e, price_day=price)


def job(ref, s, e, q=1, revenue=50000, direct=20000, earliest=None, deadline=None):
    return Request(ref, "job", pl.NEW, ((BAGR, q),), s, e, name=ref, earliest=earliest, deadline=deadline,
                   revenue=revenue, direct_cost=direct, penalty_per_day=1000)


def decision(p, ref):
    return p.requests[ref]


# ---------------------------------------------------------------------------
# Provozní proveditelnost
# ---------------------------------------------------------------------------

def test_free_capacity():
    p = pl.plan(model(), new_requests=[rent("N1", 1, 3)], period=(1, 5))
    d = decision(p, "N1")
    assert d["operational"] == pl.OK_ALONE and d["decision"].startswith("lze přijmout")
    assert p.manual == [] and p.availability["free"].tolist() == [1] * 5


def test_one_machine_two_concurrent_requests():
    m = model(var_rental=0)
    p = pl.plan(m, new_requests=[rent("A", 1, 3, price=6000), rent("B", 2, 4, price=7000)])
    g = p.groups[0]
    assert g["options"] == [("A",), ("B",)] or sorted(g["options"]) == [("A",), ("B",)]
    assert ("A", "B") not in g["options"]                               # obě současně se nevejdou
    assert g["verdict"] == pl.RECOMMEND and g["choice"] == ("B",)       # 21 000 > 18 000, data úplná
    assert all(decision(p, r)["operational"] == pl.OK_ALONE for r in ("A", "B"))


def test_multiple_units_of_same_machine():
    p = pl.plan(model(units=2), new_requests=[rent("A", 1, 3), rent("B", 2, 4)])
    assert p.groups[0]["verdict"] == pl.ALL_FIT
    p = pl.plan(model(units=2, var_rental=0),
                new_requests=[rent("A", 1, 3), rent("B", 2, 4), rent("C", 3, 5, q=1)])
    options = {tuple(sorted(o)) for o in p.groups[0]["options"]}
    assert options == {("A", "B"), ("A", "C"), ("B", "C")}               # den 3: jen 2 kusy ze 3 nároků
    # stejné ceny a délky → shodný přínos, rozhodne uživatel
    assert p.groups[0]["verdict"] == pl.EQUAL


def test_service_blocks_machine():
    m = model(units=1, maintenance=[(f"{BAGR}-01", 2, 2)])
    p = pl.plan(m, new_requests=[rent("N1", 1, 4)])
    d = decision(p, "N1")
    assert d["operational"] == pl.NOT_FEASIBLE
    assert [x[1] for x in d["problems"]] == [2, 3] and "servis" in d["problems"][0][4][0]


def test_confirmed_rental_beats_new_job_and_shift_is_offered():
    m = model(bookings=[("R1", 5, 8, 1, "confirmed")])
    p = pl.plan(m, new_requests=[job("J-NEW", 6, 7, earliest=3, deadline=12)])
    d = decision(p, "J-NEW")
    assert d["operational"] == pl.NOT_FEASIBLE and "pronájem R1" in d["problems"][0][4]
    assert [a["shift"] for a in d["alternatives"]] == [(3, 4), (9, 10), (10, 11)]     # jen v okně bez penále
    # potvrzená rezervace zůstává beze změny – žádná ruční ani automatická změna
    assert p.manual == []
    # bez okna se posun nenabízí
    p = pl.plan(m, new_requests=[job("J2", 6, 7)])
    assert decision(p, "J2")["alternatives"] == []


def test_conflict_of_two_confirmed_commitments_goes_to_manual_resolution():
    m = model(bookings=[("R1", 1, 5, 1, "confirmed"), ("R2", 4, 6, 1, "confirmed")])
    p = pl.plan(m, new_requests=[rent("N1", 10, 11)])
    assert len(p.manual) == 1 and p.manual[0]["decision"] == "k ručnímu řešení"
    assert p.manual[0]["conflict"]["days"] == [4, 5]
    assert decision(p, "N1")["operational"] == pl.OK_ALONE                # mimo konflikt dál proveditelné
    # ukázková data: oba záměrné konflikty, nabídka posunu zakázky se jen ověří, nic se nepřesouvá
    jobs = dm.legacy_jobs(BASE)
    import scheduling as sch
    plan6 = sch.evaluate(jobs, sch.optimize_schedule(jobs, dm.fleet_capacity(BASE))["starts"])
    starts = {jobs.loc[i, "job_id"]: int(r.start_day) for i, r in plan6.iterrows() if r.scheduled}
    p = pl.plan(BASE, starts)
    assert sorted(m["conflict"]["id"] for m in p.manual) == ["KOLOVY_BAGR", "SILNICNI_VALEC"]
    assert all(m["decision"] == "k ručnímu řešení" and m["alternatives"] for m in p.manual)


def test_missing_machine_count_or_quantity_cannot_be_decided():
    bad_unit = [(f"{BAGR}-09", BAGR, "owned", "x", "")]                  # neplatná dostupnost kusu
    p = pl.plan(model(extra_units=bad_unit), new_requests=[rent("N1", 1, 2)])
    d = decision(p, "N1")
    assert d["operational"] == pl.UNDETERMINED and d["decision"].startswith("nelze rozhodnout")
    assert any("počet kusů" in x for x in d["reasons"])
    p = pl.plan(model(), new_requests=[Request("N2", "rental", pl.NEW, ((BAGR, None),), 1, 2),
                                       Request("N3", "rental", pl.NEW, ((BAGR, 1),), 5, None)])
    assert decision(p, "N2")["operational"] == pl.UNDETERMINED
    assert decision(p, "N3")["operational"] == pl.UNDETERMINED
    assert any("N2" in x for x in p.missing_data) and any("N3" in x for x in p.missing_data)


# ---------------------------------------------------------------------------
# Ekonomika
# ---------------------------------------------------------------------------

def test_missing_economic_data_gives_no_false_recommendation():
    p = pl.plan(model(), new_requests=[rent("A", 1, 3, price=6000), rent("B", 2, 4, price=9000)])
    g = p.groups[0]
    assert g["verdict"] == pl.NO_DECISION and g["choice"] is None
    assert sorted(g["options"]) == [("A",), ("B",)]                     # provozní varianty zůstávají
    assert "variabilní náklad pronájmu" in g["text"]
    assert any("variabilní náklad" in x for x in p.missing_data)
    v = decision(p, "A")["economics"]
    assert v.lo is None and v.hi == 18000                                # známá částka = jen horní mez


def test_dominance_decides_even_with_incomplete_cost():
    # A: úplné (nová zakázka, náklady stroje v přímých nákladech), B: pronájem s chybějícím
    # variabilním nákladem – horní mez B (6 000) < jistá hodnota A (30 000) → A jednoznačně
    m = model(includes="ano")
    p = pl.plan(m, new_requests=[job("A", 1, 1), rent("B", 1, 1)])
    g = p.groups[0]
    assert g["verdict"] == pl.RECOMMEND and g["choice"] == ("A",)


def test_equal_economic_benefit():
    p = pl.plan(model(var_rental=500), new_requests=[rent("A", 1, 2, price=6000), rent("B", 2, 3, price=6000)])
    g = p.groups[0]
    assert g["verdict"] == pl.EQUAL and g["choice"] is None and "rozhodne uživatel" in g["text"]


def test_no_double_counting_of_costs():
    req = job("J", 1, 4, q=2, revenue=100000, direct=40000)
    # náklady strojů nejsou v přímých nákladech → odečtou se jednou (4 dny × 2 kusy × 500)
    m = model(units=2, var_job=500, includes="ne")
    v = pl.contribution(m, req)
    assert v.exact and v.lo == 100000 - 40000 - 4 * 2 * 500
    # jsou v přímých nákladech → neodečtou se znovu
    v = pl.contribution(model(units=2, var_job=500, includes="ano"), req)
    assert v.exact and v.lo == 60000
    # fixní náklady flotily se do přínosu nepřičítají (jsou jen jednou za flotilu)
    fixed = ec.fleet_fixed_costs(m)
    assert fixed and pl.contribution(m, req).lo == 56000
    # pronájem: přesně economics.rental_contribution, bez ušlého zisku navíc
    m = model(var_rental=300)
    assert pl.contribution(m, rent("R", 1, 3, q=1)).lo == ec.rental_contribution(m, BAGR, 3).value == 3 * 6200
    # porovnání kombinací sčítá jen příspěvky požadavků – žádný požadavek dvakrát
    p = pl.plan(model(var_rental=0), new_requests=[rent("A", 1, 3, price=6000), rent("B", 2, 4, price=7000)])
    vals = p.groups[0]["values"]
    assert vals[("A",)].lo == 18000 and vals[("B",)].lo == 21000


# ---------------------------------------------------------------------------
# Bezpečnost výsledku
# ---------------------------------------------------------------------------

def test_recommendations_never_exceed_capacity():
    rnd = random.Random(5)
    for _ in range(60):
        units = rnd.randint(1, 3)
        maint = [(f"{BAGR}-01", rnd.randint(1, 8), rnd.randint(1, 2))] if rnd.random() < .5 else []
        book = [("R1", s, s + rnd.randint(0, 3), 1, "confirmed") for s in [rnd.randint(1, 8)]]
        m = model(units=units, maintenance=maint, bookings=book, var_rental=rnd.choice([None, 0, 200]))
        reqs = []
        for k in range(rnd.randint(1, 6)):
            s = rnd.randint(1, 10)
            reqs.append(rent(f"N{k}", s, s + rnd.randint(0, 4), q=rnd.randint(1, 2), price=rnd.choice([5000, 6000])))
        p = pl.plan(m, new_requests=reqs)
        cal = cp.from_model(m)
        by_ref = {r.ref: r for r in reqs}
        for g in p.groups:
            for option in g["options"]:
                added = [cp.Assignment("rental", ref, BAGR, by_ref[ref].start, by_ref[ref].end,
                                       by_ref[ref].needs[0][1]) for ref in option]
                test = cp.build_calendar(cal.units, cal.assignments + added)
                for ref in option:
                    df = cp.daily_capacity(test, by_ref[ref].start, by_ref[ref].end, [BAGR])
                    assert (df["free"] >= 0).all(), (option, g)
        # neproveditelný požadavek se nikdy neobjeví v doporučené kombinaci
        bad = {ref for ref, e in p.requests.items() if e["operational"] != pl.OK_ALONE}
        assert not any(bad & set(g["choice"] or ()) for g in p.groups)


def test_tentative_booking_stays_visible_and_is_never_cancelled():
    m = model(var_rental=0, bookings=[("T1", 1, 3, 1, "tentative")])
    p = pl.plan(m, new_requests=[rent("N1", 2, 4, price=9000)])
    assert "T1" in p.requests and "zůstává" in p.requests["T1"]["note"]
    assert p.groups[0]["verdict"] == pl.RECOMMEND and p.groups[0]["choice"] == ("N1",)
    text = " ".join(pl.explain(p))
    assert "T1" in text and "zrušit" not in text and "Doporučeno (jen doporučení, nic se nemění)" in text
    assert m.rental_bookings["status"].tolist() == ["tentative"]          # model beze změny


def test_inputs_and_data_files_are_not_modified():
    files = sorted((ROOT / "data" / "model").glob("*.csv"))
    before = {f.name: hashlib.sha256(f.read_bytes()).hexdigest() for f in files}
    m = dm.load_model()
    snapshot = {n: getattr(m, n).copy() for n in dm.SCHEMA}
    reqs = [rent("N1", 13, 16), job("N2", 13, 15, earliest=13, deadline=30)]
    reqs_copy = list(reqs)
    pl.plan(m, {"J001": 1}, reqs, period=(1, 25))
    for n in dm.SCHEMA:
        pd.testing.assert_frame_equal(getattr(m, n), snapshot[n])
    assert reqs == reqs_copy
    after = {f.name: hashlib.sha256(f.read_bytes()).hexdigest() for f in files}
    assert before == after


def test_app_does_not_use_planner_yet():
    src = (ROOT / "app.py").read_text(encoding="utf-8")
    assert "planner" not in src


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK", name)
