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


# ---------------------------------------------------------------------------
# Regresní testy z auditu PR #25
# ---------------------------------------------------------------------------

def test_audit_loss_making_request_is_not_recommended_inside_a_combination():
    # B je ztrátový (cena 1 000 < variabilní náklad 5 000); {A, B} je největší kombinace,
    # ale ekonomicky nejlepší je samotné A (20 000 > 12 000 > 8 000)
    m = model(var_rental=5000)
    p = pl.plan(m, new_requests=[rent("A", 1, 2, price=15000), rent("B", 3, 4, price=1000),
                                 rent("C", 1, 4, price=7000)])
    g = p.groups[0]
    assert sorted(g["options"]) == [("A", "B"), ("C",)]                 # provozní varianty beze změny
    assert g["verdict"] == pl.RECOMMEND and g["choice"] == ("A",)
    assert "ztrátový" in p.requests["B"]["economic_note"]
    # samostatný ztrátový požadavek je provozně proveditelný, ale označený
    p = pl.plan(m, new_requests=[rent("L", 10, 11, price=1000)])
    assert p.requests["L"]["operational"] == pl.OK_ALONE and "ztrátový" in p.requests["L"]["economic_note"]


def test_audit_late_job_includes_penalty():
    m = model(var_job=0, includes="ano")
    late = Request("J", "job", pl.NEW, ((BAGR, 1),), 1, 5, deadline=3, revenue=50000, direct_cost=20000,
                   penalty_per_day=4000)
    assert pl.contribution(m, late).lo == 50000 - 20000 - 2 * 4000
    on_time = Request("K", "job", pl.NEW, ((BAGR, 1),), 1, 3, deadline=3, revenue=50000, direct_cost=20000,
                      penalty_per_day=4000)
    assert pl.contribution(m, on_time).lo == 30000
    # zakázka z dat: penále podle termínu v datech
    t = {n: getattr(BASE, n).copy() for n in dm.SCHEMA}
    t["jobs"].loc[t["jobs"]["job_id"] == "J001", "status"] = "tentative"
    md = dm.DataModel(**t)
    req = [r for r in pl.tentative_requests(md, {"J001": 9}) if r.ref == "J001"][0]   # 9–13, termín 10
    assert (req.start, req.end, req.deadline) == (9, 13, 10)
    assert pl.contribution(md, req).result.parts.get("penále") == -3 * 8000            # 3 dny po termínu


def test_audit_duplicate_request_refs_are_rejected():
    import pytest
    with pytest.raises(ValueError, match="jedinečn"):
        pl.plan(model(), new_requests=[rent("X", 1, 2), rent("X", 3, 4)])
    with pytest.raises(ValueError, match="FIKT-P06"):
        pl.plan(BASE, {}, [Request("FIKT-P06", "rental", pl.NEW, (("PASOVY_DUMPER", 1),), 21, 22)])


def test_overlapping_service_records_count_once_in_planner():
    # dva servisní záznamy téhož kusu se překrývají: kus je v servisu jednou, druhý kus zůstává volný
    m = model(units=2, maintenance=[(f"{BAGR}-01", 1, 3), (f"{BAGR}-01", 2, 3)])
    p = pl.plan(m, new_requests=[rent("N", 2, 3)])
    assert p.requests["N"]["operational"] == pl.OK_ALONE
    assert p.manual == []                                             # není to konflikt využití
    assert any("jen jednou" in n and "duplicitní" in n for n in p.notes)
    # jediný kus se dvěma servisy: požadavek dál neproveditelný, příčinou je servis
    m = model(units=1, maintenance=[(f"{BAGR}-01", 1, 3), (f"{BAGR}-01", 2, 3)])
    p = pl.plan(m, new_requests=[rent("N", 2, 3)])
    assert p.requests["N"]["operational"] == pl.NOT_FEASIBLE
    assert all(x[2] == 0 for x in p.requests["N"]["problems"])          # volno 0, ne −1


def test_audit_shift_alternatives_stay_inside_window_without_penalty():
    rnd = random.Random(3)
    for _ in range(40):
        s = rnd.randint(1, 8)
        m = model(bookings=[("R1", s, s + rnd.randint(0, 4), 1, "confirmed")])
        st = rnd.randint(1, 10); dur = rnd.randint(1, 4); earliest = rnd.randint(1, st); deadline = st + dur - 1 + rnd.randint(0, 5)
        p = pl.plan(m, new_requests=[job("J", st, st + dur - 1, earliest=earliest, deadline=deadline)])
        for a in p.requests["J"]["alternatives"]:
            s0, e0 = a["shift"]
            assert earliest <= s0 and e0 <= deadline and e0 - s0 + 1 == dur


def test_audit_independent_capacity_check_both_directions():
    """Nezávisle (prostá aritmetika z dat, bez capacity.py): každá varianta i doporučení se vejdou
    a každý požadavek, který se sám vejde, je označený jako proveditelný – a naopak."""
    rnd = random.Random(21)
    for _ in range(80):
        n = rnd.randint(1, 3)
        maint = [(f"{BAGR}-{rnd.randint(1, n):02d}", rnd.randint(1, 10), rnd.randint(1, 3))] if rnd.random() < .6 else []
        book = []
        for k in range(rnd.randint(0, 3)):
            s = rnd.randint(1, 10)
            book.append((f"R{k}", s, s + rnd.randint(0, 3), 1, rnd.choice(["confirmed", "confirmed", "tentative"])))
        m = model(units=n, maintenance=maint, bookings=book, var_rental=rnd.choice([0, 300, None]))
        reqs = [rent(f"N{k}", s, s + rnd.randint(0, 3), q=rnd.randint(1, 2), price=rnd.choice([4000, 6500, 9000]))
                for k, s in enumerate(rnd.randint(1, 10) for _ in range(rnd.randint(1, 5)))]
        try:
            p = pl.plan(m, new_requests=reqs)
        except ValueError:
            continue
        free = {d: n for d in range(1, 20)}
        for u, s, days in maint:
            for d in range(s, s + days):
                free[d] -= 1
        for _, s, e, q, st in book:
            if st == "confirmed":
                for d in range(s, e + 1):
                    free[d] -= q
        allreq = {r.ref: r for r in reqs}
        allreq.update({b[0]: rent(b[0], b[1], b[2], q=b[3], status=pl.TENTATIVE) for b in book if b[4] == "tentative"})

        def fits(refs):
            use = {}
            for ref in refs:
                r = allreq[ref]
                for d in range(r.start, r.end + 1):
                    use[d] = use.get(d, 0) + r.needs[0][1]
            return all(free[d] - u >= 0 for d, u in use.items())
        for ref, e in p.requests.items():
            assert (e["operational"] == pl.OK_ALONE) == fits([ref]), (ref, e["operational"])
        for g in p.groups:
            for option in g["options"]:
                assert fits(option), option
            if g["choice"]:
                assert fits(g["choice"])


def test_audit_twelve_concurrent_requests_stay_fast():
    import time
    t = time.time()
    p = pl.plan(model(units=3, var_rental=0), new_requests=[rent(f"R{i}", 1, 5, price=5000 + i) for i in range(12)])
    assert time.time() - t < 10
    g = p.groups[0]
    assert g["verdict"] == pl.RECOMMEND and sorted(g["choice"]) == ["R10", "R11", "R9"]   # 3 nejdražší
    p = pl.plan(model(units=3), new_requests=[rent(f"R{i}", 1, 5) for i in range(13)])
    assert p.groups[0]["verdict"] == pl.NO_DECISION


def test_new_request_ref_must_not_collide_with_data_ids():
    # E2.5b: nová zakázka „J006“ by v ekonomice převzala údaje původní zakázky J006 (tržba 320 000)
    import pytest
    with pytest.raises(ValueError, match="J006"):
        pl.plan(BASE, {}, [Request("J006", "job", pl.NEW, (("SILNICNI_VALEC", 1),), 21, 22,
                                   revenue=10000, direct_cost=5000)])
    with pytest.raises(ValueError, match="FIKT-P01"):
        pl.plan(BASE, {}, [Request("FIKT-P01", "rental", pl.NEW, ((BAGR, 1),), 30, 31)])
    # jedinečné označení: ekonomika podle zadaných údajů, ne podle původní zakázky
    req = Request("NP-1", "job", pl.NEW, (("SILNICNI_VALEC", 1),), 21, 22, revenue=10000, direct_cost=5000)
    assert pl.contribution(BASE, req).hi == 5000
    pl.plan(BASE, {}, [req])                                             # jedinečné označení projde


# ---------------------------------------------------------------------------
# E2.5c krok 1: varianta „nepřijmout nic“ (přírůstkový příspěvek 0 Kč)
# ---------------------------------------------------------------------------

def _pair(m, a, b):
    """Dva soupeřící požadavky na jediný bagr (dny 1–3 a 2–4 se překrývají)."""
    p = pl.plan(m, new_requests=[a, b])
    assert len(p.groups) == 1 and sorted(p.groups[0]["requests"]) == sorted([a.ref, b.ref])
    return p, p.groups[0]


def test_all_loss_making_variants_recommend_nothing_not_smallest_loss():
    # variabilní náklad 7 000 Kč/den > cena → oba pronájmy přesně ztrátové (−3 000 a −1 500 Kč)
    p, g = _pair(model(var_rental=7000), rent("A", 1, 3, price=6000), rent("B", 2, 4, price=6500))
    assert g["values"][("A",)].lo == -3000 and g["values"][("B",)].lo == -1500
    assert g["verdict"] == pl.NONE_PAYS and g["choice"] == ()
    assert g["values"][pl.NOTHING].exact and g["values"][pl.NOTHING].lo == 0
    assert "nevyplatí" in g["text"] and "Nic se neruší" in g["text"]
    # provozní stav se nemění – oba se samy vejdou
    assert all(p.requests[r]["operational"] == pl.OK_ALONE for r in ("A", "B"))
    assert any("nepřijmout žádný" in line for line in pl.explain(p))


def test_profitable_variant_is_still_recommended():
    p, g = _pair(model(var_rental=0), rent("A", 1, 3, price=6000), rent("B", 2, 4, price=7000))
    assert g["verdict"] == pl.RECOMMEND and g["choice"] == ("B",)
    assert g["values"][pl.NOTHING].lo == 0                    # „nic“ je mezi porovnávanými variantami
    # ziskový i ztrátový požadavek: doporučí se ziskový, ne „nic“
    p, g = _pair(model(var_rental=6200), rent("A", 1, 3, price=6000), rent("B", 2, 4, price=7000))
    assert g["values"][("A",)].lo < 0 < g["values"][("B",)].lo
    assert g["verdict"] == pl.RECOMMEND and g["choice"] == ("B",)


def test_zero_contribution_ties_with_nothing_and_is_not_recommended():
    # skutečná nula (cena 0, náklad 0) – rovnocenné s „nepřijmout nic“, rozhodne uživatel
    p, g = _pair(model(var_rental=0), rent("A", 1, 3, price=0), rent("B", 2, 4, price=0))
    assert g["verdict"] == pl.EQUAL and g["choice"] is None
    # nula proti ztrátě: nula je shodná s „nic“ → ne doporučení A, ne „žádný se nevyplatí“
    p, g = _pair(model(var_rental=500), rent("A", 1, 3, price=500), rent("B", 2, 4, price=0))
    assert g["values"][("A",)].exact and g["values"][("A",)].lo == 0
    assert g["verdict"] == pl.EQUAL and "nepřijmout nic" in g["text"]


def test_incomplete_data_never_recommend_nothing_automatically():
    # chybí variabilní náklad: známá je jen horní mez (kladná) → nelze rozhodnout, ne „nic“
    p, g = _pair(model(), rent("A", 1, 3, price=6000), rent("B", 2, 4, price=7000))
    assert g["values"][("A",)].lo is None and g["values"][("A",)].hi > 0
    assert g["verdict"] == pl.NO_DECISION and g["choice"] is None
    assert g["verdict"] != pl.NONE_PAYS
    # chybí tržba zakázky → horní mez neznámá → nelze rozhodnout
    j = Request("J", "job", pl.NEW, ((BAGR, 1),), 1, 3, name="J", revenue=None, direct_cost=1000)
    p, g = _pair(model(var_rental=7000), j, rent("B", 2, 4, price=6500))
    assert g["verdict"] == pl.NO_DECISION and p.missing_data
    # jedna varianta jistě ztrátová, druhá neúplná s kladnou horní mezí → nelze rozhodnout
    loss = job("J2", 1, 3, revenue=10000, direct=20000)
    p, g = _pair(model(), loss, rent("B", 2, 4, price=7000))
    assert g["verdict"] == pl.NO_DECISION


def test_nothing_only_when_loss_is_certain_even_with_missing_costs():
    # chybí jen náklady (snižují příspěvek), ale už tržba − známé náklady < 0 → ztráta je jistá
    a, b = job("J1", 1, 3, revenue=10000, direct=20000), job("J2", 2, 4, revenue=5000, direct=9000)
    p, g = _pair(model(), a, b)
    assert all(v.hi is not None and v.hi < 0 for k, v in g["values"].items() if k)
    assert all(v.lo is None for k, v in g["values"].items() if k)          # údaje jsou neúplné
    assert g["verdict"] == pl.NONE_PAYS


def test_nothing_variant_keeps_tentative_bookings_and_data():
    before = {n: getattr(BASE, n).copy() for n in dm.SCHEMA}
    m = model(var_rental=7000, bookings=[("FIKT-T", 1, 3, 1, pl.TENTATIVE)])
    p = pl.plan(m, new_requests=[rent("NP-1", 2, 4, price=6000)])
    g = p.groups[0]
    assert sorted(g["requests"]) == ["FIKT-T", "NP-1"] and g["verdict"] == pl.NONE_PAYS
    assert "nic se automaticky neruší" in p.requests["FIKT-T"]["note"]
    assert (m.rental_bookings["status"] == pl.TENTATIVE).all()             # model beze změny
    assert all(before[n].equals(getattr(BASE, n)) for n in dm.SCHEMA)


def test_nothing_variant_property_random():
    rnd = random.Random(25)
    seen = set()
    for _ in range(40):
        units = rnd.randint(1, 2)
        reqs = []
        for i in range(rnd.randint(2, 5)):
            s = rnd.randint(1, 6)
            reqs.append(rent(f"R{i}", s, s + rnd.randint(0, 3), q=1, price=rnd.choice([0, 1000, 2000, 4000])))
        p = pl.plan(model(units=units, var_rental=2000), new_requests=reqs)
        for g in p.groups:
            if g["verdict"] == pl.ALL_FIT:
                continue
            seen.add(g["verdict"])
            vals = {k: v for k, v in g["values"].items() if k}
            if g["verdict"] == pl.NONE_PAYS:
                assert all(v.lo < 0 for v in vals.values()) and g["choice"] == ()
            elif g["verdict"] == pl.RECOMMEND:
                assert g["values"][g["choice"]].lo > 0                      # nikdy nejmenší ztráta
                assert all(v.lo < g["values"][g["choice"]].lo for k, v in vals.items() if k != g["choice"])
            else:
                assert g["verdict"] == pl.EQUAL                             # úplná data → vždy rozhodnutelné
    assert {pl.NONE_PAYS, pl.RECOMMEND} <= seen


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK", name)
