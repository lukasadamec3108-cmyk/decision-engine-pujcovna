"""Testy ekonomické evidence závazků za období (E2.2 – K3.1, period_ledger). Jen fiktivní data.

Spuštění: python -m pytest tests/test_fleet_economics.py
"""

import ast
import sys
from datetime import date
from pathlib import Path

import pandas as pd
import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import capacity as cp  # noqa: E402
import datamodel as dm  # noqa: E402
import economics as ec  # noqa: E402
import fleet_economics as fe  # noqa: E402
import fleet_usage as fu  # noqa: E402

BAGR, DUMPER = "MINI_BAGR", "PASOVY_DUMPER"        # ceník 6 500 a 8 000 Kč/den (fiktivní)
BASE = dm.load_model()


def frame(rows, table):
    cols = list(dm.SCHEMA[table])
    return pd.DataFrame([{c: str(r.get(c, "")) for c in cols} for r in rows], columns=cols, dtype=str)


def booking(ref, t, s, e, units=1, price=None, status="confirmed"):
    price = {BAGR: 6500, DUMPER: 8000}[t] if price is None else price
    return dict(booking_id=ref, type_id=t, start_day=s, end_day=e, units=units, price_day=price, status=status)


def job(ref, duration, deadline=30, revenue=100000, cost=40000, penalty=1000, status="confirmed", earliest=1):
    return dict(job_id=ref, name=f"Zakázka {ref}", job_type="test", status=status, p_win="" if status == "confirmed"
                else "0.5", priority="Medium", risk_level="Medium", revenue=revenue, direct_cost=cost,
                duration_days=duration, earliest_start_day=earliest, deadline_day=deadline, penalty_per_day=penalty)


def need(ref, t, q=1, pos=1):
    return dict(job_id=ref, resource_kind="machine", resource_id=t, quantity=q, position=pos)


def model(units=((f"{BAGR}-01", BAGR), (f"{DUMPER}-01", DUMPER)), bookings=(), jobs=(), needs=(), service=(),
          types=None, settings=None):
    mt = BASE.machine_types.copy()
    if types:
        mt = types(mt)
    st = BASE.settings.copy()
    if settings:
        st = pd.concat([st, frame([dict(key=k, value=v) for k, v in settings.items()], "settings")],
                       ignore_index=True)
    return dm.DataModel(
        machine_types=mt, settings=st, operator_roles=BASE.operator_roles.copy(),
        rental_demand=BASE.rental_demand.copy(),
        fleet_units=frame([dict(unit_id=u, type_id=t, status="owned") for u, t in units], "fleet_units"),
        maintenance=frame([dict(unit_id=u, start_day=s, days=d) for u, s, d in service], "maintenance"),
        rental_bookings=frame(list(bookings), "rental_bookings"),
        jobs=frame(list(jobs), "jobs"), job_requirements=frame(list(needs), "job_requirements"))


def row(ledger, ref):
    hit = ledger.entries[ledger.entries["ref"] == ref]
    assert len(hit) == 1, (ref, len(hit))
    return hit.iloc[0]


def certain_refs(ledger):
    return set(ledger.entries.loc[ledger.entries["counted_as_certain"], "ref"])


def total(ledger, key_start):
    [key] = [k for k in ledger.totals if k.startswith(key_start)]
    return ledger.totals[key]


# ---------------------------------------------------------------------------
# Konflikty
# ---------------------------------------------------------------------------

def test_rental_and_job_conflict_is_not_counted_as_certain():
    m = model(bookings=[booking("P1", BAGR, 1, 5)], jobs=[job("J1", 3)], needs=[need("J1", BAGR)])
    led = fe.period_ledger(m, 1, 10, job_starts={"J1": 4})
    p, j = row(led, "P1"), row(led, "J1")
    assert p.state == j.state == fe.CONFLICT                       # obě strany, žádná priorita
    assert p.conflict_days == j.conflict_days == [4, 5]
    assert (p.md_conflict_in_period, j.md_conflict_in_period) == (2, 2)
    assert certain_refs(led) == set()
    assert total(led, "jisté – pronájmy").value == total(led, "jisté – zakázky").value == 0
    assert total(led, "v konfliktu").value == p.contribution + j.contribution == 5 * 6500 + 60000
    assert "priorita se neurčuje" in p.reason and led.status == fu.INCOMPLETE


def test_conflict_states_do_not_depend_on_order_of_records():
    rows = [booking("P1", BAGR, 1, 5), booking("P2", BAGR, 3, 4)]
    a = fe.period_ledger(model(bookings=rows), 1, 5)
    b = fe.period_ledger(model(bookings=rows[::-1]), 1, 5)
    for led in (a, b):
        assert row(led, "P1").state == row(led, "P2").state == fe.CONFLICT


def test_service_and_job_conflict():
    m = model(jobs=[job("J1", 4)], needs=[need("J1", BAGR)], service=[(f"{BAGR}-01", 3, 1)])
    led = fe.period_ledger(m, 1, 10, job_starts={"J1": 1})
    j = row(led, "J1")
    assert j.state == fe.CONFLICT and j.conflict_days == [3] and j.md_conflict_in_period == 1
    assert "J1" not in certain_refs(led)
    # servis v jiný den: zakázka je jistá
    led2 = fe.period_ledger(model(jobs=[job("J1", 4)], needs=[need("J1", BAGR)],
                                  service=[(f"{BAGR}-01", 6, 2)]), 1, 10, job_starts={"J1": 1})
    assert row(led2, "J1").state == fe.CERTAIN and certain_refs(led2) == {"J1"}


def test_partial_overlap_marks_whole_commitment_without_proration():
    m = model(bookings=[booking("P1", BAGR, 1, 5), booking("P2", BAGR, 5, 8), booking("P3", DUMPER, 1, 8)])
    led = fe.period_ledger(m, 1, 10)
    p1, p2, p3 = row(led, "P1"), row(led, "P2"), row(led, "P3")
    assert p1.state == p2.state == fe.CONFLICT and p1.conflict_days == p2.conflict_days == [5]
    assert (p1.machine_days, p1.md_conflict_in_period, p2.machine_days) == (5, 1, 4)
    # příspěvek je za celý závazek, nic se poměrně nedělí ani nepřesouvá mezi skupinami
    assert p1.contribution == 5 * 6500 and p2.contribution == 4 * 6500
    assert p3.state == fe.CERTAIN and total(led, "jisté – pronájmy").value == 8 * 8000


def test_job_with_several_machines():
    m = model(bookings=[booking("P1", DUMPER, 3, 3)], jobs=[job("J1", 4), job("J2", 2)],
              needs=[need("J1", BAGR), need("J1", DUMPER, pos=2), need("J2", BAGR)])
    led = fe.period_ledger(m, 1, 10, job_starts={"J1": 1, "J2": 5})
    j1, j2 = row(led, "J1"), row(led, "J2")
    assert j1.needs == f"{BAGR}×1; {DUMPER}×1" and j1.machine_days == 8      # 4 dny × 2 stroje
    assert j1.state == fe.CONFLICT and j1.conflict_days == [3] and j1.md_conflict_in_period == 1
    assert j2.state == fe.CERTAIN and row(led, "P1").state == fe.CONFLICT
    assert j1.machine_days == sum(ec.job_machine_days(m, "J1").values())    # stejné strojodny jako economics


# ---------------------------------------------------------------------------
# Předběžné, chybějící náklady, přesah, neplatná data
# ---------------------------------------------------------------------------

def test_tentative_bookings_and_jobs_are_only_potential():
    m = model(bookings=[booking("P1", BAGR, 1, 5), booking("P2", BAGR, 3, 7, status="tentative")],
              jobs=[job("J1", 2, status="tentative"), job("J2", 2, status="opportunity")],
              needs=[need("J1", DUMPER), need("J2", DUMPER)])
    led = fe.period_ledger(m, 1, 10, job_starts={"J1": 1})
    assert row(led, "P1").state == fe.CERTAIN                      # předběžná rezervace neblokuje
    assert row(led, "P2").state == row(led, "J1").state == row(led, "J2").state == fe.TENTATIVE
    assert row(led, "J2").date_source == fe.NO_DATE and row(led, "J2").start is None
    assert certain_refs(led) == {"P1"}
    assert total(led, "předběžné").value == 5 * 6500 + 60000 + 60000
    assert not any(r.counted_as_certain for r in led.entries.itertuples() if r.state == fe.TENTATIVE)


def test_missing_costs_stay_incomplete_never_zero():
    m = model(bookings=[booking("P1", BAGR, 1, 2)], jobs=[job("J1", 2)], needs=[need("J1", DUMPER)])
    led = fe.period_ledger(m, 1, 5, job_starts={"J1": 1})
    p, j = row(led, "P1"), row(led, "J1")
    assert p.contribution_status == j.contribution_status == "neúplný"
    assert any("variabilní náklad pronájmu" in x for x in p.missing)
    assert any("direct_cost_includes_machines" in x for x in j.missing)
    assert p.contribution == 2 * 6500                               # horní mez, ne vymyšlená částka
    assert led.status == fu.INCOMPLETE and total(led, "jisté – pronájmy").status == "neúplný"
    # skutečně zadané hodnoty (i skutečná nula) → úplný výsledek
    def costs(mt):
        mt = mt.copy()
        mt.loc[mt["type_id"] == BAGR, "variable_cost_rental_day"] = "1000"
        mt.loc[mt["type_id"] == DUMPER, "variable_cost_job_day"] = "0"
        return mt
    m2 = model(bookings=[booking("P1", BAGR, 1, 2)], jobs=[job("J1", 2)], needs=[need("J1", DUMPER)],
               types=costs, settings={"direct_cost_includes_machines": "ne"})
    led2 = fe.period_ledger(m2, 1, 5, job_starts={"J1": 1})
    assert row(led2, "P1").contribution == 2 * (6500 - 1000) and row(led2, "P1").contribution_status == "úplný"
    assert row(led2, "J1").contribution == 60000 and row(led2, "J1").contribution_status == "úplný"
    assert led2.status == fu.COMPLETE


def test_commitments_crossing_the_period_need_a_rule():
    m = model(bookings=[booking("P1", BAGR, 4, 8), booking("P2", DUMPER, 9, 12)],
              jobs=[job("J1", 5, deadline=4, penalty=2000)], needs=[need("J1", DUMPER)])
    led = fe.period_ledger(m, 1, 5, job_starts={"J1": 3})
    p1, p2, j1 = row(led, "P1"), row(led, "P2"), row(led, "J1")
    assert p1.state == j1.state == fe.CROSSES and p1.crosses_period and j1.crosses_period
    assert (p1.md_in_period, p1.md_outside) == (2, 3) and (j1.md_in_period, j1.md_outside) == (3, 2)
    assert p1.contribution == 5 * 6500                              # celý závazek, žádné poměrné dělení
    assert j1.contribution == 100000 - 40000 - 3 * 2000             # penále podle návrhu plánu (konec 7, termín 4)
    assert p2.state == fe.OUTSIDE and certain_refs(led) == set()
    assert "pravidlo časového rozlišení" in p1.reason and any("časové rozlišení" in r for r in led.open_rules)
    # explicitní období: 45 dní = 45 dní (žádný 30denní měsíc)
    led45 = fe.period_ledger(m, 1, 45, job_starts={"J1": 3})
    assert led45.usage.fleet["capacity"] == 2 * 45 and row(led45, "P1").state == fe.CERTAIN


def test_period_across_month_with_calendar_dates():
    m = model(bookings=[booking("P1", BAGR, 2, 4)])
    led = fe.period_ledger(m, date(2026, 1, 31), date(2026, 2, 3), plan_start=date(2026, 1, 30))
    p = row(led, "P1")
    assert (p.start, p.end) == (date(2026, 1, 31), date(2026, 2, 2)) and p.state == fe.CERTAIN
    assert led.usage.fleet["capacity"] == 2 * 4


def test_invalid_data_is_never_certain():
    m = model(bookings=[booking("P1", BAGR, 5, 3), booking("P2", DUMPER, 1, 2, price="abc"),
                        booking("P3", DUMPER, 6, 6, status="potvrzeno"), booking("P4", DUMPER, 8, 8, status="cancelled")],
              jobs=[job("J1", 2), job("J2", 2)], needs=[need("J1", BAGR), need("J2", BAGR)])
    led = fe.period_ledger(m, 1, 10, job_starts={"J1": 1})
    p1, p2, p3, p4 = (row(led, r) for r in ("P1", "P2", "P3", "P4"))
    assert p1.state == fe.UNKNOWN and "termín" in p1.reason and p1.contribution_status == "nelze spočítat"
    assert p2.contribution is None and p2.contribution_status == "nelze spočítat"
    assert p2.state == fe.UNKNOWN                                   # typ s neplatným záznamem P3 – nelze ověřit
    assert p3.state == fe.UNKNOWN and "neznámý stav" in p3.reason
    assert p4.state == fe.CANCELLED and not p4.counted_as_certain
    # potvrzená zakázka bez termínu v plánu: nic se nevymýšlí, ostatní na stejném stroji jsou nejisté
    j1, j2 = row(led, "J1"), row(led, "J2")
    assert j2.state == fe.UNKNOWN and j2.date_source == fe.NO_DATE
    assert any("penále nelze určit" in x for x in j2.missing) and j2.contribution_status == "neúplný"
    assert j1.state == fe.UNKNOWN and "pronájem P1" in j1.reason        # neplatný záznam na stejném stroji
    assert certain_refs(led) == set() and led.status == fu.INCOMPLETE
    # jen zakázka bez termínu: zakázka se stejným strojem nemůže být jistá
    only = fe.period_ledger(model(jobs=[job("J1", 2), job("J2", 2)], needs=[need("J1", BAGR), need("J2", BAGR)]),
                            1, 10, job_starts={"J1": 1})
    assert row(only, "J1").state == fe.UNKNOWN and "J2" in row(only, "J1").reason
    # neplatné období
    for start, end in [(None, 5), (5, 1), (1, date(2026, 1, 2)), ("1", 3)]:
        bad = fe.period_ledger(m, start, end, job_starts={"J1": 1})
        assert bad.status == fu.NOT_COMPUTABLE and bad.entries.empty and bad.notes


# ---------------------------------------------------------------------------
# Ukázková data
# ---------------------------------------------------------------------------

def sample():
    import capacity_view as cv
    import scheduling as sch
    jobs = dm.legacy_jobs(BASE)
    plan = sch.evaluate(jobs, sch.optimize_schedule(jobs, dm.fleet_capacity(BASE))["starts"])
    return cv.job_starts_from_plan(jobs, plan)


STARTS = sample()


def test_sample_data_states_and_conflicting_revenue_is_not_certain():
    led = fe.period_ledger(BASE, 1, 25, job_starts=STARTS)
    states = dict(zip(led.entries["ref"], led.entries["state"]))
    assert {r for r, s in states.items() if s == fe.CONFLICT} == {"FIKT-P05", "J005", "J006"}
    assert {r for r, s in states.items() if s == fe.TENTATIVE} == {"FIKT-P06", "FIKT-P07"}
    assert {s for r, s in states.items()} == {fe.CERTAIN, fe.CONFLICT, fe.TENTATIVE}
    assert row(led, "J006").conflict_days == [15, 16, 17, 18] and row(led, "J005").conflict_days == [11]
    assert row(led, "FIKT-P05").md_conflict_in_period == 4
    # každý závazek z dat právě jednou
    assert sorted(led.entries["ref"]) == sorted(list(BASE.rental_bookings["booking_id"]) + list(BASE.jobs["job_id"]))
    # konfliktní výnosy nejsou mezi jistými
    assert certain_refs(led).isdisjoint({"FIKT-P05", "J005", "J006", "FIKT-P06", "FIKT-P07"})
    assert total(led, "jisté – pronájmy").value == 194500
    assert total(led, "jisté – zakázky").value == 635000
    assert total(led, "v konfliktu").value == 47500 + 115000 + 142000
    assert total(led, "předběžné").value == 40000 + 37500
    # naivní součet (economics.total_result bez kapacity) by konfliktní závazky zahrnul
    naive = ec.total_result(BASE, [(r.type_id, (int(r.end_day) - int(r.start_day) + 1), float(r.price_day))
                                   for r in BASE.rental_bookings.itertuples() if r.status == "confirmed"])
    assert naive["příspěvky z pronájmů"].value == 194500 + 47500
    # nic není úplné: chybí variabilní náklady a nastavení přímých nákladů
    assert led.status == fu.INCOMPLETE and all(v.status != "úplný" for k, v in led.totals.items()
                                               if v.value)
    assert set(led.entries["unit_assignment"]) >= {fe.UNIT_AMBIGUOUS, fe.UNIT_DERIVED}
    assert row(led, "FIKT-P04").unit_assignment == fe.UNIT_AMBIGUOUS and row(led, "FIKT-P04").state == fe.CERTAIN
    assert all(r.date_source == fe.FROM_PLAN for r in led.entries.itertuples() if r.kind == "job")


def test_sample_machine_days_are_never_counted_twice():
    led = fe.period_ledger(BASE, 1, 25, job_starts=STARTS)
    by_day = led.usage.by_day.set_index(["day", "type_id"])
    claims = {}
    for r in led.entries.itertuples():
        if r.booking_status != "confirmed" or r.start is None:
            continue
        for part in r.needs.split("; "):
            t, q = part.split("×")
            for d in range(max(r.start, 1), min(r.end, 25) + 1):
                claims[(d, t)] = claims.get((d, t), 0) + int(q)
    service = led.usage.by_unit[led.usage.by_unit["sources"].str.contains("servis")]
    for (d, t), x in by_day.iterrows():
        c = claims.get((d, t), 0)
        if x.status == fu.CONFLICT:            # konfliktní den: nároky + servis = konflikt + nárok nad kapacitu
            s = int(((service["day"] == d) & (service["type_id"] == t)).sum())
            assert c + s == x.conflict + x.excess, (d, t)
        else:                                  # jinak každý strojoden pronájmu / zakázky právě jednou
            assert c == x.rental + x.job, (d, t)
    # strojodny jistých závazků v období = pronájem + zakázky mimo konfliktní dny
    certain_md = led.entries.loc[led.entries["counted_as_certain"], "md_in_period"].sum()
    conflict_free = by_day[by_day["status"] != fu.CONFLICT]
    in_conflict_entries = led.entries[led.entries["state"] == fe.CONFLICT]
    assert certain_md + (in_conflict_entries["md_in_period"] - in_conflict_entries["md_conflict_in_period"]).sum() \
        == conflict_free["rental"].sum() + conflict_free["job"].sum()


def test_job_plan_is_a_proposal_and_demand_model_is_not_used():
    led = fe.period_ledger(BASE, 1, 25)                       # bez návrhu plánu se termíny nevymýšlejí
    jobs = led.entries[led.entries["kind"] == "job"]
    assert set(jobs["state"]) == {fe.UNKNOWN} and set(jobs["date_source"]) == {fe.NO_DATE}
    assert jobs["start"].isna().all()
    assert any("návrh" in r for r in led.open_rules)
    # modelová poptávka nemá na evidenci vliv
    other = dm.DataModel(**{n: getattr(BASE, n) for n in dm.SCHEMA if n != "rental_demand"},
                         rental_demand=BASE.rental_demand.assign(expected_days="1"))
    a = fe.period_ledger(BASE, 1, 25, job_starts=STARTS).entries
    b = fe.period_ledger(other, 1, 25, job_starts=STARTS).entries
    pd.testing.assert_frame_equal(a, b)


def test_contributions_come_from_economics_unchanged():
    led = fe.period_ledger(BASE, 1, 25, job_starts=STARTS)
    for r in BASE.rental_bookings.itertuples():
        expected = ec.rental_contribution(BASE, r.type_id, int(r.end_day) - int(r.start_day) + 1, float(r.price_day))
        assert led.results[("rental", r.booking_id)] == expected
    for j in BASE.jobs.itertuples():
        late = max(0, STARTS[j.job_id] + int(j.duration_days) - 1 - int(j.deadline_day))
        assert led.results[("job", j.job_id)] == ec.job_contribution(BASE, j.job_id, late)
    # zakázky po penále celkem = známých 892 000 Kč (horní mez) – evidence je jen rozděluje do skupin
    assert sum(led.results[("job", j)].value for j in BASE.jobs["job_id"]) == 892000


def test_inputs_are_not_modified():
    before = {n: getattr(BASE, n).copy() for n in dm.SCHEMA}
    starts = dict(STARTS)
    fingerprint = dm.data_fingerprint()
    fe.period_ledger(BASE, 1, 25, job_starts=starts)
    assert starts == STARTS and dm.data_fingerprint() == fingerprint
    for n, df in before.items():
        pd.testing.assert_frame_equal(getattr(BASE, n), df)


def test_module_is_standalone_and_not_connected_to_ui():
    tree = ast.parse((ROOT / "fleet_economics.py").read_text(encoding="utf-8"))
    imported = {n.module.split(".")[0] for n in ast.walk(tree) if isinstance(n, ast.ImportFrom) and n.module}
    imported |= {a.name.split(".")[0] for n in ast.walk(tree) if isinstance(n, ast.Import) for a in n.names}
    assert "streamlit" not in imported and imported <= {"dataclasses", "pandas", "capacity", "economics", "fleet_usage",
                                                         "econ_formula", "planner", "calendar", "itertools",
                                                         "datetime"}
    names = {n.name for n in tree.body if isinstance(n, ast.FunctionDef) and not n.name.startswith("_")}
    assert names == {"period_ledger", "conflict_options", "period_result"}   # jen K3.1–K3.3, žádné další kroky
    # fleet_view.py je zobrazovací vrstva K3.4a nad tímto modulem; sama do UI napojená není
    # (hlídá tests/test_fleet_view.py) – proto se tu vynechává.
    for path in [ROOT / "app.py", *ROOT.glob("*_view.py"), *(ROOT / "stranky").glob("*.py")]:
        if path.name != "fleet_view.py":
            assert "fleet_economics" not in path.read_text(encoding="utf-8"), path


@pytest.mark.parametrize("period", [(1, 1), (25, 25), (26, 40)])
def test_edge_periods(period):
    led = fe.period_ledger(BASE, *period, job_starts=STARTS)
    assert led.status in (fu.COMPLETE, fu.INCOMPLETE) and len(led.entries) == 19
    outside = led.entries["state"] == fe.OUTSIDE
    assert (led.entries.loc[outside, "md_in_period"] == 0).all()
    assert not led.entries.loc[outside, "counted_as_certain"].any()
