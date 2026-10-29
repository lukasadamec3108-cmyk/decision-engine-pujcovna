"""Testy jednotného využití fyzických strojů za období (E2.2 – K2). Jen fiktivní testovací data.

Spuštění: python -m pytest tests/test_fleet_usage.py
"""

import ast
import random
import sys
from datetime import date
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import capacity as cp  # noqa: E402
import datamodel as dm  # noqa: E402
import fleet_usage as fu  # noqa: E402
from capacity import Assignment, OpenJob, Unit  # noqa: E402

B = "BAGR"


def units(n=1, t=B):
    return [Unit(f"{t}-{i:02d}", t) for i in range(1, n + 1)]


def rental(ref, s, e, q=1, status="confirmed", unit=None, t=B):
    return Assignment("rental", ref, t, s, e, q, status, unit)


def job(ref, s, e, q=1, status="confirmed", unit=None, t=B):
    return Assignment("job", ref, t, s, e, q, status, unit)


def service(unit, s, e, t=B):
    return Assignment("maintenance", f"{unit} od {s}", t, s, e, 1, "confirmed", unit)


def usage(fleet, items=(), start=1, end=5, open_jobs=(), **kw):
    return fu.machine_days(cp.build_calendar(fleet, items, open_jobs), start, end, **kw)


def typ(u, t=B):
    return u.by_type.set_index("type_id").loc[t]


def day(u, d, t=B):
    return u.by_day[(u.by_day["day"] == d) & (u.by_day["type_id"] == t)].iloc[0]


def state(u, d, unit_id):
    return u.by_unit[(u.by_unit["day"] == d) & (u.by_unit["unit_id"] == unit_id)].iloc[0]["state"]


def assert_each_unit_once_and_balanced(u, cal):
    """Každý kus právě jednou za den; složky dne dávají dohromady kapacitu (je-li známá)."""
    expected = len(cp.days(u.start, u.end)) * len(cal.units)
    assert len(u.by_unit) == expected
    assert not u.by_unit.duplicated(["day", "unit_id"]).any()
    for r in u.by_day.itertuples():
        assert sum(getattr(r, g) for g in fu.GROUPS) == r.capacity
        cap = int(u.by_unit[(u.by_unit["day"] == r.day) & (u.by_unit["type_id"] == r.type_id)]["in_capacity"].sum())
        assert cap == r.capacity
    for r in u.by_type.itertuples():
        assert r.balanced is (True if r.capacity_known else None)


# ---------------------------------------------------------------------------
# Základní situace
# ---------------------------------------------------------------------------

def test_one_machine_one_rental():
    u = usage(units(1), [rental("P1", 2, 4)])
    t = typ(u)
    assert (t.capacity, t.rental, t.job, t.maintenance, t.conflict, t.unknown, t.free) == (5, 3, 0, 0, 0, 0, 2)
    assert u.status == fu.COMPLETE and u.conflicts == []
    assert [state(u, d, "BAGR-01") for d in range(1, 6)] == [fu.FREE, fu.RENTAL, fu.RENTAL, fu.RENTAL, fu.FREE]
    # rezervace uvádí jen typ – jediný kus je obsazený jistě, ale konkrétní kus není v datech
    assert t.unit_level.startswith("jen typ") and any("jen typ stroje" in n for n in u.notes)
    # stejný pronájem přiřazený konkrétnímu kusu
    u2 = usage(units(1), [rental("P1", 2, 4, unit="BAGR-01")])
    assert typ(u2).rental == 3 and typ(u2).unit_level == "konkrétní kusy"


def test_rental_and_job_on_the_same_day_is_conflict_not_double_count():
    cal = cp.build_calendar(units(1), [rental("P1", 1, 3), job("J1", 3, 5)])
    u = fu.machine_days(cal, 1, 5)
    t = typ(u)
    assert (t.rental, t.job, t.conflict, t.free, t.excess) == (2, 2, 1, 0, 1)
    assert t.rental + t.job + t.conflict == t.capacity == 5          # den 3 jen jednou
    assert state(u, 3, "BAGR-01") == fu.CONFLICT
    [c] = u.conflicts
    assert c["scope"] == "typ" and c["days"] == [3] and "konflikt pronájmu a vlastní zakázky" in c["kind"]
    assert c["sources"] == "pronájem P1, vlastní zakázka J1"
    assert_each_unit_once_and_balanced(u, cal)
    # dva kusy: obojí se vejde, jen nevíme, který kus dělá co
    u2 = usage(units(2), [rental("P1", 3, 3), job("J1", 3, 3)], 3, 3)
    assert (typ(u2).rental, typ(u2).job, typ(u2).conflict) == (1, 1, 0) and u2.conflicts == []
    assert state(u2, 3, "BAGR-01") == state(u2, 3, "BAGR-02") == fu.AMBIGUOUS
    # konkrétní kus se dvěma využitími
    u3 = usage(units(2), [rental("P1", 3, 3, unit="BAGR-01"), job("J1", 3, 3, unit="BAGR-01")], 3, 3)
    assert state(u3, 3, "BAGR-01") == fu.CONFLICT and state(u3, 3, "BAGR-02") == fu.FREE
    assert u3.conflicts[0]["scope"] == "kus" and typ(u3).conflict == 1 and typ(u3).free == 1


def test_service_during_rental():
    cal = cp.build_calendar(units(1), [rental("P1", 1, 5), service("BAGR-01", 3, 4)])
    u = fu.machine_days(cal, 1, 5)
    t = typ(u)
    assert (t.rental, t.maintenance, t.conflict, t.excess) == (3, 0, 2, 2)
    assert [state(u, d, "BAGR-01") for d in range(1, 6)] == [fu.RENTAL, fu.RENTAL, fu.CONFLICT, fu.CONFLICT, fu.RENTAL]
    [c] = u.conflicts
    assert c["days"] == [3, 4] and "konflikt se servisem" in c["kind"]
    assert_each_unit_once_and_balanced(u, cal)


def test_two_units_of_same_type_with_different_use():
    cal = cp.build_calendar(units(2), [rental("P1", 1, 3, unit="BAGR-01"), service("BAGR-02", 2, 4)])
    u = fu.machine_days(cal, 1, 5)
    assert [state(u, d, "BAGR-01") for d in range(1, 6)] == [fu.RENTAL] * 3 + [fu.FREE] * 2
    assert [state(u, d, "BAGR-02") for d in range(1, 6)] == [fu.FREE, fu.SERVICE, fu.SERVICE, fu.SERVICE, fu.FREE]
    t = typ(u)
    assert (t.capacity, t.rental, t.maintenance, t.free, t.conflict) == (10, 3, 3, 4, 0)
    assert t.unit_level == "konkrétní kusy" and u.conflicts == []
    # pronájem jen na typ + servis kusu 02 → pronájem jistě na kusu 01
    u2 = usage(units(2), [rental("P1", 2, 2), service("BAGR-02", 2, 2)], 2, 2)
    assert state(u2, 2, "BAGR-01") == fu.RENTAL and state(u2, 2, "BAGR-02") == fu.SERVICE
    assert_each_unit_once_and_balanced(u, cal)


def test_overlapping_bookings():
    u = usage(units(1), [rental("P1", 1, 3), rental("P2", 3, 5)])
    [c] = u.conflicts
    assert c["kind"] == "překročení kapacity typu – překryv rezervací" and c["days"] == [3]
    assert (typ(u).rental, typ(u).conflict, typ(u).excess) == (4, 1, 1)
    # navazující rezervace (1.–3. a 4.–5.) nejsou konflikt
    u2 = usage(units(1), [rental("P1", 1, 3), rental("P2", 4, 5)])
    assert u2.conflicts == [] and typ(u2).rental == 5
    # dva překrývající se servisy téhož kusu: jeden servis za den, žádný konflikt, jen poznámka
    u3 = usage(units(1), [service("BAGR-01", 1, 3), service("BAGR-01", 2, 4)])
    assert typ(u3).maintenance == 4 and u3.conflicts == []
    assert any("překrývající se servisní záznamy" in n for n in u3.notes)


def test_unknown_unit_assignment_is_explicit():
    u = usage(units(3), [rental("P1", 1, 5, q=2)])
    t = typ(u)
    assert (t.rental, t.free, t.ambiguous) == (10, 5, 15)                 # počty jisté
    assert set(u.by_unit["state"]) == {fu.AMBIGUOUS}                      # konkrétní kus ne
    note = u.by_unit["note"].iloc[0]
    assert "obsazeno 2 z 3 kusů bez určení konkrétního kusu" in note
    assert t.unit_level == "jen typ stroje – konkrétní kus neurčen"


def test_booking_statuses():
    items = [rental("P1", 1, 2), rental("P2", 2, 3, status="tentative"), job("J1", 4, 4, status="opportunity"),
             rental("P3", 5, 5, status="cancelled")]
    u = usage(units(1), items)
    t = typ(u)
    assert (t.rental, t.job, t.free, t.tentative) == (2, 0, 3, 3)          # předběžné neblokují
    assert u.conflicts == [] and any("předběžné" in n for n in u.notes)
    # opatrnější pohled: předběžné jako blokující → den 2 je překryv
    u2 = usage(units(1), items, include_tentative=True)
    assert u2.conflicts[0]["days"] == [2] and typ(u2).job == 1 and typ(u2).tentative == 0
    # neznámý stav (např. překlep) může kapacitu blokovat → volno nelze určit
    u3 = usage(units(1), [rental("P1", 1, 2), rental("P9", 4, 4, status="potvrzeno")])
    t3 = typ(u3)
    assert u3.status == fu.INCOMPLETE and (t3.rental, t3.unknown, t3.free) == (2, 3, 0)
    assert any("neznámý stav" in n for n in u3.notes)


def test_period_across_month_end():
    cal = cp.build_calendar(units(2), [rental("P1", date(2026, 1, 30), date(2026, 2, 2)),
                                       service("BAGR-02", date(2026, 2, 1), date(2026, 2, 3))])
    u = fu.machine_days(cal, date(2026, 1, 31), date(2026, 2, 2))
    assert list(u.by_day["day"]) == [date(2026, 1, 31), date(2026, 2, 1), date(2026, 2, 2)]
    t = typ(u)
    assert (t.capacity, t.rental, t.maintenance, t.free) == (6, 3, 2, 1)
    assert state(u, date(2026, 2, 1), "BAGR-01") == fu.RENTAL
    assert_each_unit_once_and_balanced(u, cal)
    # číselné dny z dat E1 převedené na kalendářní data (den 1 = 30. 1.)
    model = dm.load_model()
    u2 = fu.from_model(model, date(2026, 1, 30), date(2026, 3, 1), plan_start=date(2026, 1, 30))
    assert u2.status != fu.NOT_COMPUTABLE and len(u2.by_day["day"].unique()) == 31


def test_invalid_or_incomplete_dates():
    cal = cp.build_calendar(units(1), [rental("P1", 1, 2)])
    for start, end, text in [(None, 5, "chybí"), (5, 1, "konec období je před začátkem"),
                             (1, date(2026, 1, 5), "míchá"), (date(2026, 1, 1), date(2026, 1, 5), "jiný druh"),
                             ("1", 5, "neplatné datum"), (True, 5, "neplatné datum")]:
        u = fu.machine_days(cal, start, end)
        assert u.status == fu.NOT_COMPUTABLE and text in u.notes[0], (start, end, u.notes)
        assert u.by_day.empty and u.by_unit.empty
    # potvrzený pronájem bez platného termínu: nezahodí se, volné strojodny jsou „nelze určit“
    u = usage(units(1), [rental("P1", 1, 2), rental("P2", None, 4), rental("P3", 5, 3)])
    t = typ(u)
    assert u.status == fu.INCOMPLETE and (t.rental, t.unknown, t.free) == (2, 3, 0)
    assert day(u, 4).status == fu.UNKNOWN and state(u, 4, "BAGR-01") == fu.UNKNOWN
    assert sum("chybí nebo je neplatný termín" in n for n in u.notes) == 2
    # kus s neplatnou dostupností → počet kusů nejistý, součet strojodnů výslovně neurčitelný
    fleet = units(1) + [Unit("BAGR-02", B, "owned", 7, 3)]
    cal2 = cp.build_calendar(fleet, [rental("P1", 1, 1)])
    u2 = fu.machine_days(cal2, 1, 2)
    assert not typ(u2).capacity_known
    assert typ(u2).balanced is None and u2.status == fu.INCOMPLETE
    assert u2.fleet["excluded"] == [B] and any("neurčitelný" in n for n in u2.notes)


def test_open_job_without_fixed_start_makes_free_days_uncertain():
    u = usage(units(2), [], open_jobs=[OpenJob("J9", B, 3, None, 2)])
    assert [day(u, d).unknown for d in range(1, 6)] == [0, 0, 1, 1, 1]
    assert [day(u, d).free for d in range(1, 6)] == [2, 2, 1, 1, 1]
    assert u.status == fu.INCOMPLETE and any("bez pevného termínu" in n for n in u.notes)


def test_unit_outside_fleet_is_not_capacity_and_use_is_conflict():
    fleet = [Unit("BAGR-01", B), Unit("BAGR-02", B, "ordered", 4, None)]
    cal = cp.build_calendar(fleet, [rental("P1", 2, 4, unit="BAGR-02")])
    u = fu.machine_days(cal, 1, 5)
    assert [day(u, d).capacity for d in range(1, 6)] == [1, 1, 1, 2, 2]
    assert state(u, 2, "BAGR-02") == fu.OUT and state(u, 4, "BAGR-02") == fu.RENTAL
    c = [c for c in u.conflicts if c["scope"] == "kus"]
    assert c and c[0]["days"] == [2, 3] and typ(u).excess == 2
    assert_each_unit_once_and_balanced(u, cal)
    # servis kusu před dodáním kapacitu neubírá ani se neobjeví mezi zdroji konfliktu typu
    cal2 = cp.build_calendar(fleet, [service("BAGR-02", 1, 1), rental("P1", 1, 1, q=2)])
    u2 = fu.machine_days(cal2, 1, 1)
    [c2] = u2.conflicts
    assert c2["sources"] == "pronájem P1" and c2["excess"] == 1 and day(u2, 1).maintenance == 0
    assert state(u2, 1, "BAGR-02") == fu.OUT


# ---------------------------------------------------------------------------
# Součty, náhodné případy a shoda s capacity.py
# ---------------------------------------------------------------------------

def test_sum_of_groups_equals_capacity_in_random_cases():
    rng = random.Random(7)
    for _ in range(150):
        n = rng.randint(0, 3)
        fleet = units(n)
        items = []
        for i in range(rng.randint(0, 6)):
            s = rng.randint(1, 10)
            e = s + rng.randint(0, 4)
            unit = rng.choice([None, None] + [x.unit_id for x in fleet])
            kind = rng.choice(["rental", "job", "maintenance"] if unit else ["rental", "job"])
            status = rng.choice(["confirmed", "confirmed", "tentative", "cancelled"])
            if kind == "maintenance":
                items.append(service(unit, s, e))
            else:
                items.append(Assignment(kind, f"X{i}", B, s, e, 1 if unit else rng.randint(1, 2), status, unit))
        cal = cp.build_calendar(fleet, items)
        u = fu.machine_days(cal, 1, 12)
        assert_each_unit_once_and_balanced(u, cal)
        assert u.fleet["balanced"]
        # bez nároků na konkrétní kus (kromě servisu) se počty shodují s capacity.daily_capacity
        if all(a.unit_id is None or a.kind == "maintenance" for a in cal.assignments) and fleet:
            dc = cp.daily_capacity(cal, 1, 12).set_index("day")
            for r in u.by_day.itertuples():
                ref = dc.loc[r.day]
                assert r.capacity == ref.units
                if ref.free >= 0:
                    assert (r.rental, r.job, r.maintenance, r.free) == (ref.rental, ref.job, ref.maintenance, ref.free)
                    assert r.status != fu.CONFLICT
                else:
                    assert r.status == fu.CONFLICT and r.conflict == ref.units and r.excess == -ref.free


def sample():
    import capacity_view as cv
    import scheduling as sch
    model = dm.load_model()
    jobs = dm.legacy_jobs(model)
    plan = sch.evaluate(jobs, sch.optimize_schedule(jobs, dm.fleet_capacity(model))["starts"])
    starts = cv.job_starts_from_plan(jobs, plan)
    return model, starts, cv.build(model, starts)


def test_sample_data_conflicts_and_consistency_with_existing_overview():
    import capacity_view as cv
    model, starts, cal = sample()
    u = fu.from_model(model, 1, 25, job_starts=starts)
    assert u.status == fu.COMPLETE
    assert [(c["id"], c["days"], c["sources"]) for c in u.conflicts] == [
        ("KOLOVY_BAGR", [11], "servis / údržba KOLOVY_BAGR-01 od dne 11, vlastní zakázka J005"),
        ("SILNICNI_VALEC", [15, 16, 17, 18], "pronájem FIKT-P05, vlastní zakázka J006"),
    ]
    # stejné konflikty jako capacity.find_conflicts (stránka Plánování kapacity)
    ref = [(c["id"], c["days"]) for c in cp.find_conflicts(cal, 1, 25)]
    assert [(c["id"], c["days"]) for c in u.conflicts] == ref
    # stejné strojodny jako přehled využití (capacity_view.utilization_table)
    names = cv.type_names(model)
    table = cv.utilization_table(cv.grid(cal, 1, 25, cal.type_ids()), names, cal.unknown_capacity)
    mine = u.by_type.set_index("type_id")
    for r in table.itertuples():
        for g in fu.GROUPS:
            assert mine.loc[r.type_id, g] == getattr(r, g), (r.type_id, g)
        assert mine.loc[r.type_id, "excess"] == r.excess_demand
    assert u.fleet["capacity"] == 9 * 25 and u.fleet["balanced"]
    assert set(u.by_type["unit_level"]) == {"jen typ stroje – konkrétní kus neurčen"}
    assert_each_unit_once_and_balanced(u, cal)


def test_inputs_and_data_are_not_modified():
    model, starts, cal = sample()
    before = {n: getattr(model, n).copy() for n in dm.SCHEMA}
    fingerprint = dm.data_fingerprint()
    snapshot = (list(cal.units), list(cal.assignments), list(cal.issues))
    fu.machine_days(cal, 1, 25)
    fu.from_model(model, 1, 25, job_starts=starts)
    assert (list(cal.units), list(cal.assignments), list(cal.issues)) == snapshot
    for n, df in before.items():
        pd.testing.assert_frame_equal(getattr(model, n), df)
    assert dm.data_fingerprint() == fingerprint


def test_module_is_standalone_and_not_connected_to_ui():
    tree = ast.parse((ROOT / "fleet_usage.py").read_text(encoding="utf-8"))
    imported = {n.module.split(".")[0] for n in ast.walk(tree) if isinstance(n, ast.ImportFrom) and n.module}
    imported |= {a.name.split(".")[0] for n in ast.walk(tree) if isinstance(n, ast.Import) for a in n.names}
    assert "streamlit" not in imported
        # K3.4c: app.py modul používá jen na stránce Ekonomika flotily (výpočet v mezipaměti); jinde nikde
    sys.path.insert(0, str(ROOT / "tests"))
    from app_pages import outside_economics_page
    assert outside_economics_page("fleet_usage", "fu") == []
    for path in [*ROOT.glob("*_view.py"), *(ROOT / "stranky").glob("*.py")]:
        if path.name != "fleet_view.py":
            assert "fleet_usage" not in path.read_text(encoding="utf-8"), path
