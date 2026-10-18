"""Testy kapacitního plánování (capacity.py). Jen fiktivní testovací data.

Spuštění: python -m pytest tests   nebo   python tests/test_capacity.py
"""

import shutil
import sys
import tempfile
from datetime import date
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import capacity as cp  # noqa: E402
import datamodel as dm  # noqa: E402
from capacity import Assignment, OpenJob, Unit  # noqa: E402

BAGR = "MINI_BAGR"


def fleet(n=1, t=BAGR):
    return [Unit(f"{t}-{i:02d}", t) for i in range(1, n + 1)]


def rental(ref, s, e, q=1, status="confirmed", t=BAGR, unit=None):
    return Assignment("rental", ref, t, s, e, q, status, unit)


def job(ref, s, e, q=1, status="confirmed", t=BAGR, unit=None):
    return Assignment("job", ref, t, s, e, q, status, unit)


def service(unit, s, e, t=BAGR):
    return Assignment("maintenance", f"S-{unit}", t, s, e, 1, "confirmed", unit)


def day_row(cal, d, t=BAGR):
    cap = cp.daily_capacity(cal, d, d, [t])
    return cap.iloc[0]


# ---------------------------------------------------------------------------
# Základní pravidla dnů
# ---------------------------------------------------------------------------

def test_days_are_inclusive_and_mixing_is_rejected():
    assert cp.days(3, 5) == [3, 4, 5]
    assert cp.days(date(2026, 1, 30), date(2026, 2, 2))[-1] == date(2026, 2, 2)
    assert cp.overlap(1, 5, 5, 8) and not cp.overlap(1, 5, 6, 8)
    try:
        cp.days(1, date(2026, 1, 1))
        assert False, "míchání čísel a dat musí selhat"
    except TypeError:
        pass


# 1) Jeden stroj, jedna rezervace
def test_one_machine_one_rental():
    cal = cp.build_calendar(fleet(1), [rental("R1", 3, 5)])
    cap = cp.daily_capacity(cal, 1, 6)
    assert cap["rental"].tolist() == [0, 0, 1, 1, 1, 0]
    assert cap["status"].tolist() == ["volno", "volno", "plno", "plno", "plno", "volno"]
    assert cp.find_conflicts(cal) == []
    r = cp.evaluate_request(cal, BAGR, 4, 4)
    assert r["status"] == cp.INFEASIBLE and r["problem_days"] == [(4, 0, 1, ["pronájem R1"])]
    assert cp.evaluate_request(cal, BAGR, 1, 2)["status"] == cp.FEASIBLE


# 2) Jeden stroj, dvě překrývající se rezervace
def test_one_machine_two_overlapping_rentals():
    cal = cp.build_calendar(fleet(1), [rental("R1", 1, 5), rental("R2", 4, 8)])
    conf = cp.find_conflicts(cal)
    assert len(conf) == 1
    c = conf[0]
    assert c["scope"] == "typ" and c["kind"] == "překročení kapacity typu – překryv rezervací"
    assert c["days"] == [4, 5] and c["sources"] == ["pronájem R1", "pronájem R2"] and c["max_over"] == 1
    assert day_row(cal, 4)["status"] == "překročeno" and day_row(cal, 4)["free"] == -1
    # stejný konkrétní kus dvakrát → konflikt kusu
    cal2 = cp.build_calendar(fleet(2), [rental("R1", 1, 5, unit="MINI_BAGR-01"),
                                        rental("R2", 5, 6, unit="MINI_BAGR-01")])
    unit_conf = [c for c in cp.find_conflicts(cal2) if c["scope"] == "kus"]
    assert unit_conf == [dict(kind="překryv rezervací", scope="kus", id="MINI_BAGR-01", days=[5],
                              sources=["pronájem R1", "pronájem R2"])]
    # celkově jsou ale volné 2 kusy → typ přetížen není
    assert not [c for c in cp.find_conflicts(cal2) if c["scope"] == "typ"]


# 3) Více kusů stejného typu
def test_multiple_units_of_same_type():
    cal = cp.build_calendar(fleet(3), [rental("R1", 1, 4, q=2)])
    row = day_row(cal, 2)
    assert row["units"] == 3 and row["rental"] == 2 and row["free"] == 1 and row["status"] == "volno"
    assert cp.evaluate_request(cal, BAGR, 2, 3, quantity=1)["status"] == cp.FEASIBLE
    r = cp.evaluate_request(cal, BAGR, 2, 6, quantity=2)
    assert r["status"] == cp.INFEASIBLE and [p[0] for p in r["problem_days"]] == [2, 3, 4]
    assert r["free_by_day"][5] == 3 and r["min_free"] == 1
    # počet kusů ≠ strojodny: 2 kusy × 4 dny = 8 strojodnů, ale v každém dni jen 2 kusy
    assert cp.daily_capacity(cal, 1, 4)["rental"].max() == 2


# 4) Pronájem + vlastní zakázka současně
def test_rental_and_job_at_same_time():
    cal = cp.build_calendar(fleet(1), [rental("R1", 1, 5), job("J1", 3, 7)])
    conf = cp.find_conflicts(cal)
    assert len(conf) == 1 and conf[0]["kind"] == "překročení kapacity typu – konflikt pronájmu a vlastní zakázky"
    assert conf[0]["days"] == [3, 4, 5] and conf[0]["sources"] == ["pronájem R1", "vlastní zakázka J1"]
    assert day_row(cal, 4)[["rental", "job"]].tolist() == [1, 1]
    # se dvěma kusy je vše v pořádku a modul neurčuje, kdo má přednost
    assert cp.find_conflicts(cp.build_calendar(fleet(2), [rental("R1", 1, 5), job("J1", 3, 7)])) == []


# 5) Stroj v servisu
def test_machine_in_service():
    units = fleet(2)
    cal = cp.build_calendar(units, [service("MINI_BAGR-01", 3, 4)])
    row = day_row(cal, 3)
    assert row["units"] == 2 and row["maintenance"] == 1 and row["free"] == 1
    occ = cp.unit_occupancy(cal, 3, 3).set_index("unit_id")["state"]
    assert occ["MINI_BAGR-01"] == "servis / údržba" and occ["MINI_BAGR-02"] == "bez přiřazení"
    r = cp.evaluate_request(cal, BAGR, 2, 5, quantity=2)
    assert r["status"] == cp.INFEASIBLE
    assert [(p[0], p[3]) for p in r["problem_days"]] == [(3, ["servis / údržba S-MINI_BAGR-01"]),
                                                          (4, ["servis / údržba S-MINI_BAGR-01"])]
    # konkrétní kus pronajatý během servisu
    cal2 = cp.build_calendar(units, [service("MINI_BAGR-01", 3, 4), rental("R1", 4, 6, unit="MINI_BAGR-01")])
    assert [c["kind"] for c in cp.find_conflicts(cal2)] == ["konflikt se servisem"]
    # na úrovni typu: jediný kus v servisu a současně pronájem
    cal3 = cp.build_calendar(fleet(1), [service("MINI_BAGR-01", 3, 4), rental("R1", 4, 6)])
    c3 = cp.find_conflicts(cal3)
    assert len(c3) == 1 and c3[0]["kind"] == "překročení kapacity typu – konflikt se servisem" and c3[0]["days"] == [4]


# 6) Přes konec měsíce
def test_across_month_end():
    cal = cp.build_calendar(fleet(1), [rental("R1", date(2026, 1, 29), date(2026, 2, 2))])
    cap = cp.daily_capacity(cal, date(2026, 1, 28), date(2026, 2, 3))
    assert len(cap) == 7 and cap["rental"].tolist() == [0, 1, 1, 1, 1, 1, 0]
    r = cp.evaluate_request(cal, BAGR, date(2026, 2, 1), date(2026, 2, 5))
    assert r["status"] == cp.INFEASIBLE
    assert [p[0] for p in r["problem_days"]] == [date(2026, 2, 1), date(2026, 2, 2)]
    # únor v přestupném roce
    leap = cp.build_calendar(fleet(1), [rental("R2", date(2028, 2, 28), date(2028, 3, 1))])
    assert len(cp.daily_capacity(leap, date(2028, 2, 28), date(2028, 3, 1))) == 3


# 7) Přes konec roku
def test_across_year_end():
    cal = cp.build_calendar(fleet(1), [rental("R1", date(2026, 12, 30), date(2027, 1, 2)),
                                       job("J1", date(2027, 1, 1), date(2027, 1, 3))])
    conf = cp.find_conflicts(cal)
    assert len(conf) == 1 and conf[0]["days"] == [date(2027, 1, 1), date(2027, 1, 2)]
    assert cp.evaluate_request(cal, BAGR, date(2027, 1, 4), date(2027, 1, 10))["status"] == cp.FEASIBLE
    assert cp.evaluate_request(cal, BAGR, date(2026, 12, 31), date(2026, 12, 31))["status"] == cp.INFEASIBLE


# 8) Požadavek přesně navazující na předchozí rezervaci
def test_request_adjacent_to_previous_booking():
    cal = cp.build_calendar(fleet(1), [rental("R1", 1, 5)])
    assert cp.evaluate_request(cal, BAGR, 6, 9)["status"] == cp.FEASIBLE       # den po skončení
    r = cp.evaluate_request(cal, BAGR, 5, 9)                                   # předání ve stejný den
    assert r["status"] == cp.INFEASIBLE and [p[0] for p in r["problem_days"]] == [5]
    calm = cp.build_calendar(fleet(1), [rental("R1", date(2026, 3, 1), date(2026, 3, 31))])
    assert cp.evaluate_request(calm, BAGR, date(2026, 4, 1), date(2026, 4, 3))["status"] == cp.FEASIBLE


# 9) Překročení kapacity o jeden kus
def test_over_capacity_by_one():
    cal = cp.build_calendar(fleet(2), [rental("R1", 1, 5), job("J1", 1, 5)])
    assert cp.find_conflicts(cal) == []
    r = cp.evaluate_request(cal, BAGR, 3, 3)
    assert r["status"] == cp.INFEASIBLE and r["problem_days"] == [(3, 0, 1, ["pronájem R1", "vlastní zakázka J1"])]
    cal2 = cp.build_calendar(fleet(2), [rental("R1", 1, 5), job("J1", 1, 5), rental("R2", 3, 3)])
    conf = cp.find_conflicts(cal2)
    assert len(conf) == 1 and conf[0]["max_over"] == 1 and conf[0]["days"] == [3]


# 10) Neúplná / neplatná data → nikdy „volno“
def test_incomplete_or_invalid_data_is_never_reported_as_free():
    bad = [rental("R1", 1, None), rental("R2", 8, 3), rental("R3", 1, 2, q=0),
           rental("R4", 1, 2, t="NEZNAMY"), Assignment("rental", "R5", BAGR, 1, 2, 1, "potvrzeno")]
    for record in bad:
        cal = cp.build_calendar(fleet(3), [record])
        assert cal.issues, record
        r = cp.evaluate_request(cal, BAGR, 1, 2)
        if record.type_id == BAGR and record.status == "confirmed":
            assert r["status"] == cp.UNKNOWN and r["reasons"], record
            assert set(cp.daily_capacity(cal, 1, 2)["status"]) == {cp.UNKNOWN}
        assert r["status"] != "volno"
    # neplatný kus flotily: neznámý počet kusů → nelze určit, ani když je kapacita „jinak“ málo
    cal = cp.build_calendar([Unit("MINI_BAGR-01", BAGR), Unit("MINI_BAGR-02", BAGR, "owned", "x")],
                            [rental("R1", 1, 3)])
    r = cp.evaluate_request(cal, BAGR, 2, 2)
    assert r["status"] == cp.UNKNOWN and any("neplatné datum" in x for x in r["reasons"])
    # neplatný požadavek
    assert cp.evaluate_request(cal, BAGR, 5, 1)["status"] == cp.UNKNOWN
    assert cp.evaluate_request(cal, BAGR, 1, 2, quantity=0)["status"] == cp.UNKNOWN
    # neúplný záznam nepomůže, když kusy chybí už podle jistých údajů
    cal2 = cp.build_calendar(fleet(1), [rental("R1", 1, 5), rental("R2", 1, None)])
    assert cp.evaluate_request(cal2, BAGR, 2, 2)["status"] == cp.INFEASIBLE


def test_open_job_without_date_makes_result_uncertain():
    # potvrzená zakázka bez termínu v okně 1–10 → nevíme, kdy stroj zabere
    cal = cp.build_calendar(fleet(1), open_jobs=[OpenJob("J1", BAGR, 1, 10, 3)])
    r = cp.evaluate_request(cal, BAGR, 4, 5)
    assert r["status"] == cp.UNKNOWN and "J1" in r["reasons"][0]
    assert day_row(cal, 4)["status"] == cp.UNKNOWN and day_row(cal, 11)["status"] == "volno"
    assert cp.evaluate_request(cal, BAGR, 11, 12)["status"] == cp.FEASIBLE
    # s rezervou kusů je jisté i v nejhorším případě
    cal2 = cp.build_calendar(fleet(2), open_jobs=[OpenJob("J1", BAGR, 1, 10, 3)])
    assert cp.evaluate_request(cal2, BAGR, 4, 5)["status"] == cp.FEASIBLE
    # dvě zakázky bez termínu se v nejhorším případě sčítají
    cal3 = cp.build_calendar(fleet(2), open_jobs=[OpenJob("J1", BAGR, 1, 10, 3), OpenJob("J2", BAGR, 1, 10, 3)])
    assert cp.evaluate_request(cal3, BAGR, 4, 5)["status"] == cp.UNKNOWN


# 11) Zrušená a předběžná rezervace
def test_cancelled_and_tentative_bookings():
    cal = cp.build_calendar(fleet(1), [rental("R1", 1, 5, status="cancelled")])
    assert cal.assignments == [] and not cal.issues
    assert cp.evaluate_request(cal, BAGR, 1, 5)["status"] == cp.FEASIBLE
    cal2 = cp.build_calendar(fleet(1), [rental("R1", 1, 5, status="tentative")])
    r = cp.evaluate_request(cal2, BAGR, 3, 3)
    assert r["status"] == cp.FEASIBLE and r["warnings"] and "R1" in r["warnings"][0][1]
    assert day_row(cal2, 3)["tentative"] == 1 and day_row(cal2, 3)["free"] == 1
    assert cp.evaluate_request(cal2, BAGR, 3, 3, include_tentative=True)["status"] == cp.INFEASIBLE
    assert cp.find_conflicts(cp.build_calendar(fleet(1), [rental("R1", 1, 5), rental("R2", 1, 5, status="tentative")])) == []


# 12) Známý jen typ, ne konkrétní kus
def test_only_type_known_not_unit():
    cal = cp.build_calendar(fleet(2), [rental("R1", 1, 3)])
    occ = cp.unit_occupancy(cal, 2, 2)
    assert set(occ["state"]) == {"bez přiřazení"}               # žádnému kusu se rezervace nepřiřadí
    assert day_row(cal, 2)["free"] == 1
    # rezervace na typ + servis konkrétního kusu: kapacita typu se počítá správně
    cal2 = cp.build_calendar(fleet(2), [rental("R1", 1, 3), service("MINI_BAGR-02", 2, 2)])
    assert day_row(cal2, 2)["free"] == 0 and cp.find_conflicts(cal2) == []
    assert cp.evaluate_request(cal2, BAGR, 2, 2)["status"] == cp.INFEASIBLE


def test_unit_availability_and_ordered_units():
    units = [Unit("MINI_BAGR-01", BAGR), Unit("MINI_BAGR-02", BAGR, "ordered", 10),
             Unit("MINI_BAGR-03", BAGR, "ordered"), Unit("MINI_BAGR-04", BAGR, "leased", None, 5)]
    cal = cp.build_calendar(units)
    cap = cp.daily_capacity(cal, 4, 11).set_index("day")["units"]
    assert cap[4] == 2 and cap[5] == 2 and cap[6] == 1 and cap[10] == 2   # objednaný bez data se nepočítá


def test_inputs_are_not_modified_and_nothing_is_rescheduled():
    items = [rental("R1", 1, 5), rental("R2", 3, 6)]
    snapshot = list(items)
    cal = cp.build_calendar(fleet(1), items)
    cp.find_conflicts(cal)
    cp.evaluate_request(cal, BAGR, 1, 9)
    assert items == snapshot and cal.assignments == snapshot


# ---------------------------------------------------------------------------
# Napojení na datový model E1
# ---------------------------------------------------------------------------

def model_with(**edits):
    tmp = tempfile.mkdtemp()
    shutil.copytree(dm.DEFAULT_DIR, tmp, dirs_exist_ok=True)
    for table, change in edits.items():
        path = Path(tmp) / f"{table}.csv"
        df = pd.read_csv(path, dtype=str, keep_default_na=False)
        change(df).to_csv(path, index=False)
    model = dm.load_model(tmp)
    shutil.rmtree(tmp)
    return model


def test_from_demo_model():
    model = dm.load_model()
    cal = cp.from_model(model)
    assert not cal.issues
    assert len(cal.units) == len(model.fleet_units)
    confirmed = model.jobs[model.jobs["status"] == "confirmed"]
    assert {j.ref for j in cal.open_jobs} == set(confirmed["job_id"])     # zakázky nemají pevný termín
    # zakázky bez termínu → jistota nikde, kde mohou proběhnout
    r = cp.evaluate_request(cal, BAGR, 1, 3)
    assert r["status"] == cp.UNKNOWN
    # s termíny z plánu (např. výstup tab 6) se stanou pevným využitím
    starts = {j: 1 for j in confirmed["job_id"]}
    cal2 = cp.from_model(model, job_starts=starts)
    assert cal2.open_jobs == [] and any(a.kind == "job" for a in cal2.assignments)


def test_from_model_with_dates_and_bookings():
    def bookings(df):
        rows = [dict(booking_id="B1", type_id=BAGR, start_day="30", end_day="33", units="1",
                     price_day="6500", with_operator="", status="confirmed"),
                dict(booking_id="B2", type_id=BAGR, start_day="31", end_day="31", units="1",
                     price_day="6500", with_operator="", status="tentative")]
        return pd.concat([df, pd.DataFrame(rows)], ignore_index=True)

    def service_rows(df):
        return pd.concat([df, pd.DataFrame([dict(unit_id="MINI_BAGR-01", start_day="40", days="2", note="")])],
                         ignore_index=True)

    model = model_with(rental_bookings=bookings, maintenance=service_rows)
    cal = cp.from_model(model, plan_start=date(2026, 12, 1))
    b1 = [a for a in cal.assignments if a.ref == "B1"][0]
    assert (b1.start, b1.end, b1.unit_id) == (date(2026, 12, 30), date(2027, 1, 2), None)
    s = [a for a in cal.assignments if a.kind == "maintenance"][0]
    assert (s.start, s.end) == (date(2027, 1, 9), date(2027, 1, 10))
    n = int((model.fleet_units["type_id"] == BAGR).sum())
    row = day_row(cal, date(2027, 1, 1))
    assert row["units"] == n and row["rental"] == 1
    assert day_row(cal, date(2026, 12, 31))["tentative"] == 1


def test_from_model_invalid_booking_is_flagged():
    def bookings(df):
        row = dict(booking_id="B1", type_id=BAGR, start_day="5", end_day="", units="1",
                   price_day="6500", with_operator="", status="confirmed")
        return pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    cal = cp.from_model(model_with(rental_bookings=bookings), job_starts={})
    assert BAGR in cal.unknown_occupancy and cal.issues_for(BAGR)
    assert cp.evaluate_request(cal, BAGR, 50, 51)["status"] == cp.UNKNOWN


def test_app_does_not_use_capacity_module_yet():
    import re
    src = (ROOT / "app.py").read_text(encoding="utf-8")
    assert not re.search(r"^\s*(import capacity|from capacity |.*reload\(capacity)", src, re.M)


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK", name)
