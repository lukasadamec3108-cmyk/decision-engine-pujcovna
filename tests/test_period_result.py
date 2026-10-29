"""Testy ekonomického výsledku flotily za období (E2.2 – K3.3, fleet_economics.period_result).

Jen fiktivní data. Spuštění: python -m pytest tests/test_period_result.py
"""

import sys
from datetime import date
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))
import economics as ec  # noqa: E402
import fleet_economics as fe  # noqa: E402
import fleet_usage as fu  # noqa: E402
from test_fleet_economics import BAGR, BASE, DUMPER, STARTS, booking, job, model, need  # noqa: E402

UNITS = ((f"{BAGR}-01", BAGR), (f"{DUMPER}-01", DUMPER))
FIXED_MONTH = 45000 + 12000 + 55000 + 15000          # fixní náklady + průměrný servis (bagr, dumper) / měsíc


def complete_types(mt):
    mt = mt.copy()
    mt.loc[mt["type_id"] == BAGR, ["variable_cost_rental_day", "variable_cost_job_day"]] = ["500", "300"]
    mt.loc[mt["type_id"] == DUMPER, ["variable_cost_rental_day", "variable_cost_job_day"]] = ["400", "0"]
    return mt


def full(overhead="20000", **kw):
    """Model s úplnými údaji (sazby, nastavení přímých nákladů, režie; overhead=None = nezadaná režie)."""
    m = model(types=complete_types, settings={"direct_cost_includes_machines": "ne"}, **kw)
    st = m.settings.copy()                       # režie už v nastavení je (prázdná) – přepsat, nepřidávat
    st.loc[st["key"] == "overhead_cost_month", "value"] = "" if overhead is None else overhead
    m.settings = st
    return m


def plain_model(**kw):
    """Model bez režie v nastavení (ukázková data ji mají prázdnou)."""
    return model(**kw)


def result_without_conflicts(m, starts=None, months=1.0, start=1, end=10):
    return fe.period_result(m, start, end, starts, fixed_cost_method="months", months=months)


# ---------------------------------------------------------------------------
# Bez konfliktů a soulad s economics.total_result
# ---------------------------------------------------------------------------

def test_no_conflicts_matches_total_result_with_complete_inputs():
    m = full(bookings=[booking("P1", BAGR, 1, 4), booking("P2", DUMPER, 6, 8, price=7000)],
             jobs=[job("J1", 3, deadline=4, penalty=1500)], needs=[need("J1", DUMPER)])
    r = result_without_conflicts(m, {"J1": 1}, months=1.0)
    ref = ec.total_result(m, rentals=[(BAGR, 4, 6500.0), (DUMPER, 3, 7000.0)], jobs=[("J1", 0)], months=1.0)
    assert r.status == fu.COMPLETE and r.result.complete and not r.missing
    assert r.result.value == ref["výsledek"].value
    assert r.rentals.value == ref["příspěvky z pronájmů"].value == 4 * (6500 - 500) + 3 * (7000 - 400)
    assert r.jobs.value == ref["příspěvky ze zakázek"].value == 100000 - 40000 - 3 * 0
    assert r.fixed_costs.value == ref["fixní náklady flotily"].value == FIXED_MONTH
    assert r.overhead.value == ref["režie"].value == 20000
    assert r.range == (r.result.value, r.result.value) and r.conflicts == []
    assert not any("NEJDE" in n for n in r.notes)


def test_no_conflicts_matches_total_result_with_incomplete_inputs():
    m = plain_model(bookings=[booking("P1", BAGR, 1, 4)], jobs=[job("J1", 3)], needs=[need("J1", DUMPER)])
    r = result_without_conflicts(m, {"J1": 1}, months=0.5)
    ref = ec.total_result(m, rentals=[(BAGR, 4, 6500.0)], jobs=[("J1", 0)], months=0.5)
    assert r.result.value == ref["výsledek"].value and set(r.result.missing) == set(ref["výsledek"].missing)
    assert r.status == fu.INCOMPLETE and any("NEJDE o přesný čistý zisk" in n for n in r.notes)


# ---------------------------------------------------------------------------
# Konflikty
# ---------------------------------------------------------------------------

def test_rental_vs_job_conflict_is_not_counted_and_gives_range():
    m = full(bookings=[booking("P1", BAGR, 1, 5), booking("P2", DUMPER, 1, 2)],
             jobs=[job("J1", 3, deadline=10)], needs=[need("J1", BAGR)])
    r = result_without_conflicts(m, {"J1": 4}, months=1.0, end=12)
    assert r.rentals.value == 2 * (8000 - 400)                          # jen P2 – P1 je v konfliktu
    assert r.jobs.value == 0
    assert r.status == fu.INCOMPLETE and any("konfliktní závazky nejsou započteny (J1, P1)" in x for x in r.missing)
    assert [c["id"] for c in r.conflicts] == ["K1"] and {x["ref"] for x in r.conflicts[0]["commitments"]} == {"P1", "J1"}
    base = r.result.value
    shift = [s for s in r.scenarios if s["variants"] == {"K1": "posun J1"}][0]
    expected = base + 5 * (6500 - 500) + (100000 - 40000 - 3 * 300)     # celé závazky, každý jednou
    assert shift["valid"] and shift["in_period"] and shift["bounds"] == (expected, expected)
    cancel = [s for s in r.scenarios if s["variants"] == {"K1": "zrušení P1"}][0]
    assert cancel["bounds"] == (None, base + 100000 - 40000 - 3 * 300)   # smluvní náhrada neznámá
    assert r.range == (None, expected)


def test_service_vs_job():
    m = full(jobs=[job("J1", 4, deadline=5, penalty=2000)], needs=[need("J1", BAGR)],
             service=[(f"{BAGR}-01", 3, 1)])
    r = result_without_conflicts(m, {"J1": 1}, end=10)
    assert r.jobs.value == 0 and r.conflicts[0]["services"] == [f"servis / údržba {BAGR}-01 od dne 3"]
    [s] = r.scenarios                                                   # jediná vyčíslitelná varianta: posun J1
    late = 2 * 2000
    assert s["variants"] == {"K1": "posun J1"}
    assert s["bounds"][0] == r.result.value + 100000 - 40000 - 4 * 300 - late


def test_partial_overlap_and_multi_machine_job():
    m = full(bookings=[booking("P1", BAGR, 1, 5), booking("P2", BAGR, 5, 8), booking("P3", DUMPER, 2, 3)],
             jobs=[job("J1", 3, deadline=20)], needs=[need("J1", BAGR), need("J1", DUMPER, pos=2)])
    r = result_without_conflicts(m, {"J1": 10}, end=20)
    assert {x["ref"] for c in r.conflicts for x in c["commitments"]} == {"P1", "P2"}
    assert r.jobs.value == 100000 - 40000 - 3 * 300 - 3 * 0             # zakázka s dvěma stroji je jistá
    assert r.rentals.value == 2 * (8000 - 400)
    keys = {s["variants"]["K1"] for s in r.scenarios}
    assert {"zkrácení P1", "zkrácení P2", "zrušení P1", "zrušení P2"} <= keys


def test_several_conflicts_are_combined_and_checked_together():
    m = full(bookings=[booking("P1", BAGR, 1, 3), booking("P2", DUMPER, 1, 3)],
             jobs=[job("J1", 2, deadline=30), job("J2", 2, deadline=30)],
             needs=[need("J1", BAGR), need("J2", DUMPER)])
    r = result_without_conflicts(m, {"J1": 3, "J2": 3}, end=20)
    assert len(r.conflicts) == 2
    assert len(r.scenarios) == 3 * 3                                     # 3 varianty × 3 varianty
    assert all(set(s["variants"]) == {"K1", "K2"} for s in r.scenarios)
    both_shift = [s for s in r.scenarios if s["variants"] == {"K1": "posun J1", "K2": "posun J2"}][0]
    assert both_shift["valid"] and both_shift["in_period"]
    his = [s["bounds"][1] for s in r.scenarios if s["valid"] and s["in_period"]]
    assert r.range[1] == max(his) and r.range[0] is None


def test_combined_variants_that_collide_are_not_counted():
    # dvě zakázky na jednom stroji v konfliktu s pronájmy; posun obou na stejný první volný termín koliduje
    m = full(bookings=[booking("P1", BAGR, 1, 2), booking("P2", BAGR, 10, 11)],
             jobs=[job("J1", 2, deadline=30), job("J2", 2, deadline=30)],
             needs=[need("J1", BAGR), need("J2", BAGR)])
    r = result_without_conflicts(m, {"J1": 2, "J2": 11}, end=20)
    assert len(r.conflicts) == 2
    clash = [s for s in r.scenarios if s["variants"] == {"K1": "posun J1", "K2": "posun J2"}][0]
    assert clash["valid"] is False and "nevyřeší" in clash["reason"]


def test_shift_beyond_period_is_not_in_range():
    m = full(bookings=[booking("P1", BAGR, 1, 5)], jobs=[job("J1", 3, deadline=10)], needs=[need("J1", BAGR)])
    r = result_without_conflicts(m, {"J1": 4}, end=7)                    # posun na 6–8 přesáhne den 7
    shift = [s for s in r.scenarios if s["variants"] == {"K1": "posun J1"}][0]
    assert shift["valid"] and not shift["in_period"] and "přesah" in shift["reason"]
    assert r.range[1] != shift["bounds"][1]


# ---------------------------------------------------------------------------
# Chybějící údaje, skutečná nula, přepočet fixních nákladů
# ---------------------------------------------------------------------------

def test_missing_costs_and_overhead_are_never_zero():
    m = plain_model(bookings=[booking("P1", BAGR, 1, 4)])
    r = result_without_conflicts(m)
    assert r.overhead.missing == ("společná režie/měs.",) and r.overhead.status == "neúplný"
    assert any("variabilní náklad pronájmu" in x for x in r.missing) and r.status == fu.INCOMPLETE
    assert r.result.value == 4 * 6500 - FIXED_MONTH                       # horní mez: známé položky
    assert r.range[0] is None                                           # dolní mez neznámá – nevymýšlí se


def test_missing_value_differs_from_real_zero():
    zero = result_without_conflicts(full(overhead="0", bookings=[booking("P1", BAGR, 1, 4)]))
    empty = result_without_conflicts(full(overhead=None, bookings=[booking("P1", BAGR, 1, 4)]))
    assert zero.overhead.value == 0 and zero.overhead.complete and zero.status == fu.COMPLETE
    assert empty.overhead.value == 0 and not empty.overhead.complete and empty.status == fu.INCOMPLETE
    assert zero.range == (zero.result.value, zero.result.value)
    assert empty.range == (None, empty.result.value)


def test_fixed_costs_need_an_explicit_method():
    m = full(bookings=[booking("P1", BAGR, 1, 4)])
    none = fe.period_result(m, 1, 10)
    assert none.months is None and none.fixed_costs.value == 0 and not none.fixed_costs.complete
    assert any("metoda přepočtu" in x for x in none.missing) and none.status == fu.INCOMPLETE
    half = fe.period_result(m, 1, 10, fixed_cost_method="months", months=0.5)
    assert half.fixed_costs.value == FIXED_MONTH * 0.5 and half.overhead.value == 20000 * 0.5
    assert "0.5000 měsíce" in half.fixed_cost_basis
    for bad in [dict(fixed_cost_method="months"), dict(fixed_cost_method="months", months=0),
                dict(fixed_cost_method="months", months=-1), dict(fixed_cost_method="30 dní"),
                dict(fixed_cost_method="calendar_days")]:
        assert fe.period_result(m, 1, 10, **bad).status == fu.NOT_COMPUTABLE, bad


def test_fixed_costs_across_month_boundary_by_calendar_days():
    m = full(bookings=[booking("P1", BAGR, 1, 4)])
    # den 1 = 30. 1. 2026; dny 1–10 = 30. 1.–8. 2. → 2/31 ledna + 8/28 února
    r = fe.period_result(m, 1, 10, fixed_cost_method="calendar_days", plan_start=date(2026, 1, 30))
    months = 2 / 31 + 8 / 28
    assert r.months == pytest.approx(months) and r.fixed_costs.value == pytest.approx(FIXED_MONTH * months)
    assert "2/31 (1/2026) + 8/28 (2/2026)" in r.fixed_cost_basis
    whole = fe.period_result(m, 1, 31, fixed_cost_method="calendar_days", plan_start=date(2026, 3, 1))
    assert whole.months == pytest.approx(1.0) and whole.fixed_costs.value == pytest.approx(FIXED_MONTH)


def test_fixed_costs_are_subtracted_once_and_service_not_again():
    m = full(bookings=[booking("P1", BAGR, 1, 4)], service=[(f"{BAGR}-01", 6, 2)])
    r = result_without_conflicts(m)
    assert r.fixed_costs.value == ec.fleet_fixed_costs(m, 1.0)["celkem"].value == FIXED_MONTH
    assert r.result.value == r.rentals.value + r.jobs.value - FIXED_MONTH - 20000


# ---------------------------------------------------------------------------
# Ukázková data a dvojí započtení
# ---------------------------------------------------------------------------

def test_sample_data():
    r = fe.period_result(BASE, 1, 25, STARTS, fixed_cost_method="months", months=25 / 30)
    assert (r.rentals.value, r.jobs.value) == (194500, 635000)
    assert r.fixed_costs.value == pytest.approx(440000) and r.fixed_costs.complete
    assert r.overhead.missing == ("společná režie/měs.",)
    assert r.result.value == pytest.approx(194500 + 635000 - 440000) and r.status == fu.INCOMPLETE
    assert {x["ref"] for c in r.conflicts for x in c["commitments"]} == {"FIKT-P05", "J005", "J006"}
    assert [t["ref"] for t in r.tentative] == ["FIKT-P06", "FIKT-P07"] and r.uncertain == []
    # posuny J005 / J006 přesahují den 25 → rozpětí za období 1–25 nelze určit
    assert r.range == (None, None) and all(not s["in_period"] for s in r.scenarios)
    # delší období: posuny jsou celé v období → horní mez rozpětí
    r31 = fe.period_result(BASE, 1, 31, STARTS, fixed_cost_method="months", months=25 / 30)
    assert r31.range == (None, pytest.approx(194500 + 635000 - 440000 + 151500 - 20000))
    assert any("NEJDE o přesný čistý zisk" in n for n in r.notes)


def test_every_commitment_counted_once():
    r = fe.period_result(BASE, 1, 25, STARTS, fixed_cost_method="months", months=25 / 30)
    led = r.options.ledger.entries
    certain = set(led.loc[led["counted_as_certain"], "ref"])
    in_conflict = {x["ref"] for c in r.conflicts for x in c["commitments"]}
    tentative = {t["ref"] for t in r.tentative}
    assert certain.isdisjoint(in_conflict) and certain.isdisjoint(tentative) and in_conflict.isdisjoint(tentative)
    assert len(certain | in_conflict | tentative) == len(led)              # každý závazek z dat právě jednou
    # vzorec jistých příspěvků = součet z evidence (bez konfliktních, předběžných)
    f = fe._certain_formula(BASE, r.options.ledger and fe.cp.from_model(BASE, job_starts=STARTS), led)
    assert f.scenarios["ano"].known == r.rentals.value + r.jobs.value


def test_demand_model_is_not_used():
    from test_fleet_economics import dm
    other = dm.DataModel(**{n: getattr(BASE, n) for n in dm.SCHEMA if n != "rental_demand"},
                         rental_demand=BASE.rental_demand.assign(expected_days="30"))
    a = fe.period_result(BASE, 1, 25, STARTS, fixed_cost_method="months", months=1)
    b = fe.period_result(other, 1, 25, STARTS, fixed_cost_method="months", months=1)
    assert a.result == b.result and a.range == b.range


def test_invalid_period():
    for start, end in [(None, 5), (5, 1), ("1", 3)]:
        assert fe.period_result(BASE, start, end, STARTS).status == fu.NOT_COMPUTABLE
