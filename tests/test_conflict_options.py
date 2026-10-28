"""Testy variant řešení konfliktů (E2.2 – K3.2, fleet_economics.conflict_options). Jen fiktivní data.

Spuštění: python -m pytest tests/test_conflict_options.py
"""

import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))
import datamodel as dm  # noqa: E402
import econ_formula as ef  # noqa: E402
import fleet_economics as fe  # noqa: E402
import fleet_usage as fu  # noqa: E402
from test_fleet_economics import BAGR, BASE, DUMPER, STARTS, booking, job, model, need  # noqa: E402


def variant(case, key):
    hit = [v for v in case["variants"] if v["key"] == key]
    assert len(hit) == 1, (key, [v["key"] for v in case["variants"]])
    return hit[0]


def only_case(opts):
    assert len(opts.cases) == 1, [c["id"] for c in opts.cases]
    return opts.cases[0]


def known(v, scenario="ano"):
    """Známá část vzorce varianty ve scénáři „přímé náklady obsahují stroje“ (jinak jediný scénář)."""
    sc = v["formula"].scenarios
    return (sc.get(scenario) or sc[ef.ONLY]).known


def assert_never_selected(opts):
    for c in opts.cases:
        assert c["selected"] is None and "rozhodne majitel" in c["note"]
        assert any(v["feasibility"] == fe.MANUAL for v in c["variants"])     # ruční rozhodnutí vždy


# ---------------------------------------------------------------------------
# Druhy konfliktů
# ---------------------------------------------------------------------------

def test_rental_vs_job_variants():
    m = model(bookings=[booking("P1", BAGR, 1, 5)], jobs=[job("J1", 3, deadline=10)], needs=[need("J1", BAGR)])
    opts = fe.conflict_options(m, 1, 10, {"J1": 4})
    case = only_case(opts)
    assert case["types"] == [BAGR] and case["days"] == {BAGR: [4, 5]}
    assert {c["ref"] for c in case["commitments"]} == {"P1", "J1"}
    shift = variant(case, "posun J1")
    assert shift["feasibility"] == fe.FEASIBLE and shift["resolves"] == fe.RESOLVED
    assert (shift["changes"]["new_start"], shift["changes"]["new_end"], shift["changes"]["new_days_late"]) == (6, 8, 0)
    # celé dotčené závazky, každý jednou: pronájem 5 dní + zakázka 60 000 Kč
    assert known(shift) == 5 * 6500 + 60000 == shift["result"].value
    cancel, shorten = variant(case, "zrušení P1"), variant(case, "zkrácení P1")
    assert cancel["feasibility"] == shorten["feasibility"] == fe.CONDITIONAL
    assert (shorten["changes"]["new_start"], shorten["changes"]["new_end"]) == (1, 3)
    assert known(cancel) == 60000 and known(shorten) == 3 * 6500 + 60000
    assert any("smluvní náhrada" in x for x in cancel["missing"]) and any("souhlas zákazníka" in x for x in cancel["missing"])
    # podmíněné varianty nikdy nejsou jistě lepší – závisí na neznámé smluvní náhradě
    for comp in case["comparisons"]:
        if "zrušení P1" in (comp["a"], comp["b"]) or "zkrácení P1" in (comp["a"], comp["b"]):
            assert comp["status"] == ef.UNDECIDABLE
    assert_never_selected(opts)


def test_service_vs_job():
    m = model(jobs=[job("J1", 4, deadline=5, penalty=2000)], needs=[need("J1", BAGR)], service=[(f"{BAGR}-01", 3, 1)])
    opts = fe.conflict_options(m, 1, 10, {"J1": 1})
    case = only_case(opts)
    assert case["services"] == [f"servis / údržba {BAGR}-01 od dne 3"]
    shift = variant(case, "posun J1")
    assert shift["changes"]["new_start"] == 4 and shift["changes"]["new_days_late"] == 2
    assert shift["changes"]["contribution_change"] == -2 * 2000                # penále z economics.py
    assert shift["resolves"] == fe.RESOLVED
    service = [v for v in case["variants"] if v["kind"] == "posun servisu"]
    assert len(service) == 1 and service[0]["feasibility"] == fe.NOT_EVALUABLE
    assert any("servis přesunout" in x for x in service[0]["missing"])
    assert not any(v["kind"].startswith("zachovat zakázku") for v in case["variants"])  # žádný pronájem
    assert all(v["feasibility"] == fe.NOT_EVALUABLE for v in case["variants"] if v["kind"].startswith("externí"))


def test_partial_overlap_of_bookings():
    m = model(bookings=[booking("P1", BAGR, 1, 5), booking("P2", BAGR, 5, 8)])
    case = only_case(fe.conflict_options(m, 1, 10))
    assert case["days"] == {BAGR: [5]}
    s1, s2 = variant(case, "zkrácení P1"), variant(case, "zkrácení P2")
    assert (s1["changes"]["new_start"], s1["changes"]["new_end"]) == (1, 4)
    assert (s2["changes"]["new_start"], s2["changes"]["new_end"]) == (6, 8)
    assert s1["changes"]["released_machine_days"] == s2["changes"]["released_machine_days"] == 1
    assert s1["resolves"] == s2["resolves"] == fe.RESOLVED
    assert not any(v["kind"].startswith("zachovat ostatní") for v in case["variants"])   # žádná zakázka
    # konflikt uprostřed pronájmu: zkrácení by pronájem rozdělilo → nevyhodnotitelné
    m2 = model(bookings=[booking("P1", BAGR, 1, 9), booking("P2", BAGR, 4, 5)])
    inner = variant(only_case(fe.conflict_options(m2, 1, 10)), "zkrácení P1")
    assert inner["feasibility"] == fe.NOT_EVALUABLE and any("rozdělit" in x for x in inner["missing"])


def test_multi_machine_job_needs_all_machines_at_once():
    m = model(bookings=[booking("P1", DUMPER, 2, 3), booking("P2", BAGR, 4, 5)],
              jobs=[job("J1", 3, deadline=4)], needs=[need("J1", BAGR), need("J1", DUMPER, pos=2)])
    case = only_case(fe.conflict_options(m, 1, 10, {"J1": 1}))
    assert case["types"] == [DUMPER]
    shift = variant(case, "posun J1")
    # dumper volný od dne 4, bagr až od dne 6 → zakázka se vejde celá až od dne 6
    assert (shift["changes"]["new_start"], shift["changes"]["new_end"]) == (6, 8)
    assert shift["changes"]["new_days_late"] == 4
    assert "pronájem P2" in {b["source"] for b in shift["blockers"]}
    assert shift["resolves"] == fe.RESOLVED


def test_cascade_and_partial_resolution_is_never_resolved():
    # J1 koliduje s P1; nejbližší termíny blokuje zakázka J2 → posun až za ni
    m = model(bookings=[booking("P1", BAGR, 2, 3)], jobs=[job("J1", 3), job("J2", 3)],
              needs=[need("J1", BAGR), need("J2", BAGR)])
    shift = variant(only_case(fe.conflict_options(m, 1, 12, {"J1": 1, "J2": 4})), "posun J1")
    assert shift["changes"]["new_start"] == 7 and "vlastní zakázka J2" in {b["source"] for b in shift["blockers"]}
    # řetěz P1 × J1 (den 3) a J1 × J2 (den 4) je jeden případ
    m2 = model(bookings=[booking("P1", BAGR, 1, 3)], jobs=[job("J1", 2), job("J2", 2)],
               needs=[need("J1", BAGR), need("J2", BAGR)])
    opts = fe.conflict_options(m2, 1, 12, {"J1": 3, "J2": 4})
    case = only_case(opts)
    assert {c["ref"] for c in case["commitments"]} == {"P1", "J1", "J2"}
    j1, j2, p1 = variant(case, "posun J1"), variant(case, "posun J2"), variant(case, "zrušení P1")
    assert j1["resolves"] == fe.RESOLVED and j1["feasibility"] == fe.FEASIBLE
    for v in (j2, p1):
        assert v["resolves"] == fe.PARTLY and v["feasibility"] == fe.PARTIAL
        assert any(x.startswith("zbývá konflikt") for x in v["missing"])
    compared = {k for c in case["comparisons"] for k in (c["a"], c["b"])}
    assert "posun J2" not in compared and "zrušení P1" not in compared           # částečné se nesrovnávají


def test_tentative_bookings_only_warn():
    m = model(bookings=[booking("P1", BAGR, 1, 5), booking("P9", BAGR, 6, 8, status="tentative")],
              jobs=[job("J1", 3)], needs=[need("J1", BAGR)])
    case = only_case(fe.conflict_options(m, 1, 10, {"J1": 4}))
    assert "P9" not in {c["ref"] for c in case["commitments"]}                  # předběžná nic neblokuje
    shift = variant(case, "posun J1")
    assert shift["changes"]["new_start"] == 6
    assert shift["warnings"] and all("P9" in w for w in shift["warnings"])
    # jen předběžná rezervace přes potvrzenou: žádný konflikt
    alone = fe.conflict_options(model(bookings=[booking("P1", BAGR, 1, 5),
                                                booking("P9", BAGR, 3, 4, status="tentative")]), 1, 10)
    assert alone.cases == [] and any("nejsou konflikty" in n for n in alone.notes)


# ---------------------------------------------------------------------------
# Chybějící sazby a neplatná data
# ---------------------------------------------------------------------------

def test_missing_rates_stay_unknown_never_zero():
    m = model(bookings=[booking("P1", BAGR, 1, 5)], jobs=[job("J1", 3)], needs=[need("J1", BAGR)])
    shift = variant(only_case(fe.conflict_options(m, 1, 10, {"J1": 4})), "posun J1")
    assert shift["econ_status"] == fe.ECON_INCOMPLETE
    assert set(shift["formula"].scenarios) == {"ano", "ne"}                  # nastavení přímých nákladů chybí
    terms = {k for f in shift["formula"].scenarios.values() for k, _ in f.terms}
    assert ("variable_cost_rental_day", BAGR) in terms and ("variable_cost_job_day", BAGR) in terms
    assert shift["bounds"][0] is None                                       # dolní mez neznámá, nevymýšlí se
    assert any("variabilní náklad" in t for t in shift["formula_text"])

    def costs(mt):
        mt = mt.copy()
        mt.loc[mt["type_id"] == BAGR, "variable_cost_rental_day"] = "500"
        mt.loc[mt["type_id"] == BAGR, "variable_cost_job_day"] = "0"
        return mt
    m2 = model(bookings=[booking("P1", BAGR, 1, 5)], jobs=[job("J1", 3)], needs=[need("J1", BAGR)],
               types=costs, settings={"direct_cost_includes_machines": "ne"})
    case = only_case(fe.conflict_options(m2, 1, 10, {"J1": 4}))
    shift2 = variant(case, "posun J1")
    assert shift2["econ_status"] == fe.ECON_COMPLETE and shift2["result"].value == 5 * (6500 - 500) + 60000
    assert shift2["bounds"] == (shift2["result"].value, shift2["result"].value)
    assert variant(case, "zrušení P1")["econ_status"] == fe.ECON_INCOMPLETE     # smluvní náhrada stále neznámá


def test_invalid_data_and_period():
    m = model(bookings=[booking("P1", BAGR, 1, 5)], jobs=[job("J1", 3, revenue="abc")], needs=[need("J1", BAGR)])
    case = only_case(fe.conflict_options(m, 1, 10, {"J1": 4}))
    shift = variant(case, "posun J1")
    assert shift["econ_status"] == fe.ECON_NONE and not shift["formula"].valid
    assert all("posun J1" not in (c["a"], c["b"]) for c in case["comparisons"])
    for start, end in [(None, 5), (5, 1), ("1", 3), (1.0, 3)]:
        bad = fe.conflict_options(m, start, end, {"J1": 4})
        assert bad.status == fu.NOT_COMPUTABLE and bad.cases == []
    # zakázka bez termínu v plánu na stejném stroji → posun nelze spolehlivě vyhodnotit
    m3 = model(bookings=[booking("P1", BAGR, 1, 5)], jobs=[job("J1", 3), job("J2", 2)],
               needs=[need("J1", BAGR), need("J2", BAGR)])
    shift3 = variant(only_case(fe.conflict_options(m3, 1, 10, {"J1": 4})), "posun J1")
    assert shift3["feasibility"] == fe.NOT_EVALUABLE and shift3["missing"]


# ---------------------------------------------------------------------------
# Ukázková data
# ---------------------------------------------------------------------------

OPTS = fe.conflict_options(BASE, 1, 25, STARTS)


def case_with(ref):
    [c] = [c for c in OPTS.cases if ref in {x["ref"] for x in c["commitments"]}]
    return c


def test_sample_p05_vs_j006():
    case = case_with("J006")
    assert {c["ref"] for c in case["commitments"]} == {"FIKT-P05", "J006"}
    assert case["days"] == {"SILNICNI_VALEC": [15, 16, 17, 18]}
    shift = variant(case, "posun J006")
    ch = shift["changes"]
    assert (ch["new_start"], ch["new_end"], ch["days_late"], ch["new_days_late"]) == (20, 27, 2, 11)
    assert ch["contribution_change"] == -81000 and ch["penalty_free_starts"] == []
    assert shift["feasibility"] == fe.FEASIBLE and shift["resolves"] == fe.RESOLVED
    assert known(shift) == 47500 + 61000 and shift["bounds"] == (None, 108500)
    assert any("FIKT-P07" in w for w in shift["warnings"])                    # předběžná rezervace smykového nakladače
    cancel, shorten = variant(case, "zrušení FIKT-P05"), variant(case, "zkrácení FIKT-P05")
    assert cancel["feasibility"] == shorten["feasibility"] == fe.CONDITIONAL
    assert known(cancel) == 142000 and known(shorten) == 142000 + 9500
    assert (shorten["changes"]["new_start"], shorten["changes"]["new_end"]) == (19, 19)
    assert all(c["status"] == ef.UNDECIDABLE for c in case["comparisons"])


def test_sample_service_vs_j005():
    case = case_with("J005")
    assert [c["ref"] for c in case["commitments"]] == ["J005"]
    assert case["services"] == ["servis / údržba KOLOVY_BAGR-01 od dne 11"]
    shift = variant(case, "posun J005")
    ch = shift["changes"]
    assert (ch["new_start"], ch["new_end"], ch["new_days_late"]) == (21, 26, 15)
    assert ch["contribution_change"] == -135000 and shift["bounds"] == (None, -20000)
    blockers = {b["source"] for b in shift["blockers"]}
    assert {"pronájem FIKT-P03", "vlastní zakázka J007", "vlastní zakázka J009", "vlastní zakázka J011"} <= blockers
    assert any("FIKT-P06" in w for w in shift["warnings"])
    assert [v["kind"] for v in case["variants"] if v["feasibility"] == fe.NOT_EVALUABLE] == [
        "posun servisu", "externí zapůjčení / náhradní stroj"]


def test_sample_overall():
    assert len(OPTS.cases) == 2 and OPTS.status == fu.INCOMPLETE
    assert_never_selected(OPTS)
    # každá varianta obsahuje všechny dotčené závazky právě jednou (vzorec = součet z economics)
    for case in OPTS.cases:
        for v in case["variants"]:
            if v["formula"] is not None and v["formula"].valid:
                assert abs(known(v) - v["result"].value) < 1e-6, v["key"]
    # evidence K3.1 se nemění
    pd.testing.assert_frame_equal(OPTS.ledger.entries, fe.period_ledger(BASE, 1, 25, job_starts=STARTS).entries)


def test_inputs_are_not_modified():
    before = {n: getattr(BASE, n).copy() for n in dm.SCHEMA}
    starts = dict(STARTS)
    fingerprint = dm.data_fingerprint()
    fe.conflict_options(BASE, 1, 25, starts)
    assert starts == STARTS and dm.data_fingerprint() == fingerprint
    for n, df in before.items():
        pd.testing.assert_frame_equal(getattr(BASE, n), df)
