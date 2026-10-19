"""Testy kapacitního kalendáře (capacity_view.py) a fiktivních ukázkových dat E2.3B.

Spuštění: python -m pytest tests   nebo   python tests/test_capacity_view.py
"""

import re
import shutil
import sys
import tempfile
from datetime import date
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import capacity as cp  # noqa: E402
import capacity_view as cv  # noqa: E402
import datamodel as dm  # noqa: E402
import scheduling as sch  # noqa: E402
from capacity import Assignment, OpenJob, Unit  # noqa: E402

BAGR = "MINI_BAGR"


def tab6_plan(model):
    """Stejný plán jako záložka 6 (app.plan_jobs)."""
    jobs = dm.legacy_jobs(model)
    plan = sch.evaluate(jobs, sch.optimize_schedule(jobs, dm.fleet_capacity(model))["starts"])
    return jobs, plan


MODEL = dm.load_model()
JOBS, PLAN = tab6_plan(MODEL)
CAL = cv.build(MODEL, cv.job_starts_from_plan(JOBS, PLAN))
NAMES = cv.type_names(MODEL)
INFO = cv.job_info(MODEL)


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


def cells(html_text):
    return re.findall(r'<td class="cc(?: [^"]*)?" .*?</td>', html_text)


# ---------------------------------------------------------------------------
# Fiktivní ukázková data
# ---------------------------------------------------------------------------

def test_demo_data_is_valid_and_marked_fictitious():
    assert dm.errors(dm.validate(MODEL)) == []
    assert all(b.startswith("FIKT-") for b in MODEL.rental_bookings["booking_id"])
    assert all("fiktivní" in n for n in MODEL.maintenance["note"])
    assert set(MODEL.rental_bookings["status"]) == {"confirmed", "tentative"}


def test_demo_data_does_not_change_inputs_of_other_tabs():
    # Záložky 1–6 čtou jen typy strojů, kusy flotily a zakázky – pronájmy a servis ne.
    bare = model_with(rental_bookings=lambda df: df.iloc[0:0], maintenance=lambda df: df.iloc[0:0])
    pd.testing.assert_frame_equal(dm.legacy_machines(MODEL), dm.legacy_machines(bare))
    pd.testing.assert_frame_equal(dm.legacy_jobs(MODEL), dm.legacy_jobs(bare))
    assert dm.fleet_capacity(MODEL) == dm.fleet_capacity(bare)
    pd.testing.assert_frame_equal(PLAN, tab6_plan(bare)[1])


def test_demo_data_has_exactly_the_two_intended_conflicts():
    conf = cp.find_conflicts(CAL)
    found = sorted((c["id"], c["kind"], tuple(c["days"])) for c in conf)
    assert found == [
        ("KOLOVY_BAGR", "překročení kapacity typu – konflikt se servisem", (11,)),
        ("SILNICNI_VALEC", "překročení kapacity typu – konflikt pronájmu a vlastní zakázky", (15, 16, 17, 18)),
    ]


# ---------------------------------------------------------------------------
# Termíny z plánu záložky 6
# ---------------------------------------------------------------------------

def test_job_starts_come_from_tab6_plan_including_late_jobs():
    starts = cv.job_starts_from_plan(JOBS, PLAN)
    assert len(starts) == int(PLAN["scheduled"].sum()) == 12
    late = PLAN[PLAN["days_late"] > 0]
    assert len(late) >= 1
    for i, r in late.iterrows():
        job_id = JOBS.loc[i, "job_id"]
        a = [x for x in CAL.assignments if x.ref == job_id][0]
        assert (a.start, a.end) == (int(r.start_day), int(r.end_day)) and a.end > int(r.deadline)
    late_ids = {x[0] for x in cv.late_jobs(CAL, INFO)}
    assert late_ids == {JOBS.loc[i, "job_id"] for i in late.index}
    # zpožděná zakázka (Výkop bazénu) drží mini bagr ve dnech po termínu
    detail = cv.day_detail(CAL, BAGR, 18, INFO)
    assert any(detail["Poznámka"].str.startswith("po termínu"))


def test_unscheduled_job_is_shown_as_unknown_not_free():
    plan = PLAN.copy()
    plan.loc[plan.index[0], "scheduled"] = False          # J001 bez termínu v plánu
    cal = cv.build(MODEL, cv.job_starts_from_plan(JOBS, plan))
    assert [j.ref for j in cal.open_jobs] == ["J001", "J001"]
    df = cv.grid(cal, 21, 30, [BAGR])                       # dny bez jiného využití
    assert set(df["status"]) == {cp.UNKNOWN}
    for c in cells(cv.grid_html(df, NAMES, cv.fleet_units(cal))):
        assert "cc unknown" in c and ">?<" in c and "seg free" not in c
    detail = cv.day_detail(cal, BAGR, 25, INFO)
    assert "bez termínu" in detail.iloc[0]["Stav"]
    r = cv.check_request(cal, BAGR, 22, 23, 1, False, INFO)
    assert r["status"] == cp.UNKNOWN and "Nejde o potvrzení volna" in r["headline"]


# ---------------------------------------------------------------------------
# Mřížka
# ---------------------------------------------------------------------------

def test_grid_values_come_from_capacity_module():
    df = cv.grid(CAL, 1, 25, list(NAMES))
    ref = cp.daily_capacity(CAL, 1, 25, list(NAMES))
    pd.testing.assert_frame_equal(df[cp.COLUMNS], ref)
    assert (df["used"] == ref["maintenance"] + ref["rental"] + ref["job"]).all()


def test_grid_html_distinguishes_all_states():
    cal = cp.build_calendar(
        [Unit("A-1", "A"), Unit("A-2", "A"), Unit("B-1", "B"), Unit("C-1", "C")],
        [Assignment("rental", "R1", "A", 1, 1), Assignment("job", "J1", "A", 2, 2),
         Assignment("maintenance", "S1", "A", 3, 3, unit_id="A-1"),
         Assignment("rental", "R2", "B", 1, 1), Assignment("job", "J2", "B", 1, 1),
         Assignment("rental", "T1", "B", 2, 2, status="tentative")],
        open_jobs=[OpenJob("J9", "C", 1, None, 2)])
    df = cv.grid(cal, 1, 3, ["A", "B", "C"])
    html_text = cv.grid_html(df, {"A": "Stroj A", "B": "Stroj B", "C": "Stroj C"}, cv.fleet_units(cal))
    row = {t: cells(part) for t, part in zip("ABC", html_text.split("<tr>")[2:])}
    assert "seg rental" in row["A"][0] and "seg free" in row["A"][0] and ">1<small>/2" in row["A"][0]
    assert "seg job" in row["A"][1]
    assert "seg maintenance" in row["A"][2]
    assert 'class="cc over"' in row["B"][0] and "+1" in row["B"][0] and "seg over" in row["B"][0]
    assert 'class="cc free"' in row["B"][1] and "cc-tent" in row["B"][1]       # předběžná = jen značka
    assert all('class="cc unknown"' in c and "seg free" not in c for c in row["C"])
    for label in ("volný kus", "pronájem", "vlastní zakázka", "servis / údržba", "konflikt", "nelze určit",
                  "předběžná rezervace"):
        assert label in html_text


def test_grid_rows_follow_machine_order_and_names_are_escaped():
    df = cv.grid(CAL, 1, 3, list(NAMES))
    html_text = cv.grid_html(df, NAMES, cv.fleet_units(CAL))
    positions = [html_text.index(f">{n}<small>") for n in NAMES.values()]
    assert positions == sorted(positions)
    evil = cv.grid_html(df[df["type_id"] == BAGR], {BAGR: "<script>x</script>"}, {BAGR: 1})
    assert "<script>x" not in evil and "&lt;script&gt;" in evil


def test_day_labels_with_dates_across_year_end():
    assert cv.day_label(5) == "den 5"
    start = date(2026, 12, 30)
    assert cv.day_label(3, start, short=True) == "Pá 1. 1."
    assert cv.day_label(3, start) == "den 3 · Pá 1. 1. 2027"
    df = cv.grid(CAL, 1, 4, [BAGR])
    head = cv.grid_html(df, NAMES, cv.fleet_units(CAL), start)
    assert "St 30. 12." in head and "Čt 31. 12." in head and "Pá 1. 1." in head and "So 2. 1." in head


# ---------------------------------------------------------------------------
# Souhrn, detail, konflikty, požadavek
# ---------------------------------------------------------------------------

def test_summary_counts():
    s = cv.summary(cv.grid(CAL, 1, 25, list(NAMES)), NAMES)
    assert s["conflict_days"] == [11, 15, 16, 17, 18] and s["unknown_cells"] == 0
    assert s["busiest"] in NAMES.values() and 0 < s["busiest_share"] <= 100
    s2 = cv.summary(cv.grid(CAL, 1, 10, [BAGR]), NAMES)
    assert s2["conflict_days"] == [] and s2["busiest"] == "Mini bagr" and s2["busiest_share"] == 100


def test_conflicts_table_with_names_and_filter():
    t = cv.conflicts_table(CAL, 1, 25, list(NAMES), NAMES, INFO)
    assert t["Stroj"].tolist() == ["Kolový bagr", "Silniční válec"]
    assert "J006 – Příprava parkovací plochy" in t.iloc[1]["Příčina"] and "FIKT-P05" in t.iloc[1]["Příčina"]
    assert t.iloc[1]["Dny"] == "den 15 – den 18"
    assert cv.conflicts_table(CAL, 1, 25, [BAGR], NAMES, INFO).empty
    assert cv.conflicts_table(CAL, 1, 10, list(NAMES), NAMES, INFO).empty


def test_day_detail_and_unit_states():
    d = cv.day_detail(CAL, BAGR, 14, INFO)
    assert d["Označení"].tolist() == ["FIKT-P01"] and d.iloc[0]["Kus"] == "kterýkoli kus typu"
    d = cv.day_detail(CAL, "VIBRACNI_DESKA", 12, INFO)
    assert d["Druh"].tolist() == ["servis / údržba"] and d.iloc[0]["Kus"] == "VIBRACNI_DESKA-02"
    u = cv.unit_states(CAL, "VIBRACNI_DESKA", 12)
    assert u.set_index("Kus").loc["VIBRACNI_DESKA-02", "Stav kusu"] == "servis / údržba"
    assert "bez přiřazení" in u.set_index("Kus").loc["VIBRACNI_DESKA-01", "Stav kusu"]
    d = cv.day_detail(CAL, "PASOVY_DUMPER", 22, INFO)
    assert d.iloc[0]["Stav"].startswith("předběžně")


def test_check_request():
    r = cv.check_request(CAL, BAGR, 13, 16, 1, False, INFO)
    assert r["status"] == cp.INFEASIBLE and r["tone"] == "bad" and len(r["problems"]) == 4
    assert set(r["problems"]["Co stroj drží"]) == {"pronájem FIKT-P01"}
    r = cv.check_request(CAL, BAGR, 21, 25, 1, False, INFO)
    assert r["status"] == cp.FEASIBLE and r["tone"] == "good"
    r = cv.check_request(CAL, "PASOVY_DUMPER", 21, 22, 1, False, INFO)
    assert r["status"] == cp.FEASIBLE and r["warnings"]                       # předběžná jen varuje
    r = cv.check_request(CAL, "PASOVY_DUMPER", 21, 22, 1, True, INFO)
    assert r["status"] == cp.INFEASIBLE                                        # opatrnější pohled
    r = cv.check_request(CAL, "VIBRACNI_DESKA", 1, 2, 3, False, INFO)
    assert r["status"] == cp.INFEASIBLE                                        # 3 kusy, flotila má 2
    r = cv.check_request(CAL, BAGR, 9, 5, 1, False, INFO)
    assert r["status"] == cp.UNKNOWN and r["headline"].startswith("Nelze vyhodnotit")
    r = cv.check_request(CAL, BAGR, 1, 2, 1, False, INFO)
    assert "J001 – Základy rodinného domu" in r["problems"].iloc[0]["Co stroj drží"]


def test_invalid_data_is_listed_and_never_free():
    def bookings(df):
        row = dict(booking_id="FIKT-X", type_id=BAGR, start_day="30", end_day="", units="1",
                   price_day="1", with_operator="", status="confirmed")
        return pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    model = model_with(rental_bookings=bookings)
    cal = cv.build(model, cv.job_starts_from_plan(JOBS, PLAN))
    assert cal.issues
    df = cv.grid(cal, 21, 25, [BAGR])
    assert set(df["status"]) == {cp.UNKNOWN}
    assert all("seg free" not in c for c in cells(cv.grid_html(df, NAMES, cv.fleet_units(cal))))


# ---------------------------------------------------------------------------
# Regresní testy z auditu PR #22
# ---------------------------------------------------------------------------

def test_audit_summary_share_never_exceeds_100_percent():
    # konflikt (2 nároky na 1 kus) nesmí vypadat jako „obsazeno 200 %“ – přetížení se hlásí zvlášť
    s = cv.summary(cv.grid(CAL, 15, 18, ["SILNICNI_VALEC"]), NAMES)
    assert s["busiest_share"] == 100 and s["conflict_days"] == [15, 16, 17, 18]


def test_audit_day_status_text_never_shows_negative_free_units():
    df = cv.grid(CAL, 15, 15, ["SILNICNI_VALEC", BAGR, "VIBRACNI_DESKA"]).set_index("type_id")
    over = cv.day_status_text(df.loc["SILNICNI_VALEC"])
    assert "-" not in over.split("·")[0] and "o 1 ks víc" in over and "konflikt" in over
    assert cv.day_status_text(df.loc[BAGR]).startswith("volné 0 z 1 ks")
    assert cv.day_status_text(df.loc["VIBRACNI_DESKA"]).startswith("volné 2 z 2 ks")
    unknown = df.loc[BAGR].copy()
    unknown["status"] = cp.UNKNOWN
    assert cv.day_status_text(unknown).startswith("počet volných kusů nelze určit")


def test_audit_late_jobs_follow_type_and_period_filter():
    assert {x[0] for x in cv.late_jobs(CAL, INFO)} == {"J006", "J009", "J011"}
    assert cv.late_jobs(CAL, INFO, ["REZACKA_ASFALTU"]) == []                 # řezačku žádná nedrží
    assert [x[0] for x in cv.late_jobs(CAL, INFO, [BAGR])] == ["J011"]
    assert [x[0] for x in cv.late_jobs(CAL, INFO, None, 1, 10)] == []          # zpoždění až po dni 10


def test_audit_days_text_uses_dates_when_plan_start_is_set():
    assert cv.days_text([11, 15, 16]) == "11, 15, 16"
    assert cv.days_text([11, 15], date(2026, 12, 20)) == "St 30. 12., Ne 3. 1."
    assert cv.days_text(list(range(1, 12))).endswith("…")


def test_audit_default_period_is_limited_for_long_data():
    cal = cp.build_calendar(cv.build(MODEL, {}).units, [Assignment("rental", "R", BAGR, 300, 360)])
    end, last = cv.period_bounds(cal)
    assert end == cv.DEFAULT_PERIOD_DAYS and last == 365          # celý rok jde zvolit posuvníkem
    assert cv.period_bounds(CAL)[0] == 25                         # ukázková data: celé období


def test_period_bounds_cover_all_known_use():
    end, last = cv.period_bounds(CAL)
    assert end == 25 and last >= 30


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK", name)
