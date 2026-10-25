"""Testy přehledu využití strojového parku (E2.4). Jen fiktivní testovací data.

Spuštění: python -m pytest tests   nebo   python tests/test_utilization.py
"""

import os
import random
import secrets
import sys
from pathlib import Path

from streamlit.testing.v1 import AppTest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))
import capacity as cp  # noqa: E402
import capacity_view as cv  # noqa: E402
from capacity import Assignment, OpenJob, Unit  # noqa: E402

A = "A"
NAMES = {A: "Stroj A"}


def table(units, items=(), open_jobs=(), start=1, end=10):
    cal = cp.build_calendar(units, items, open_jobs)
    grid = cv.grid(cal, start, end, cal.type_ids())
    t = cv.utilization_table(grid, {**NAMES, **{x: x for x in cal.type_ids() if x != A}}, cal.unknown_capacity)
    return t, cv.fleet_utilization(t, grid), grid


def row(t, type_id=A):
    return t.set_index("type_id").loc[type_id]


def units(n, type_id=A):
    return [Unit(f"{type_id}-{i}", type_id) for i in range(n)]


def rental(s, e, q=1, status="confirmed", t=A):
    return Assignment("rental", f"R{s}{e}{q}{status}", t, s, e, q, status)


def job(s, e, q=1, t=A):
    return Assignment("job", f"J{s}{e}{q}", t, s, e, q)


def service(unit, s, e, t=A):
    return Assignment("maintenance", f"S-{unit}", t, s, e, 1, "confirmed", unit)


# ---------------------------------------------------------------------------
# Kapacita a využití
# ---------------------------------------------------------------------------

def test_capacity_one_unit():
    t, f, _ = table(units(1), [rental(1, 3)])
    r = row(t)
    assert (r.capacity, r.rental, r.free, r.util_min, r.util_max) == (10, 3, 7, 30.0, 30.0)
    assert cv.share_text(r.util_min, r.util_max) == "30,0 %"


def test_capacity_multiple_units_counts_unit_days_not_machines():
    t, f, _ = table(units(3), [job(1, 5, q=2)])
    r = row(t)
    assert r.capacity == 30 and r.units_max == 3                # 3 kusy × 10 dní
    assert r.job == 10 and r.free == 20                          # 2 kusy × 5 dní
    assert round(r.util_min, 4) == round(100 * 10 / 30, 4)


def test_ordered_unit_adds_capacity_only_from_delivery():
    t, _, _ = table([Unit("A-1", A), Unit("A-2", A, "ordered", 6)])
    assert row(t).capacity == 10 + 5 and row(t).units_max == 2


def test_rental_and_job_overlap_is_not_double_counted():
    t, f, _ = table(units(1), [rental(1, 5), job(4, 8)])
    r = row(t)
    # dny 4–5: 1 kus, 2 nároky → konflikt; rozdělení se nevymýšlí
    assert (r.rental, r.job, r.conflict, r.free) == (3, 3, 2, 2)
    assert r.conflict_days == 2 and r.excess_demand == 2
    assert r.productive_min == r.productive_max == 8              # stroj je produktivně obsazený vždy
    assert r.util_max <= 100 and r.util_min == 80.0
    assert f["conflict_days"] == [4, 5]


def test_overbooking_never_exceeds_100_percent():
    t, f, _ = table(units(1), [rental(1, 10, q=3)])
    r = row(t)
    assert r.conflict == 10 and r.excess_demand == 20 and r.util_max == 100.0
    assert f["util_max"] == 100.0


def test_service_is_not_utilization_and_conflict_with_service_gives_range():
    t, _, _ = table(units(2), [service("A-0", 1, 10), job(1, 10)])
    r = row(t)
    assert (r.maintenance, r.job, r.free, r.util_min) == (10, 10, 0, 50.0)   # servis ≠ využití
    t, _, _ = table(units(1), [service("A-0", 1, 2), job(1, 2)], start=1, end=4)
    r = row(t)
    assert r.conflict == 2 and r.maintenance == 0 and r.job == 0
    assert (r.productive_min, r.productive_max) == (0, 2)        # servis, nebo zakázka – nelze určit
    assert cv.share_text(r.util_min, r.util_max) == "0,0–50,0 %"


def test_tentative_bookings_do_not_block_or_count():
    t, f, _ = table(units(1), [rental(1, 5, status="tentative")])
    r = row(t)
    assert (r.rental, r.free, r.util_min, r.tentative) == (0, 10, 0.0, 5)
    assert f["tentative"] == 5 and f["conflict_days"] == []


def test_zero_capacity_and_missing_data_are_not_zero():
    t, f, _ = table([Unit("A-1", A, "ordered")])                # objednaný bez data dodání
    r = row(t)
    assert r.capacity == 0 and r.util_min is None and r.rating == "nelze určit"
    assert cv.share_text(None, None) == "nelze určit"
    assert f["util_min"] is None
    # zakázka bez termínu → zbytek kapacity „nelze určit“, ne volno
    t, f, _ = table(units(1), open_jobs=[OpenJob("J9", A, 1, None, 3)])
    r = row(t)
    assert (r.unknown, r.free, r.util_min, r.util_max) == (10, 0, 0.0, 100.0)
    assert cv.share_text(r.util_min, r.util_max) == "0,0–100,0 %" and r.rating == "nelze určit"
    # neplatný kus → počet kusů nejistý: kapacita „nelze určit“, typ mimo součty flotily
    t, f, _ = table([Unit("A-1", A), Unit("A-2", A, "owned", "x"), Unit("B-1", "B")])
    assert not row(t).capacity_known and row(t).rating == "nelze určit"
    shown = cv.utilization_display(t).set_index("Stroj").loc["Stroj A"]
    assert shown["Kapacita (strojodny)"] == "nelze určit" and shown["Využití"] == "nelze určit"
    assert f["excluded"] == ["Stroj A"] and f["capacity"] == 10


def test_zero_capacity_type_next_to_others_shows_cannot_determine_not_nan():
    # objednaný kus bez data dodání (0 kapacity) a rezervace na něj vedle běžného stroje
    t, f, _ = table([Unit("A-1", A, "ordered"), Unit("B-1", "B")], [rental(1, 2)])
    shown = cv.utilization_display(t).set_index("Stroj")
    assert shown.loc["Stroj A", "Využití"] == "nelze určit" and "nan" not in shown.to_string()
    assert row(t).conflict_days == 2 and row(t).excess_demand == 2      # nárok bez kusu = konflikt
    assert f["capacity"] == 10 and f["conflict_days"] == [1, 2]


def test_audit_uncertain_fleet_overbooking_is_not_reported_as_certain_conflict():
    # Kalendář ukazuje „nelze určit“ (počet kusů není jistý) – přehled nesmí tvrdit jistý konflikt.
    t, f, grid = table([Unit("A-1", A), Unit("A-2", A, "owned", "x")], [rental(1, 2, q=2)], end=3)
    assert set(grid["status"]) == {cp.UNKNOWN}
    r = row(t)
    assert r.conflict == 0 and r.conflict_days == 0 and r.excess_demand == 0
    assert r.unknown == r.capacity == 3 and r.rating == "nelze určit"
    assert (t[cv.UTIL_GROUPS].sum(axis=1) == t["capacity"]).all()
    # přetížení při jistém počtu kusů zůstává konfliktem i u typu s jinými neúplnými údaji
    t, f, grid = table(units(1), [rental(1, 2, q=2), rental(5, None)], end=3)
    assert row(t).conflict_days == 2 and f["conflict_days"] == [1, 2]


def test_groups_sum_to_capacity_and_fleet_is_sum_of_types():
    rnd = random.Random(11)
    for _ in range(200):
        types = ["A", "B", "C"]
        fleet = [Unit(f"{t}-{i}", t, rnd.choice(["owned", "leased", "ordered"]), rnd.choice([None, 4]))
                 for t in types for i in range(rnd.randint(1, 3))]
        items = []
        for k in range(rnd.randint(0, 8)):
            t = rnd.choice(types); s = rnd.randint(1, 12); e = s + rnd.randint(0, 4)
            kind = rnd.choice(["rental", "job", "maintenance"])
            unit = rnd.choice([u.unit_id for u in fleet if u.type_id == t]) if kind == "maintenance" else None
            items.append(Assignment(kind, f"X{k}", t, s, e, 1 if unit else rnd.randint(1, 2),
                                    rnd.choice(["confirmed", "confirmed", "tentative"]), unit))
        cal = cp.build_calendar(fleet, items)
        grid = cv.grid(cal, 1, 15, types)
        t = cv.utilization_table(grid, {x: x for x in types}, cal.unknown_capacity)
        f = cv.fleet_utilization(t, grid)
        assert (t[cv.UTIL_GROUPS].sum(axis=1) == t["capacity"]).all()
        # konfliktní dny přehledu = dny „překročeno“ v kalendáři (stejný kapacitní model)
        over = grid[grid["status"] == cp.OVER]
        assert t.set_index("type_id")["conflict_days"].to_dict() == {
            x: int((over["type_id"] == x).sum()) for x in types}
        assert f["conflict_days"] == sorted(int(d) for d in over.loc[
            over["type_id"].isin(t.loc[t["capacity_known"], "type_id"]), "day"].unique())
        for g in ["capacity"] + cv.UTIL_GROUPS:
            assert f[g] == int(t[g].sum())
        for r in t.itertuples():
            if r.capacity:
                assert 0 <= r.util_min <= r.util_max <= 100
            else:
                assert cv.share_text(r.util_min, r.util_max) == "nelze určit"
            # produktivní obsazení nikdy nepřesáhne kusy v jednotlivých dnech
            assert r.productive_max <= r.capacity


def test_demo_data_numbers():
    from test_capacity_view import CAL, NAMES as DEMO_NAMES
    grid = cv.grid(CAL, 1, 25, list(DEMO_NAMES))
    t = cv.utilization_table(grid, DEMO_NAMES, CAL.unknown_capacity)
    f = cv.fleet_utilization(t, grid)
    assert {k: f[k] for k in ["capacity"] + cv.UTIL_GROUPS} == dict(
        capacity=225, rental=25, job=110, maintenance=8, conflict=5, unknown=0, free=77)
    assert f["conflict_days"] == [11, 15, 16, 17, 18] and f["excess_demand"] == 5 and f["tentative"] == 10
    assert cv.share_text(f["util_min"], f["util_max"]) == "61,8–62,2 %"
    valec = row(t, "SILNICNI_VALEC")
    assert (valec.rental, valec.job, valec.conflict, valec.util_min) == (1, 10, 4, 60.0)  # ne 5 + 14
    assert t.set_index("type_id")["rating"].to_dict()["NAKLADNI_AUTO"] == "vysoké"
    assert t.set_index("type_id")["rating"].to_dict()["REZACKA_ASFALTU"] == "nízké"
    fig, height = cv.utilization_figure(t, f)
    assert fig.data[0].y[0].startswith("Celá flotila") and height > 300
    assert {tr.name for tr in fig.data} == {"pronájem", "vlastní zakázky", "servis",
                                            "konflikt – rozdělení nelze určit", "volná kapacita"}


# ---------------------------------------------------------------------------
# Aplikace (AppTest)
# ---------------------------------------------------------------------------

PASSWORD = secrets.token_urlsafe(12)


def overview_app():
    os.chdir(ROOT)
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300)
    at.secrets["password"] = PASSWORD
    at.run()
    at.text_input[0].set_value(PASSWORD).run()
    at.switch_page("stranky/stranka_kapacita.py").run()             # dříve záložka 7
    assert at.radio(key="cap_view").value == "Kalendář dostupnosti"   # výchozí zůstává kalendář
    at.radio(key="cap_view").set_value("Přehled využití").run()
    assert not at.exception, [e.value for e in at.exception]
    return at


def capacity_line(at):
    return [m.value for m in at.main.markdown if "Kapacita flotily" in m.value][0]


def test_overview_in_app_follows_period_and_type_filters():
    at = overview_app()
    assert "225 strojodnů" in capacity_line(at)
    kpis = " ".join(m.value for m in at.main.markdown if "de-kpi" in m.value)
    assert "61,8–62,2 %" in kpis and "77 strojodnů" in kpis and "Ukázková data" in " ".join(
        m.value for m in at.main.markdown)
    at.slider(key="cap_period").set_value((1, 10)).run()
    assert "90 strojodnů" in capacity_line(at)                     # 9 kusů × 10 dní
    at.multiselect(key="cap_types").set_value(["SILNICNI_VALEC"]).run()
    assert "10 strojodnů" in capacity_line(at)
    shown = [d.value for d in at.main.dataframe][0]
    assert shown["Stroj"].tolist() == ["Silniční válec"]
    assert not at.exception


def test_audit_occupancy_vs_productive_use_and_orientation_thresholds_are_explicit():
    at = overview_app()
    texts = " ".join(m.value for m in at.main.markdown)
    assert "servis se nepočítá" in texts and "orientační" in texts
    expl = " ".join(m.value for m in at.main.expander[0].markdown) if at.main.expander else ""
    assert "Hodnocení" in expl and "orientační" in expl and "obsazenost" in expl.lower()
    at.radio(key="cap_view").set_value("Kalendář dostupnosti").run()
    kal = [m.value for m in at.main.markdown if "Nejvytíženější stroj" in m.value][0]
    assert "vč. servisu" in kal                                      # obsazenost ≠ využití


def test_overview_button_opens_type_in_calendar():
    at = overview_app()
    at.selectbox(key="use_detail_type").set_value("KOLOVY_BAGR").run()
    at.button(key="use_to_calendar").click().run()
    assert at.radio(key="cap_view").value == "Kalendář dostupnosti"
    assert at.multiselect(key="cap_types").value == ["KOLOVY_BAGR"]
    assert at.selectbox(key="cap_detail_type").value == "KOLOVY_BAGR"
    html_text = [m.value for m in at.main.markdown if "cc-table" in m.value][0]
    assert "Kolový bagr" in html_text and "Mini bagr" not in html_text
    assert not at.exception


# Obnovení po změně datového souboru (kopie aplikace v samostatném procesu – viz test_data_refresh)
REFRESH_SCRIPT = r'''
import json, os, sys, logging, warnings
logging.disable(logging.CRITICAL); warnings.filterwarnings("ignore")
app_dir, password = sys.argv[1], sys.argv[2]
sys.path.insert(0, app_dir); os.chdir(app_dir)
from streamlit.testing.v1 import AppTest
at = AppTest.from_file(os.path.join(app_dir, "app.py"), default_timeout=300)
at.secrets["password"] = password
at.run(); at.text_input[0].set_value(password).run()
at.switch_page("stranky/stranka_kapacita.py").run()
at.radio(key="cap_view").set_value("Přehled využití").run()
def kpis():
    return " ".join(m.value for m in at.main.markdown if "de-kpi" in m.value)
out = {"start": kpis()}
path = os.path.join(app_dir, "data", "model", "rental_bookings.csv")
lines = open(path, encoding="utf-8").read().splitlines()
open(path, "w", encoding="utf-8").write("\n".join(l for l in lines if not l.startswith("FIKT-P02")) + "\n")
at.run(); out["bez_FIKT_P02"] = kpis(); out["exc"] = [e.value for e in at.exception]
print("VYSLEDEK" + json.dumps(out, ensure_ascii=False))
'''


def test_overview_refreshes_after_data_file_change():
    from test_data_refresh import run_copy
    out = run_copy(script=REFRESH_SCRIPT)
    assert out["exc"] == []
    assert "25 strojodnů" in out["start"]                          # pronájmy 25 strojodnů
    # FIKT-P02 (nákladní auto, dny 17–24): v období 1–25 ubere 8 strojodnů pronájmu
    assert "17 strojodnů" in out["bez_FIKT_P02"] and "25 strojodnů" not in out["bez_FIKT_P02"].split("Pronájmy")[1][:80]


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK", name)
