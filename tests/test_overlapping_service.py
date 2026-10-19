"""Překrývající se servisní záznamy téhož kusu: kus je v servisu nejvýše jednou za den.

Spuštění: python -m pytest tests   nebo   python tests/test_overlapping_service.py
"""

import sys
from datetime import date
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))
import capacity as cp  # noqa: E402
import capacity_view as cv  # noqa: E402
from capacity import Assignment, Unit  # noqa: E402

B = "B"


def svc(ref, unit, s, e):
    return Assignment("maintenance", ref, B, s, e, 1, "confirmed", unit)


def cap(cal, s, e):
    return cp.daily_capacity(cal, s, e, [B]).set_index("day")


def test_same_unit_counted_once_two_units_separately():
    units = [Unit("B-1", B), Unit("B-2", B)]
    cal = cp.build_calendar(units, [svc("S1", "B-1", 1, 3), svc("S2", "B-1", 2, 4)])
    c = cap(cal, 1, 5)
    assert c["maintenance"].tolist() == [1, 1, 1, 1, 0] and c["free"].tolist() == [1, 1, 1, 1, 2]
    assert cp.evaluate_request(cal, B, 2, 3)["status"] == cp.FEASIBLE
    cal = cp.build_calendar(units, [svc("S1", "B-1", 1, 3), svc("S2", "B-2", 2, 4)])
    assert cap(cal, 1, 5)["maintenance"].tolist() == [1, 2, 2, 1, 0]          # dva různé kusy
    assert cp.evaluate_request(cal, B, 2, 3)["status"] == cp.INFEASIBLE


def test_three_overlapping_records_and_zero_availability():
    cal = cp.build_calendar([Unit("B-1", B)], [svc("S1", "B-1", 1, 5), svc("S2", "B-1", 2, 4), svc("S3", "B-1", 3, 3)])
    c = cap(cal, 1, 6)
    assert c["maintenance"].tolist() == [1, 1, 1, 1, 1, 0]
    assert c["free"].min() == 0 and cp.OVER not in set(c["status"])               # plno, ne překročeno
    assert cp.find_conflicts(cal) == []                                           # nejde o souběh využití
    r = cp.evaluate_request(cal, B, 3, 3)
    assert r["status"] == cp.INFEASIBLE and r["problem_days"][0][1] == 0          # volno 0, ne −2


def test_period_boundaries_and_adjacent_records():
    units = [Unit("B-1", B), Unit("B-2", B)]
    # záznamy přesahují začátek i konec sledovaného období
    cal = cp.build_calendar(units, [svc("S1", "B-1", 1, 10), svc("S2", "B-1", 5, 20)])
    assert cap(cal, 8, 12)["maintenance"].tolist() == [1] * 5
    ov = cp.overlapping_maintenance(cal, 8, 12)
    assert len(ov) == 1 and ov[0]["days"] == [8, 9, 10] and ov[0]["unit_id"] == "B-1"
    assert cp.overlapping_maintenance(cal, 11, 12) == []
    # navazující záznamy (1–3 a 4–6) se nepřekrývají
    cal = cp.build_calendar(units, [svc("S1", "B-1", 1, 3), svc("S2", "B-1", 4, 6)])
    assert cp.overlapping_maintenance(cal) == [] and cap(cal, 1, 6)["maintenance"].tolist() == [1] * 6


def test_across_year_end_with_dates():
    cal = cp.build_calendar([Unit("B-1", B), Unit("B-2", B)],
                            [svc("S1", "B-1", date(2026, 12, 30), date(2027, 1, 2)),
                             svc("S2", "B-1", date(2026, 12, 31), date(2027, 1, 1))])
    c = cp.daily_capacity(cal, date(2026, 12, 30), date(2027, 1, 2), [B])
    assert c["maintenance"].tolist() == [1, 1, 1, 1]
    assert cp.overlapping_maintenance(cal)[0]["days"] == [date(2026, 12, 31), date(2027, 1, 1)]


def test_unit_outside_fleet_and_other_conflicts_unchanged():
    # vrácený leasing: servis kusu mimo flotilu kapacitu neubírá (ani dvakrát)
    units = [Unit("B-1", B, "leased", None, 2), Unit("B-2", B)]
    cal = cp.build_calendar(units, [svc("S1", "B-1", 1, 4), svc("S2", "B-1", 2, 4)])
    assert cap(cal, 1, 4)["maintenance"].tolist() == [1, 1, 0, 0]
    # pronájem konkrétního kusu v servisu zůstává konfliktem se servisem
    cal = cp.build_calendar([Unit("B-1", B), Unit("B-2", B)],
                            [svc("S1", "B-1", 1, 3), svc("S2", "B-1", 2, 3),
                             Assignment("rental", "R1", B, 3, 4, 1, "confirmed", "B-1")])
    kinds = [(c["kind"], c["scope"]) for c in cp.find_conflicts(cal)]
    assert kinds == [("konflikt se servisem", "kus"), ("konflikt se servisem", "kus")]   # S1×R1, S2×R1
    # předběžná rezervace dál kapacitu neblokuje
    cal = cp.build_calendar([Unit("B-1", B)], [svc("S1", "B-1", 1, 2), svc("S2", "B-1", 1, 2),
                                              Assignment("rental", "T", B, 3, 3, 1, "tentative")])
    c = cap(cal, 1, 3)
    assert c.loc[3, "free"] == 1 and c.loc[3, "tentative"] == 1


def test_unit_state_and_utilization_count_service_once():
    cal = cp.build_calendar([Unit("B-1", B), Unit("B-2", B)], [svc("S1", "B-1", 1, 3), svc("S2", "B-1", 2, 3)])
    assert cp.unit_occupancy(cal, 2, 2).set_index("unit_id").loc["B-1", "state"] == "servis / údržba"
    grid = cv.grid(cal, 1, 4, [B])
    t = cv.utilization_table(grid, {B: "B"}, cal.unknown_capacity)
    assert int(t.maintenance[0]) == 3 and int(t.free[0]) == 5 and int(t.conflict_days[0]) == 0


def test_from_model_with_duplicate_service_rows():
    from test_capacity import model_with
    def dup(df):
        extra = pd.DataFrame([dict(unit_id="REZACKA_ASFALTU-01", start_day="11", days="2", note="duplicitní")])
        return pd.concat([df, extra], ignore_index=True)
    m = model_with(maintenance=dup)
    cal = cp.from_model(m)
    c = cp.daily_capacity(cal, 9, 13, ["REZACKA_ASFALTU"]).set_index("day")
    assert c["maintenance"].tolist() == [0, 1, 1, 1, 0]
    assert len(cp.overlapping_maintenance(cal)) == 1


# Celý průchod aplikací: kopie aplikace, do dat se přidá duplicitní servis řezačky (dny 11–12).
APP_SCRIPT = r'''
import json, os, sys, logging, warnings
logging.disable(logging.CRITICAL); warnings.filterwarnings("ignore")
app_dir, password = sys.argv[1], sys.argv[2]
sys.path.insert(0, app_dir); os.chdir(app_dir)
path = os.path.join(app_dir, "data", "model", "maintenance.csv")
open(path, "a", encoding="utf-8").write("REZACKA_ASFALTU-01,11,2,fiktivní – duplicitní záznam\n")
from streamlit.testing.v1 import AppTest
at = AppTest.from_file(os.path.join(app_dir, "app.py"), default_timeout=300)
at.secrets["password"] = password
at.run(); at.text_input[0].set_value(password).run()
t = at.tabs[6]
out = dict(kpi=[m.value for m in t.markdown if "Dny s konfliktem" in m.value],
           conflicts=[d.value["Stroj"].tolist() for d in t.dataframe if "Konflikt" in d.value.columns],
           info=[m.value for m in t.markdown if "Načtená data" in m.value],
           exc=[e.value for e in at.exception])
at.selectbox(key="cap_detail_type").set_value("REZACKA_ASFALTU")
at.number_input(key="cap_detail_day").set_value(11).run()
out["detail"] = [m.value for m in at.tabs[6].markdown if m.value.startswith("**Řezačka")]
print("VYSLEDEK" + json.dumps(out, ensure_ascii=False))
'''


def test_calendar_in_app_counts_duplicate_service_once():
    from test_data_refresh import run_copy
    out = run_copy(script=APP_SCRIPT)
    assert out["exc"] == []
    assert "4 servisní odstávky" in out["info"][0]                      # záznam načtený
    assert ">5<" in out["kpi"][0].replace(" ", "")                       # stále 5 konfliktních dní
    assert out["conflicts"] == [["Kolový bagr", "Silniční válec"]]       # řezačka bez falešného konfliktu
    assert "volné 0 z 1 ks" in out["detail"][0] and "plno" in out["detail"][0]


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK", name)
