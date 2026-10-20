"""Testy záložky Rozhodovací plánovač (planner_view.py + app.py, E2.5b). Jen fiktivní data.

Spuštění: python -m pytest tests   nebo   python tests/test_planner_view.py
"""

import hashlib
import os
import pickle
import secrets
import sys
from pathlib import Path

import pandas as pd
from streamlit.testing.v1 import AppTest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))
import capacity as cp  # noqa: E402
import datamodel as dm  # noqa: E402
import planner_view as pv  # noqa: E402
import scheduling as sch  # noqa: E402
from test_planner import BAGR, model  # noqa: E402

DEMO = dm.load_model()
_JOBS = dm.legacy_jobs(DEMO)
_PLAN = sch.evaluate(_JOBS, sch.optimize_schedule(_JOBS, dm.fleet_capacity(DEMO))["starts"])
STARTS = {_JOBS.loc[i, "job_id"]: int(r.start_day) for i, r in _PLAN.iterrows() if r.scheduled}


def files_hash():
    return {f.name: hashlib.sha256(f.read_bytes()).hexdigest() for f in sorted((ROOT / "data" / "model").glob("*.csv"))}


def card(res, ref):
    return next(c for c in res["cards"] if c["ref"] == ref)


# ---------------------------------------------------------------------------
# Vstupy
# ---------------------------------------------------------------------------

def test_refs_are_unique_and_never_collide_with_data():
    assert pv.next_ref(DEMO, []) == "NP-1"
    assert pv.next_ref(DEMO, ["NP-1", "NP-2"]) == "NP-3"
    ids = set(DEMO.jobs["job_id"]) | set(DEMO.rental_bookings["booking_id"])
    assert pv.next_ref(DEMO, []) not in ids


def test_job_template_copies_data_and_keeps_missing_values_empty():
    t = pv.job_template(DEMO, "J006")
    assert t["needs"] == [("SMYKOVY_NAKLADAC", 1), ("SILNICNI_VALEC", 1)]
    assert (t["duration"], t["earliest"], t["deadline"], t["revenue"], t["direct_cost"]) == (8, 1, 16, 320000.0, 160000.0)
    assert "kopie údajů J006" in t["name"]
    tables = {n: getattr(DEMO, n).copy() for n in dm.SCHEMA}
    tables["jobs"].loc[tables["jobs"]["job_id"] == "J006", ["revenue", "penalty_per_day"]] = ""
    t = pv.job_template(dm.DataModel(**tables), "J006")
    assert t["revenue"] is None and t["penalty_per_day"] is None              # nikdy 0


def test_request_builders_and_form_checks():
    r = pv.rental_request("NP-1", BAGR, 2, 3, 5)
    assert r["needs"] == ((BAGR, 2),) and r["price_day"] is None and pv.form_problems(r) == []
    j = pv.job_request("NP-2", "", [(BAGR, 1), (None, 1), (BAGR, 1)], 4, 3, 2, 10, 50000, None, None)
    assert j["needs"] == ((BAGR, 2),) and (j["start"], j["end"]) == (4, 6) and j["direct_cost"] is None
    assert pv.form_problems(pv.rental_request("X", BAGR, 1, 5, 3)) == ["Konec je před začátkem."]
    assert "Vyberte aspoň jeden stroj." in pv.form_problems(pv.job_request("Y", "", [(None, 1)], 1, 1))
    assert pv.unfreeze(pv.freeze([r, j]))[1]["needs"] == ((BAGR, 2),)


# ---------------------------------------------------------------------------
# Štítky a vysvětlení
# ---------------------------------------------------------------------------

def test_lone_request_is_feasible_never_recommended_even_with_complete_economics():
    res = pv.evaluate(model(var_rental=0), {}, [pv.rental_request("NP-1", BAGR, 1, 1, 3, 6000)])
    c = card(res, "NP-1")
    assert c["badge"] == pv.FEASIBLE and "není s čím porovnat" in c["why"]
    assert "příspěvek 18 000 Kč" in c["economic"]                          # ekonomika zvlášť


def test_recommended_and_feasible_in_competition():
    reqs = [pv.rental_request("NP-1", BAGR, 1, 1, 3, 6000), pv.rental_request("NP-2", BAGR, 1, 2, 4, 9000)]
    res = pv.evaluate(model(var_rental=0), {}, reqs)
    assert card(res, "NP-2")["badge"] == pv.RECOMMENDED and "lepší i v nejhorším případě" in card(res, "NP-2")["why"]
    assert card(res, "NP-1")["badge"] == pv.FEASIBLE and "jiná varianta" in card(res, "NP-1")["why"]
    assert res["groups"][0]["choice"] == ["NP-2"]


def test_equal_benefit_is_feasible_and_missing_economics_cannot_decide():
    reqs = [pv.rental_request("NP-1", BAGR, 1, 1, 2, 6000), pv.rental_request("NP-2", BAGR, 1, 2, 3, 6000)]
    res = pv.evaluate(model(var_rental=0), {}, reqs)
    assert {c["badge"] for c in res["cards"]} == {pv.FEASIBLE} and "shodný" in card(res, "NP-1")["why"]
    res = pv.evaluate(model(), {}, reqs)                                     # chybí variabilní náklad
    assert {c["badge"] for c in res["cards"]} == {pv.UNDECIDED}
    econ = card(res, "NP-1")["economic"]
    assert "nejvýše" in econ and "chybí" in econ and "příspěvek 0" not in econ
    assert any("variabilní náklad" in m for m in res["missing"])


def test_not_feasible_and_undetermined_have_reasons():
    m = model(bookings=[("R1", 1, 5, 1, "confirmed")])
    res = pv.evaluate(m, {}, [pv.rental_request("NP-1", BAGR, 1, 3, 4)])
    c = card(res, "NP-1")
    assert c["badge"] == pv.NOT_FEASIBLE and c["blocking"][0]["Co stroj drží"] == "pronájem R1"
    assert "nehodnotí" in c["economic"]
    bad = [(f"{BAGR}-09", BAGR, "owned", "x", "")]
    res = pv.evaluate(model(extra_units=bad), {}, [pv.rental_request("NP-1", BAGR, 1, 3, 4)])
    c = card(res, "NP-1")
    assert c["badge"] == pv.UNDECIDED and "počet kusů" in c["operational"]
    for c in res["cards"]:
        assert c["why"] and c["operational"] and c["economic"]


def test_template_job_does_not_double_count_original():
    t = pv.job_template(DEMO, "J006")
    # kopie ve dnech původní J006 (11–18): blokuje ji sama původní zakázka – započtená jednou
    req = pv.job_request("NP-1", t["name"], t["needs"], 11, t["duration"], None, None, t["revenue"], t["direct_cost"])
    res = pv.evaluate(DEMO, STARTS, [req])
    c = card(res, "NP-1")
    assert c["badge"] == pv.NOT_FEASIBLE and any("J006" in b["Co stroj drží"] for b in c["blocking"])
    cal = cp.from_model(DEMO, job_starts=STARTS)
    assert cp.daily_capacity(cal, 12, 12, ["SMYKOVY_NAKLADAC"]).iloc[0]["job"] == 1
    assert all(c["ref"] != "J006" for c in res["cards"])                     # původní zakázka se nehodnotí
    # ve volných dnech je kopie proveditelná a ekonomika vychází z jejích vlastních údajů
    req = pv.job_request("NP-1", t["name"], t["needs"], 27, t["duration"], None, None, 10000, 5000)
    c = card(pv.evaluate(DEMO, STARTS, [req]), "NP-1")
    assert c["badge"] == pv.FEASIBLE and "nejvýše 5 000 Kč" in c["economic"]


def test_result_is_plain_and_inputs_unchanged():
    before = files_hash()
    snapshot = {n: getattr(DEMO, n).copy() for n in dm.SCHEMA}
    res = pv.evaluate(DEMO, STARTS, [pv.rental_request("NP-1", "PASOVY_DUMPER", 1, 22, 23)])
    pickle.dumps(res)                                                        # jen obyčejné typy
    for n in dm.SCHEMA:
        pd.testing.assert_frame_equal(getattr(DEMO, n), snapshot[n])
    assert files_hash() == before
    assert card(res, "FIKT-P06")["note"].startswith("předběžná rezervace zůstává")


def test_card_html_escapes_text():
    c = dict(ref="NP-1", title="<b>x</b>", kind="rental", own=True, name="<script>", period=(1, 2),
             needs=[(BAGR, 1)], badge=pv.FEASIBLE, why="<i>", operational="a", economic="b",
             blocking=[], alternatives=[], note="")
    out = pv.card_html(c, {BAGR: "<img>"})
    assert "<script>" not in out and "<img>" not in out and "&lt;script&gt;" in out


# ---------------------------------------------------------------------------
# Aplikace (AppTest)
# ---------------------------------------------------------------------------

PASSWORD = secrets.token_urlsafe(12)


def app():
    os.chdir(ROOT)
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300)
    at.secrets["password"] = PASSWORD
    at.run()
    at.text_input[0].set_value(PASSWORD).run()
    assert not at.exception, [e.value for e in at.exception]
    return at


def click(at, label):
    [b for b in at.tabs[7].button if b.label == label][0].click().run()
    assert not at.exception, [e.value for e in at.exception]


def badges(at):
    import re
    return re.findall(r'dp-badge \w+">([^<]+)<', " ".join(m.value for m in at.tabs[7].markdown))


def test_tab_in_app_full_flow_without_changing_data():
    before = files_hash()
    at = app()
    assert [t.label for t in at.tabs][-1] == "🧭 Rozhodovací plánovač" and len(at.tabs) == 8
    text = " ".join(m.value for m in at.tabs[7].markdown)
    assert "Ukázková data" in text and "nikam neukládají" in text
    # 1) samostatný pronájem ve volném termínu
    at.number_input(key="dp_r_from").set_value(21); at.number_input(key="dp_r_to").set_value(24)
    click(at, "Vyhodnotit samostatně")
    assert badges(at) == [pv.FEASIBLE]
    # 2) obsazený termín → neproveditelné s příčinou
    at.number_input(key="dp_r_from").set_value(13); at.number_input(key="dp_r_to").set_value(16)
    click(at, "Vyhodnotit samostatně")
    assert badges(at) == [pv.NOT_FEASIBLE]
    blocking = [d.value for d in at.tabs[7].dataframe if "Co stroj drží" in d.value.columns][0]
    assert set(blocking["Co stroj drží"]) == {"pronájem FIKT-P01"}
    # 3) soupeření s předběžnou rezervací – nelze rozhodnout, rezervace zůstává
    at.selectbox(key="dp_r_type").set_value("PASOVY_DUMPER")
    at.number_input(key="dp_r_from").set_value(22); at.number_input(key="dp_r_to").set_value(23)
    click(at, "Přidat k porovnání")
    assert badges(at) == [pv.NOT_FEASIBLE, pv.UNDECIDED, pv.UNDECIDED]
    assert any("předběžná rezervace z dat – neruší se" in m.value for m in at.tabs[7].markdown)
    # 4) odebrání a vymazání
    at.button(key="dp_rm_NP-1").click().run()
    assert badges(at) == [pv.UNDECIDED, pv.UNDECIDED]
    at.button(key="dp_clear").click().run()
    assert badges(at) == []
    # 5) zakázka podle šablony, chybná data formuláře
    at.radio(key="dp_kind").set_value("Vlastní zakázka").run()
    at.selectbox(key="dp_template").set_value("J006").run()
    assert at.number_input(key="dp_j_revenue").value == 320000.0 and not at.warning
    at.number_input(key="dp_j_start").set_value(27)
    click(at, "Vyhodnotit samostatně")
    assert badges(at) == [pv.FEASIBLE]
    at.selectbox(key="dp_j_type0").set_value("—"); at.selectbox(key="dp_j_type1").set_value("—")
    click(at, "Vyhodnotit samostatně")
    assert any("Vyberte aspoň jeden stroj" in e.value for e in at.error)
    assert files_hash() == before                                            # CSV beze změny


def test_tab_limits_number_of_compared_requests():
    at = app()
    for i in range(pv.MAX_REQUESTS):
        at.number_input(key="dp_r_from").set_value(30 + 3 * i); at.number_input(key="dp_r_to").set_value(31 + 3 * i)
        click(at, "Přidat k porovnání")
    click(at, "Přidat k porovnání")
    assert any("nejvýš" in w.value for w in at.tabs[7].warning)
    assert len(at.session_state["dp_requests"]) == pv.MAX_REQUESTS


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK", name)
