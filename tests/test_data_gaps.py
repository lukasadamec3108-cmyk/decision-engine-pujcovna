"""E3.2 – Přehled chybějících údajů (data_gaps.evaluate, data_gaps_view.build, stránka „Chybějící údaje“).

Stránka je jen pro čtení: nic nedoplňuje ani neodhaduje. Testy hlídají rozpoznání chybějících, nulových
a neplatných hodnot, přiřazení k záznamům, souhrnné počty, priority, prázdná a duplicitní data, zobrazení
stránky a to, že se data ani výpočty nemění. Hesla jsou náhodná.

Spuštění: python -m pytest tests/test_data_gaps.py
"""

import ast
import dataclasses
import os
import re
import secrets
import sys
from pathlib import Path

import pandas as pd
import pytest
from streamlit.testing.v1 import AppTest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))
import capacity_view as cv  # noqa: E402
import data_gaps as dg  # noqa: E402
import data_gaps_view as dgv  # noqa: E402
import datamodel as dm  # noqa: E402
import economics as ec  # noqa: E402
import fleet_economics as fe  # noqa: E402
import navigace  # noqa: E402
from app_pages import open_page  # noqa: E402

PASSWORD = secrets.token_urlsafe(12)
TYPES = ["MINI_BAGR", "PASOVY_DUMPER", "SMYKOVY_NAKLADAC", "VIBRACNI_DESKA", "KOLOVY_BAGR", "SILNICNI_VALEC",
         "REZACKA_ASFALTU", "NAKLADNI_AUTO"]


def model():
    return dm.load_model()


def changed(m, table, key_col, key, **values):
    df = getattr(m, table).copy()
    for col, value in values.items():
        df.loc[df[key_col] == key, col] = value
    return dataclasses.replace(m, **{table: df})


def with_setting(m, key, value):
    s = m.settings[m.settings["key"] != key].copy()
    s.loc[len(s) + 100] = {c: "" for c in s.columns} | {"key": key, "value": value}
    return dataclasses.replace(m, settings=s.reset_index(drop=True))


def all_types(m, **values):
    for t in TYPES:
        m = changed(m, "machine_types", "type_id", t, **values)
    return m


def gaps(m, **filt):
    return [g for g in dg.evaluate(m).gaps if all(getattr(g, k) == v for k, v in filt.items())]


def one(m, column, record_id):
    found = [g for g in dg.evaluate(m).gaps if g.column == column and g.record_id == record_id]
    assert len(found) <= 1, found
    return found[0] if found else None


# ---------------------------------------------------------------------------
# 1. Rozpoznání chybějících hodnot (ukázková data)
# ---------------------------------------------------------------------------

def test_demo_data_missing_items():
    r = dg.evaluate(model())
    cols = [(g.column, g.record_id) for g in r.gaps]
    for t in TYPES:
        assert ("variable_cost_rental_day", t) in cols and ("variable_cost_job_day", t) in cols
        assert ("ownership_cost_month", t) in cols                       # rozdělení fixních nákladů (nízká)
    assert ("direct_cost_includes_machines", "direct_cost_includes_machines") in cols
    assert ("overhead_cost_month", "overhead_cost_month") in cols
    assert len(r.gaps) == 26 and all(g.status == dg.MISSING for g in r.gaps)
    # povinné údaje v ukázkových datech nechybějí
    assert not [g for g in r.gaps if g.table in ("jobs", "rental_bookings", "maintenance", "job_requirements")]
    assert r.fictitious


def test_missing_required_values_are_found():
    m = changed(model(), "jobs", "job_id", "J003", revenue="")
    m = changed(m, "rental_bookings", "booking_id", "FIKT-P02", price_day="")
    m = changed(m, "machine_types", "type_id", "VIBRACNI_DESKA", rental_price_day="")
    assert one(m, "revenue", "J003").status == dg.MISSING
    assert one(m, "price_day", "FIKT-P02").status == dg.MISSING
    assert one(m, "rental_price_day", "VIBRACNI_DESKA").status == dg.MISSING
    assert one(m, "revenue", "J003").rule == "H2"


def test_missing_setting_row_and_empty_value():
    g = one(model(), "direct_cost_includes_machines", "direct_cost_includes_machines")
    assert g.status == dg.MISSING and "řádek zatím vůbec není" in g.why            # řádek v datech chybí
    g = one(with_setting(model(), "direct_cost_includes_machines", ""), "direct_cost_includes_machines",
            "direct_cost_includes_machines")
    assert g.status == dg.MISSING and "řádek zatím vůbec není" not in g.why


def test_missing_column_reported_once_for_whole_table():
    m = model()
    m = dataclasses.replace(m, machine_types=m.machine_types.drop(columns=["variable_cost_rental_day"]))
    found = [g for g in dg.evaluate(m).gaps if g.column == "variable_cost_rental_day"]
    assert len(found) == 1 and found[0].record_id == "celá tabulka" and found[0].priority == dg.HIGH


# ---------------------------------------------------------------------------
# 2. Skutečná nula není chybějící údaj
# ---------------------------------------------------------------------------

def test_real_zero_is_not_missing():
    m = with_setting(all_types(model(), variable_cost_rental_day="0", variable_cost_job_day="0"),
                     "overhead_cost_month", "0")
    found = {g.column for g in dg.evaluate(m).gaps}
    assert "variable_cost_rental_day" not in found and "variable_cost_job_day" not in found
    assert "overhead_cost_month" not in found
    assert ec.overhead(m).complete and ec.overhead(m).value == 0                 # shodně s economics.py


def test_setting_yes_no_is_valid_and_job_costs_not_needed_when_included():
    m = with_setting(model(), "direct_cost_includes_machines", "ano")
    found = {g.column for g in dg.evaluate(m).gaps}
    assert "direct_cost_includes_machines" not in found
    assert "variable_cost_job_day" not in found                                  # v přímých nákladech – nepoužije se
    assert "variable_cost_rental_day" in found
    m = with_setting(model(), "direct_cost_includes_machines", "ne")
    found = {g.column for g in dg.evaluate(m).gaps}
    assert "direct_cost_includes_machines" not in found and "variable_cost_job_day" in found


# ---------------------------------------------------------------------------
# 3. Neplatné hodnoty a datové typy
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("table,key_col,key,column,value", [
    ("machine_types", "type_id", "MINI_BAGR", "variable_cost_rental_day", "hodně"),
    ("machine_types", "type_id", "MINI_BAGR", "variable_cost_job_day", "-50"),
    ("machine_types", "type_id", "MINI_BAGR", "failure_prob", "1.5"),
    ("jobs", "job_id", "J004", "duration_days", "2.5"),
    ("jobs", "job_id", "J004", "duration_days", "0"),
    ("jobs", "job_id", "J004", "direct_cost", "abc"),
    ("rental_bookings", "booking_id", "FIKT-P03", "units", "dva"),
])
def test_invalid_values(table, key_col, key, column, value):
    g = one(changed(model(), table, key_col, key, **{column: value}), column, key)
    assert g.status == dg.INVALID and g.value == value and g.rule == "H1" and g.priority == dg.HIGH


def test_invalid_setting_values():
    g = one(with_setting(model(), "direct_cost_includes_machines", "možná"), "direct_cost_includes_machines",
            "direct_cost_includes_machines")
    assert g.status == dg.INVALID and g.value == "možná"
    g = one(with_setting(model(), "overhead_cost_month", "spousta"), "overhead_cost_month", "overhead_cost_month")
    assert g.status == dg.INVALID and g.priority == dg.HIGH


def test_invalid_fixed_cost_and_requirement_quantity():
    g = one(changed(model(), "machine_types", "type_id", "KOLOVY_BAGR", fixed_cost_month_unsplit="x"),
            "fixed_cost_month_unsplit", "KOLOVY_BAGR")
    assert g.status == dg.INVALID
    m = changed(model(), "job_requirements", "job_id", "J010", quantity="0")
    found = [g for g in dg.evaluate(m).gaps if g.table == "job_requirements"]
    assert len(found) == 1 and found[0].record_id == "J010" and found[0].status == dg.INVALID


# ---------------------------------------------------------------------------
# 4. Přiřazení ke stroji / zakázce / pronájmu
# ---------------------------------------------------------------------------

def test_items_are_assigned_to_records():
    m = changed(model(), "jobs", "job_id", "J007", penalty_per_day="")
    m = changed(m, "rental_bookings", "booking_id", "FIKT-P04", start_day="")
    g = one(m, "penalty_per_day", "J007")
    assert (g.record_kind, g.record_name) == ("zakázka", "Demolice malé stavby")
    g = one(m, "start_day", "FIKT-P04")
    assert g.record_kind == "pronájem"
    g = one(m, "variable_cost_rental_day", "SILNICNI_VALEC")
    assert (g.record_kind, g.record_name) == ("stroj", "Silniční válec")
    g = one(m, "direct_cost_includes_machines", "direct_cost_includes_machines")
    assert g.record_kind == "nastavení"


def test_job_without_machines_and_machine_without_demand():
    m = model()
    m = dataclasses.replace(m, job_requirements=m.job_requirements[m.job_requirements["job_id"] != "J001"])
    g = [x for x in dg.evaluate(m).gaps if x.record_id == "J001" and x.table == "job_requirements"]
    assert len(g) == 1 and g[0].label == "Stroje potřebné pro zakázku" and g[0].status == dg.MISSING
    m = model()
    m = dataclasses.replace(m, rental_demand=m.rental_demand[m.rental_demand["type_id"] != "REZACKA_ASFALTU"])
    g = one(m, "expected_days", "REZACKA_ASFALTU")
    assert g.status == dg.MISSING and g.areas == ("investice",)


# ---------------------------------------------------------------------------
# 5. Souhrnné počty odpovídají seznamu
# ---------------------------------------------------------------------------

def test_summary_counts_match_detail():
    m = changed(model(), "jobs", "job_id", "J002", revenue="abc")
    r = dg.evaluate(m)
    s = r.summary
    assert s["total"] == len(r.gaps) == s["missing"] + s["invalid"] == sum(s["by_priority"].values())
    assert s["invalid"] == 1
    assert s["machines"] == sorted({g.record_id for g in r.gaps if g.record_kind == "stroj"})
    assert set(s["areas"]) == {a for g in r.gaps for a in g.areas}
    assert s["not_tracked"] == len(dg.NOT_TRACKED_ITEMS)                         # počítají se zvlášť
    v = dgv.build(r, cv.type_names(m))
    assert len(v["table"]) == s["total"]
    assert sum(len(gr["items"]) for p in v["priorities"] for gr in p["groups"]) == s["total"]
    assert [p["count"] for p in v["priorities"]] == [s["by_priority"][p] for p in dg.PRIORITIES]


def test_demo_summary_values():
    s = dg.evaluate(model()).summary
    assert s["by_priority"] == {dg.HIGH: 1, dg.MEDIUM: 17, dg.LOW: 8}
    assert len(s["machines"]) == 8 and s["jobs_total"] == 12 and len(s["jobs"]) == 12
    assert s["areas"] == ["zakazky", "pronajmy", "flotila", "srovnani", "planovac"]


def test_affected_jobs_follow_economics():
    m = with_setting(all_types(model(), variable_cost_rental_day="100", variable_cost_job_day="200"),
                     "direct_cost_includes_machines", "ne")
    r = dg.evaluate(m)
    assert r.jobs_incomplete == [] and r.summary["jobs"] == []                  # příspěvky jsou úplné
    assert all(ec.job_contribution(m, j, 0).complete for j in m.jobs["job_id"])


# ---------------------------------------------------------------------------
# 6. Pravidla priorit
# ---------------------------------------------------------------------------

def test_priority_rules():
    r = dg.evaluate(changed(model(), "jobs", "job_id", "J005", revenue="x"))
    for g in r.gaps:
        assert g.priority == dg.PRIORITY_RULES[g.rule][0]
        if g.status == dg.INVALID:
            assert g.rule == "H1"
    rules = {g.column: g.rule for g in r.gaps}
    assert rules["direct_cost_includes_machines"] == "H3"
    assert rules["variable_cost_rental_day"] == rules["variable_cost_job_day"] == rules["overhead_cost_month"] == "M1"
    assert rules["ownership_cost_month"] == "L1"
    assert [g.priority for g in r.gaps] == sorted((g.priority for g in r.gaps), key=dg.PRIORITIES.index)


def test_first_to_fill_order():
    v = dgv.build(dg.evaluate(model()), cv.type_names(model()))
    assert [f["text"] for f in v["first"]] == [
        "Zda přímé náklady zakázek už obsahují náklady vlastních strojů – nastavení dat",
        "Variabilní náklad na den pronájmu – 8 strojů",
        "Variabilní náklad na den práce na vlastní zakázce – 8 strojů"]


# ---------------------------------------------------------------------------
# 7. Prázdná data a neúplné záznamy
# ---------------------------------------------------------------------------

def test_empty_tables_do_not_crash():
    m = model()
    empty = {t: getattr(m, t).iloc[0:0] for t in ("machine_types", "jobs", "rental_bookings", "maintenance",
                                                   "job_requirements", "rental_demand", "settings")}
    r = dg.evaluate(dataclasses.replace(m, **empty))
    assert r.summary["total"] == 2                                     # jen dvě nastavení (řádky chybějí)
    assert r.summary["jobs_total"] == 0 and r.summary["machines"] == []
    assert "V datech nejsou žádné stroje." in r.notes and "V datech nejsou žádné vlastní zakázky." in r.notes
    v = dgv.build(r, {})
    assert v["kpis"][2]["value"] == "0" and len(v["table"]) == 2


def test_record_with_many_empty_cells():
    m = changed(model(), "jobs", "job_id", "J008", revenue="", direct_cost="", penalty_per_day="", deadline_day="")
    found = [g for g in dg.evaluate(m).gaps if g.record_id == "J008"]
    assert {g.column for g in found} == {"revenue", "direct_cost", "penalty_per_day", "deadline_day"}


# ---------------------------------------------------------------------------
# 8. Duplicitní záznamy
# ---------------------------------------------------------------------------

def test_duplicate_records_reported_once_and_not_double_counted():
    m = model()
    dup = pd.concat([m.machine_types, m.machine_types[m.machine_types["type_id"] == "MINI_BAGR"]],
                    ignore_index=True)
    r = dg.evaluate(dataclasses.replace(m, machine_types=dup))
    dups = [g for g in r.gaps if g.label == "Duplicitní záznam"]
    assert len(dups) == 1 and dups[0].record_id == "MINI_BAGR" and dups[0].status == dg.INVALID
    assert len([g for g in r.gaps if g.record_id == "MINI_BAGR" and g.column == "variable_cost_rental_day"]) == 1
    assert len(r.gaps) == 27                                                     # 26 + jedna duplicita
    keys = [g.key for g in r.gaps]
    assert len(keys) == len(set(keys))


def test_duplicate_setting_row():
    m = with_setting(model(), "overhead_cost_month", "1000")
    s = pd.concat([m.settings, m.settings[m.settings["key"] == "overhead_cost_month"]], ignore_index=True)
    r = dg.evaluate(dataclasses.replace(m, settings=s))
    assert [g.record_id for g in r.gaps if g.label == "Duplicitní záznam"] == ["overhead_cost_month"]


# ---------------------------------------------------------------------------
# „Zatím není evidováno“ – ověřeno v kódu, ne vymyšleno
# ---------------------------------------------------------------------------

def test_not_tracked_items_match_the_code():
    cols = {c for t in dm.SCHEMA.values() for c in t}
    assert not [c for c in cols if "transport" in c or "doprav" in c]           # doprava nemá pole
    assert "daily_cost" in dm.SCHEMA["operator_roles"]                          # obsluha pole má…
    calc = "".join((ROOT / f).read_text(encoding="utf-8") for f in (
        "economics.py", "fleet_economics.py", "planner.py", "econ_formula.py", "capacity.py", "scheduling.py",
        "portfolio.py", "investice.py", "fleet_usage.py"))
    for unused in ("operator_roles", "daily_cost", "with_operator", "p_win", "working_days_per_week",
                   "planning_horizon_days"):
        assert unused not in calc, unused                                        # …ale výpočty ho nepoužívají
    exists = {t["key"]: t["field_exists"] for t in dg.NOT_TRACKED_ITEMS}
    assert exists == {"doprava": False, "obsluha": True, "poptavka": False, "kusy": False, "nahrada": False,
                      "p_win": True, "kalendar": True}
    r = dg.evaluate(model())
    assert all(t["status"] == dg.NOT_TRACKED for t in r.not_tracked)
    assert not [g for g in r.gaps if g.status == dg.NOT_TRACKED]


# ---------------------------------------------------------------------------
# Jen čtení: data a výpočty se nemění, nic se nedoplňuje
# ---------------------------------------------------------------------------

def test_evaluation_does_not_change_model_or_results():
    m = model()
    before = {t: getattr(m, t).copy() for t in dm.SCHEMA}
    jobs_before = [ec.job_contribution(m, j, 0) for j in m.jobs["job_id"]]
    rows_before = fe.job_vs_rental(m)
    dg.evaluate(m)
    for t, df in before.items():
        pd.testing.assert_frame_equal(getattr(m, t), df)
    assert [ec.job_contribution(m, j, 0) for j in m.jobs["job_id"]] == jobs_before
    assert fe.job_vs_rental(m) == rows_before


def test_modules_are_read_only_and_without_streamlit():
    for name in ("data_gaps.py", "data_gaps_view.py"):
        src = (ROOT / name).read_text(encoding="utf-8")
        tree = ast.parse(src)
        imported = {a.name for n in ast.walk(tree) if isinstance(n, ast.Import) for a in n.names}
        imported |= {n.module for n in ast.walk(tree) if isinstance(n, ast.ImportFrom) and n.module}
        assert "streamlit" not in imported, name
        for forbidden in ("to_csv", "open(", "write", "secrets", "print("):
            assert forbidden not in src, (name, forbidden)


# ---------------------------------------------------------------------------
# 9. Stránka
# ---------------------------------------------------------------------------

def page(key="chybejici_udaje"):
    os.chdir(ROOT)
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300)
    at.secrets["password"] = PASSWORD
    at.run()
    at.text_input[0].set_value(PASSWORD).run()
    open_page(at, key)
    assert not at.exception, [e.value for e in at.exception]
    return at


def plain(at):
    raw = " ".join(m.value for m in at.main.markdown if not m.value.lstrip().startswith("<style"))
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", raw).replace("\xa0", " "))


def test_page_registered_in_system_group():
    assert navigace.groups()["Systém"] == ["nastaveni", "chybejici_udaje"]
    entry = next(p for p in navigace.PAGES if p[0] == "chybejici_udaje")
    assert entry[2] == "Chybějící údaje" and entry[5] == "chybejici-udaje"
    assert navigace.page_file("chybejici_udaje").is_file()


def test_page_renders_summary_priorities_and_explanation():
    at = page()
    text = plain(at)
    for snippet in ("Chybějící údaje", "Ukázková (fiktivní) data", "Souhrn", "Chybějící nebo neplatné údaje 26",
                    "Dotčené stroje 8", "Dotčené zakázky 12 z 12", "Omezené oblasti výpočtů 5 z 7",
                    "Co doplnit jako první", "Vysoká priorita 1 údaj", "Střední priorita 17 údajů",
                    "Nízká priorita 8 údajů", "Zatím není evidováno", "Náklady na dopravu stroje",
                    "Skutečná nula", "Chybějící údaj", "Neplatná hodnota", "Kde doplnit:", "Ovlivňuje:"):
        assert snippet in text, snippet
    assert not at.warning and not at.error
    [exp] = [e for e in at.main.expander if e.label.startswith("Všechny položky")]
    assert exp.label == "Všechny položky v tabulce (26)"
    [table] = at.main.dataframe
    assert len(table.value) == 26 and list(table.value.columns) == [
        "Priorita", "Údaj", "Záznam", "Stav", "Proč je potřebný", "Ovlivňuje", "Kde doplnit"]


def test_page_has_no_inputs_or_forms():
    at = page()
    kinds = ("text_input", "number_input", "selectbox", "multiselect", "slider", "checkbox", "radio", "date_input",
             "button", "toggle", "text_area")
    assert not [w for k in kinds for w in at.main.get(k)]                       # jen pro čtení


def test_settings_page_links_to_missing_data():
    at = page("nastaveni")
    assert "Chybějící údaje – co doplnit a proč" in [p.label for p in at.main.get("page_link")]


def test_priority_section_shows_rules_that_apply():
    v = dgv.build(dg.evaluate(model()), cv.type_names(model()))
    high = next(p for p in v["priorities"] if p["priority"] == dg.HIGH)
    assert high["rules"] == [dg.PRIORITY_RULES["H3"][1]]                    # na ukázkových datech jen H3
    v = dgv.build(dg.evaluate(changed(model(), "jobs", "job_id", "J001", revenue="x")), {})
    high = next(p for p in v["priorities"] if p["priority"] == dg.HIGH)
    assert high["rules"] == [dg.PRIORITY_RULES["H1"][1], dg.PRIORITY_RULES["H3"][1]]
