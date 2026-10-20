"""Testy ekonomického modulu (economics.py). Používají jen fiktivní testovací data.

Spuštění: python -m pytest tests   nebo   python tests/test_economics.py
"""

import shutil
import sys
import tempfile
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import datamodel as dm  # noqa: E402
import economics as ec  # noqa: E402


def model_with(**edits):
    """Kopie ukázkových dat s úpravami: edits = {tabulka: funkce(df) -> df}."""
    tmp = tempfile.mkdtemp()
    shutil.copytree(dm.DEFAULT_DIR, tmp, dirs_exist_ok=True)
    for table, change in edits.items():
        path = Path(tmp) / f"{table}.csv"
        df = pd.read_csv(path, dtype=str, keep_default_na=False)
        change(df).to_csv(path, index=False)
    model = dm.load_model(tmp)
    shutil.rmtree(tmp)
    return model


def set_cells(row_filter_col, row_value, **values):
    def change(df):
        mask = df[row_filter_col] == row_value
        for col, v in values.items():
            df.loc[mask, col] = v
        return df
    return change


def set_setting(key, value):
    def change(df):
        if (df["key"] == key).any():
            df.loc[df["key"] == key, "value"] = value
        else:
            df = pd.concat([df, pd.DataFrame([dict(key=key, value=value, unit="", note="test")])],
                           ignore_index=True)
        return df
    return change


# Testovací (fiktivní) sazby pro Mini bagr: variabilní náklad pronájmu 500 Kč/den, na zakázce 800 Kč/den.
MINI_COSTS = set_cells("type_id", "MINI_BAGR", variable_cost_rental_day="500", variable_cost_job_day="800")


# --- 1. Příspěvek z pronájmu ---------------------------------------------------

def test_rental_contribution():
    m = model_with(machine_types=MINI_COSTS)
    r = ec.rental_contribution(m, "MINI_BAGR", 10)          # ceník 6 500 Kč/den
    assert r.value == 10 * 6500 - 10 * 500 and r.complete
    assert ec.rental_contribution(m, "MINI_BAGR", 10, price_day=7000).value == 10 * 7000 - 10 * 500
    assert ec.rental_day_contribution(m, "MINI_BAGR").value == 6000


def test_rental_contribution_does_not_include_fixed_costs():
    m = model_with(machine_types=MINI_COSTS)
    assert ec.rental_contribution(m, "MINI_BAGR", 0).value == 0     # stání = žádný příspěvek ani náklad


# --- 2. Příspěvek ze zakázky ----------------------------------------------------

def test_job_contribution_machines_not_in_direct_cost():
    # J001: tržba 180 000, přímé náklady 85 000, 5 dní, Mini bagr + Pásový dumper, penále 8 000/den
    m = model_with(machine_types=lambda d: set_cells("type_id", "PASOVY_DUMPER", variable_cost_job_day="1000")(
        MINI_COSTS(d)), settings=set_setting("direct_cost_includes_machines", "ne"))
    r = ec.job_contribution(m, "J001", days_late=2)
    assert r.value == 180000 - 85000 - (5 * 800 + 5 * 1000) - 2 * 8000 and r.complete


def test_job_contribution_machines_already_in_direct_cost():
    m = model_with(machine_types=MINI_COSTS, settings=set_setting("direct_cost_includes_machines", "ano"))
    r = ec.job_contribution(m, "J001")
    assert r.value == 180000 - 85000 and r.complete      # stroje se neodečtou podruhé


def test_job_machine_days_respect_quantity():
    m = model_with(job_requirements=set_cells("job_id", "J008", quantity="2"))
    assert ec.job_machine_days(m, "J008") == {"VIBRACNI_DESKA": 6.0}   # 3 dny × 2 kusy


# --- 3. Oddělení fixních a variabilních nákladů ----------------------------------

def test_fixed_costs_by_group_and_total():
    m = dm.load_model()
    f = ec.fleet_fixed_costs(m)
    # Ukázková data: nerozdělené fixní 414 000 (deska 2 ks) + servis 114 000.
    assert f["nerozdělené fixní"].value == 414000
    assert f["provozní fixní"].value == 114000
    assert f["celkem"].value == 528000 and f["celkem"].complete
    assert not f["vlastnické"].complete      # rozdělení zatím nezadané – rozpad není přesný


def test_split_fixed_costs():
    m = model_with(machine_types=set_cells("type_id", "MINI_BAGR", ownership_cost_month="30000",
                                           fixed_operating_cost_month="15000", fixed_cost_month_unsplit=""))
    c = ec.unit_fixed_costs(m, "MINI_BAGR")
    assert c["vlastnické"].value == 30000 and c["vlastnické"].complete
    assert c["provozní fixní"].value == 15000 + 12000          # + servis
    assert ec.unit_fixed_total(m, "MINI_BAGR").value == 57000


def test_variable_costs_only_by_usage():
    m = model_with(machine_types=MINI_COSTS)
    v = ec.variable_costs(m, {"MINI_BAGR": {"rental_days": 10, "job_days": 5}})
    assert v["pronájem"].value == 5000 and v["zakázky"].value == 4000 and v["celkem"].value == 9000
    assert ec.fleet_fixed_costs(m)["celkem"].value == 528000     # variabilní nejsou ve fixních


def test_ordered_units_are_not_charged():
    m = model_with(fleet_units=set_cells("unit_id", "MINI_BAGR-01", status="ordered"))
    assert ec.fleet_fixed_costs(m)["celkem"].value == 528000 - 57000
    assert ec.fleet_fixed_costs(m, statuses=("owned", "leased", "ordered"))["celkem"].value == 528000


# --- 4. Žádné dvojí započtení --------------------------------------------------

def test_double_counting_of_fixed_costs_is_invalid():
    m = model_with(machine_types=set_cells("type_id", "MINI_BAGR", ownership_cost_month="30000"))
    assert ec.unit_fixed_total(m, "MINI_BAGR").status == "nelze spočítat"
    assert ec.fleet_fixed_costs(m)["celkem"].value is None


def test_each_cost_counted_once_in_total_result():
    # Změna variabilního nákladu o 100 Kč/den při 10 dnech pronájmu změní výsledek přesně o 1 000 Kč.
    s = set_setting("overhead_cost_month", "20000")
    a = model_with(machine_types=MINI_COSTS, settings=s)
    b = model_with(machine_types=set_cells("type_id", "MINI_BAGR", variable_cost_rental_day="600",
                                           variable_cost_job_day="800"), settings=s)
    ra = ec.total_result(a, rentals=[("MINI_BAGR", 10)])["výsledek"]
    rb = ec.total_result(b, rentals=[("MINI_BAGR", 10)])["výsledek"]
    assert ra.value - rb.value == 1000
    # Výsledek = příspěvek − fixní náklady flotily − režie (fixní náklady jen jednou).
    assert ra.value == (65000 - 5000) - 528000 - 20000


def test_total_result_identity():
    m = model_with(machine_types=MINI_COSTS, settings=lambda d: set_setting("overhead_cost_month", "20000")(
        set_setting("direct_cost_includes_machines", "ano")(d)))
    t = ec.total_result(m, rentals=[("MINI_BAGR", 10)], jobs=[("J001", 0), ("J002", 3)])
    expect = t["příspěvky z pronájmů"].value + t["příspěvky ze zakázek"].value \
        - t["fixní náklady flotily"].value - t["režie"].value
    assert t["výsledek"].value == expect and t["výsledek"].complete


# --- 5. Chybějící údaje ----------------------------------------------------------

def test_missing_variable_cost_marks_result_incomplete():
    m = dm.load_model()                                      # variabilní náklady nezadané
    r = ec.rental_contribution(m, "MINI_BAGR", 10)
    assert not r.complete and r.status == "neúplný"
    assert any("variabilní náklad pronájmu" in x for x in r.missing)
    assert "Neúplný výpočet" in r.describe()
    assert r.value == 65000                                  # jen známé položky, nikdy „přesně“


def test_unknown_machine_inclusion_marks_job_incomplete():
    r = ec.job_contribution(dm.load_model(), "J001")
    assert not r.complete and any("direct_cost_includes_machines" in x for x in r.missing)


def test_missing_overhead_marks_total_incomplete():
    t = ec.total_result(dm.load_model())
    assert not t["výsledek"].complete and "společná režie/měs." in t["výsledek"].missing


def test_invalid_value_cannot_be_computed():
    m = model_with(machine_types=set_cells("type_id", "MINI_BAGR", variable_cost_rental_day="neznámo"))
    r = ec.rental_contribution(m, "MINI_BAGR", 3)
    assert r.value is None and r.status == "nelze spočítat" and "neplatné" in r.describe()
    neg = model_with(machine_types=set_cells("type_id", "MINI_BAGR", variable_cost_rental_day="-5"))
    assert ec.rental_contribution(neg, "MINI_BAGR", 3).value is None


def test_unknown_codes_raise():
    m = dm.load_model()
    for call in (lambda: ec.rental_contribution(m, "Mini bagr", 1),   # název místo kódu
                 lambda: ec.job_contribution(m, "J999")):
        try:
            call()
        except KeyError:
            pass
        else:
            raise AssertionError("neznámý kód nebyl odmítnut")


# --- 6. Skutečná nula -------------------------------------------------------------

def test_real_zero_is_complete():
    m = model_with(machine_types=set_cells("type_id", "MINI_BAGR", variable_cost_rental_day="0"),
                   settings=set_setting("overhead_cost_month", "0"))
    r = ec.rental_contribution(m, "MINI_BAGR", 10)
    assert r.value == 65000 and r.complete                 # 0 ≠ nezadáno
    assert ec.overhead(m).complete and ec.overhead(m).value == 0


def test_zero_days_need_no_rate():
    r = ec.rental_contribution(dm.load_model(), "MINI_BAGR", 0)
    assert r.value == 0 and r.complete


# --- 7. Záporné výsledky a hraniční scénáře ----------------------------------------

def test_negative_results():
    m = model_with(machine_types=MINI_COSTS, settings=set_setting("direct_cost_includes_machines", "ne"))
    j = ec.job_contribution(m, "J011", days_late=30)        # 125 000 − 60 000 − 3 200 − 150 000
    assert j.value == 125000 - 60000 - 4 * 800 - 30 * 5000 and j.value < 0
    assert not j.complete          # sazba dumperu na zakázce je nezadaná → výsledek není přesný
    t = ec.total_result(m)                                   # nic se nepronajalo ani nerealizovalo
    assert t["výsledek"].value == -528000


def test_invalid_arguments():
    m = dm.load_model()
    for call in (lambda: ec.rental_contribution(m, "MINI_BAGR", -1),
                 lambda: ec.job_contribution(m, "J001", days_late=-1)):
        try:
            call()
        except ValueError:
            pass
        else:
            raise AssertionError("záporný vstup nebyl odmítnut")


def test_opportunity_cost():
    m = model_with(machine_types=MINI_COSTS)
    assert ec.opportunity_cost(m, {"MINI_BAGR": 4}).value == 4 * 6000
    assert ec.opportunity_cost(m, {}).value == 0
    basis = ec.opportunity_cost_basis(m)
    assert set(basis) == set(m.machine_types["type_id"])
    assert basis["MINI_BAGR"]["příspěvek dne pronájmu"].value == 6000
    assert basis["MINI_BAGR"]["očekávaná poptávka (dny/období)"].value == 18


# --- 8. Kompatibilita s datovým modelem E1 a fiktivní data ---------------------------

def test_works_on_demo_model_with_codes():
    m = dm.load_model()
    for t in m.machine_types["type_id"]:
        assert ec.rental_contribution(m, t, 1).value is not None
        assert ec.unit_fixed_total(m, t).complete
    for j in m.jobs["job_id"]:
        assert ec.job_contribution(m, j).value is not None


def test_fictitious_flag():
    assert ec.rental_contribution(dm.load_model(), "MINI_BAGR", 1).fictitious
    real = model_with(settings=set_setting("data_origin", "údaje majitele"))
    assert not ec.rental_contribution(real, "MINI_BAGR", 1).fictitious
    assert "(fiktivní data)" in ec.rental_contribution(dm.load_model(), "MINI_BAGR", 1).describe()


def test_module_does_not_modify_data():
    m = dm.load_model()
    before = {t: getattr(m, t).copy() for t in dm.SCHEMA}
    ec.total_result(m, rentals=[("MINI_BAGR", 5)], jobs=[("J001", 1)])
    ec.opportunity_cost_basis(m)
    for t in dm.SCHEMA:
        pd.testing.assert_frame_equal(before[t], getattr(m, t))


def test_app_uses_economics_only_through_planner_tab():
    # E2.1: výpočty záložek 1–7 ekonomický modul nepoužívají. Od E2.5b ho využívá jen záložka 8
    # (Rozhodovací plánovač) nepřímo přes planner.py – app.py ho přímo neimportuje.
    import re
    src = (ROOT / "app.py").read_text(encoding="utf-8")
    assert not re.search(r"^\s*(import economics|from economics)", src, re.M)
    before_tab8 = src[:src.index("# ZÁLOŽKA 8 – ROZHODOVACÍ PLÁNOVAČ")]
    for line in before_tab8.splitlines():
        if line.startswith(("import planner", "import econ_formula", "pl = importlib", "ef = importlib",
                            "pv = importlib", "CODE_MODULES")):
            continue
        assert "pl." not in line and "pv." not in line, line


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK", name)
