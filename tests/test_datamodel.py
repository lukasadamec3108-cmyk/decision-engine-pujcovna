"""Testy společné datové vrstvy (datamodel.py).

Spuštění: python -m pytest tests   nebo   python tests/test_datamodel.py
"""

import shutil
import sys
import tempfile
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import datamodel as dm  # noqa: E402

LEGACY = ROOT / "tests" / "fixtures" / "legacy"
LEGACY_CAPACITY = {
    "Mini bagr": 1, "Pásový dumper": 1, "Smykový nakladač": 1, "Vibrační deska": 2,
    "Kolový bagr": 1, "Silniční válec": 1, "Řezačka asfaltu": 1, "Nákladní auto": 1,
}


def edited(table: str, change) -> list:
    """Zkopíruje ukázková data, upraví jednu tabulku a vrátí chyby kontroly."""
    with tempfile.TemporaryDirectory() as tmp:
        shutil.copytree(dm.DEFAULT_DIR, tmp, dirs_exist_ok=True)
        path = Path(tmp) / f"{table}.csv"
        df = pd.read_csv(path, dtype=str, keep_default_na=False)
        df = change(df)
        df.to_csv(path, index=False)
        return [str(i) for i in dm.errors(dm.validate(dm.load_model(tmp)))]


def any_contains(messages, text):
    return any(text in m for m in messages)


# --- Ukázková data -----------------------------------------------------------

def test_demo_data_is_valid():
    issues = dm.validate(dm.load_model())
    assert dm.errors(issues) == []
    # Jediné upozornění: rozdělení fixních nákladů zatím chybí (doplní majitel).
    assert all("rozdělení fixních nákladů" in i.message for i in issues)


def test_demo_data_is_marked_fictitious():
    s = dm.load_model().settings.set_index("key")["value"]
    assert "fiktivní" in s["data_origin"]
    assert set(dm.load_model().rental_demand["source"]) == {"fiktivní ukázka"}


def test_unique_ids_and_names():
    m = dm.load_model()
    assert m.machine_types["type_id"].is_unique and m.machine_types["name"].is_unique
    assert m.fleet_units["unit_id"].is_unique and m.jobs["job_id"].is_unique
    assert m.machine_types["type_id"].str.fullmatch(r"[A-Z0-9_]+").all()


# --- Zpětná kompatibilita ----------------------------------------------------

def test_legacy_machines_identical():
    pd.testing.assert_frame_equal(
        dm.legacy_machines(dm.load_model()), pd.read_csv(LEGACY / "machines_demo.csv")
    )


def test_legacy_jobs_identical():
    pd.testing.assert_frame_equal(
        dm.legacy_jobs(dm.load_model()), pd.read_csv(LEGACY / "jobs_demo.csv")
    )


def test_fleet_capacity_identical_including_order():
    cap = dm.fleet_capacity(dm.load_model())
    assert cap == LEGACY_CAPACITY and list(cap) == list(LEGACY_CAPACITY)


# --- Vazby mezi tabulkami ----------------------------------------------------

def test_unknown_machine_type_in_fleet():
    errs = edited("fleet_units", lambda d: pd.concat(
        [d, pd.DataFrame([dict(unit_id="X-01", type_id="JERAB", status="owned")])], ignore_index=True).fillna(""))
    assert any_contains(errs, "neexistující typ stroje: ['JERAB']")


def test_requirement_by_name_instead_of_id_is_rejected():
    def change(d):
        d.loc[0, "resource_id"] = "Mini bagr"
        return d
    assert any_contains(edited("job_requirements", change), "neexistující typ stroje")


def test_requirement_for_unknown_job():
    errs = edited("job_requirements", lambda d: pd.concat(
        [d, pd.DataFrame([dict(job_id="J999", resource_kind="machine", resource_id="MINI_BAGR",
                               quantity="1", position="1")])], ignore_index=True))
    assert any_contains(errs, "neexistující zakázku: ['J999']")


def test_maintenance_for_unknown_unit():
    errs = edited("maintenance", lambda d: pd.DataFrame(
        [dict(unit_id="MINI_BAGR-09", start_day="3", days="2", note="")]))
    assert any_contains(errs, "neexistující kus flotily")


def test_job_without_machine():
    errs = edited("job_requirements", lambda d: d[d["job_id"] != "J008"])
    assert any_contains(errs, "zakázky bez požadovaného stroje: ['J008']")


def test_duplicate_ids():
    errs = edited("jobs", lambda d: pd.concat([d, d.iloc[[0]]], ignore_index=True))
    assert any_contains(errs, "duplicitní klíč")


# --- Validace hodnot ---------------------------------------------------------

def test_cost_double_counting_is_rejected():
    def change(d):
        d.loc[0, "ownership_cost_month"] = "30000"
        return d
    assert any_contains(edited("machine_types", change), "dvojí započtení")


def test_split_costs_are_accepted_and_summed():
    with tempfile.TemporaryDirectory() as tmp:
        shutil.copytree(dm.DEFAULT_DIR, tmp, dirs_exist_ok=True)
        path = Path(tmp) / "machine_types.csv"
        d = pd.read_csv(path, dtype=str, keep_default_na=False)
        d.loc[0, ["ownership_cost_month", "fixed_operating_cost_month", "fixed_cost_month_unsplit"]] = ["30000", "15000", ""]
        d.to_csv(path, index=False)
        m = dm.load_model(tmp)
        assert dm.errors(dm.validate(m)) == []
        assert dm.legacy_machines(m).loc[0, "monthly_fixed_cost"] == 45000


def test_invalid_values():
    def prob(d):
        d.loc[0, "failure_prob"] = "1.5"
        return d
    def price(d):
        d.loc[1, "rental_price_day"] = "-100"
        return d
    def status(d):
        d.loc[0, "status"] = "hotovo"
        return d
    assert any_contains(edited("machine_types", prob), "pravděpodobnost musí být 0–1")
    assert any_contains(edited("machine_types", price), "záporná hodnota")
    assert any_contains(edited("jobs", status), "nepovolená hodnota")


def test_booking_checks():
    def bookings(rows):
        return lambda d: pd.DataFrame(rows)
    base = dict(type_id="MINI_BAGR", units="1", price_day="6500", with_operator="ne", status="confirmed")
    errs = edited("rental_bookings", bookings([dict(base, booking_id="R1", start_day="5", end_day="3")]))
    assert any_contains(errs, "konec před začátkem")
    errs = edited("rental_bookings", bookings([dict(base, booking_id="R1", start_day="1", end_day="3"),
                                               dict(base, booking_id="R2", start_day="3", end_day="4")]))
    assert any_contains(errs, "překračují počet kusů")
    errs = edited("rental_bookings", bookings([dict(base, booking_id="R1", start_day="1", end_day="3"),
                                               dict(base, booking_id="R2", start_day="4", end_day="6")]))
    assert errs == []


def test_missing_file_is_reported():
    with tempfile.TemporaryDirectory() as tmp:
        shutil.copytree(dm.DEFAULT_DIR, tmp, dirs_exist_ok=True)
        (Path(tmp) / "jobs.csv").unlink()
        try:
            dm.load_model(tmp)
        except FileNotFoundError as e:
            assert "jobs.csv" in str(e)
        else:
            raise AssertionError("chybějící soubor nebyl nahlášen")


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK", name)
