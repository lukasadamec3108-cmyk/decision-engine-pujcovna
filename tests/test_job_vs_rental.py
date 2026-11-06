"""E3.1 – „Zakázka, nebo pronájem?“ (fleet_economics.job_vs_rental, fleet_view.job_rental, stránka Ekonomika flotily).

Jen informace: příspěvek zakázky bez penále × pronájem stejných strojodnů za ceník a hranice využití pronájmu.
Testy hlídají shodu s economics.py, jednotky, nulové / chybějící / neplatné údaje a ochranu proti dvojímu
započtení. Hesla jsou náhodná.

Spuštění: python -m pytest tests/test_job_vs_rental.py
"""

import dataclasses
import os
import re
import secrets
import sys
from pathlib import Path

import pytest
from streamlit.testing.v1 import AppTest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))
import capacity_view as cv  # noqa: E402
import datamodel as dm  # noqa: E402
import economics as ec  # noqa: E402
import fleet_economics as fe  # noqa: E402
import fleet_view as fv  # noqa: E402
from app_pages import open_page  # noqa: E402

PASSWORD = secrets.token_urlsafe(12)
TYPES = ["MINI_BAGR", "PASOVY_DUMPER", "SMYKOVY_NAKLADAC", "VIBRACNI_DESKA", "KOLOVY_BAGR", "SILNICNI_VALEC",
         "REZACKA_ASFALTU", "NAKLADNI_AUTO"]


def model():
    return dm.load_model()


def changed(m, table, key_col, key, **values):
    """Kopie modelu s upravenými buňkami jednoho řádku (data na disku se nemění)."""
    df = getattr(m, table).copy()
    for col, value in values.items():
        df[col] = df[col].astype(object)
        df.loc[df[key_col] == key, col] = value
    return dataclasses.replace(m, **{table: df})


def with_setting(m, key, value):
    s = m.settings.copy()
    s = s[s["key"] != key]
    s.loc[len(s)] = {c: "" for c in s.columns} | {"key": key, "value": value}
    return dataclasses.replace(m, settings=s.reset_index(drop=True))


def all_types(m, **values):
    for t in TYPES:
        m = changed(m, "machine_types", "type_id", t, **values)
    return m


def rows(m):
    return {r["ref"]: r for r in fe.job_vs_rental(m)}


def price(m, t):
    return float(m.machine_types.set_index("type_id").loc[t, "rental_price_day"])


# ---------------------------------------------------------------------------
# Výpočet na ukázkových datech – shoda s economics.py a jednotky
# ---------------------------------------------------------------------------

def test_every_job_once_in_data_order():
    m = model()
    out = fe.job_vs_rental(m)
    assert [r["ref"] for r in out] == m.jobs["job_id"].tolist()


def test_demo_values_match_economics_and_units():
    m = model()
    for r in fe.job_vs_rental(m):
        days = ec.job_machine_days(m, r["ref"])                         # strojodny = délka × kusy
        assert dict(r["machine_days"]) == days and r["total_days"] == sum(days.values())
        job = m.jobs.set_index("job_id").loc[r["ref"]]
        assert r["total_days"] == float(job["duration_days"]) * len(days)
        assert r["state"] == fe.JR_ORIENTATION and not r["machine_costs"]
        # příspěvek zakázky bez penále a bez nákladů strojů = economics.job_contribution (tržba − přímé náklady)
        assert r["contribution"] == ec.job_contribution(m, r["ref"], 0, includes_machines=True).value
        assert r["contribution"] == float(job["revenue"]) - float(job["direct_cost"])
        # tržba z pronájmu = Σ strojodny × ceník (Kč)
        assert r["rental_revenue"] == pytest.approx(sum(d * price(m, t) for t, d in days.items()))
        assert r["rental_contribution"] == r["rental_revenue"]          # orientačně bez variabilních nákladů
        assert r["threshold"] == pytest.approx(r["contribution"] / r["rental_revenue"])


def test_demo_reference_values():
    r = rows(model())
    assert (r["J001"]["contribution"], r["J001"]["rental_revenue"]) == (95000.0, 72500.0)
    assert r["J007"]["threshold"] == pytest.approx(70000 / 105000)       # 67 %
    assert r["J009"]["threshold"] == pytest.approx(60000 / 68000)        # 88 %
    assert r["J005"]["threshold"] == pytest.approx(115000 / 120000)      # 96 %
    assert sorted(k for k, v in r.items() if v["threshold"] < 1) == ["J005", "J007", "J009"]


def test_missing_inputs_are_listed_never_zero():
    r = rows(model())["J001"]
    assert "Mini bagr: variabilní náklad pronájmu/den" in r["missing"]
    assert "Pásový dumper: variabilní náklad na zakázce/den" in r["missing"]
    assert fe.JR_MACHINES_IN_DIRECT in r["missing"]


# ---------------------------------------------------------------------------
# Ochrana proti dvojímu započtení a nezapočítané položky
# ---------------------------------------------------------------------------

def test_penalty_is_not_counted():
    base = rows(model())
    m = model()
    for j in m.jobs["job_id"]:
        m = changed(m, "jobs", "job_id", j, penalty_per_day="999999")
    assert rows(m) == base


def test_fixed_costs_service_and_overhead_are_not_counted():
    base = rows(model())
    m = all_types(model(), fixed_cost_month_unsplit="1", service_cost_month_mean="999999",
                  ownership_cost_month="5", fixed_operating_cost_month="7")
    m = with_setting(m, "overhead_cost_month", "1000000")
    assert rows(m) == base


def test_rental_demand_is_not_used():
    base = rows(model())
    m = model()
    m = dataclasses.replace(m, rental_demand=m.rental_demand.assign(expected_days="0"))
    assert rows(m) == base


def test_full_comparison_with_machine_costs_on_both_sides():
    m = with_setting(all_types(model(), variable_cost_rental_day="500", variable_cost_job_day="800"),
                     "direct_cost_includes_machines", "ne")
    r = rows(m)["J001"]                                                   # 5 dní mini bagr + 5 dní dumper
    assert r["state"] == fe.JR_FULL and r["machine_costs"] and r["missing"] == ()
    assert r["contribution"] == 180000 - 85000 - 10 * 800 == ec.job_contribution(m, "J001", 0).value
    assert r["rental_revenue"] == 72500
    assert r["rental_contribution"] == 72500 - 10 * 500
    assert r["threshold"] == pytest.approx((180000 - 85000 - 8000) / (72500 - 5000))


def test_machine_costs_in_direct_costs_are_not_subtracted_again():
    m = with_setting(all_types(model(), variable_cost_rental_day="500", variable_cost_job_day="800"),
                     "direct_cost_includes_machines", "ano")
    r = rows(m)["J001"]
    assert r["state"] == fe.JR_FULL
    assert r["contribution"] == 180000 - 85000                           # stroje už jsou v přímých nákladech
    assert r["rental_contribution"] == 72500 - 5000


def test_real_zero_costs_are_used_as_zero_not_as_missing():
    m = with_setting(all_types(model(), variable_cost_rental_day="0", variable_cost_job_day="0"),
                     "direct_cost_includes_machines", "ne")
    r = rows(m)["J001"]
    assert r["state"] == fe.JR_FULL and r["missing"] == ()                # nula je údaj, ne „chybí“
    assert r["threshold"] == pytest.approx(95000 / 72500)


def test_partially_entered_machine_costs_are_not_evaluated():
    m = all_types(model(), variable_cost_rental_day="500")               # na zakázce chybí
    r = rows(m)["J001"]
    assert r["state"] == fe.JR_NONE and r["threshold"] is None and r["contribution"] is None
    assert any("jen zčásti" in x for x in r["invalid"])
    assert "Mini bagr: variabilní náklad na zakázce/den" in r["missing"]


# ---------------------------------------------------------------------------
# Nulové, chybějící a neplatné údaje
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("table,key_col,key,values,expect", [
    ("jobs", "job_id", "J001", dict(revenue=""), "tržba zakázky"),
    ("jobs", "job_id", "J001", dict(direct_cost=""), "přímé náklady zakázky"),
    ("jobs", "job_id", "J001", dict(revenue="hodně"), "tržba zakázky"),
    ("jobs", "job_id", "J001", dict(direct_cost="-5"), "přímé náklady"),
    ("jobs", "job_id", "J001", dict(duration_days="abc"), "délka zakázky"),
    ("jobs", "job_id", "J001", dict(duration_days="0"), "délka zakázky"),
    ("machine_types", "type_id", "MINI_BAGR", dict(rental_price_day=""), "cena pronájmu/den"),
    ("machine_types", "type_id", "MINI_BAGR", dict(rental_price_day="-100"), "cena pronájmu/den"),
    ("machine_types", "type_id", "MINI_BAGR", dict(variable_cost_rental_day="x"), "variabilní náklad pronájmu"),
    ("job_requirements", "job_id", "J001", dict(quantity="0"), "počet kusů"),
])
def test_missing_or_invalid_input_is_not_evaluable(table, key_col, key, values, expect):
    out = fe.job_vs_rental(changed(model(), table, key_col, key, **values))
    r = {x["ref"]: x for x in out}["J001"]
    assert len(out) == 12                                                 # řádek nezmizí
    assert r["state"] == fe.JR_NONE and r["threshold"] is None and r["contribution"] is None
    assert any(expect in x for x in r["invalid"] + r["missing"]), (r["invalid"], r["missing"])
    card = fv.job_rental_card(r)
    assert dict(card["rows"])["Hranice využití pronájmu"] == fv.JR_NOT_EVALUABLE
    assert dict(card["rows"])["Příspěvek zakázky (bez penále)"] == fv.JR_NOT_EVALUABLE
    assert "0 Kč" not in " ".join(v for _, v in card["rows"])


def test_job_without_machines_and_unknown_machine():
    m = model()
    req = m.job_requirements
    r = rows(dataclasses.replace(m, job_requirements=req[req["job_id"] != "J001"]))["J001"]
    assert r["state"] == fe.JR_NONE and "zakázka nemá v datech žádné stroje" in r["invalid"]
    r = rows(changed(m, "job_requirements", "job_id", "J010", resource_id="NEZNAMY"))["J010"]
    assert r["state"] == fe.JR_NONE and any("neznámý stroj" in x for x in r["invalid"])


def test_zero_price_means_threshold_cannot_be_determined():
    r = rows(changed(model(), "machine_types", "type_id", "REZACKA_ASFALTU", rental_price_day="0"))["J010"]
    assert r["state"] == fe.JR_NONE and r["rental_revenue"] == 0.0 and r["threshold"] is None
    assert any("nepřinese kladný příspěvek" in x for x in r["invalid"])


def test_zero_revenue_job_gives_zero_threshold_and_honest_text():
    r = rows(changed(model(), "jobs", "job_id", "J001", revenue="0"))["J001"]
    assert r["state"] == fe.JR_ORIENTATION and r["contribution"] == -85000.0 and r["threshold"] == 0.0
    card = fv.job_rental_card(r)
    assert "není kladný" in card["why"][0]
    assert dict(card["rows"])["Příspěvek zakázky (bez penále)"] == "nejvýše −85\xa0000\xa0Kč"


def test_quantity_two_doubles_machine_days_and_rental_revenue():
    r = rows(changed(model(), "job_requirements", "job_id", "J010", quantity="2"))["J010"]
    assert r["total_days"] == 4 and r["rental_revenue"] == 4 * price(model(), "REZACKA_ASFALTU")


# ---------------------------------------------------------------------------
# Zobrazovací vrstva – texty bez doporučení, stavy, souhrn chybějících údajů
# ---------------------------------------------------------------------------

def view():
    m = model()
    return fv.job_rental(fe.job_vs_rental(m), cv.type_names(m))


def test_cards_show_orientational_values_and_units():
    cards = {c["ref"]: c for c in view()["cards"]}
    c = cards["J007"]
    assert c["title"] == "J007 – Demolice malé stavby" and c["tag"] == fe.JR_ORIENTATION
    assert c["sub"] == "potvrzená zakázka · 10 strojodnů: Kolový bagr 5, Nákladní auto 5"
    assert dict(c["rows"]) == {"Příspěvek zakázky (bez penále)": "nejvýše 70\xa0000\xa0Kč",
                               "Tržba z pronájmu stejných strojodnů": "105\xa0000\xa0Kč",
                               "Hranice využití pronájmu": "orientačně 67\xa0%"}
    assert "alespoň na 67\xa0% z 10 strojodnů" in c["why"][0]
    assert "Ani kdyby se stroje pronajaly na všech 10 strojodnů" in cards["J001"]["why"][0]


def test_block_texts_are_information_not_recommendation():
    v = view()
    text = " ".join([v["title"], v["subtitle"], v["intro"], v["disclaimer"], *v["notes"], *v["method"]]
                    + [x for c in v["cards"] for x in [c["title"], c["sub"], c["tag"], *c["why"]]
                       + [val for _, val in c["rows"]]])
    for word in ("vyplatí se", "zvolte pronájem", "zvolte zakázku", "nejlepší"):
        assert word not in text.lower(), word
    assert "zisk" not in text.replace("nejde o čistý zisk", "")
    assert "orientační modelové srovnání na ukázkových datech" in v["subtitle"]
    assert "nejde o čistý zisk ani o rozhodnutí" in v["disclaimer"]
    assert "doporuč" not in text.lower()                                  # stejná ochrana jako zbytek stránky


def test_missing_summary_is_grouped():
    v = view()
    assert v["missing"][0].startswith("variabilní náklad pronájmu/den – Mini bagr, ")
    assert v["missing"][1].startswith("variabilní náklad na zakázce/den – ")
    assert v["missing"][2] == fv.MACHINES_SETTING_MISSING
    assert v["counts"] == {fe.JR_FULL: 0, fe.JR_ORIENTATION: 12, fe.JR_NONE: 0}


# ---------------------------------------------------------------------------
# Stránka Ekonomika flotily
# ---------------------------------------------------------------------------

def page():
    os.chdir(ROOT)
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300)
    at.secrets["password"] = PASSWORD
    at.run()
    at.text_input[0].set_value(PASSWORD).run()
    open_page(at, "ekonomika")
    assert not at.exception, [e.value for e in at.exception]
    return at


def plain(at):
    raw = " ".join(m.value for m in at.main.markdown if not m.value.lstrip().startswith("<style"))
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", raw).replace("\xa0", " "))


def test_page_shows_block_with_all_jobs():
    at = page()
    text = plain(at)
    assert "Zakázka, nebo pronájem?" in text and "nejde o čistý zisk ani o rozhodnutí" in text
    text = text[text.index("Zakázka, nebo pronájem?"):]
    assert text.count(" orientační ") >= 12
    order = [text.index(f"J0{n:02d} – ") for n in range(1, 13)]
    assert order == sorted(order)                                         # pořadí podle dat, žádné řazení
    assert "Hranice využití pronájmu orientačně 67 %" in text
    raw = " ".join(m.value for m in at.main.markdown)               # V5: detail HTML v rozbalovací sekci
    assert '<details class="de-more"><summary>Jak se srovnání zakázky a pronájmu počítá' in raw
    assert "Zakázka, nebo pronájem?" in [e.label for e in at.main.expander]
    assert not at.warning and not at.error


def test_block_is_shown_even_when_period_result_cannot_be_computed():
    at = page()
    labels = [r.options for r in at.radio if r.key == "fe_method"][0]
    at.radio(key="fe_method").set_value(labels[2]).run()                 # podle kalendáře bez data dne 1
    assert not at.exception
    assert "Zakázka, nebo pronájem?" in plain(at) and "orientačně 67 %" in plain(at)
