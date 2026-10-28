"""E2.2 / K1 – testy ekonomické konzistence mezi moduly a regrese současných výsledků.

Co hlídají (bez změny jakéhokoli výpočtu):
  1. Shoda současných výpočtů s economics.py – stejné vstupy musí dát stejné částky:
     plán zakázek (scheduling), nákup strojů / Strojový park (portfolio), Investice (investice),
     Rozhodovací plánovač (planner) a převod dat (datamodel.legacy_*).
  2. Ochrana před dvojím započtením – fixní náklady a servis jsou v každém výpočtu jen jednou,
     příspěvky (pronájem, zakázka) fixní náklady neobsahují.
  3. Známé rozdíly (E2.2) – zdokumentované testy „ZNÁMÝ ROZDÍL“: popisují současné chování,
     které se od economics.py záměrně liší. Neopravují ho – změna vyžaduje schválení majitele
     (viz docs/EKONOMIKA.md). Pokud se rozdíl odstraní, test selže a musí se vědomě upravit.
  4. Regrese současných výsledků a exportu DOCX proti uložené referenci
     (tests/fixtures/econ_regression_reference.json, vytvořena z main bfb7723).

Referenci obnovte JEN po schválené změně výpočtů: python tests/test_econ_consistency.py --update
"""

import hashlib
import io
import json
import sys
import zipfile
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))
import datamodel as dm  # noqa: E402
import economics as ec  # noqa: E402
import investice as inv  # noqa: E402
import planner as pl  # noqa: E402
import portfolio as pf  # noqa: E402
import report_docx as rd  # noqa: E402
import scheduling as sch  # noqa: E402

REFERENCE = ROOT / "tests" / "fixtures" / "econ_regression_reference.json"
_CACHE = {}


def model():
    if "m" not in _CACHE:
        _CACHE["m"] = dm.load_model()
    return _CACHE["m"]


def machines():
    return dm.legacy_machines(model())


def type_id_of(name: str) -> str:
    mt = model().machine_types
    return mt.loc[mt["name"] == name, "type_id"].iloc[0]


def plan():
    """Plán zakázek přesně jako app.plan_jobs (optimalizace → vyhodnocení → přínos zakázek)."""
    if "plan" not in _CACHE:
        jobs = dm.legacy_jobs(model())
        cap = dm.fleet_capacity(model())
        result = sch.optimize_schedule(jobs, cap)
        new_plan = sch.evaluate(jobs, result["starts"])
        contrib = sch.job_contributions(jobs, cap, sch.summary(new_plan)["penalty"])
        _CACHE["plan"] = (jobs, result, new_plan, contrib)
    return _CACHE["plan"]


# ---------------------------------------------------------------------------
# 1. Převod dat (datamodel) ↔ economics.py: stejné nákladové položky
# ---------------------------------------------------------------------------

def test_legacy_fixed_and_service_costs_equal_economics_unit_costs():
    """Fixní náklad strojů v Investicích/Nákupu strojů (legacy) = nerozdělený fixní náklad z dat
    a fixní náklad kusu v economics = ten samý fixní náklad + průměrný servis (každá položka jednou)."""
    m = model()
    for row in machines().itertuples():
        t = type_id_of(row.machine_type)
        raw = m.machine_types.set_index("type_id").loc[t]
        assert row.monthly_fixed_cost == float(raw["fixed_cost_month_unsplit"])
        assert row.service_cost_mean == float(raw["service_cost_month_mean"])
        unit = ec.unit_fixed_total(m, t)
        assert unit.complete and unit.value == row.monthly_fixed_cost + row.service_cost_mean


def test_fleet_units_are_the_same_everywhere():
    """Počty kusů: plán zakázek a kapacita (fleet_capacity) = kusy, za které economics platí fixní náklady."""
    m = model()
    cap = dm.fleet_capacity(m)
    for t, name in m.machine_types[["type_id", "name"]].itertuples(index=False):
        assert ec.units_in_fleet(m, t) == cap[name]
    assert sum(cap.values()) == 9 and len(cap) == 8


def test_month_length_is_the_same_in_data_and_purchase_model():
    assert float(ec.setting(model(), "days_per_month")) == pf.DAYS_PER_MONTH == 30


# ---------------------------------------------------------------------------
# 2. Nákup strojů / Strojový park (portfolio.type_month) ↔ economics.py
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("units", [1, 2, 3])
def test_purchase_model_costs_equal_economics_fixed_costs(units):
    """Náklady v nákupním modelu = počet kusů × (fixní + servis) z economics – nic navíc, nic dvakrát."""
    m = model()
    for _, row in machines().iterrows():
        unit = ec.unit_fixed_total(m, type_id_of(row["machine_type"]))
        assert pf.type_month(row, units)["costs"] == pytest.approx(units * unit.value, abs=1e-9)


def test_purchase_model_revenue_equals_economics_rental_revenue():
    """Tržby nákupního modelu = economics.rental_contribution při stejných dnech a ceně.

    economics navíc odečítá variabilní náklad pronájmu – v datech chybí, proto je jeho výsledek
    NEÚPLNÝ a jeho známá částka (horní mez) se rovná tržbám. Nákupní model variabilní náklady
    nezná vůbec (ZNÁMÝ ROZDÍL 2, docs/EKONOMIKA.md)."""
    m = model()
    for _, row in machines().iterrows():
        t = type_id_of(row["machine_type"])
        one = pf.type_month(row, 1)
        contrib = ec.rental_contribution(m, t, one["days"])
        assert contrib.value == pytest.approx(one["revenue"], abs=1e-6)
        assert not contrib.complete
        assert any("variabilní náklad pronájmu" in x for x in contrib.missing)


def test_purchase_model_profit_already_contains_fixed_costs_of_fleet():
    """Ochrana před dvojím započtením: zisk nákupního modelu pro kusy flotily už obsahuje přesně
    fixní náklady flotily z economics. Sčítat „zisky strojů“ a k tomu odečítat fixní náklady
    flotily by je započítalo dvakrát."""
    m = model()
    cap = dm.fleet_capacity(m)
    fleet = ec.fleet_fixed_costs(m)
    model_costs = 0.0
    for _, row in machines().iterrows():
        t = type_id_of(row["machine_type"])
        costs = pf.type_month(row, cap[row["machine_type"]])["costs"]
        assert costs == pytest.approx(fleet["podle typu"][t].value)
        model_costs += costs
    assert model_costs == pytest.approx(fleet["celkem"].value) == pytest.approx(528000.0)


# ---------------------------------------------------------------------------
# 3. Vlastní zakázky (scheduling) ↔ economics.job_contribution
# ---------------------------------------------------------------------------

def test_job_profit_in_plan_equals_economics_contribution():
    """Zisk zakázky v plánu (po penále) = economics.job_contribution, pokud přímé náklady zakázky
    už obsahují vlastní stroje (includes_machines=True). Penále = dny po termínu × sazba."""
    m = model()
    jobs, _, new_plan, _ = plan()
    ids = dict(zip(jobs["job_name"], jobs["job_id"]))
    for r in new_plan[new_plan["scheduled"]].itertuples():
        job_id = ids[r.job_name]
        before = ec.job_contribution(m, job_id, days_late=0, includes_machines=True)
        after = ec.job_contribution(m, job_id, days_late=r.days_late, includes_machines=True)
        assert before.complete and before.value == pytest.approx(r.profit)
        assert after.complete and after.value == pytest.approx(r.profit_after_penalty)
        assert after.parts["penále"] == pytest.approx(-r.penalty)


def test_plan_totals_equal_sum_of_economics_contributions():
    m = model()
    jobs, _, new_plan, _ = plan()
    s = sch.summary(new_plan)
    ids = dict(zip(jobs["job_name"], jobs["job_id"]))
    done = new_plan[new_plan["scheduled"]]
    total = ec.total(ec.job_contribution(m, ids[r.job_name], r.days_late, includes_machines=True)
                     for r in done.itertuples())
    assert total.complete and total.value == pytest.approx(s["profit_after_penalty"]) == pytest.approx(892000.0)
    assert s["revenue"] - s["costs"] == pytest.approx(s["profit"])


def test_job_benefit_uses_same_job_profit():
    """Přínos zakázky pro plán (nejpřínosnější zakázka) vychází ze stejného zisku zakázky."""
    m = model()
    jobs, _, _, contrib = plan()
    ids = dict(zip(jobs["job_name"], jobs["job_id"]))
    for r in contrib.itertuples():
        assert r.profit == pytest.approx(ec.job_contribution(m, ids[r.job_name], includes_machines=True).value)
        assert r.contribution == pytest.approx(r.profit - r.added_penalty)


def test_known_difference_plan_does_not_know_machine_costs():
    """ZNÁMÝ ROZDÍL 3: plán zakázek bere přímé náklady zakázky jako úplné. economics bez nastavení
    direct_cost_includes_machines označí příspěvek zakázky jako NEÚPLNÝ (neví, zda náklady strojů
    v přímých nákladech jsou). Částka 892 000 Kč je proto „zisk zakázek po penále“, ne čistý zisk."""
    m = model()
    assert ec.direct_cost_includes_machines(m) is None                 # v datech nezadáno
    r = ec.job_contribution(m, "J001")
    assert not r.complete and any("direct_cost_includes_machines" in x for x in r.missing)
    # kdyby přímé náklady stroje NEobsahovaly, chybí variabilní náklady strojů (v datech nezadané)
    r = ec.job_contribution(m, "J001", includes_machines=False)
    assert not r.complete and any("variabilní náklad na zakázce" in x for x in r.missing)


# ---------------------------------------------------------------------------
# 4. Investice (investice.py) ↔ economics.py
# ---------------------------------------------------------------------------

def _rows(name_1="Mini bagr", name_2="Pásový dumper"):
    df = machines()
    return df[df["machine_type"] == name_1].iloc[0], df[df["machine_type"] == name_2].iloc[0]


def test_investment_quick_estimate_equals_economics_when_fixed_cost_from_data():
    """Rychlý odhad v Investicích = ceník × dny − fixní náklad − průměrný servis.
    Se stejným fixním nákladem jako v datech dává totéž co economics: tržby − fixní náklad kusu."""
    m = model()
    row_1, row_2 = _rows()
    for row in (row_1, row_2):
        t = type_id_of(row["machine_type"])
        days = 18
        r = inv.evaluate(row, row_2 if row is row_1 else row_1, row["machine_type"], "x",
                         row["monthly_fixed_cost"], 15, 25, days, 200)
        revenue = ec.rental_contribution(m, t, days).value               # známá část = tržby
        assert r["expected_profit_mini_real"] == pytest.approx(revenue - ec.unit_fixed_total(m, t).value)


def test_investment_simulation_subtracts_fixed_cost_once_per_month():
    """Ochrana před dvojím započtením: simulovaný zisk = tržby − fixní náklad − servis, každé jednou."""
    row_1, row_2 = _rows()
    sim_1, _ = inv.simulate(row_1, row_2, 100000, 15, 25, 500)
    assert sim_1["profit"] == pytest.approx(sim_1["revenue"] - 100000 - sim_1["service_cost"])


def test_known_difference_investment_fixed_cost_comes_from_panel():
    """ZNÁMÝ ROZDÍL 1: Investice (záložky 1–4, Dashboard, DOCX) berou fixní náklad z panelu
    „Nastavení investičního scénáře“ – výchozí 100 000 Kč pro oba stroje. Nákup strojů, Strojový park
    a economics berou individuální fixní náklad z dat. Stejný stroj má proto dva různé náklady."""
    row_1, row_2 = _rows()
    panel_default = 100000                                              # výchozí hodnota v app.py
    assert row_1["monthly_fixed_cost"] == 45000 and row_2["monthly_fixed_cost"] == 55000
    assert panel_default not in (row_1["monthly_fixed_cost"], row_2["monthly_fixed_cost"])


def test_known_difference_break_even_ignores_service_but_minimum_price_includes_it():
    """ZNÁMÝ ROZDÍL 4 (popsaný i v aplikaci): bod zvratu = fixní náklad ÷ cena (bez servisu),
    minimální cena = (fixní náklad + servis) ÷ dny. Při bodu zvratu je zisk se servisem záporný."""
    row_1, row_2 = _rows()
    r = inv.evaluate(row_1, row_2, "Mini bagr", "Pásový dumper", 100000, 15, 25, 18, 200)
    assert r["break_even_mini"] * row_1["daily_price"] == pytest.approx(100000)
    assert r["required_price_1"] * 18 == pytest.approx(100000 + row_1["service_cost_mean"])


# ---------------------------------------------------------------------------
# 5. Rozhodovací plánovač (Pronájem, Rozhodovací centrum) používá economics.py
# ---------------------------------------------------------------------------

def test_planner_rental_value_equals_economics_rental_contribution():
    m = model()
    req = pl.Request("NP-1", "rental", pl.NEW, (("PASOVY_DUMPER", 2),), 22, 23, price_day=9000.0)
    v = pl.contribution(m, req)
    exp = ec.rental_contribution(m, "PASOVY_DUMPER", 2, 9000.0).scaled(2)
    assert v.hi == pytest.approx(exp.value) == pytest.approx(36000.0)
    assert v.lo is None                                                 # chybí variabilní náklad → jen horní mez


def test_planner_job_value_equals_economics_job_contribution():
    m = model()
    job = m.jobs.set_index("job_id").loc["J001"]
    end = int(job["deadline_day"]) + 2                                  # 2 dny po termínu
    req = pl.Request("J001", "job", pl.TENTATIVE, (("MINI_BAGR", 1), ("PASOVY_DUMPER", 1)),
                     end - int(job["duration_days"]) + 1, end, deadline=int(job["deadline_day"]), in_data=True)
    v = pl.contribution(m, req)
    exp = ec.job_contribution(m, "J001", days_late=2)
    assert v.hi == pytest.approx(exp.value) == pytest.approx(180000 - 85000 - 2 * 8000)
    assert v.lo is None and not exp.complete


def test_contributions_never_contain_fixed_costs():
    """Ochrana před dvojím započtením: příspěvek pronájmu ani zakázky nezávisí na fixních nákladech
    strojů (ty jsou jen jednou za flotilu)."""
    m = model()
    mt = m.machine_types.copy()
    mt["fixed_cost_month_unsplit"] = "999999"
    tables = {n: getattr(m, n) for n in dm.SCHEMA}
    tables["machine_types"] = mt
    m2 = dm.DataModel(**tables)
    assert ec.rental_contribution(m2, "MINI_BAGR", 5).value == ec.rental_contribution(m, "MINI_BAGR", 5).value
    assert ec.job_contribution(m2, "J001", 1).value == ec.job_contribution(m, "J001", 1).value
    assert ec.fleet_fixed_costs(m2)["celkem"].value != ec.fleet_fixed_costs(m)["celkem"].value


def test_total_result_is_incomplete_on_current_data():
    """Celkový výsledek flotily se v aplikaci nezobrazuje: na současných datech je neúplný
    (chybí variabilní náklady, režie a nastavení nákladů strojů v zakázkách)."""
    m = model()
    r = ec.total_result(m, rentals=[("MINI_BAGR", 4, 6500.0)], jobs=[("J001", 0)])
    assert r["výsledek"].status == "neúplný" and r["výsledek"].fictitious
    assert any("režie" in x for x in r["výsledek"].missing)


# ---------------------------------------------------------------------------
# 6. Regrese současných výsledků a exportu DOCX
# ---------------------------------------------------------------------------

def _docx_parts(data: bytes) -> dict:
    z = zipfile.ZipFile(io.BytesIO(data))
    return {n: hashlib.sha256(z.read(n)).hexdigest() for n in sorted(z.namelist())}


def current_values() -> dict:
    """Současné ekonomické výsledky aplikace (zaokrouhleno na haléře) – pro referenci."""
    m = model()
    jobs, result, new_plan, contrib = plan()
    s = sch.summary(new_plan)
    out = {
        "plan": {k: round(float(v), 2) for k, v in s.items()},
        "plan_optimal": bool(result["optimal"]),
        "plan_jobs": [[r.job_name, int(r.start_day), int(r.end_day), int(r.days_late), round(r.penalty, 2),
                       round(r.profit_after_penalty, 2)] for r in new_plan[new_plan["scheduled"]].itertuples()],
        "job_contributions": [[r.job_name, round(r.profit, 2), round(r.added_penalty, 2), round(r.contribution, 2)]
                              for r in contrib.itertuples()],
        "unit_economics": [[row["machine_type"]] + [round(pf.type_month(row, 1)[k], 2)
                                                    for k in ("days", "revenue", "costs", "profit")]
                           for _, row in machines().iterrows()],
        "purchase_best_3m": {k: (round(v, 2) if isinstance(v, float) else v)
                             for k, v in pf.best_portfolio(pf.all_portfolios(machines()), 3000000).items()
                             if k in ("counts", "units", "investment", "revenue", "costs", "profit")},
        "fleet_fixed_month": round(ec.fleet_fixed_costs(m)["celkem"].value, 2),
        "unit_fixed_month": {t: round(ec.unit_fixed_total(m, t).value, 2) for t in m.machine_types["type_id"]},
    }
    from test_report_docx import inputs                                  # stejné vstupy jako aplikace
    out["docx_vychozi"] = _docx_parts(rd.build_report(**inputs("vychozi")))
    return out


def test_current_results_match_reference():
    ref = json.loads(REFERENCE.read_text(encoding="utf-8"))
    now = json.loads(json.dumps(current_values(), ensure_ascii=False))
    for key in ref:
        assert now[key] == ref[key], key
    assert set(now) == set(ref)


def test_headline_values_are_unchanged():
    """Hlavní částky zobrazené v aplikaci (Dashboard, Vlastní zakázky, Strojový park)."""
    ref = json.loads(REFERENCE.read_text(encoding="utf-8"))
    assert ref["plan"]["profit_after_penalty"] == 892000.0 and ref["plan"]["penalty"] == 108000.0
    assert ref["plan"]["revenue"] == 1955000.0 and ref["plan"]["costs"] == 955000.0
    assert ref["fleet_fixed_month"] == 528000.0
    assert ref["purchase_best_3m"]["profit"] == 199215.0
    assert [r[0] for r in ref["unit_economics"]][:2] == ["Mini bagr", "Pásový dumper"]


if __name__ == "__main__":
    if "--update" in sys.argv:
        REFERENCE.write_text(json.dumps(current_values(), ensure_ascii=False, indent=1, sort_keys=True),
                             encoding="utf-8")
        print(f"Reference uložena: {REFERENCE}")
    else:
        print(__doc__)
