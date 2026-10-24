"""Testy modulu investice (etapa M1) – investiční výpočty záložek 1–4.

Referenční hodnoty (tests/fixtures/investice_reference.json) vznikly z modulu po ověření, že dává
bit po bitu stejné výsledky jako původní kód v app.py (main 5fef75c, 5 scénářů včetně celých polí
simulace a DOCX). Test hlídá přesné hodnoty: otisky polí z jejich bajtů a čísla s plnou přesností.
Po ZÁMĚRNÉ a schválené změně výpočtu obnovte referenci: python tests/test_investice.py --update
"""

import hashlib
import json
import math
import sys
from pathlib import Path

import numpy as np
import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import datamodel as dm  # noqa: E402
import investice as inv  # noqa: E402

REFERENCE = ROOT / "tests" / "fixtures" / "investice_reference.json"
# (stroj 1, stroj 2, scénář, fixní náklad, dny pro rychlý odhad, počet simulací)
CASES = {
    "vychozi": ("Mini bagr", "Pásový dumper", "Realistický", 100000, 18, 1000),
    "pesimisticky": ("Mini bagr", "Pásový dumper", "Pesimistický", 120000, 18, 1000),
    "optimisticky": ("Kolový bagr", "Silniční válec", "Optimistický", 60000, 25, 2000),
    "oba_ztratove": ("Vibrační deska", "Řezačka asfaltu", "Pesimistický", 150000, 5, 300),
}


def machines():
    return dm.legacy_machines(dm.load_model())


def run(case):
    m1, m2, scenario, fixed, days, n = CASES[case]
    df = machines()
    row_1 = df[df["machine_type"] == m1].iloc[0]
    row_2 = df[df["machine_type"] == m2].iloc[0]
    lo, hi = inv.scenario_days(scenario)
    return inv.evaluate(row_1, row_2, m1, m2, fixed, lo, hi, days, n)


def _fingerprint(value):
    """Přesný otisk hodnoty: pole podle bajtů, čísla jako hexadecimální zápis floatu."""
    if isinstance(value, np.ndarray):
        return {"sha256": hashlib.sha256(value.tobytes()).hexdigest(), "dtype": str(value.dtype),
                "shape": list(value.shape)}
    if isinstance(value, (float, np.floating)):
        return "nan" if math.isnan(value) else float(value).hex()
    if isinstance(value, (np.integer,)):
        return int(value)
    if isinstance(value, (list, tuple)):
        return [_fingerprint(v) for v in value]
    if isinstance(value, dict):
        return {k: _fingerprint(v) for k, v in value.items()}
    return value


def snapshot(case) -> dict:
    r = run(case)
    keys = sorted(k for k in r if k not in ("sim_1", "sim_2"))
    out = {k: _fingerprint(r[k]) for k in keys}
    out["sim_1"] = _fingerprint(r["sim_1"])
    out["sim_2"] = _fingerprint(r["sim_2"])
    return out


# ---------------------------------------------------------------------------
# Přesná shoda s referencí (pořadí losování, vzorce)
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("case", sorted(CASES))
def test_results_match_reference_exactly(case):
    ref = json.loads(REFERENCE.read_text(encoding="utf-8"))[case]
    got = snapshot(case)
    diff = sorted(k for k in set(ref) | set(got) if ref.get(k) != got.get(k))
    assert not diff, f"{case}: liší se {diff}"


def test_same_inputs_give_identical_results():
    a, b = run("vychozi"), run("vychozi")
    assert a["profit_mini"].tobytes() == b["profit_mini"].tobytes()
    assert a["profit_dumper"].tobytes() == b["profit_dumper"].tobytes()


def test_random_draw_order_is_machine_1_then_machine_2():
    df = machines()
    r1, r2 = df.iloc[0], df.iloc[1]
    rng = np.random.RandomState(inv.SIMULATION_SEED)
    first = inv.simulate_machine(rng, r1, float(r1["daily_price"]), 100000, 15, 25, 500)
    second = inv.simulate_machine(rng, r2, float(r2["daily_price"]), 100000, 15, 25, 500)
    s1, s2 = inv.simulate(r1, r2, 100000, 15, 25, 500)
    assert first["profit"].tobytes() == s1["profit"].tobytes()
    assert second["profit"].tobytes() == s2["profit"].tobytes()
    # jiný seed = jiné výsledky (seed se opravdu používá)
    o1, _ = inv.simulate(r1, r2, 100000, 15, 25, 500, seed=inv.SIMULATION_SEED + 1)
    assert o1["profit"].tobytes() != s1["profit"].tobytes()


# ---------------------------------------------------------------------------
# Pravidla a vzorce
# ---------------------------------------------------------------------------

def test_scenario_days():
    assert inv.scenario_days("Pesimistický") == (8, 15)
    assert inv.scenario_days("Realistický") == (15, 25)
    assert inv.scenario_days("Optimistický") == (22, 30)
    assert inv.scenario_days("cokoli jiného") == (22, 30)          # jako dříve větev „else“


def test_risk_light_thresholds():
    assert inv.risk_light(0.0) == "🟢 Nízké riziko"
    assert inv.risk_light(0.0499) == "🟢 Nízké riziko"
    assert inv.risk_light(0.05) == "🟡 Střední riziko"
    assert inv.risk_light(0.1499) == "🟡 Střední riziko"
    assert inv.risk_light(0.15) == "🔴 Vysoké riziko"


def test_better_index_and_score_rows():
    assert inv.better_index(2, 1, True) == 0 and inv.better_index(2, 1, False) == 1
    assert inv.better_index(1, 1, True) is None                      # při shodě bod nikdo
    rows = inv.score_rows(10, 20, 0.1, 0.05, -5, -3, None, 12.0)
    assert [r[0] for r in rows] == ["profit", "loss", "var95", "payback"]
    assert [r[3] for r in rows] == [1, 1, 1, None]                   # návratnost jen když oba vydělávají


def test_scorecard_counts_points():
    r = run("vychozi")
    points = [best for *_, best in r["score_rows"]]
    assert r["scorecard_mini"] == points.count(0) and r["scorecard_dumper"] == points.count(1)


def test_optimize_days_prefers_more_profitable_machine():
    o = inv.optimize_days(1000.0, 500.0)
    assert o["best_d1"] == 30 and o["best_d2"] == 0 and math.isclose(o["x1_opt"], 30)
    assert math.isclose(o["best_dp_profit"], 30000.0)
    o = inv.optimize_days(-10.0, 200.0)
    assert o["best_d1"] == 0 and o["best_d2"] == 30


def test_sensitivity_sorted_and_price_equals_days_effect():
    r = run("vychozi")
    vals = [v for _, v in r["sensitivity_data"]]
    assert len(vals) == 8 and vals == sorted(vals, reverse=True)
    eff = dict(r["sensitivity_data"])
    for m in ("Mini bagr", "Pásový dumper"):                         # tržba = cena × dny
        assert math.isclose(eff[f"Denní cena – {m}"], eff[f"Vytížení (dny/měsíc) – {m}"])
    assert r["labels"] == [k for k, _ in r["sensitivity_data"]]


def test_break_even_map_and_minimum_price_formulas():
    m1, m2, scenario, fixed, days, n = CASES["vychozi"]
    df = machines()
    row_1 = df[df["machine_type"] == m1].iloc[0]
    r = run("vychozi")
    price = float(row_1["daily_price"])
    assert r["break_even_mini"] == fixed / price
    assert r["profit_curve_mini"][10] == price * 10 - fixed
    assert r["profit_matrix"].shape == (len(r["price_range"]), len(r["days_range"]))
    assert r["profit_matrix"][0, 0] == r["price_range"][0] * r["days_range"][0] - fixed
    assert r["required_price_1"] == (fixed + row_1["service_cost_mean"]) / days


def test_break_even_is_nan_for_zero_price():
    # Nulová cena u stroje 2: bod zvratu nelze určit (NaN), jako v původním kódu. (Nulová cena u stroje 1
    # skončí chybou už v původním kódu při sestavení mapy ziskovosti – viz otevřené body etapy M1.)
    df = machines()
    r2 = df.iloc[1].copy()
    r2["daily_price"] = 0
    r = inv.evaluate(df.iloc[0], r2, "A", "B", 100000, 15, 25, 18, 50)
    assert math.isnan(r["break_even_dumper"])


def test_module_has_no_streamlit_dependency():
    import ast
    tree = ast.parse((ROOT / "investice.py").read_text(encoding="utf-8"))
    imported = {a.name.split(".")[0] for n in ast.walk(tree) if isinstance(n, ast.Import) for a in n.names}
    imported |= {n.module.split(".")[0] for n in ast.walk(tree) if isinstance(n, ast.ImportFrom) and n.module}
    assert "streamlit" not in imported and imported <= {"numpy", "scipy"}, imported


if __name__ == "__main__":
    if "--update" in sys.argv:
        REFERENCE.write_text(json.dumps({c: snapshot(c) for c in sorted(CASES)}, ensure_ascii=False, indent=1,
                                        sort_keys=True), encoding="utf-8")
        print(f"Reference uložena: {REFERENCE}")
    else:
        print(__doc__)
