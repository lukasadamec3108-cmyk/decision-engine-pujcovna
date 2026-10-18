"""Testy plánu zakázek (scheduling.py).

Spuštění: python -m pytest tests   nebo   python tests/test_scheduling.py
"""

import itertools
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import scheduling as sch  # noqa: E402

DEMO = pd.read_csv(ROOT / "data" / "jobs_demo.csv")
DEMO_CAPACITY = {
    "Mini bagr": 1, "Pásový dumper": 1, "Smykový nakladač": 1, "Vibrační deska": 2,
    "Kolový bagr": 1, "Silniční válec": 1, "Řezačka asfaltu": 1, "Nákladní auto": 1,
}


def jobs(*rows) -> pd.DataFrame:
    """rows = (název, stroje, délka, termín, penále/den)."""
    return pd.DataFrame(
        [dict(job_name=n, required_machines=m, duration_days=d, deadline_days=t,
              delay_penalty_per_day=p, expected_revenue=100_000, cost_base=60_000,
              priority="Medium", risk_level="Low") for n, m, d, t, p in rows]
    )


def plan_of(df, capacity):
    res = sch.optimize_schedule(df, capacity)
    return res, sch.evaluate(df, res["starts"])


def test_missing_machine_is_not_scheduled():
    df = jobs(("A", "Bagr", 2, 5, 1000), ("B", "Neexistující stroj", 2, 5, 1000))
    res, plan = plan_of(df, {"Bagr": 1})
    assert plan.loc[0, "scheduled"] and not plan.loc[1, "scheduled"]
    assert sch.summary(plan)["waiting"] == 1


def test_same_machine_jobs_do_not_overlap():
    df = jobs(("A", "Bagr", 3, 3, 1000), ("B", "Bagr", 3, 6, 1000), ("C", "Bagr;Auto", 2, 8, 1000))
    res, plan = plan_of(df, {"Bagr": 1, "Auto": 1})
    assert sch.conflicts(plan, {"Bagr": 1, "Auto": 1}) == []
    assert sch.summary(plan)["length"] == 8 and sch.summary(plan)["penalty"] == 0


def test_capacity_two_allows_parallel_jobs():
    df = jobs(("A", "Deska", 3, 3, 1000), ("B", "Deska", 3, 3, 1000))
    _, plan2 = plan_of(df, {"Deska": 2})
    _, plan1 = plan_of(df, {"Deska": 1})
    assert sch.summary(plan2)["length"] == 3 and sch.summary(plan2)["penalty"] == 0
    assert sch.summary(plan1)["length"] == 6 and sch.summary(plan1)["penalty"] == 3000


def test_plan_without_delay():
    df = jobs(("A", "Bagr", 2, 10, 5000), ("B", "Auto", 4, 10, 5000))
    res, plan = plan_of(df, {"Bagr": 1, "Auto": 1})
    assert res["optimal"] and sch.summary(plan)["late"] == 0 and sch.summary(plan)["penalty"] == 0


def test_matches_brute_force_on_small_case():
    df = jobs(("A", "Bagr", 3, 3, 2000), ("B", "Bagr", 2, 2, 9000),
              ("C", "Bagr;Auto", 2, 4, 1000), ("D", "Auto", 4, 4, 3000))
    cap = {"Bagr": 1, "Auto": 1}
    res, plan = plan_of(df, cap)
    best = None
    horizon = int(df["duration_days"].sum())
    for starts in itertools.product(range(1, horizon + 1), repeat=len(df)):
        cand = sch.evaluate(df, dict(enumerate(starts)))
        if not sch.conflicts(cand, cap):
            pen = sch.summary(cand)["penalty"]
            best = pen if best is None else min(best, pen)
    assert res["optimal"] and sch.summary(plan)["penalty"] == best


def test_demo_plan_is_valid_and_not_worse_than_rule():
    res, plan = plan_of(DEMO, DEMO_CAPACITY)
    s = sch.summary(plan)
    assert s["scheduled"] == 12 and sch.conflicts(plan, DEMO_CAPACITY) == []
    assert s["profit_after_penalty"] == s["profit"] - s["penalty"]
    rule = sch.summary(sch.evaluate(DEMO, sch.greedy_schedule(DEMO, DEMO_CAPACITY)))
    assert s["penalty"] <= rule["penalty"]


def test_profit_objective_equals_penalty_objective_when_all_jobs_planned():
    # Zisk všech zakázek je pevný – zisk po penále se liší jen o penále.
    _, plan = plan_of(DEMO, DEMO_CAPACITY)
    rule = sch.evaluate(DEMO, sch.greedy_schedule(DEMO, DEMO_CAPACITY))
    a, b = sch.summary(plan), sch.summary(rule)
    assert a["profit"] == b["profit"]
    assert (a["profit_after_penalty"] - b["profit_after_penalty"]) == (b["penalty"] - a["penalty"])


def test_contributions():
    _, plan = plan_of(DEMO, DEMO_CAPACITY)
    c = sch.job_contributions(DEMO, DEMO_CAPACITY, sch.summary(plan)["penalty"])
    assert len(c) == 12
    assert ((c["contribution"] - (c["profit"] - c["added_penalty"])).abs() < 1e-9).all()
    assert (c["added_penalty"] >= 0).all()


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK", name)
