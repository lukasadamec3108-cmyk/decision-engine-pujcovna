"""Testy výběru nákupní sestavy (portfolio.py).

Spuštění: python -m pytest tests   nebo   python tests/test_portfolio.py
"""

import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import portfolio as pf  # noqa: E402

DATA = pd.read_csv(ROOT / "data" / "machines_demo.csv")


def machine(**kw) -> pd.Series:
    base = dict(machine_type="Test", daily_price=1000, purchase_price=100_000,
                monthly_fixed_cost=5_000, service_cost_mean=1_000, demand_days_mean=18,
                failure_prob=0.0, failure_days_mean=0)
    base.update(kw)
    return pd.Series(base)


def test_second_unit_does_not_multiply_demand():
    one = pf.type_month(machine(demand_days_mean=18), 1)
    two = pf.type_month(machine(demand_days_mean=18), 2)
    assert one["days"] == 18 and two["days"] == 18          # ne 36 dní
    assert two["costs"] == 2 * one["costs"]                  # náklady nese každý kus
    assert two["profit"] < one["profit"]


def test_demand_above_capacity_is_split():
    assert pf.type_month(machine(demand_days_mean=45), 1)["days"] == 30
    assert pf.type_month(machine(demand_days_mean=45), 2)["days"] == 45
    assert pf.type_month(machine(demand_days_mean=80), 2)["days"] == 60   # max 30 na kus


def test_failures_reduce_days():
    m = machine(demand_days_mean=18, failure_prob=0.1, failure_days_mean=2)
    assert abs(pf.type_month(m, 1)["days"] - 17.8) < 1e-9


def test_nothing_bought_when_all_loss_making():
    df = DATA.copy()
    df["demand_days_mean"] = 0
    options = pf.all_portfolios(df)
    best = pf.best_portfolio(options, 5_000_000)
    assert best["units"] == 0 and best["profit"] == 0
    assert pf.cheapest_profitable(options) is None


def test_best_portfolio_respects_budget_and_is_optimal():
    options = pf.all_portfolios(DATA)
    assert len(options) == 3 ** len(DATA)
    for budget in (100_000, 500_000, 2_200_000, 3_000_000, 5_000_000):
        best = pf.best_portfolio(options, budget)
        assert best["investment"] <= budget
        assert best["profit"] >= 0
        assert all(p["profit"] <= best["profit"] + 1e-9 for p in options if p["investment"] <= budget)


def test_frontier_matches_best_portfolio():
    options = pf.all_portfolios(DATA)
    steps = pf.budget_frontier(options)
    for budget in range(0, 10_000_001, 50_000):
        expected = pf.best_portfolio(options, budget)["profit"]
        step = [p for x, p in steps if x <= budget][-1]
        assert abs(step["profit"] - expected) < 1e-6


def test_payback():
    assert pf.payback_months(dict(profit=10_000, investment=120_000)) == 12
    assert pf.payback_months(dict(profit=0, investment=120_000)) is None


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK", name)
