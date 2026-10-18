"""Výběr nákupní sestavy strojů v rámci rozpočtu (záložka „Nákup strojů“).

Čisté funkce bez Streamlitu, aby šly samostatně testovat.

Model jednoho průměrného měsíce pro typ stroje s `n` kusy:
    dny pronájmu = min(poptávka, 30 × n) − očekávané dny poruch
    tržby        = denní cena × dny pronájmu
    náklady      = n × (měsíční fixní náklad stroje + průměrný servis)
    zisk         = tržby − náklady

Poptávka (`demand_days_mean`) je celková měsíční poptávka po daném typu stroje.
Další kus stejného typu ji nenásobí – pokryje jen tu část, na kterou kapacita
předchozích kusů (30 dní na kus) nestačí.
"""

from itertools import product
from math import ceil

DAYS_PER_MONTH = 30
MAX_UNITS = 2


def type_month(row, units: int) -> dict:
    """Očekávaný průměrný měsíc pro `units` kusů jednoho typu stroje."""
    demand = max(float(row["demand_days_mean"]), 0.0)
    if units <= 0:
        return dict(units=0, investment=0.0, days=0.0, revenue=0.0, costs=0.0, profit=0.0)
    capacity = DAYS_PER_MONTH * units
    # Poruchy postihují kusy, které jsou potřeba k pokrytí poptávky.
    busy_units = min(units, ceil(demand / DAYS_PER_MONTH)) if demand > 0 else 0
    lost_days = busy_units * float(row["failure_prob"]) * float(row["failure_days_mean"])
    days = max(min(demand, capacity) - lost_days, 0.0)
    revenue = float(row["daily_price"]) * days
    costs = units * (float(row["monthly_fixed_cost"]) + float(row["service_cost_mean"]))
    return dict(
        units=units,
        investment=units * float(row["purchase_price"]),
        days=days,
        revenue=revenue,
        costs=costs,
        profit=revenue - costs,
    )


def all_portfolios(df, max_units: int = MAX_UNITS) -> list:
    """Všechny kombinace 0–max_units kusů u všech typů strojů (včetně „nic nekupovat“)."""
    names = df["machine_type"].tolist()
    options = [
        [type_month(row, k) for k in range(max_units + 1)] for _, row in df.iterrows()
    ]
    result = []
    for counts in product(range(max_units + 1), repeat=len(names)):
        parts = [options[i][k] for i, k in enumerate(counts)]
        result.append(
            dict(
                counts={names[i]: k for i, k in enumerate(counts) if k > 0},
                investment=sum(p["investment"] for p in parts),
                revenue=sum(p["revenue"] for p in parts),
                costs=sum(p["costs"] for p in parts),
                profit=sum(p["profit"] for p in parts),
                units=sum(counts),
                lines=[(names[i], parts[i]) for i, k in enumerate(counts) if k > 0],
            )
        )
    return result


def _rank(p: dict):
    # Vyšší zisk; při shodě nižší investice a méně kusů.
    return (round(p["profit"], 6), -p["investment"], -p["units"])


def best_portfolio(portfolios: list, budget: float) -> dict:
    """Nejlepší sestava v rozpočtu. Prázdná sestava (zisk 0) vyhraje, pokud nic nevydělává."""
    feasible = [p for p in portfolios if p["investment"] <= budget + 1e-9]
    return max(feasible, key=_rank)


def cheapest_profitable(portfolios: list):
    """Nejlevnější sestava s kladným ziskem (bez ohledu na rozpočet), nebo None."""
    profitable = [p for p in portfolios if p["profit"] > 1e-9]
    return min(profitable, key=lambda p: (p["investment"], -p["profit"])) if profitable else None


def budget_frontier(portfolios: list) -> list:
    """Přesná schodovitá křivka: rozpočty, od kterých se nejlepší sestava zlepší.

    Vrací seznam (rozpočet, nejlepší sestava). Mezi dvěma body je nejlepší
    sestava stejná jako v levém bodě – stejná logika jako best_portfolio().
    """
    steps, current = [], None
    for p in sorted(portfolios, key=lambda p: (p["investment"], [-x for x in _rank(p)])):
        if current is None or _rank(p) > _rank(current):
            current = p
            steps.append((p["investment"], p))
    return steps


def payback_months(portfolio: dict):
    """Prostá návratnost v měsících (investice ÷ měsíční provozní zisk), jen při kladném zisku."""
    if portfolio["profit"] <= 0 or portfolio["investment"] <= 0:
        return None
    return portfolio["investment"] / portfolio["profit"]


def profit_with_demand(df, counts: dict, factor: float) -> float:
    """Zisk stejné sestavy, pokud by poptávka byla `factor`× očekávané (zátěžový test)."""
    total = 0.0
    for _, row in df.iterrows():
        k = counts.get(row["machine_type"], 0)
        if k:
            changed = row.copy()
            changed["demand_days_mean"] = float(row["demand_days_mean"]) * factor
            total += type_month(changed, k)["profit"]
    return total
