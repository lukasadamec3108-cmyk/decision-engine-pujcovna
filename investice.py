"""Investiční výpočty záložek 1–4 (etapa M1): porovnání dvou strojů.

Simulace Monte Carlo, odvozené metriky, optimalizace LP/DP, doplňkové bodové hodnocení,
citlivost, křivka bodu zvratu, mapa ziskovosti a minimální cena pronájmu.

Kód je přesunutý z app.py BEZ ZMĚNY vzorců i pořadí operací. Simulace losuje náhodná čísla
ve stejném pořadí jako dříve (nejdřív všechna čísla stroje 1, pak stroje 2, se stejným zdrojem
náhodných čísel), takže při stejných vstupech dává bit po bitu stejné výsledky.

Modul nezávisí na Streamlitu a nic neformátuje – texty a zobrazení zůstávají v app.py.
Výsledek je slovník se stejnými názvy hodnot, jaké dříve používal app.py.
"""

import numpy as np
from scipy.optimize import linprog

# Pevný náhodný základ: stejné vstupy dají vždy stejné výsledky. Pro řízený přepočet
# s jiným základem stačí změnit tuto konstantu (případně ji později nabídnout v panelu).
SIMULATION_SEED = 42

# Scénář trhu → rozsah dní pronájmu za měsíc (min, max). Jiná hodnota = optimistický (jako dříve).
SCENARIO_DAYS = {"Pesimistický": (8, 15), "Realistický": (15, 25), "Optimistický": (22, 30)}

# Údaje pro citlivost (±10 %): (popisek, kód)
DRIVERS = [
    ("Denní cena", "price"),
    ("Vytížení (dny/měsíc)", "days"),
    ("Fixní náklad", "fixed"),
    ("Servisní náklad", "service"),
]


def scenario_days(scenario: str) -> tuple:
    """Rozsah dní pronájmu měsíčně pro scénář trhu."""
    if scenario == "Pesimistický":
        return SCENARIO_DAYS["Pesimistický"]
    if scenario == "Realistický":
        return SCENARIO_DAYS["Realistický"]
    return SCENARIO_DAYS["Optimistický"]


def risk_light(prob_loss: float) -> str:
    """Vrátí textový semafor podle pravděpodobnosti ztráty."""
    if prob_loss < 0.05:
        return "🟢 Nízké riziko"
    if prob_loss < 0.15:
        return "🟡 Střední riziko"
    return "🔴 Vysoké riziko"


# ---------------------------------------------------------------------------
# Simulace Monte Carlo
# ---------------------------------------------------------------------------

def simulate_machine(rng, row, avg_daily, fixed_cost, min_days, max_days, n) -> dict:
    """Simulace n měsíců jednoho stroje. Pořadí volání rng je pevné (dny, servis, porucha,
    délka poruchy, cena) – mění-li se, změní se i výsledky."""
    demand_mean = row["demand_days_mean"]
    mode_days = np.clip(demand_mean, min_days, max_days)
    util_days = rng.triangular(min_days, mode_days, max_days, n)

    service_cost = rng.normal(
        row["service_cost_mean"],
        row["service_cost_std"],
        n,
    )
    service_cost = np.clip(service_cost, 0, None)

    failure_happens = rng.rand(n) < row["failure_prob"]
    failure_days = np.where(
        failure_happens,
        rng.poisson(row["failure_days_mean"], n),
        0,
    )
    real_days = np.maximum(util_days - failure_days, 0)

    price_std = avg_daily * 0.08
    daily_price = rng.normal(avg_daily, price_std, n)
    daily_price = np.clip(daily_price, avg_daily * 0.8, avg_daily * 1.2)

    rev = real_days * daily_price
    profit = rev - fixed_cost - service_cost
    return dict(util_days=util_days, service_cost=service_cost, failure_days=failure_days,
                real_days=real_days, daily_price=daily_price, revenue=rev, profit=profit)


def simulate(row_1, row_2, fixed_cost, min_days, max_days, n, seed: int = SIMULATION_SEED) -> tuple:
    """Simulace obou strojů se společným zdrojem náhodných čísel (nejdřív stroj 1, pak stroj 2)."""
    rng = np.random.RandomState(seed)
    sim_1 = simulate_machine(rng, row_1, float(row_1["daily_price"]), fixed_cost, min_days, max_days, n)
    sim_2 = simulate_machine(rng, row_2, float(row_2["daily_price"]), fixed_cost, min_days, max_days, n)
    return sim_1, sim_2


# ---------------------------------------------------------------------------
# Optimalizace LP a DP (rozdělení 30 dní mezi dva stroje)
# ---------------------------------------------------------------------------

def optimize_days(profit_mini_day: float, profit_dumper_day: float) -> dict:
    c = [-profit_mini_day, -profit_dumper_day]
    A = [[1, 1]]
    b = [30]
    bounds = [(0, 30), (0, 30)]

    lp_result = linprog(c, A_ub=A, b_ub=b, bounds=bounds, method="highs")
    if lp_result.success:
        x1_opt, x2_opt = lp_result.x
        lp_profit = -lp_result.fun
    else:
        x1_opt, x2_opt = 0, 0
        lp_profit = 0

    best_dp_profit = -np.inf
    best_d1, best_d2 = 0, 0

    for d1 in range(31):
        d2 = 30 - d1
        dp_profit = d1 * profit_mini_day + d2 * profit_dumper_day

        if dp_profit > best_dp_profit:
            best_dp_profit = dp_profit
            best_d1, best_d2 = d1, d2

    return dict(x1_opt=x1_opt, x2_opt=x2_opt, lp_profit=lp_profit,
                best_d1=best_d1, best_d2=best_d2, best_dp_profit=best_dp_profit)


# ---------------------------------------------------------------------------
# Doplňkové bodové hodnocení
# ---------------------------------------------------------------------------
# Bod získá jen stroj, který je v ukazateli skutečně lepší; při shodě nezíská bod nikdo.
# Hlavní doporučení (better_machine) určuje vyšší očekávaný zisk – body ho jen doplňují.

def better_index(a, b, higher_is_better: bool):
    if a == b:
        return None
    return (0 if a > b else 1) if higher_is_better else (0 if a < b else 1)


def score_rows(expected_profit_mini, expected_profit_dumper, prob_loss_mini, prob_loss_dumper,
               var95_mini, var95_dumper, roi_mini, roi_dumper) -> list:
    """Ukazatele bodového hodnocení: (kód, hodnota stroje 1, hodnota stroje 2, index lepšího / None)."""
    return [
        ("profit", expected_profit_mini, expected_profit_dumper,
         better_index(expected_profit_mini, expected_profit_dumper, True)),
        ("loss", prob_loss_mini, prob_loss_dumper,
         better_index(prob_loss_mini, prob_loss_dumper, False)),
        ("var95", var95_mini, var95_dumper,
         better_index(var95_mini, var95_dumper, True)),
        ("payback", roi_mini, roi_dumper,
         better_index(roi_mini, roi_dumper, False)
         if roi_mini is not None and roi_dumper is not None else None),
    ]


# ---------------------------------------------------------------------------
# Citlivost (±10 %)
# ---------------------------------------------------------------------------

def sensitivity(avg_daily_mini, avg_daily_dumper, row_1, row_2, fixed_cost, expected_days,
                machine_1, machine_2) -> list:
    """Rozdíl zisku mezi −10 % a +10 % jednoho údaje (rychlý odhad), seřazený od největšího vlivu."""
    sensitivity_data = []
    for label, driver in DRIVERS:
        if driver == "price":
            low_1 = (
                (avg_daily_mini * 0.9) * expected_days
                - fixed_cost
                - row_1["service_cost_mean"]
            )
            high_1 = (
                (avg_daily_mini * 1.1) * expected_days
                - fixed_cost
                - row_1["service_cost_mean"]
            )
            low_2 = (
                (avg_daily_dumper * 0.9) * expected_days
                - fixed_cost
                - row_2["service_cost_mean"]
            )
            high_2 = (
                (avg_daily_dumper * 1.1) * expected_days
                - fixed_cost
                - row_2["service_cost_mean"]
            )
        elif driver == "days":
            low_1 = (
                avg_daily_mini * (expected_days * 0.9)
                - fixed_cost
                - row_1["service_cost_mean"]
            )
            high_1 = (
                avg_daily_mini * (expected_days * 1.1)
                - fixed_cost
                - row_1["service_cost_mean"]
            )
            low_2 = (
                avg_daily_dumper * (expected_days * 0.9)
                - fixed_cost
                - row_2["service_cost_mean"]
            )
            high_2 = (
                avg_daily_dumper * (expected_days * 1.1)
                - fixed_cost
                - row_2["service_cost_mean"]
            )
        elif driver == "fixed":
            low_1 = (
                avg_daily_mini * expected_days
                - (fixed_cost * 0.9)
                - row_1["service_cost_mean"]
            )
            high_1 = (
                avg_daily_mini * expected_days
                - (fixed_cost * 1.1)
                - row_1["service_cost_mean"]
            )
            low_2 = (
                avg_daily_dumper * expected_days
                - (fixed_cost * 0.9)
                - row_2["service_cost_mean"]
            )
            high_2 = (
                avg_daily_dumper * expected_days
                - (fixed_cost * 1.1)
                - row_2["service_cost_mean"]
            )
        else:
            low_1 = (
                avg_daily_mini * expected_days
                - fixed_cost
                - (row_1["service_cost_mean"] * 0.9)
            )
            high_1 = (
                avg_daily_mini * expected_days
                - fixed_cost
                - (row_1["service_cost_mean"] * 1.1)
            )
            low_2 = (
                avg_daily_dumper * expected_days
                - fixed_cost
                - (row_2["service_cost_mean"] * 0.9)
            )
            high_2 = (
                avg_daily_dumper * expected_days
                - fixed_cost
                - (row_2["service_cost_mean"] * 1.1)
            )

        sensitivity_data.append((f"{label} – {machine_1}", abs(high_1 - low_1)))
        sensitivity_data.append((f"{label} – {machine_2}", abs(high_2 - low_2)))

    return sorted(sensitivity_data, key=lambda x: x[1], reverse=True)


# ---------------------------------------------------------------------------
# Celé vyhodnocení (stejné pořadí jako dříve v app.py)
# ---------------------------------------------------------------------------

def evaluate(row_1, row_2, machine_1: str, machine_2: str, fixed_cost, min_days: int, max_days: int,
             expected_days: int, n: int, seed: int = SIMULATION_SEED) -> dict:
    """Všechny investiční výsledky pro dvojici strojů. Klíče odpovídají dřívějším proměnným v app.py."""
    avg_daily_mini = float(row_1["daily_price"])
    avg_daily_dumper = float(row_2["daily_price"])

    # --- Monte Carlo simulace ---
    sim_1, sim_2 = simulate(row_1, row_2, fixed_cost, min_days, max_days, n, seed)
    profit_mini = sim_1["profit"]
    profit_dumper = sim_2["profit"]

    # --- Realistický scénář při očekávané poptávce (rychlý odhad) ---
    expected_profit_mini_real = (
        avg_daily_mini * expected_days - fixed_cost - row_1["service_cost_mean"]
    )
    expected_profit_dumper_real = (
        avg_daily_dumper * expected_days - fixed_cost - row_2["service_cost_mean"]
    )

    # --- Základní metriky ---
    prob_loss_mini = float(np.mean(profit_mini < 0))
    prob_loss_dumper = float(np.mean(profit_dumper < 0))

    risk_loss_mini = prob_loss_mini * 100
    risk_loss_dumper = prob_loss_dumper * 100

    break_even_mini = fixed_cost / avg_daily_mini if avg_daily_mini > 0 else np.nan
    break_even_dumper = fixed_cost / avg_daily_dumper if avg_daily_dumper > 0 else np.nan

    expected_profit_mini = float(np.mean(profit_mini))
    expected_profit_dumper = float(np.mean(profit_dumper))

    # --- Percentily / VaR / intervaly ---
    p10_mini = float(np.percentile(profit_mini, 10))
    p10_dumper = float(np.percentile(profit_dumper, 10))
    p50_mini = float(np.percentile(profit_mini, 50))
    p50_dumper = float(np.percentile(profit_dumper, 50))
    p90_mini = float(np.percentile(profit_mini, 90))
    p90_dumper = float(np.percentile(profit_dumper, 90))

    var95_mini = float(np.percentile(profit_mini, 5))
    var95_dumper = float(np.percentile(profit_dumper, 5))

    ci_mini = (p10_mini, p90_mini)
    ci_dumper = (p10_dumper, p90_dumper)

    prob_dumper_better = float(np.mean(profit_dumper > profit_mini))

    # --- Prostá návratnost (měsíce) ---
    investment_cost_mini = row_1["purchase_price"]
    investment_cost_dumper = row_2["purchase_price"]

    roi_mini = (
        investment_cost_mini / expected_profit_mini if expected_profit_mini > 0 else None
    )
    roi_dumper = (
        investment_cost_dumper / expected_profit_dumper
        if expected_profit_dumper > 0
        else None
    )

    # --- Rozdíly pro shrnutí ---
    profit_diff = expected_profit_dumper - expected_profit_mini
    risk_diff = (prob_loss_mini - prob_loss_dumper) * 100

    better_machine = (
        machine_2 if expected_profit_dumper > expected_profit_mini else machine_1
    )

    # --- LP a DP ---
    profit_mini_day = avg_daily_mini - row_1["service_cost_mean"] / 30 - fixed_cost / 30
    profit_dumper_day = avg_daily_dumper - row_2["service_cost_mean"] / 30 - fixed_cost / 30
    opt = optimize_days(profit_mini_day, profit_dumper_day)
    x1_opt, x2_opt = opt["x1_opt"], opt["x2_opt"]
    best_d1, best_d2 = opt["best_d1"], opt["best_d2"]

    mc_better = machine_2 if expected_profit_dumper > expected_profit_mini else machine_1
    lp_better = (
        machine_1
        if x1_opt > x2_opt
        else machine_2 if x2_opt > x1_opt else "oba stroje podobně"
    )
    dp_better = (
        machine_1
        if best_d1 > best_d2
        else machine_2 if best_d2 > best_d1 else "oba stroje podobně"
    )

    # --- Finální investiční verdikt (původní body; aplikace je nezobrazuje, zachováno beze změny) ---
    verdict_score_mini = 0
    verdict_score_dumper = 0

    if expected_profit_mini > expected_profit_dumper:
        verdict_score_mini += 1
    else:
        verdict_score_dumper += 1

    if prob_loss_mini < prob_loss_dumper:
        verdict_score_mini += 1
    else:
        verdict_score_dumper += 1

    if roi_mini is not None and roi_dumper is not None:
        if roi_mini < roi_dumper:
            verdict_score_mini += 1
        else:
            verdict_score_dumper += 1

    if x1_opt > x2_opt:
        verdict_score_mini += 1
    elif x2_opt > x1_opt:
        verdict_score_dumper += 1

    if best_d1 > best_d2:
        verdict_score_mini += 1
    elif best_d2 > best_d1:
        verdict_score_dumper += 1

    # --- Doplňkové bodové hodnocení ---
    rows = score_rows(expected_profit_mini, expected_profit_dumper, prob_loss_mini, prob_loss_dumper,
                      var95_mini, var95_dumper, roi_mini, roi_dumper)
    scorecard_mini = sum(1 for *_, best in rows if best == 0)
    scorecard_dumper = sum(1 for *_, best in rows if best == 1)

    # --- Citlivost: křivka bodu zvratu (jen fixní náklad) ---
    days_grid = np.arange(0, 32)
    profit_curve_mini = avg_daily_mini * days_grid - fixed_cost
    profit_curve_dumper = avg_daily_dumper * days_grid - fixed_cost

    # --- Citlivost vstupů (±10 %) ---
    sensitivity_data = sensitivity(avg_daily_mini, avg_daily_dumper, row_1, row_2, fixed_cost,
                                   expected_days, machine_1, machine_2)
    labels = [x[0] for x in sensitivity_data]
    values = [x[1] for x in sensitivity_data]

    # --- Riziko × výnos ---
    rr_x1 = prob_loss_mini * 100
    rr_x2 = prob_loss_dumper * 100
    rr_y1 = expected_profit_mini
    rr_y2 = expected_profit_dumper

    # --- Mapa ziskovosti (stroj 1, jen fixní náklad) ---
    days_range = np.arange(5, 31, 1)
    price_range = np.arange(
        avg_daily_mini * 0.6, avg_daily_mini * 1.4, avg_daily_mini * 0.05
    )
    profit_matrix = np.zeros((len(price_range), len(days_range)))

    for i, price in enumerate(price_range):
        for j, days in enumerate(days_range):
            profit_matrix[i, j] = price * days - fixed_cost

    # --- Minimální cena pronájmu ---
    required_price_1 = (fixed_cost + row_1["service_cost_mean"]) / max(expected_days, 1)
    required_price_2 = (fixed_cost + row_2["service_cost_mean"]) / max(expected_days, 1)

    return dict(
        sim_1=sim_1, sim_2=sim_2, profit_mini=profit_mini, profit_dumper=profit_dumper,
        expected_profit_mini_real=expected_profit_mini_real, expected_profit_dumper_real=expected_profit_dumper_real,
        prob_loss_mini=prob_loss_mini, prob_loss_dumper=prob_loss_dumper,
        risk_loss_mini=risk_loss_mini, risk_loss_dumper=risk_loss_dumper,
        break_even_mini=break_even_mini, break_even_dumper=break_even_dumper,
        expected_profit_mini=expected_profit_mini, expected_profit_dumper=expected_profit_dumper,
        p10_mini=p10_mini, p10_dumper=p10_dumper, p50_mini=p50_mini, p50_dumper=p50_dumper,
        p90_mini=p90_mini, p90_dumper=p90_dumper, var95_mini=var95_mini, var95_dumper=var95_dumper,
        ci_mini=ci_mini, ci_dumper=ci_dumper, prob_dumper_better=prob_dumper_better,
        investment_cost_mini=investment_cost_mini, investment_cost_dumper=investment_cost_dumper,
        roi_mini=roi_mini, roi_dumper=roi_dumper, profit_diff=profit_diff, risk_diff=risk_diff,
        better_machine=better_machine, profit_mini_day=profit_mini_day, profit_dumper_day=profit_dumper_day,
        **opt, mc_better=mc_better, lp_better=lp_better, dp_better=dp_better,
        verdict_score_mini=verdict_score_mini, verdict_score_dumper=verdict_score_dumper,
        score_rows=rows, scorecard_mini=scorecard_mini, scorecard_dumper=scorecard_dumper,
        days_grid=days_grid, profit_curve_mini=profit_curve_mini, profit_curve_dumper=profit_curve_dumper,
        drivers=DRIVERS, sensitivity_data=sensitivity_data, labels=labels, values=values,
        rr_x1=rr_x1, rr_x2=rr_x2, rr_y1=rr_y1, rr_y2=rr_y2,
        days_range=days_range, price_range=price_range, profit_matrix=profit_matrix,
        required_price_1=required_price_1, required_price_2=required_price_2,
    )
