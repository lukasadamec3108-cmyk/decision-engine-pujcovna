"""Plán zakázek (záložka „Plán zakázek“). Čisté funkce bez Streamlitu.

Model:
- Den 1 je první den plánu, všechny zakázky mohou začít hned (nejsou známá jiná data).
- Zakázka trvá `duration_days` po sobě jdoucích dní a po celou dobu potřebuje všechny
  své stroje (`required_machines`). Typ stroje má `capacity` kusů – v žádném dni nesmí
  běžet víc zakázek se stejným strojem, než je kusů.
- Zpoždění = max(0, poslední den zakázky − termín); penále = zpoždění × denní penále.
- Tržba, náklady (`cost_base`) a zisk zakázky nezávisí na pořadí. Plánují-li se všechny
  zakázky, je proto maximalizace zisku po penále totéž jako minimalizace penále.

Plán se hledá celočíselnou optimalizací (časově indexovaný model, SciPy `milp` / HiGHS).
Pokud řešič v časovém limitu optimum neprokáže, vrátí se nejlepší nalezený plán
(nebo plán z jednoduchého pravidla) a výsledek se tak také označí.
"""

import numpy as np
import pandas as pd
from scipy.optimize import LinearConstraint, milp
from scipy.sparse import lil_matrix


def machines_of(value) -> list:
    return [m.strip() for m in str(value).split(";") if m.strip()]


def missing_machines(jobs: pd.DataFrame, capacity: dict) -> dict:
    """Zakázky, které nelze naplánovat, protože některý stroj chybí (kapacita 0)."""
    out = {}
    for idx, row in jobs.iterrows():
        miss = [m for m in machines_of(row["required_machines"]) if capacity.get(m, 0) <= 0]
        if miss:
            out[idx] = miss
    return out


# ---------------------------------------------------------------------------
# Vyhodnocení a kontrola plánu
# ---------------------------------------------------------------------------

def evaluate(jobs: pd.DataFrame, starts: dict) -> pd.DataFrame:
    """Tabulka plánu. `starts` = {index zakázky: první den (1, 2, …)}; chybějící = nenaplánováno."""
    rows = []
    for idx, row in jobs.iterrows():
        revenue = float(row["expected_revenue"])
        costs = float(row["cost_base"])
        rec = dict(
            job_index=idx,
            job_name=row["job_name"],
            machines=machines_of(row["required_machines"]),
            duration=int(row["duration_days"]),
            deadline=int(row["deadline_days"]),
            penalty_per_day=float(row["delay_penalty_per_day"]),
            priority=row["priority"],
            risk_level=row["risk_level"],
            revenue=revenue,
            costs=costs,
            profit=revenue - costs,
        )
        if idx in starts:
            start = int(starts[idx])
            end = start + rec["duration"] - 1
            late = max(0, end - rec["deadline"])
            rec.update(scheduled=True, start_day=start, end_day=end, days_late=late,
                       penalty=late * rec["penalty_per_day"])
        else:
            rec.update(scheduled=False, start_day=None, end_day=None, days_late=0, penalty=0.0)
        rec["profit_after_penalty"] = rec["profit"] - rec["penalty"] if rec["scheduled"] else 0.0
        rows.append(rec)
    return pd.DataFrame(rows).set_index("job_index")


def summary(plan: pd.DataFrame) -> dict:
    done = plan[plan["scheduled"]]
    return dict(
        scheduled=int(len(done)),
        waiting=int((~plan["scheduled"]).sum()),
        late=int((done["days_late"] > 0).sum()),
        penalty=float(done["penalty"].sum()),
        revenue=float(done["revenue"].sum()),
        costs=float(done["costs"].sum()),
        profit=float(done["profit"].sum()),
        profit_after_penalty=float(done["profit"].sum() - done["penalty"].sum()),
        length=int(done["end_day"].max()) if len(done) else 0,
    )


def conflicts(plan: pd.DataFrame, capacity: dict) -> list:
    """Dny, kdy je stroj přidělen více zakázkám, než má kusů: [(stroj, den, použito, kapacita)]."""
    use = {}
    for _, r in plan[plan["scheduled"]].iterrows():
        for day in range(int(r["start_day"]), int(r["end_day"]) + 1):
            for m in r["machines"]:
                use[(m, day)] = use.get((m, day), 0) + 1
    return sorted(
        (m, day, n, capacity.get(m, 0)) for (m, day), n in use.items() if n > capacity.get(m, 0)
    )


def utilization(plan: pd.DataFrame, capacity: dict) -> pd.DataFrame:
    """Vytížení = obsazené dny stroje ÷ (počet kusů × délka plánu)."""
    length = summary(plan)["length"]
    rows = []
    for m, cap in capacity.items():
        used = int(sum(r["duration"] for _, r in plan[plan["scheduled"]].iterrows() if m in r["machines"]))
        share = used / (cap * length) * 100 if cap > 0 and length > 0 else 0.0
        rows.append(dict(machine=m, units=cap, used_days=used, utilization=share))
    return pd.DataFrame(rows).sort_values("utilization", ascending=False)


# ---------------------------------------------------------------------------
# Jednoduché pravidlo (záloha, když řešič nedoběhne)
# ---------------------------------------------------------------------------

def greedy_schedule(jobs: pd.DataFrame, capacity: dict) -> dict:
    """Zakázky podle termínu (při shodě vyšší penále), každá do nejdřívějšího volného okna."""
    skip = missing_machines(jobs, capacity)
    order = jobs.drop(index=list(skip)).sort_values(
        ["deadline_days", "delay_penalty_per_day"], ascending=[True, False]
    ).index
    use, starts = {}, {}
    for idx in order:
        row = jobs.loc[idx]
        dur, need = int(row["duration_days"]), machines_of(row["required_machines"])
        start = 1
        while any(use.get((m, d), 0) >= capacity[m] for d in range(start, start + dur) for m in need):
            start += 1
        for d in range(start, start + dur):
            for m in need:
                use[(m, d)] = use.get((m, d), 0) + 1
        starts[idx] = start
    return starts


# ---------------------------------------------------------------------------
# Celočíselná optimalizace
# ---------------------------------------------------------------------------

def optimize_schedule(jobs: pd.DataFrame, capacity: dict, time_limit: float = 20.0) -> dict:
    """Plán s nejnižším celkovým penále; při shodě co nejdříve dokončené zakázky.

    Vrací dict(starts, optimal, method). `optimal=True` jen tehdy, když řešič prokázal
    optimum (nulová mezera) – v rámci výše popsaného modelu.
    """
    skip = missing_machines(jobs, capacity)
    idxs = [i for i in jobs.index if i not in skip]
    if not idxs:
        return dict(starts={}, optimal=True, method="nic k naplánování")

    dur = {i: int(jobs.loc[i, "duration_days"]) for i in idxs}
    # Existuje optimální plán bez dne, kdy nic neběží, takže stačí horizont = součet délek.
    horizon = sum(dur.values())
    var = [(i, s) for i in idxs for s in range(1, horizon - dur[i] + 2)]
    n = len(var)

    late_cost = np.array([
        max(0, s + dur[i] - 1 - int(jobs.loc[i, "deadline_days"]))
        * float(jobs.loc[i, "delay_penalty_per_day"])
        for i, s in var
    ])
    finish = np.array([s + dur[i] - 1 for i, s in var], dtype=float)

    # Každá zakázka začne právě jednou.
    one = lil_matrix((len(idxs), n))
    pos = {i: k for k, i in enumerate(idxs)}
    for k, (i, s) in enumerate(var):
        one[pos[i], k] = 1
    # Kapacita strojů v každém dni.
    rows = [(m, d) for m in capacity for d in range(1, horizon + 1)]
    rpos = {r: k for k, r in enumerate(rows)}
    cap = lil_matrix((len(rows), n))
    for k, (i, s) in enumerate(var):
        for m in machines_of(jobs.loc[i, "required_machines"]):
            for d in range(s, s + dur[i]):
                cap[rpos[(m, d)], k] = 1
    cap_ub = np.array([capacity[m] for m, _ in rows], dtype=float)
    base = [
        LinearConstraint(one.tocsr(), 1, 1),
        LinearConstraint(cap.tocsr(), -np.inf, cap_ub),
    ]
    opts = dict(time_limit=time_limit, mip_rel_gap=0)

    res1 = milp(late_cost, constraints=base, integrality=np.ones(n), bounds=(0, 1), options=opts)
    if res1.x is None:
        return dict(starts=greedy_schedule(jobs, capacity), optimal=False,
                    method="pravidlo podle termínu (řešič nedoběhl)")
    best_penalty = float(late_cost @ np.round(res1.x))
    optimal = res1.status == 0

    # 2. krok: stejné (nejnižší) penále, ale zakázky co nejdříve – kompaktní plán.
    res2 = milp(
        finish,
        constraints=base + [LinearConstraint(late_cost.reshape(1, -1), -np.inf, best_penalty + 0.5)],
        integrality=np.ones(n), bounds=(0, 1), options=opts,
    )
    x = np.round(res2.x if res2.x is not None else res1.x)
    starts = {i: s for (i, s), v in zip(var, x) if v > 0.5}
    return dict(starts=starts, optimal=optimal,
                method="celočíselná optimalizace" + ("" if optimal else " (optimum neprokázáno)"))


def job_contributions(jobs: pd.DataFrame, capacity: dict, base_penalty: float,
                      time_limit: float = 10.0) -> pd.DataFrame:
    """Přínos každé zakázky pro celý plán.

    přínos = zisk zakázky − (nejnižší penále plánu se zakázkou − nejnižší penále bez ní).
    Rozdíl penále zahrnuje zpoždění zakázky samotné i zpoždění, které způsobí ostatním.
    """
    skip = missing_machines(jobs, capacity)
    rows = []
    for idx in jobs.index:
        if idx in skip:
            continue
        rest = jobs.drop(index=idx)
        result = optimize_schedule(rest, capacity, time_limit)
        pen_without = summary(evaluate(rest, result["starts"]))["penalty"]
        profit = float(jobs.loc[idx, "expected_revenue"] - jobs.loc[idx, "cost_base"])
        added = base_penalty - pen_without
        rows.append(dict(job_index=idx, job_name=jobs.loc[idx, "job_name"], profit=profit,
                         added_penalty=added, contribution=profit - added,
                         exact=result["optimal"]))
    return pd.DataFrame(rows).sort_values(["contribution", "profit"], ascending=False)
