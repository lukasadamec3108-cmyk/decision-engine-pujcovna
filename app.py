import os
from io import BytesIO
from itertools import product

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
from docx import Document
from docx.shared import Inches
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from scipy.optimize import linprog

# ==================================================
# POMOCNÉ FUNKCE
# ==================================================


def format_czk(value) -> str:
    """Naformátuje částku s mezerou jako oddělovačem tisíců a jednotkou Kč."""
    return f"{value:,.0f}".replace(",", " ") + " Kč"


def risk_light(prob_loss: float) -> str:
    """Vrátí textový semafor podle pravděpodobnosti ztráty."""
    if prob_loss < 0.05:
        return "🟢 Nízké riziko"
    if prob_loss < 0.15:
        return "🟡 Střední riziko"
    return "🔴 Vysoké riziko"


def build_mgmt_text(machine_1, machine_2, expected_profit_mini, expected_profit_dumper) -> str:
    """Sestaví shrnutí pro management podle toho, který stroj má vyšší očekávaný zisk."""
    if expected_profit_dumper > expected_profit_mini:
        return (
            f"Pro management aktuálně vychází lépe {machine_2}. "
            f"Nabízí vyšší očekávaný zisk, nižší riziko ztráty "
            f"a lepší bod zvratu než {machine_1}. "
            f"Přesto je investice citlivá na cenu pronájmu a vytížení."
        )
    elif expected_profit_mini > expected_profit_dumper:
        return (
            f"Pro management aktuálně vychází lépe {machine_1}. "
            f"Nabízí lepší kombinaci očekávaného zisku, rizika "
            f"a návratnosti než {machine_2}."
        )
    else:
        return (
            "Obě varianty vycházejí velmi podobně. "
            "Pro finální rozhodnutí doporučujeme doplnit provozní data."
        )


@st.cache_data
def load_data() -> pd.DataFrame:
    """Načtení demonstračních dat o strojích."""
    return pd.read_csv("data/machines_demo.csv")


# ==================================================
# KONFIGURACE APLIKACE A PŘÍSTUP
# ==================================================

st.set_page_config(page_title="Decision Engine – Půjčovna", layout="centered")

PASSWORD = st.secrets["password"]
password = st.text_input("Zadejte přístupové heslo", type="password")
if password != PASSWORD:
    st.stop()


# ==================================================
# HLAVIČKA APLIKACE
# ==================================================

st.title("Decision Engine – Půjčovna malostrojů")
st.caption(
    "Interaktivní nástroj pro porovnání investic do stavebních strojů "
    "pomocí simulace nejistoty (Monte Carlo) a optimalizačního modelu (LP, DP)."
)
st.write(
    "Porovnejte dva stroje z hlediska zisku, rizika a návratnosti "
    "při různých scénářích poptávky."
)


# ==================================================
# DATA A SIDEBAR VSTUPY
# ==================================================

df = load_data()

jobs_df = pd.read_csv("data/jobs_demo.csv")

risk_multiplier = {"Low": 0.95, "Medium": 0.85, "High": 0.70}

priority_score = {"High": 3, "Medium": 2, "Low": 1}
risk_score = {"Low": 0, "Medium": 1, "High": 2}

jobs_df["expected_profit"] = jobs_df["expected_revenue"] - jobs_df["cost_base"]

jobs_df["risk_adjusted_profit"] = jobs_df.apply(
    lambda row: row["expected_profit"] * risk_multiplier[row["risk_level"]], axis=1
)

jobs_df["priority_score"] = jobs_df["priority"].map(priority_score)

available_machines = set(df["machine_type"].tolist())

st.sidebar.subheader("Kapacitní scénáře")

extra_dumper = st.sidebar.checkbox("Koupit druhý pásový dumper")

machine_capacity = {
    "Mini bagr": 1,
    "Pásový dumper": 1,
    "Smykový nakladač": 1,
    "Vibrační deska": 2,
    "Kolový bagr": 1,
    "Silniční válec": 1,
    "Řezačka asfaltu": 1,
    "Nákladní auto": 1,
}

if extra_dumper:
    machine_capacity["Pásový dumper"] = 2


def check_machine_availability(required_machines):
    required = [m.strip() for m in required_machines.split(";")]
    missing = [m for m in required if m not in available_machines]

    if missing:
        return "Chybí: " + ", ".join(missing)
    else:
        return "OK"


def build_job_schedule(jobs):
    machine_busy_until = {machine: 0 for machine in machine_capacity.keys()}
    scheduled_rows = []

    for _, row in jobs.iterrows():
        required = [m.strip() for m in row["required_machines"].split(";")]
        duration = int(row["duration_days"])

        # Pokud stroj není v kapacitním seznamu, považujeme ho za nedostupný
        missing_machines = [
            machine for machine in required if machine not in machine_busy_until
        ]

        row_data = row.copy()

        if missing_machines:
            row_data["schedule_status"] = "🔴 Čeká"
            row_data["schedule_note"] = "Chybí: " + ", ".join(missing_machines)
            row_data["start_day"] = None
            row_data["end_day"] = None
        else:
            # Zakázka začne až ve chvíli, kdy jsou volné všechny potřebné stroje
            start_day = max(machine_busy_until[machine] for machine in required)
            end_day = start_day + duration

            row_data["schedule_status"] = "🟢 Naplánováno"
            row_data["schedule_note"] = (
                f"Start den {start_day + 1}, konec den {end_day}"
            )
            row_data["start_day"] = start_day + 1
            row_data["end_day"] = end_day

            # Všechny použité stroje jsou obsazené do konce zakázky
            for machine in required:
                machine_busy_until[machine] = end_day

        scheduled_rows.append(row_data)

    return pd.DataFrame(scheduled_rows)


jobs_df["machine_status"] = jobs_df["required_machines"].apply(
    check_machine_availability
)

jobs_df["machine_ready_score"] = jobs_df["machine_status"].apply(
    lambda x: 1 if x == "OK" else 0
)

jobs_df["decision_score"] = (
    jobs_df["risk_adjusted_profit"]
    + jobs_df["priority_score"] * 20000
    + jobs_df["machine_ready_score"] * 30000
    - jobs_df["delay_penalty_per_day"] * 2
)

recommended_jobs = jobs_df.sort_values("decision_score", ascending=False)

scheduled_jobs = build_job_schedule(recommended_jobs)

machine_options = df["machine_type"].tolist()

machine_1 = st.sidebar.selectbox("Stroj 1", machine_options, index=0)
machine_2 = st.sidebar.selectbox("Stroj 2", machine_options, index=1)

row_1 = df[df["machine_type"] == machine_1].iloc[0]
row_2 = df[df["machine_type"] == machine_2].iloc[0]

avg_daily_mini = float(row_1["daily_price"])
avg_daily_dumper = float(row_2["daily_price"])

st.sidebar.header("Nastavení simulace")

fixed_cost = st.sidebar.number_input(
    "Fixní náklad (Kč / měsíc)",
    min_value=0,
    value=120000,
    step=5000,
)

scenario = st.sidebar.selectbox(
    "Scénář trhu",
    ["Pesimistický", "Realistický", "Optimistický"],
)

if scenario == "Pesimistický":
    min_days, max_days = 8, 15
    scenario_note = "Slabší poptávka, opatrný scénář."
elif scenario == "Realistický":
    min_days, max_days = 15, 25
    scenario_note = "Běžný provoz firmy."
else:
    min_days, max_days = 22, 30
    scenario_note = "Silná poptávka a vysoké vytížení."

st.sidebar.caption(
    f"{scenario_note} Rozmezí vytížení: {min_days} až {max_days} dní / měsíc."
)

expected_days = st.sidebar.slider("Odhad reálné poptávky (dny/měsíc)", 0, 31, 18)
n = st.sidebar.slider("Počet simulací (měsíců)", 100, 10000, 1000)

st.sidebar.subheader("Optimalizace portfolia")
budget = st.sidebar.number_input(
    "Investiční rozpočet (Kč)",
    min_value=100000,
    value=3000000,
    step=100000,
)


# ==================================================
# MONTE CARLO SIMULACE
# ==================================================

# --- Stroj 1 ---
demand_mean_1 = row_1["demand_days_mean"]
mode_days_1 = np.clip(demand_mean_1, min_days, max_days)
util_days_1 = np.random.triangular(min_days, mode_days_1, max_days, n)

service_cost_1 = np.random.normal(
    row_1["service_cost_mean"],
    row_1["service_cost_std"],
    n,
)
service_cost_1 = np.clip(service_cost_1, 0, None)

failure_happens_1 = np.random.rand(n) < row_1["failure_prob"]
failure_days_1 = np.where(
    failure_happens_1,
    np.random.poisson(row_1["failure_days_mean"], n),
    0,
)
real_days_1 = np.maximum(util_days_1 - failure_days_1, 0)

price_std_1 = avg_daily_mini * 0.08
daily_price_1 = np.random.normal(avg_daily_mini, price_std_1, n)
daily_price_1 = np.clip(daily_price_1, avg_daily_mini * 0.8, avg_daily_mini * 1.2)

rev_mini = real_days_1 * daily_price_1
profit_mini = rev_mini - fixed_cost - service_cost_1

# --- Stroj 2 ---
demand_mean_2 = row_2["demand_days_mean"]
mode_days_2 = np.clip(demand_mean_2, min_days, max_days)
util_days_2 = np.random.triangular(min_days, mode_days_2, max_days, n)

service_cost_2 = np.random.normal(
    row_2["service_cost_mean"],
    row_2["service_cost_std"],
    n,
)
service_cost_2 = np.clip(service_cost_2, 0, None)

failure_happens_2 = np.random.rand(n) < row_2["failure_prob"]
failure_days_2 = np.where(
    failure_happens_2,
    np.random.poisson(row_2["failure_days_mean"], n),
    0,
)
real_days_2 = np.maximum(util_days_2 - failure_days_2, 0)

price_std_2 = avg_daily_dumper * 0.08
daily_price_2 = np.random.normal(avg_daily_dumper, price_std_2, n)
daily_price_2 = np.clip(daily_price_2, avg_daily_dumper * 0.8, avg_daily_dumper * 1.2)

rev_dumper = real_days_2 * daily_price_2
profit_dumper = rev_dumper - fixed_cost - service_cost_2


# ==================================================
# ODVOZENÉ METRIKY A ROZHODOVACÍ LOGIKA
# ==================================================

# --- Realistický scénář při očekávané poptávce ---
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

# --- ROI ---
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

# --- Executive summary metriky ---
profit_diff = expected_profit_dumper - expected_profit_mini
risk_diff = (prob_loss_mini - prob_loss_dumper) * 100
risk_diff_pp = risk_diff
be_diff = break_even_mini - break_even_dumper
risk_adv_pct = risk_diff

better_machine = (
    machine_2 if expected_profit_dumper > expected_profit_mini else machine_1
)

mgmt_text = build_mgmt_text(machine_1, machine_2, expected_profit_mini, expected_profit_dumper)

if (prob_loss_dumper < prob_loss_mini) and (break_even_dumper < break_even_mini):
    recommendation = f"✅ Investičně vychází lépe: {machine_2}"
else:
    recommendation = "⚠️ Nelze dát jednoznačné doporučení – doplň data"

if profit_diff > 0:
    profit_text = f"{machine_2} vydělá o {profit_diff:,.0f} Kč více než {machine_1}"
else:
    profit_text = (
        f"{machine_1} vydělá o {abs(profit_diff):,.0f} Kč více než {machine_2}"
    )

if risk_diff > 0:
    risk_text = f"{machine_2} má o {risk_diff:.1f} p.b. nižší riziko ztráty"
else:
    risk_text = f"{machine_1} má o {abs(risk_diff):.1f} p.b. nižší riziko ztráty"


# ==================================================
# OPTIMALIZAČNÍ MODELY – LP A DP
# ==================================================

profit_mini_day = avg_daily_mini - row_1["service_cost_mean"] / 30 - fixed_cost / 30
profit_dumper_day = avg_daily_dumper - row_2["service_cost_mean"] / 30 - fixed_cost / 30

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


# ==================================================
# SCORE MODELY
# ==================================================

# --- Finální investiční verdikt ---
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

# --- Investiční scorecard ---
scorecard_mini = 0
scorecard_dumper = 0

if expected_profit_mini > expected_profit_dumper:
    scorecard_mini += 1
else:
    scorecard_dumper += 1

if prob_loss_mini < prob_loss_dumper:
    scorecard_mini += 1
else:
    scorecard_dumper += 1

if var95_mini > var95_dumper:
    scorecard_mini += 1
else:
    scorecard_dumper += 1

if roi_mini is not None and roi_dumper is not None:
    if roi_mini < roi_dumper:
        scorecard_mini += 1
    else:
        scorecard_dumper += 1

# --- AI doporučení ---
if scorecard_dumper > scorecard_mini:
    ai_text = f"""
**Proč vychází lépe {machine_2}:**
- Očekávaný měsíční zisk: **{machine_2} {expected_profit_dumper:,.0f} Kč** vs. **{machine_1} {expected_profit_mini:,.0f} Kč**
- {machine_2} má **o {risk_adv_pct:.1f} p.b. nižší riziko ztráty**
- Break-even: **{machine_2} {break_even_dumper:.2f} dní** vs. **{machine_1} {break_even_mini:.2f} dní**

**Scénáře zisku (Kč) – aby bylo jasné riziko:**
- 10% pesimistický scénář: {machine_2} **{p10_dumper:,.0f}**, {machine_1} **{p10_mini:,.0f}**
- Medián: {machine_2} **{p50_dumper:,.0f}**, {machine_1} **{p50_mini:,.0f}**
- 90% optimistický scénář: {machine_2} **{p90_dumper:,.0f}**, {machine_1} **{p90_mini:,.0f}**

**Co je potřeba doplnit před finálním nákupem:**
- reálnou poptávku po {machine_2}
- servisní náklady / odstávky / pojištění
- sezónnost (léto vs. zima)
"""
elif scorecard_mini > scorecard_dumper:
    ai_text = f"""
**Proč vychází lépe {machine_1}:**
- Očekávaný měsíční zisk: **{machine_1} {expected_profit_mini:,.0f} Kč** vs. **{machine_2} {expected_profit_dumper:,.0f} Kč**
- {machine_1} má **o {abs(risk_adv_pct):.1f} p.b. nižší riziko ztráty**
- Break-even: **{machine_1} {break_even_mini:.2f} dní** vs. **{machine_2} {break_even_dumper:.2f} dní**

**Scénáře zisku (Kč) – aby bylo jasné riziko:**
- 10% pesimistický scénář: {machine_1} **{p10_mini:,.0f}**, {machine_2} **{p10_dumper:,.0f}**
- Medián: {machine_1} **{p50_mini:,.0f}**, {machine_2} **{p50_dumper:,.0f}**
- 90% optimistický scénář: {machine_1} **{p90_mini:,.0f}**, {machine_2} **{p90_dumper:,.0f}**

**Co je potřeba doplnit před finálním nákupem:**
- reálnou poptávku po {machine_1}
- servisní náklady / odstávky / pojištění
- sezónnost (léto vs. zima)
"""
else:
    ai_text = f"""
**Proč nejde dát jasné „kup“:**
- Rozdíl není dostatečně jednoznačný při zvolených parametrech.

**Rychlá fakta:**
- Očekávaný zisk {machine_2}: **{expected_profit_dumper:,.0f} Kč**
- Očekávaný zisk {machine_1}: **{expected_profit_mini:,.0f} Kč**
- Riziko ztráty {machine_2}: **{prob_loss_dumper * 100:.1f} %**
- Riziko ztráty {machine_1}: **{prob_loss_mini * 100:.1f} %**

**Další krok:**
Doplň data o poptávce, servisních nákladech a sezónnosti.
"""


# ==================================================
# CITLIVOST A PORTFOLIO – VÝPOČTY
# ==================================================

# --- Citlivost: break-even křivka ---
days_grid = np.arange(0, 32)
profit_curve_mini = avg_daily_mini * days_grid - fixed_cost
profit_curve_dumper = avg_daily_dumper * days_grid - fixed_cost

# --- Citlivost vstupů ---
base_profit_1 = avg_daily_mini * expected_days - fixed_cost - row_1["service_cost_mean"]
base_profit_2 = (
    avg_daily_dumper * expected_days - fixed_cost - row_2["service_cost_mean"]
)

sensitivity_data = []
drivers = [
    ("Denní cena", "price"),
    ("Vytížení (dny/měsíc)", "days"),
    ("Fixní náklad", "fixed"),
    ("Servisní náklad", "service"),
]

for label, driver in drivers:
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

sensitivity_data = sorted(sensitivity_data, key=lambda x: x[1], reverse=True)
labels = [x[0] for x in sensitivity_data]
values = [x[1] for x in sensitivity_data]

# --- Risk vs Return ---
rr_x1 = prob_loss_mini * 100
rr_x2 = prob_loss_dumper * 100
rr_y1 = expected_profit_mini
rr_y2 = expected_profit_dumper

# --- Heatmapa ---
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

# --- Optimalizace portfolia ---
machine_summary = []
for _, row in df.iterrows():
    expected_unit_profit = (
        row["daily_price"] * row["demand_days_mean"]
        - row["monthly_fixed_cost"]
        - row["service_cost_mean"]
    )
    machine_summary.append(
        {
            "machine_type": row["machine_type"],
            "purchase_price": row["purchase_price"],
            "expected_unit_profit": expected_unit_profit,
        }
    )

machines = sorted(
    machine_summary,
    key=lambda x: x["expected_unit_profit"],
    reverse=True,
)[:5]

best_profit = -1e18
best_combo = None
n_machines = len(machines)
max_units = 2

for counts in product(range(max_units + 1), repeat=n_machines):
    total_cost = sum(
        counts[i] * machines[i]["purchase_price"] for i in range(n_machines)
    )
    total_profit = sum(
        counts[i] * machines[i]["expected_unit_profit"] for i in range(n_machines)
    )

    if total_cost <= budget and total_profit > best_profit and sum(counts) > 0:
        best_profit = total_profit
        best_combo = counts

# --- Budget vs profit křivka ---
budget_values = np.arange(500000, 3000001, 500000)
max_profits = []

for test_budget in budget_values:
    best_profit_curve = -1e18

    for counts in product(range(3), repeat=n_machines):
        total_cost = sum(
            counts[i] * machines[i]["purchase_price"] for i in range(n_machines)
        )
        total_profit = sum(
            counts[i] * machines[i]["expected_unit_profit"] for i in range(n_machines)
        )

        if total_cost <= test_budget and total_profit > best_profit_curve:
            best_profit_curve = total_profit

    if best_profit_curve == -1e18:
        best_profit_curve = 0

    max_profits.append(best_profit_curve)

# Složka pro export grafů
os.makedirs("outputs", exist_ok=True)


# ==================================================
# EXPORT REPORTU – FUNKCE
# ==================================================


def fig_to_png_bytes(fig) -> bytes:
    img = BytesIO()
    fig.savefig(img, format="png", dpi=200, bbox_inches="tight")
    img.seek(0)
    return img.read()


def build_docx_report() -> bytes:
    doc = Document()
    
    # Základní písmo dokumentu
    normal_style = doc.styles["Normal"]
    normal_style.font.name = "Aptos"
    normal_style.font.size = Pt(10.5)

    # Nadpis 1
    heading_1 = doc.styles["Heading 1"]
    heading_1.font.name = "Aptos Display"
    heading_1.font.size = Pt(16)
    heading_1.font.bold = True
    heading_1.font.color.rgb = RGBColor(31, 78, 121)

    # Nadpis 2
    heading_2 = doc.styles["Heading 2"]
    heading_2.font.name = "Aptos Display"
    heading_2.font.size = Pt(13)
    heading_2.font.bold = True
    heading_2.font.color.rgb = RGBColor(31, 78, 121)

    # Nadpis 3, pokud ho používáš
    heading_3 = doc.styles["Heading 3"]
    heading_3.font.name = "Aptos"
    heading_3.font.size = Pt(11)
    heading_3.font.bold = True
    heading_3.font.color.rgb = RGBColor(31, 78, 121)

    title = doc.add_heading("Decision Engine – investiční report", level=1)

    title_run = title.runs[0]
    title_run.font.size = Pt(20)
    title_run.font.bold = True
    title_run.font.name = "Aptos Display"
    title_run.font.color.rgb = RGBColor(31, 78, 121)

    doc.add_paragraph(f"Porovnání: {machine_1} vs. {machine_2}")

    doc.add_heading("Shrnutí výsledků", level=1)
    doc.add_paragraph(f"Fixní náklad (Kč/měsíc): {fixed_cost:,.0f}")
    doc.add_paragraph(f"Rozsah vytížení: {min_days}-{max_days} dní/měsíc")
    doc.add_paragraph(f"Počet simulací: {n:,}")
    doc.add_paragraph(f"Odhad reálné poptávky: {expected_days} dní/měsíc")

    doc.add_heading("Klíčové investiční ukazatele", level=1)
    table = doc.add_table(rows=1, cols=3)
    hdr = table.rows[0].cells
    hdr[0].text = "Metrika"
    hdr[1].text = machine_1
    hdr[2].text = machine_2

    rows = [
        (
            "Průměrná denní cena",
            f"{avg_daily_mini:,.0f} Kč",
            f"{avg_daily_dumper:,.0f} Kč",
        ),
        ("Bod zvratu", f"{break_even_mini:.2f} dní", f"{break_even_dumper:.2f} dní"),
        ("Riziko ztráty", f"{risk_loss_mini:.1f} %", f"{risk_loss_dumper:.1f} %"),
        (
            "Oček. zisk (Monte Carlo)",
            f"{expected_profit_mini:,.0f} Kč",
            f"{expected_profit_dumper:,.0f} Kč",
        ),
    ]
    for metric, v1, v2 in rows:
        cells = table.add_row().cells
        cells[0].text = metric
        cells[1].text = v1
        cells[2].text = v2

    better_machine = (
    machine_1
    if expected_profit_mini > expected_profit_dumper
    else machine_2
    )

    doc.add_heading("Relativně výhodnější varianta", level=1)

    verdict = doc.add_paragraph()
    verdict.add_run("Relativně lépe vychází: ").bold = True
    verdict.add_run(f"{better_machine}.\n")

    if expected_profit_mini < 0 and expected_profit_dumper < 0:
        verdict.add_run(
            "Při aktuálním nastavení jsou však obě investiční varianty "
            "ztrátové. Investici proto nedoporučujeme realizovat bez zvýšení "
            "ceny pronájmu, vyššího vytížení nebo snížení nákladů."
        )
    else:
        verdict.add_run(
            f"{better_machine} nabízí příznivější kombinaci očekávaného "
            "výsledku, rizika ztráty a bodu zvratu."
        )

    doc.add_heading("Distribuce zisku (Monte Carlo)", level=1)

    doc.add_paragraph(
        "Následující rizikové metriky shrnují rozdělení možných scénářů "
        "měsíčního výsledku na základě Monte Carlo simulace."
    )
    if os.path.exists("outputs/porovnani_zisku.png"):
        pass
    # doc.add_picture("outputs/porovnani_zisku.png", width=Inches(6))

    doc.add_heading("Rizikové metriky", level=1)
    doc.add_paragraph(
        f"P10 (10 % nejhorších scénářů): {machine_1} = {p10_mini:,.0f} Kč, {machine_2} = {p10_dumper:,.0f} Kč"
    )
    doc.add_paragraph(
        f"Medián zisku: {machine_1} = {p50_mini:,.0f} Kč, {machine_2} = {p50_dumper:,.0f} Kč"
    )
    doc.add_paragraph(
        f"P90 (optimistický scénář): {machine_1} = {p90_mini:,.0f} Kč, {machine_2} = {p90_dumper:,.0f} Kč"
    )

    price_heading = doc.add_heading("Doporučená minimální cena pronájmu",level=1)

    price_heading.paragraph_format.page_break_before = True

    price_table = doc.add_table(rows=1, cols=3)
    price_table.style = "Table Grid"

    header_cells = price_table.rows[0].cells
    header_cells[0].text = "Stroj"
    header_cells[1].text = "Minimální cena"
    header_cells[2].text = "Hodnocení aktuální ceny"

    row = price_table.add_row().cells
    row[0].text = str(machine_1)
    row[1].text = f"{required_price_1:,.0f} Kč/den"
    row[2].text = (
        "Cena je dostačující"
        if avg_daily_mini >= required_price_1
        else "Cena je pod minimem pro zisk"
    )

    row = price_table.add_row().cells
    row[0].text = str(machine_2)
    row[1].text = f"{required_price_2:,.0f} Kč/den"
    row[2].text = (
        "Cena je dostačující"
        if avg_daily_dumper >= required_price_2
        else "Cena je pod minimem pro zisk"
    )

    doc.add_paragraph(
        "Minimální cena ukazuje, při jaké denní sazbě se při zadané "
        "poptávce pokryjí fixní a servisní náklady."
    )

    doc.add_heading("Interpretace výsledků", level=1)
    doc.add_paragraph(
        "Model simuluje možné scénáře poptávky po strojích a zohledňuje náklady na servis, "
        "fixní náklady a pravděpodobnost poruch. Výsledkem je odhad očekávaného měsíčního "
        "zisku a rizika ztráty."
    )
    doc.add_paragraph(
        "Mapa ziskovosti ukazuje, při jaké kombinaci ceny pronájmu a počtu pronajatých dní "
        "se investice dostává do zisku."
    )
    doc.add_paragraph(
        "Optimalizační model testuje možné kombinace nákupu strojů a vybírá portfolio, "
        "které maximalizuje očekávaný měsíční zisk při daném investičním rozpočtu."
    )

    doc.add_heading("Doporučený další postup", level=1)

    steps = [
        f"Ověřit skutečnou měsíční poptávku po stroji {better_machine}.",
        "Porovnat aktuální cenu pronájmu s vypočtenou minimální denní sazbou.",
        "Doplnit reálné servisní náklady, odstávky, pojištění a sezónnost.",
        "Znovu vyhodnotit investici v realistickém scénáři.",
    ]

    for step in steps:
        doc.add_paragraph(step, style="List Number")

    doc.add_heading("Závěr pro management", level=1)

    better_machine = (
        machine_1
        if expected_profit_mini > expected_profit_dumper
        else machine_2
    )

    if expected_profit_mini < 0 and expected_profit_dumper < 0:
        final_report_text = (
            f"Z porovnávaných variant vychází relativně lépe {better_machine}. "
            "Při aktuálním nastavení jsou však obě varianty ztrátové. "
            "Investici proto nedoporučujeme realizovat bez úpravy ceny pronájmu, "
            "zvýšení vytížení nebo snížení fixních a servisních nákladů."
        )
    else:
        final_report_text = (
            f"Z porovnávaných variant vychází lépe {better_machine}. "
            "Nabízí příznivější kombinaci očekávaného výsledku, rizika ztráty "
            "a bodu zvratu."
        )

    doc.add_paragraph(final_report_text)

    doc.add_heading("Podmínky ziskovosti", level=1)
    doc.add_paragraph(
        "Investice by byla výrazně atraktivnější při splnění některých z následujících podmínek:"
    )
    doc.add_paragraph(
        "- zvýšení denní ceny pronájmu\n"
        "- vyšší vytížení strojů (více dní pronájmu za měsíc)\n"
        "- snížení fixních nákladů nebo servisních nákladů"
    )
    
    out = BytesIO()
    doc.save(out)
    out.seek(0)
    return out.read()


# ==================================================
# TABS
# ==================================================

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
    [
        "📊 Shrnutí",
        "📈 Monte Carlo",
        "⚠ Riziko",
        "📉 Citlivost",
        "🗺 Optimalizace",
        "📋 Zakázky",
    ]
)


# ==================================================
# TAB 1 – SHRNUTÍ
# ==================================================

with tab1:
    st.subheader("Shrnutí pro management")

    row1_col1, row1_col2 = st.columns(2)

    with row1_col1:
        st.metric(
            f"Očekávaný zisk ({machine_1})",
            f"{expected_profit_mini:,.0f} Kč",
        )

    with row1_col2:
        st.metric(
            f"Očekávaný zisk ({machine_2})",
            f"{expected_profit_dumper:,.0f} Kč",
        )

    row2_col1, row2_col2 = st.columns(2)

    with row2_col1:
        st.metric(
            f"Pravděpodobnost ztráty ({machine_2})",
            f"{risk_loss_dumper:.1f} %",
        )

    with row2_col2:
        st.metric(
            f"Bod zvratu ({machine_2})",
            f"{break_even_dumper:.2f} dní",
        )

    if break_even_mini is not None and break_even_dumper is not None:
        break_even_diff = break_even_dumper - break_even_mini
        break_even_diff_text = f"{abs(break_even_diff):.2f} dne"
    else:
        break_even_diff_text = "Nelze spočítat"

    st.markdown(
        """
        <style>
        div[data-testid="stMetricLabel"] {
            overflow: visible !important;
            min-height: 1.8rem !important;
        }

        div[data-testid="stMetricLabel"] p {
            white-space: normal !important;
            overflow: visible !important;
            text-overflow: clip !important;
            line-height: 1.1 !important;
            font-size: 0.82rem !important;
            margin-bottom: 0 !important;
        }

        div[data-testid="stMetricValue"] {
            overflow: visible !important;
        }

        div[data-testid="stMetricValue"] > div {
            white-space: normal !important;
            overflow: visible !important;
            text-overflow: clip !important;
            line-height: 1.05 !important;
            font-size: 1.55rem !important;
        }

        div[data-testid="stMetric"] {
            overflow: visible !important;
            min-height: auto !important;
            padding-bottom: 0.4rem !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("## 🏆 Preferovaná investiční varianta")
    
    if expected_profit_mini < 0 and expected_profit_dumper < 0:
        st.warning(
            "Při aktuálním nastavení jsou obě investiční varianty ztrátové. "
            "Zvažte zvýšení ceny pronájmu, vyšší vytížení nebo snížení nákladů."
        )
    st.info("Největší vliv na zisk mají denní cena pronájmu a počet dní využití.")

    col_dec1, col_dec2, col_dec3, col_dec4 = st.columns(
        [1.65, 1.15, 1.05, 1.25],
        gap="medium",
    )

    with col_dec1:
        st.metric("Vybraný stroj", better_machine)

    with col_dec2:
        st.metric(
            "Rozdíl v zisku",
            f"{abs(profit_diff):,.0f} Kč",
        )

    with col_dec3:
        st.metric(
            "Rozdíl v riziku",
            f"{abs(risk_diff_pp):.1f} p.b.",
        )

    with col_dec4:
        st.metric(
            "Rozdíl v bodu zvratu",
            break_even_diff_text,
        )

    st.success(f"Investičně vychází lépe: {better_machine}")

    mgmt_text = build_mgmt_text(machine_1, machine_2, expected_profit_mini, expected_profit_dumper)
    st.info(mgmt_text)

    show_roi = False

    if show_roi and roi_mini is not None and roi_dumper is not None:

        st.subheader("💰 Návratnost investice")

        col_roi1, col_roi2 = st.columns(2)

        with col_roi1:
            st.metric(
                machine_1,
                f"{roi_mini:.1f} měsíců",
            )

        with col_roi2:
            st.metric(
                machine_2,
                f"{roi_dumper:.1f} měsíců",
            )

        if roi_dumper < roi_mini:
            st.success(f"{machine_2} se vrátí rychleji než {machine_1}.")
        else:
            st.success(f"{machine_1} se vrátí rychleji než {machine_2}.")

    with st.expander("🏁 Detailní investiční scorecard"):
        col_s1, col_s2 = st.columns(2)

        with col_s1:
            st.metric(
                f"Score – {machine_1}",
                f"{scorecard_mini} / 4",
            )

        with col_s2:
            st.metric(
                f"Score – {machine_2}",
                f"{scorecard_dumper} / 4",
            )

        if scorecard_mini > scorecard_dumper:
            st.info(f"Scorecard favorizuje {machine_1}.")

        elif scorecard_dumper > scorecard_mini:
            st.info(f"Scorecard favorizuje {machine_2}.")

        else:
            st.info("Scorecard vychází nerozhodně.")

    st.subheader("Realistický scénář při očekávané poptávce")
    col3, col4 = st.columns(2)
    with col3:
        st.metric(
            f"{machine_1} 📊 očekávaný zisk", f"{expected_profit_mini_real:,.0f} Kč"
        )
    with col4:
        st.metric(
            f"{machine_2} 📊 očekávaný zisk", f"{expected_profit_dumper_real:,.0f} Kč"
        )

    if expected_profit_mini_real < 0 and expected_profit_dumper_real < 0:
        st.warning(
            "Při aktuálních parametrech jsou oba stroje ztrátové. "
            "Zvažte vyšší poptávku, nižší fixní náklady nebo vyšší cenu pronájmu."
        )
    elif expected_profit_mini_real < 0 or expected_profit_dumper_real < 0:
        st.info(
            "Jeden ze strojů je při aktuálních parametrech ztrátový. "
            "Doporučuje se porovnat scénáře poptávky a fixních nákladů."
        )

    with st.expander("🧠 Detailní vysvětlení rozhodnutí"):
        st.markdown(ai_text)

    
    st.subheader("📋 Doporučený další postup")

    st.markdown(
    f"""
    1. Ověřit skutečnou měsíční poptávku po stroji **{better_machine}**.
    2. Porovnat aktuální cenu pronájmu s vypočtenou minimální denní sazbou.
    3. Doplnit reálné servisní náklady, odstávky, pojištění a sezónnost.
    4. Znovu vyhodnotit investici v realistickém scénáři.
    """
    )

    st.markdown("---")

    st.subheader("📄 Export investičního reportu")

    docx_file = build_docx_report()

    st.download_button(
        label="📄 Stáhnout report (DOCX)",
        data=docx_file,
        file_name="decision_report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


# ==================================================
# TAB 2 – MONTE CARLO
# ==================================================

with tab2:
    st.subheader("Klíčové metriky (Monte Carlo)")
    col1, col2 = st.columns(2)

    st.caption(
    "Simulace ukazuje rozdělení možných měsíčních výsledků při nejisté poptávce "
    "a nákladech. P10 představuje nepříznivý scénář, medián typický výsledek "
    "a P90 příznivý scénář."
    )

    with col1:
        st.markdown(f"### {machine_1}")
        st.metric("Očekávaný zisk", f"{expected_profit_mini:,.0f} Kč")
        st.metric("Medián zisku", f"{p50_mini:,.0f} Kč")
        st.metric("P10", f"{p10_mini:,.0f} Kč")
        st.metric("P90", f"{p90_mini:,.0f} Kč")
        st.metric("Pravděpodobnost ztráty", f"{risk_loss_mini:.1f} %")

    with col2:
        st.markdown(f"### {machine_2}")
        st.metric("Očekávaný zisk", f"{expected_profit_dumper:,.0f} Kč")
        st.metric("Medián zisku", f"{p50_dumper:,.0f} Kč")
        st.metric("P10", f"{p10_dumper:,.0f} Kč")
        st.metric("P90", f"{p90_dumper:,.0f} Kč")
        st.metric("Pravděpodobnost ztráty", f"{risk_loss_dumper:.1f} %")

    st.subheader("Porovnání distribuce zisku (Monte Carlo)")
    fig_mc = plt.figure()
    plt.hist(profit_mini, bins=25, alpha=0.5, label=machine_1)
    plt.hist(profit_dumper, bins=25, alpha=0.5, label=machine_2)
    plt.axvline(0)
    
    plt.axvline(p50_mini, linestyle=":")
    plt.axvline(p50_dumper, linestyle=":")
   
    plt.xlabel("Zisk / ztráta (Kč)")
    plt.ylabel("Počet simulovaných měsíců")
    plt.title("Distribuce zisku")
    plt.legend()
    fig_mc.savefig("outputs/porovnani_zisku.png", dpi=200, bbox_inches="tight")
    st.pyplot(fig_mc)

    st.caption(
    "Tečkované čáry označují medián simulovaného měsíčního výsledku jednotlivých strojů. "
    "Plná svislá čára označuje hranici nulového zisku."
    )
    plt.close(fig_mc)

    st.subheader("Kumulativní pravděpodobnost výsledků (CDF)")
    fig_cdf, ax_cdf = plt.subplots(figsize=(10, 6))
    sorted_mini = np.sort(profit_mini)
    sorted_dumper = np.sort(profit_dumper)
    cdf_mini = np.arange(1, len(sorted_mini) + 1) / len(sorted_mini)
    cdf_dumper = np.arange(1, len(sorted_dumper) + 1) / len(sorted_dumper)

    ax_cdf.plot(sorted_mini, cdf_mini, label=machine_1)
    ax_cdf.plot(sorted_dumper, cdf_dumper, label=machine_2)
    ax_cdf.axvline(0, color="black", linestyle="--", linewidth=1)
    ax_cdf.text(
    1500,
    0.05,
    "0 Kč – hranice zisku",
    rotation=90,
    va="bottom")
    
    ax_cdf.set_title("Kumulativní pravděpodobnost výsledků")
    ax_cdf.set_xlabel("Zisk / ztráta (Kč)")
    ax_cdf.set_ylabel("Pravděpodobnost")
    ax_cdf.legend()
    ax_cdf.grid(alpha=0.3)
    st.pyplot(fig_cdf)
    plt.close(fig_cdf)

    st.caption(
    "CDF graf ukazuje pravděpodobnost, že měsíční výsledek bude roven určité hodnotě "
    "nebo bude horší. Hodnota křivky u 0 Kč odpovídá pravděpodobnosti ztráty."
    )

    st.subheader("Interval nejistoty (Monte Carlo)")
    col_ci1, col_ci2 = st.columns(2)

    with col_ci1:
        st.metric(
            label=machine_1,
            value=f"{ci_mini[0]:,.0f} až {ci_mini[1]:,.0f} Kč",
        )

    with col_ci2:
        st.metric(
            label=machine_2,
            value=f"{ci_dumper[0]:,.0f} až {ci_dumper[1]:,.0f} Kč",
        )

    st.caption(
    "Interval obsahuje prostředních 80 % simulovaných měsíčních výsledků. "
    "Čím výše je celý interval posunutý, tím příznivější je očekávaný výsledek."
    )

    st.success(
    f"Pravděpodobnost, že **{machine_2}** dosáhne vyššího výsledku než "
    f"**{machine_1}**, je **{prob_dumper_better * 100:.1f} %**."
    )


# ==================================================
# TAB 3 – RIZIKO
# ==================================================

with tab3:
    st.subheader("🚦 Semafor rizika ztráty")
    col_r1, col_r2 = st.columns(2)
    with col_r1:
        st.write(f"*{machine_1}*")
        st.write(risk_light(prob_loss_mini))
    with col_r2:
        st.write(f"*{machine_2}*")
        st.write(risk_light(prob_loss_dumper))
    st.caption("Semafor vychází z pravděpodobnosti, že měsíční zisk bude záporný.")

    st.subheader("Provozní riziko")
    col_fail1, col_fail2 = st.columns(2)

    with col_fail1:
        st.metric(
            label=machine_1,
            value=f"{row_1['failure_prob'] * 100:.1f} %",
        )

    with col_fail2:
        st.metric(
            label=machine_2,
            value=f"{row_2['failure_prob'] * 100:.1f} %",
    )
    st.caption("Porucha znamená odstávku stroje a snížení počtu pronajatých dnů.")

    st.subheader("Hodnota v riziku (VaR)")

    col_var1, col_var2 = st.columns(2)

    with col_var1:
        st.metric(
            label=machine_1,
            value=f"{var95_mini:,.0f} Kč",
        )

    with col_var2:
        st.metric(
            label=machine_2,
            value=f"{var95_dumper:,.0f} Kč",
        )

    st.caption(
        "VaR 5 % ukazuje hranici výsledku v nejhorších 5 % simulovaných scénářů."
    )

    st.subheader("Mapa rizika vs. výnosu")
    fig_rr, ax_rr = plt.subplots()
    ax_rr.scatter(rr_x1, rr_y1, s=120)
    ax_rr.scatter(rr_x2, rr_y2, s=120)
    ax_rr.annotate(
    machine_1,
    (rr_x1, rr_y1),
    xytext=(-55, 8),
    textcoords="offset points"
    )

    ax_rr.annotate(
    machine_2,
    (rr_x2, rr_y2),
    xytext=(8, 8),
    textcoords="offset points"
    )
    ax_rr.set_xlabel("Pravděpodobnost ztráty (%)")
    ax_rr.set_ylabel("Očekávaný zisk (Kč)")
    ax_rr.set_title("Riziko vs. výnos")
    ax_rr.grid(True, alpha=0.3)
    ax_rr.axhline(0, linestyle="--")
    st.pyplot(fig_rr)

    better_risk_machine = (machine_1
        if (rr_x1 < rr_x2 and rr_y1 > rr_y2)
        else machine_2)
    
    st.caption("Příznivější varianta se nachází více vlevo a výše – má nižší pravděpodobnost ztráty "
    "a současně vyšší očekávaný výsledek.")

    st.info(f"Z pohledu rizika vychází relativně lépe **{better_risk_machine}**, "
    f"protože má nižší pravděpodobnost ztráty a příznivější očekávaný výsledek. "
    f"Při aktuálním nastavení však zůstává riziko obou variant velmi vysoké.")

    plt.close(fig_rr)


# ==================================================
# TAB 4 – CITLIVOST
# ==================================================

with tab4:
    st.subheader("Citlivostní analýza: zisk vs. vytížení (bod zvratu)")
    fig_be, ax_be = plt.subplots()
    ax_be.plot(days_grid, profit_curve_mini, label=machine_1)
    ax_be.plot(days_grid, profit_curve_dumper, label=machine_2)
    ax_be.axhline(0)
    ax_be.axvline(break_even_mini, linestyle="--")
    ax_be.axvline(break_even_dumper, linestyle="--")
    ax_be.set_xlabel("Vytížení (dny/měsíc)")
    ax_be.set_ylabel("Zisk / ztráta (Kč)")
    ax_be.set_title("Zisk podle vytížení (fixed_cost = zvolený fixní náklad)")
    ax_be.legend()
    st.pyplot(fig_be)
    plt.close(fig_be)

    st.caption(
        f"Bod zvratu {machine_1} = {break_even_mini:.2f} dní, {machine_2} = {break_even_dumper:.2f} dní\n"
        "Nad těmito hodnotami je stroj v průměru v zisku."
    )

    st.subheader("Citlivost zisku na změnu vstupů")
    fig_tornado, ax_tornado = plt.subplots(figsize=(10, 6))
    ax_tornado.barh(labels, values, alpha=0.8)
    ax_tornado.invert_yaxis()
    ax_tornado.set_title("Citlivost zisku na změnu vstupů")
    ax_tornado.set_xlabel("Dopad na očekávaný zisk (Kč)")
    ax_tornado.set_ylabel("Parametr")
    ax_tornado.grid(axis="x", linestyle="--", alpha=0.4)
    st.pyplot(fig_tornado)
    plt.close(fig_tornado)

    st.caption(
        "Tornado graf ukazuje, které vstupy mají největší vliv na očekávaný zisk. Čím delší pruh, tím citlivější je výsledek na změnu daného parametru."
    )

    st.subheader(f"Heatmapa ziskovosti – {machine_2}")
    fig_heat, ax_heat = plt.subplots()
    cbar_data = ax_heat.imshow(
        profit_matrix,
        aspect="auto",
        origin="lower",
        extent=[
            days_range.min(),
            days_range.max(),
            price_range.min(),
            price_range.max(),
        ],
    )
    ax_heat.contour(
        days_range,
        price_range,
        profit_matrix,
        levels=[0],
        linewidths=2,
        linestyles="--",
    )
    if profit_matrix.max() > 0:
        ax_heat.contourf(
            days_range,
            price_range,
            profit_matrix,
            levels=[0, profit_matrix.max()],
            colors=["green"],
            alpha=0.3,
        )
    ax_heat.set_xlabel("Poptávka (dny/měsíc)")
    ax_heat.set_ylabel("Cena pronájmu (Kč/den)")
    ax_heat.set_title(f"Zisk podle ceny a poptávky – {machine_2}")
    fig_heat.colorbar(cbar_data, ax=ax_heat, label="Zisk (Kč)")
    st.pyplot(fig_heat)
    plt.close(fig_heat)

    if profit_matrix.max() <= 0:
        st.warning(
            "Při aktuálních parametrech není v zadaném rozsahu ceny a poptávky žádná zisková kombinace."
        )
        st.caption(
            "Pro dosažení zisku je nutné zvýšit cenu pronájmu nebo dosáhnout vyššího vytížení stroje."
        )

    st.caption(
        "Heatmapa ukazuje, při jaké kombinaci ceny pronájmu a počtu pronajatých dní se investice dostává do zisku."
    )
    st.caption(
        "Barevná škála ukazuje očekávaný měsíční zisk. Tmavé barvy znamenají ztrátu, světle zelené až žluté oblasti vysoký zisk."
    )
    st.caption(
        "Přerušovaná čára označuje hranici break-even, tedy kombinace ceny a poptávky, kde je očekávaný zisk roven nule."
    )

    st.subheader("Doporučená minimální cena pronájmu")

    st.markdown(
    "Orientační denní sazba potřebná k dosažení bodu zvratu "
    "při aktuálně nastavené poptávce a nákladech."
    )

    col_price1, col_price2 = st.columns(2)

    with col_price1:
        price_1_ok = avg_daily_mini >= required_price_1

        st.metric(
            label=machine_1,
            value=f"{required_price_1:,.0f} Kč/den",
            delta=(
                "Aktuální cena je dostačující"
                if price_1_ok
                else "Aktuální cena je pod minimem"
            ),
            delta_color="normal" if price_1_ok else "inverse",
        )

    with col_price2:
        price_2_ok = avg_daily_dumper >= required_price_2

        st.metric(
            label=machine_2,
            value=f"{required_price_2:,.0f} Kč/den",
            delta=(
                "Aktuální cena je dostačující"
                if price_2_ok
                else "Aktuální cena je pod minimem"
            ),
            delta_color="normal" if price_2_ok else "inverse",
        )

    st.caption("Minimální denní sazba představuje orientační hranici bodu zvratu "
    "při aktuálně nastavené poptávce a nákladech. Vyšší sazba znamená "
    "větší rezervu pro tvorbu zisku.")

# ==================================================
# TAB 5 – OPTIMALIZACE
# ==================================================

with tab5:
    st.header("Optimalizační modely")
    st.caption(
        "Sekce ukazuje výsledky optimalizačních modelů (LP a DP) a následně optimální portfolio strojů při zadaném rozpočtu."
    )

    st.subheader("📐 Deterministická optimalizace (LP)")
    col_lp1, col_lp2, col_lp3 = st.columns(3)
    with col_lp1:
        st.metric(f"Optimální dny – {machine_1}", f"{x1_opt:.1f}")
    with col_lp2:
        st.metric(f"Optimální dny – {machine_2}", f"{x2_opt:.1f}")
    with col_lp3:
        st.metric("Max. deterministický zisk", f"{lp_profit:,.0f} Kč")

    if x1_opt > x2_opt:
        st.info(f"Deterministický model preferuje více využívat {machine_1}.")
    elif x2_opt > x1_opt:
        st.info(f"Deterministický model preferuje více využívat {machine_2}.")
    else:
        st.info("Deterministický model rozděluje využití mezi oba stroje podobně.")

    st.subheader("🧠 Diskrétní optimalizace (DP)")
    col_dp1, col_dp2, col_dp3 = st.columns(3)
    with col_dp1:
        st.metric(f"Optimální dny – {machine_1}", f"{best_d1}")
    with col_dp2:
        st.metric(f"Optimální dny – {machine_2}", f"{best_d2}")
    with col_dp3:
        st.metric("Max. diskrétní zisk", f"{best_dp_profit:,.0f} Kč")

    if best_d1 > best_d2:
        st.info(f"DP model preferuje více využívat {machine_1}.")
    elif best_d2 > best_d1:
        st.info(f"DP model preferuje více využívat {machine_2}.")
    else:
        st.info("DP model rozděluje využití mezi oba stroje podobně.")

    st.subheader("🧠 Porovnání rozhodovacích modelů")
    st.write(
        "**Monte Carlo** pracuje s nejistotou a ukazuje rozdělení možných výsledků.")
    st.write(
        "**LP optimalizace** hledá nejlepší řešení při pevných deterministických předpokladech.")
    st.write(
    "**DP optimalizace** hledá nejlepší řešení v diskrétních, celočíselných variantách.")

    col_model1, col_model2, col_model3 = st.columns(3)

    with col_model1:
        st.metric(
            label="Monte Carlo",
            value=mc_better
        )

    with col_model2:
        st.metric(
            label="LP optimalizace",
            value=lp_better
        )

    with col_model3:
        st.metric(
            label="DP optimalizace",
            value=dp_better
        )

    if mc_better == lp_better == dp_better:
        st.success(f"Všechny tři rozhodovací modely se shodují: "
        f"preferovanou variantou je **{mc_better}**.")
    else:
        st.warning("Rozhodovací modely se neshodují. "
            "Výsledek je vhodné posoudit také podle rizika a citlivosti vstupů.")

    st.subheader("📦 Optimalizace nákupu strojů")
    st.caption(
        "Optimalizační model testuje možné kombinace nákupu strojů a vybírá portfolio, které maximalizuje očekávaný měsíční zisk při daném investičním rozpočtu."
    )

    if best_combo is not None:
        st.success(f"Nejlepší portfolio při rozpočtu {budget:,.0f} Kč")

        active_machines = [
        (machines[i]["machine_type"], best_combo[i])
        for i in range(n_machines)
        if best_combo[i] > 0
    ]

        if active_machines:
            cols = st.columns(len(active_machines))

            for col, (machine_name, count) in zip(cols, active_machines):
                with col:
                    st.metric(
                        label=machine_name,
                        value=f"{count} ks"
                    )

            total_cost = sum(
                best_combo[i] * machines[i]["purchase_price"]
                for i in range(n_machines)
            )

            st.markdown("#### Souhrn portfolia")

        col_port1, col_port2 = st.columns(2)

        with col_port1:
            st.metric(
                label="Celková investice",
                value=f"{total_cost:,.0f} Kč"
            )

        with col_port2:
            st.metric(
                label="Očekávaný měsíční zisk",
                value=f"{best_profit:,.0f} Kč"
            )

        if total_cost > 0:
            portfolio_roi = best_profit / total_cost * 100

            st.caption(
            f"Očekávaný měsíční výnos portfolia odpovídá přibližně "
            f"{portfolio_roi:.2f} % hodnoty investice."
        )

        else:
            st.warning("Při daném rozpočtu se nepodařilo najít vhodnou kombinaci.")

    st.caption(
        "Optimalizace vybírá nejlepší kombinaci strojů z celého datasetu při zadaném investičním rozpočtu."
)

    st.subheader("📈 Rozpočet vs maximální zisk portfolia")

    fig_budget, ax_budget = plt.subplots(figsize=(10, 6))

    ax_budget.plot(
        budget_values,
        max_profits,
        marker="o",
        linewidth=2
    )

    # aktuálně zvolený rozpočet
    current_idx = np.argmin(np.abs(np.array(budget_values) - budget))
    current_budget_profit = max_profits[current_idx]

    ax_budget.scatter(
        budget_values[current_idx],
        current_budget_profit,
        s=120,
        zorder=5
    )

    ax_budget.annotate(
        f"Aktuální rozpočet\n{budget_values[current_idx]/1_000_000:.1f} mil. Kč\n"
        f"Zisk {current_budget_profit:,.0f} Kč/měs.",
        (budget_values[current_idx], current_budget_profit),
        xytext=(-130, -5),
        textcoords="offset points",
        arrowprops=dict(arrowstyle="->"),
    )

    ax_budget.set_xlabel("Investiční rozpočet (Kč)")
    ax_budget.set_ylabel("Maximální očekávaný měsíční zisk (Kč)")
    ax_budget.set_title("Vliv investičního rozpočtu na maximální zisk portfolia")

    ax_budget.grid(alpha=0.3)

    # hezčí formát osy X – bez zápisu 1e6
    ax_budget.ticklabel_format(style="plain", axis="x")

    fig_budget.tight_layout()

    from matplotlib.ticker import FuncFormatter

    ax_budget.xaxis.set_major_formatter(
    FuncFormatter(lambda x, _: f"{x/1_000_000:.1f} mil.")
)
    st.pyplot(fig_budget)
    plt.close(fig_budget)

    st.caption(
    "Graf ukazuje, jak se s rostoucím investičním rozpočtem mění maximální "
    "očekávaný měsíční zisk optimálního portfolia. Zvýrazněný bod odpovídá "
    "aktuálně zvolenému investičnímu rozpočtu.")

# ==================================================
# TAB 6 – ZAKÁZKY
# ==================================================

with tab6:

    st.header("📋 Optimalizace zásobníku zakázek")

    st.caption("Seřazení zakázek podle očekávaného zisku, rizika a priority.")

    planned_count = (
        scheduled_jobs["schedule_status"].str.contains("Naplánováno")
    ).sum()

    waiting_count = (scheduled_jobs["schedule_status"].str.contains("Čeká")).sum()

    total_plan_days = int(scheduled_jobs["end_day"].max())

    baseline_days = 30
    saved_days = baseline_days - total_plan_days

    daily_profit = 15000
    extra_profit = saved_days * daily_profit

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric("🟢 Naplánované", planned_count)

    with col2:
        st.metric("🔴 Čekající", waiting_count)

    with col3:
        st.metric("📅 Délka plánu", f"{total_plan_days} dní")

    with col4:
        st.metric("💰 Přínos", f"{extra_profit:,.0f} Kč")

    st.info(
        f"⚡ Optimalizací bylo ušetřeno {saved_days} dní "
        f"oproti referenčnímu plánu {baseline_days} dní."
    )

    st.subheader("🗓 Harmonogram zakázek")

    st.markdown("---")

    gantt_df = scheduled_jobs.copy()

    recommended_jobs = scheduled_jobs[
    scheduled_jobs["schedule_status"].str.contains("Naplánováno")
    ].copy()

    recommended_jobs["priority_score"] = (
        recommended_jobs["priority"]
        .map(priority_score)
        .fillna(0)
    )

    recommended_jobs["risk_score"] = (
        recommended_jobs["risk_level"]
        .map(risk_score)
        .fillna(0)
    )

    recommended_jobs["decision_score"] = (
        recommended_jobs["priority_score"] * 100000
        + recommended_jobs["expected_revenue"]
        - recommended_jobs["risk_score"] * 50000
    )

    recommended_job_name = (
        recommended_jobs
        .sort_values(
            by=["decision_score", "risk_score", "expected_revenue"],
            ascending=[False, True, False]
        )
        .iloc[0]["job_name"]
    )

    gantt_df = gantt_df[gantt_df["schedule_status"].str.contains("Naplánováno")]

    gantt_df["start_date"] = pd.to_datetime(
        gantt_df["start_day"], unit="D", origin="2025-01-01"
    )

    gantt_df["end_date"] = pd.to_datetime(
        gantt_df["end_day"], unit="D", origin="2025-01-01"
    )

    gantt_df["expected_profit_cz"] = (
    gantt_df["expected_profit"]
    .apply(format_czk)
)

    fig = px.timeline(
        gantt_df,
        x_start="start_date",
        x_end="end_date",
        y="job_name",
        color="priority",
        hover_data={
                "expected_profit": False,
                "expected_profit_cz": True,
                "risk_level": True,
                "required_machines": True,
                "start_date": False,
                "end_date": False,
            },

        labels={
            "job_name": "Zakázka",
            "priority": "Priorita",
            "expected_profit_cz": "Očekávaný zisk",
            "risk_level": "Riziko",
            "required_machines": "Potřebné stroje",
        },
    )

    fig.update_layout(
        yaxis_title="Zakázka",
        legend_title_text="Priorita"
)

    fig.update_yaxes(autorange="reversed")

    recommended_row = gantt_df[
    gantt_df["job_name"] == recommended_job_name
].iloc[0]

    fig.add_annotation(
        x=recommended_row["start_date"],
        y=recommended_job_name,
        text="⭐ Doporučená zakázka",
        showarrow=True,
        arrowhead=2,
        ax=-90,
        ay=-25
    )

    st.plotly_chart(fig, use_container_width=True)

    st.subheader("📄 Detail zakázek")

    jobs_display = scheduled_jobs[
    ["job_name", "priority", "duration_days", "expected_revenue"]
    ].rename(columns={
    "job_name": "Zakázka",
    "priority": "Priorita",
    "duration_days": "Délka (dny)",
    "expected_revenue": "Očekávaný výnos"
})

    jobs_display["Očekávaný výnos"] = (
    jobs_display["Očekávaný výnos"]
    .apply(format_czk)
)
    
    st.dataframe(
    jobs_display,
    use_container_width=True,
)
    machine_usage = []

    for machine in machine_capacity.keys():
        used_days = 0

        for _, job in scheduled_jobs.iterrows():
            required = [m.strip() for m in job["required_machines"].split(";")]

            if machine in required:
                used_days += int(job["duration_days"])

        utilization = used_days / total_plan_days * 100 if total_plan_days > 0 else 0

        machine_usage.append(
            {
                "Stroj": machine,
                "Využité dny": used_days,
                "Vytížení (%)": round(utilization, 1),
            }
        )

    machine_usage_df = pd.DataFrame(machine_usage)

    st.subheader("⚙️ Vytížení strojů")

    st.bar_chart(machine_usage_df.set_index("Stroj")["Vytížení (%)"])

    st.caption(
    "Vyšší vytížení znamená intenzivnější využití stroje v plánovaných zakázkách. "
    "Stroj s nejvyšším vytížením může představovat kapacitní omezení a měl by být sledován při plánování dalších zakázek.")

    st.dataframe(machine_usage_df,use_container_width=True,hide_index=True)

    most_used_machine = machine_usage_df.sort_values(
        "Vytížení (%)", ascending=False
    ).iloc[0]

    bottleneck_machine = most_used_machine["Stroj"]
    bottleneck_utilization = most_used_machine["Vytížení (%)"]

    st.info(
        f"Nejvíce vytížený stroj je **{most_used_machine['Stroj']}** "
        f"s vytížením **{most_used_machine['Vytížení (%)']} %**."
    )

    if bottleneck_utilization >= 70:
        st.warning(
            f"🚨 Úzké hrdlo systému: "
            f"**{bottleneck_machine}** "
            f"({bottleneck_utilization:.0f} % vytížení)"
        )
    else:
        st.success(
            f"✅ Kapacity jsou vyvážené. "
            f"Nejvíce vytížený stroj je "
            f"**{bottleneck_machine}** "
            f"({bottleneck_utilization:.0f} % vytížení)"
        )

    available_jobs = scheduled_jobs[
        scheduled_jobs["schedule_status"].str.contains("Naplánováno")
    ]

    if not available_jobs.empty:

        available_jobs = available_jobs.copy()

        available_jobs["priority_score"] = (
            available_jobs["priority"]
            .map(priority_score)
            .fillna(0)
        )

        available_jobs["risk_score"] = (
            available_jobs["risk_level"]
            .map(risk_score)
            .fillna(0)
        )

        available_jobs["decision_score"] = (
            available_jobs["priority_score"] * 100000
            + available_jobs["expected_revenue"]
            - available_jobs["risk_score"] * 50000
        )

        best_job = available_jobs.sort_values(by=["decision_score", "risk_score", "expected_revenue"],
        ascending=[False, True, False]).iloc[0]

        st.success(f"""
🏆 Doporučená prioritní zakázka

**{best_job['job_name']}**

💰 Očekávaný výnos: {format_czk(best_job['expected_revenue'])}
Priorita: **{best_job['priority']}**

Riziko: **{best_job['risk_level']}**

Rozhodovací skóre: **{best_job['decision_score']:,.0f} bodů**

{best_job['schedule_note']}
""")

        score_table = available_jobs[
            [
                "job_name",
                "priority",
                "risk_level",
                "risk_score",
                "expected_revenue",
                "decision_score"
            ]
        ].copy()

        score_table = score_table.sort_values(
        by=["decision_score", "risk_score", "expected_revenue"],
        ascending=[False, True, False])

        score_table = score_table.rename(columns={
            "job_name": "Zakázka",
            "priority": "Priorita",
            "risk_level": "Riziko",
            "expected_revenue": "Očekávaný výnos",
            "decision_score": "Rozhodovací skóre"
        })

        score_table = score_table.drop(columns=["risk_score"])

        st.subheader("📊 Pořadí zakázek podle rozhodovacího skóre")

        def highlight_priority(val):
            if val == "High":
                return "background-color: #C6EFCE; color: #006100; font-weight: 600"
            elif val == "Medium":
                return "background-color: #FFEB9C; color: #9C6500; font-weight: 600"
            elif val == "Low":
                return "background-color: #E7E6E6; color: #595959"
            return ""


        def highlight_risk(val):
            if val == "High":
                return "background-color: #FFC7CE; color: #9C0006; font-weight: 600"
            elif val == "Medium":
                return "background-color: #FCE4D6; color: #C65911; font-weight: 600"
            elif val == "Low":
                return "background-color: #C6EFCE; color: #006100; font-weight: 600"
            return ""

        styled_score_table = (
            score_table.style
            .format({
                "Očekávaný výnos": format_czk,
                "Rozhodovací skóre": lambda x: f"{x:,.0f}".replace(",", " ")
            })
            .map(highlight_priority, subset=["Priorita"])
            .map(highlight_risk, subset=["Riziko"])
        )

        st.dataframe(
            styled_score_table,
            use_container_width=True,
            hide_index=True
        )
        
    else:
        st.error("⚠️ Žádnou zakázku nelze naplánovat kvůli kapacitě strojů.")