import hmac
import importlib
import os
from io import BytesIO
from itertools import product

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from docx import Document
from docx.shared import Inches
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from scipy.optimize import linprog

import ui_theme as ui

# Streamlit při novém nasazení spustí app.py znovu, ale už načtený modul ui_theme
# ponechá v paměti ve staré verzi. Explicitní reload zajistí, že se vždy použije
# aktuální soubor (jinak hrozí AttributeError u nově přidaných funkcí).
ui = importlib.reload(ui)
from ui_theme import fmt_czk, fmt_days, fmt_num, fmt_pct

# ==================================================
# POMOCNÉ FUNKCE
# ==================================================


def format_czk(value) -> str:
    """Naformátuje částku v českém formátu s jednotkou Kč."""
    return fmt_czk(value)


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

st.set_page_config(
    page_title="Decision Engine – Investiční analýza půjčovny",
    page_icon="📊",
    layout="wide",
)
ui.inject_css()

APP_TITLE = "Decision Engine"
APP_SUBTITLE = "Investiční analýza půjčovny"
APP_FOCUS = "Půjčovna malostrojů"
DEMO_BADGE = "Ukázková data"

# Heslo se načítá pouze ze Streamlit secrets (klíč "password" na nejvyšší úrovni).
# str() a strip() zajistí, že nevadí číselná hodnota bez uvozovek ani mezery navíc.
try:
    PASSWORD = str(st.secrets.get("password", "")).strip()
except Exception:
    PASSWORD = ""

if not PASSWORD:
    st.error(
        "Přístupové heslo není nastaveno. Doplňte klíč `password` "
        "do Streamlit secrets (Settings → Secrets)."
    )
    st.stop()

# Po úspěšném přihlášení se přihlašovací pole skryje (stav je uložen v relaci).
if not st.session_state.get("authenticated", False):
    _, login_col, _ = st.columns([1, 1.2, 1])
    with login_col:
        ui.login_intro(APP_TITLE, APP_SUBTITLE, "Pro pokračování zadejte přístupové heslo.")
        password = st.text_input(
            "Přístupové heslo",
            type="password",
            help="Heslo je nastaveno ve Streamlit secrets. Po zadání stiskněte Enter.",
        )
        # Tlačítko jen potvrdí zadané heslo (stejně jako Enter); ověření níže je beze změny.
        st.button("Přihlásit se", type="primary", width="stretch")
        if password and hmac.compare_digest(
            password.strip().encode("utf-8"), PASSWORD.encode("utf-8")
        ):
            st.session_state["authenticated"] = True
            st.rerun()
        if password:
            st.error("Nesprávné heslo.")
        st.caption(
            "Pilotní prototyp. Všechny stroje, ceny, náklady a zakázky jsou fiktivní ukázková data."
        )
    st.stop()


# ==================================================
# HLAVIČKA APLIKACE
# ==================================================

ui.header(APP_TITLE, APP_SUBTITLE, DEMO_BADGE)


# ==================================================
# DATA A SIDEBAR VSTUPY
# ==================================================

# Sekce postranního panelu se vytvoří dopředu, aby šly plnit na místech,
# kde je kód potřebuje – pořadí výpočtů tak zůstává beze změny.
ui.sidebar_brand(APP_TITLE, APP_FOCUS)
st.sidebar.markdown(
    '<div class="de-demo"><b>Ukázková data</b> · všechny hodnoty jsou fiktivní</div>',
    unsafe_allow_html=True,
)
sb_machines = st.sidebar.container()
sb_costs = st.sidebar.container()
sb_market = st.sidebar.container()
sb_budget = st.sidebar.container()
sb_jobs = st.sidebar.container()
sb_advanced = st.sidebar.expander("Pokročilé nastavení")

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

# Volba druhého pásového dumperu zatím není zapojena (plánování nepracuje s počtem
# kusů stroje), proto se v panelu nenabízí a hodnota zůstává False jako dosud.
extra_dumper = False
sb_jobs.markdown(
    '<div class="de-soon">Plánování zakázek – připravujeme</div>', unsafe_allow_html=True
)

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

sb_machines.markdown("### 1 · Porovnávané stroje")
machine_1 = sb_machines.selectbox(
    "Stroj 1", machine_options, index=0,
    help="První stroj, o jehož nákupu uvažujete.",
)
machine_2 = sb_machines.selectbox(
    "Stroj 2", machine_options, index=1,
    help="Druhý stroj pro srovnání. Musí se lišit od stroje 1.",
)

if machine_1 == machine_2:
    st.warning(
        "Pro porovnání vyberte v postranním panelu dva **různé** stroje."
    )
    st.stop()

row_1 = df[df["machine_type"] == machine_1].iloc[0]
row_2 = df[df["machine_type"] == machine_2].iloc[0]

avg_daily_mini = float(row_1["daily_price"])
avg_daily_dumper = float(row_2["daily_price"])

sb_costs.markdown("### 2 · Náklady")
fixed_cost = sb_costs.number_input(
    "Měsíční fixní náklady na jeden stroj",
    min_value=0,
    value=100000,
    step=5000,
    help="Částka v Kč, kterou stroj stojí každý měsíc bez ohledu na počet dní pronájmu "
    "(např. splátka, pojištění). Odečítá se od výsledku každého z obou porovnávaných strojů. "
    "Servisní náklady se počítají zvlášť z údajů o stroji. Výchozí hodnota je ukázková.",
)
ui.sidebar_amount(sb_costs, fmt_czk(fixed_cost), "měsíčně na jeden stroj")

sb_market.markdown("### 3 · Trh")
scenario = sb_market.selectbox(
    "Scénář trhu",
    ["Pesimistický", "Realistický", "Optimistický"],
    index=1,
    help="Určuje, kolik dní v měsíci může být stroj pronajatý. "
    "Pesimistický 8–15 dní, realistický 15–25 dní, optimistický 22–30 dní.",
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

sb_market.caption(
    f"{scenario_note} Simulace počítá s {min_days}–{max_days} dny pronájmu měsíčně."
)

expected_days = sb_advanced.slider(
    "Odhad poptávky pro rychlý výpočet (dny/měsíc)", 0, 31, 18,
    help="Používá se pro rychlý výpočet bez simulace (karta „Rychlý odhad“), "
    "pro citlivost a pro minimální cenu pronájmu. Simulace pracuje s rozsahem "
    "dní podle zvoleného scénáře trhu.",
)
n = sb_advanced.slider(
    "Počet simulací (měsíců)", 100, 10000, 1000,
    help="Kolik možných měsíců aplikace nasimuluje. Více simulací = stabilnější "
    "výsledky, ale pomalejší přepočet.",
)

sb_budget.markdown("### 4 · Investiční rozpočet")
budget = sb_budget.number_input(
    "Investiční rozpočet",
    min_value=100000,
    value=3000000,
    step=100000,
    help="Kolik peněz je k dispozici na nákup strojů. Využívá záložka „Nákup strojů“.",
)
ui.sidebar_amount(sb_budget, fmt_czk(budget), "k dispozici na nákup strojů")


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


def build_docx_report() -> bytes:
    doc = Document()

    # Základní písmo dokumentu
    normal_style = doc.styles["Normal"]
    normal_style.font.name = "Aptos"
    normal_style.font.size = Pt(10.5)

    for style_name, size in (("Heading 1", 16), ("Heading 2", 13), ("Heading 3", 11)):
        heading = doc.styles[style_name]
        heading.font.name = "Aptos Display"
        heading.font.size = Pt(size)
        heading.font.bold = True
        heading.font.color.rgb = RGBColor(20, 48, 79)

    title = doc.add_heading("Decision Engine – investiční report", level=1)
    title_run = title.runs[0]
    title_run.font.size = Pt(20)

    doc.add_paragraph(f"Porovnání: {machine_1} vs. {machine_2}")
    demo = doc.add_paragraph()
    demo_run = demo.add_run(
        "Ukázkový výstup pilotního prototypu. Všechny vstupy i výsledky vycházejí "
        "z fiktivních demonstračních dat."
    )
    demo_run.italic = True
    demo_run.font.color.rgb = RGBColor(122, 74, 6)

    doc.add_heading("Zadané předpoklady", level=1)
    doc.add_paragraph(f"Fixní náklad na stroj: {fmt_czk(fixed_cost)} měsíčně")
    doc.add_paragraph(f"Scénář trhu: {scenario} ({min_days}–{max_days} dní pronájmu měsíčně)")
    doc.add_paragraph(f"Počet simulovaných měsíců: {fmt_num(n)}")
    doc.add_paragraph(f"Odhad poptávky pro rychlý výpočet: {expected_days} dní/měsíc")

    doc.add_heading("Klíčové investiční ukazatele", level=1)
    table = doc.add_table(rows=1, cols=3)
    table.style = "Table Grid"
    hdr = table.rows[0].cells
    hdr[0].text = "Ukazatel"
    hdr[1].text = machine_1
    hdr[2].text = machine_2

    rows = [
        ("Denní cena pronájmu", fmt_czk(avg_daily_mini), fmt_czk(avg_daily_dumper)),
        ("Bod zvratu", fmt_days(break_even_mini, 2), fmt_days(break_even_dumper, 2)),
        ("Šance ztrátového měsíce", fmt_pct(risk_loss_mini), fmt_pct(risk_loss_dumper)),
        ("Očekávaný měsíční zisk (simulace)", fmt_czk(expected_profit_mini), fmt_czk(expected_profit_dumper)),
    ]
    for metric, v1, v2 in rows:
        cells = table.add_row().cells
        cells[0].text = metric
        cells[1].text = v1
        cells[2].text = v2

    doc.add_heading("Relativně výhodnější varianta", level=1)
    verdict_p = doc.add_paragraph()
    verdict_p.add_run("Relativně lépe vychází: ").bold = True
    verdict_p.add_run(f"{better_machine}.\n")
    if expected_profit_mini < 0 and expected_profit_dumper < 0:
        verdict_p.add_run(
            "Při zadaném nastavení jsou však obě investiční varianty ztrátové. "
            "Investici proto nedoporučujeme realizovat bez zvýšení ceny pronájmu, "
            "vyššího vytížení nebo snížení nákladů."
        )
    else:
        verdict_p.add_run(
            f"{better_machine} nabízí vyšší očekávaný měsíční výsledek "
            "v simulaci. Podrobné srovnání rizika a bodu zvratu je v tabulce výše."
        )

    doc.add_heading("Rozpětí možných výsledků (simulace)", level=1)
    doc.add_paragraph(
        "Simulace vytváří tisíce možných měsíců s různou poptávkou, servisními náklady, "
        "poruchami a cenou. Následující hodnoty shrnují rozpětí měsíčního výsledku."
    )
    doc.add_paragraph(
        f"Slabý měsíc (P10 – jen 1 z 10 měsíců je horší): {machine_1} {fmt_czk(p10_mini)}, "
        f"{machine_2} {fmt_czk(p10_dumper)}"
    )
    doc.add_paragraph(
        f"Typický měsíc (medián): {machine_1} {fmt_czk(p50_mini)}, {machine_2} {fmt_czk(p50_dumper)}"
    )
    doc.add_paragraph(
        f"Dobrý měsíc (P90 – jen 1 z 10 měsíců je lepší): {machine_1} {fmt_czk(p90_mini)}, "
        f"{machine_2} {fmt_czk(p90_dumper)}"
    )

    price_heading = doc.add_heading("Doporučená minimální cena pronájmu", level=1)
    price_heading.paragraph_format.page_break_before = True

    price_table = doc.add_table(rows=1, cols=3)
    price_table.style = "Table Grid"
    header_cells = price_table.rows[0].cells
    header_cells[0].text = "Stroj"
    header_cells[1].text = "Minimální cena"
    header_cells[2].text = "Hodnocení aktuální ceny"

    for name, required, actual in (
        (machine_1, required_price_1, avg_daily_mini),
        (machine_2, required_price_2, avg_daily_dumper),
    ):
        row = price_table.add_row().cells
        row[0].text = str(name)
        row[1].text = f"{fmt_czk(required)}/den"
        row[2].text = "Cena je dostačující" if actual >= required else "Cena je pod minimem pro zisk"

    doc.add_paragraph(
        f"Minimální cena ukazuje, při jaké denní sazbě se při poptávce {expected_days} dní "
        "měsíčně pokryjí fixní a servisní náklady."
    )

    doc.add_heading("Interpretace výsledků", level=1)
    doc.add_paragraph(
        "Model simuluje možné scénáře poptávky po strojích a zohledňuje náklady na servis, "
        "fixní náklady a pravděpodobnost poruch. Výsledkem je odhad očekávaného měsíčního "
        "zisku a rizika ztráty."
    )
    doc.add_paragraph(
        "Optimalizační model testuje možné kombinace nákupu strojů a vybírá sestavu, "
        "která maximalizuje očekávaný měsíční zisk při daném investičním rozpočtu."
    )
    doc.add_paragraph(
        "Doporučení jsou sestavena pevně danými pravidly z vypočtených ukazatelů."
    )

    doc.add_heading("Doporučený další postup", level=1)
    for step in [
        f"Ověřit skutečnou měsíční poptávku po stroji {better_machine}.",
        "Porovnat aktuální cenu pronájmu s vypočtenou minimální denní sazbou.",
        "Doplnit reálné servisní náklady, odstávky, pojištění a sezónnost.",
        "Znovu vyhodnotit investici v realistickém scénáři.",
    ]:
        doc.add_paragraph(step, style="List Number")

    doc.add_heading("Podmínky ziskovosti", level=1)
    doc.add_paragraph(
        "Investice by byla výrazně atraktivnější při splnění některých z následujících podmínek:"
    )
    for item in [
        "zvýšení denní ceny pronájmu",
        "vyšší vytížení strojů (více dní pronájmu za měsíc)",
        "snížení fixních nebo servisních nákladů",
    ]:
        doc.add_paragraph(item, style="List Bullet")

    out = BytesIO()
    doc.save(out)
    out.seek(0)
    return out.read()


# ==================================================
# POMOCNÉ FUNKCE PRO ZOBRAZENÍ
# ==================================================

MACHINE_COLOR = {machine_1: ui.SERIES[0], machine_2: ui.SERIES[1]}
both_loss = expected_profit_mini < 0 and expected_profit_dumper < 0
models_agree = mc_better == lp_better == dp_better


def loss_share_text(prob: float) -> str:
    """Převede pravděpodobnost ztráty na srozumitelnou větu „x z 10 měsíců“."""
    if prob < 0.05:
        return "méně než 1 z 10 měsíců"
    if prob > 0.95:
        return "téměř každý měsíc"
    return f"přibližně {max(1, round(prob * 10))} z 10 měsíců"


def chart(fig) -> None:
    st.plotly_chart(fig, width="stretch", config={"displaylogo": False, "locale": "cs"})


# ==================================================
# ZÁLOŽKY
# ==================================================

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
    [
        "🏆 Doporučení",
        "📈 Možné výsledky",
        "⚠️ Riziko",
        "🎯 Co ovlivňuje zisk",
        "🛒 Nákup strojů",
        "🗓️ Plán zakázek",
    ]
)


# ==================================================
# ZÁLOŽKA 1 – DOPORUČENÍ
# ==================================================

with tab1:
    # Hodnoty doporučeného a druhého stroje (jen výběr z již vypočtených proměnných).
    rec_is_m2 = better_machine == machine_2
    other_machine = machine_1 if rec_is_m2 else machine_2
    rec_profit = expected_profit_dumper if rec_is_m2 else expected_profit_mini
    other_profit = expected_profit_mini if rec_is_m2 else expected_profit_dumper
    rec_prob = prob_loss_dumper if rec_is_m2 else prob_loss_mini
    other_prob = prob_loss_mini if rec_is_m2 else prob_loss_dumper
    rec_be = break_even_dumper if rec_is_m2 else break_even_mini
    other_be = break_even_mini if rec_is_m2 else break_even_dumper
    rec_price = avg_daily_dumper if rec_is_m2 else avg_daily_mini
    rec_required = required_price_2 if rec_is_m2 else required_price_1
    rec_better_share = prob_dumper_better if rec_is_m2 else 1 - prob_dumper_better

    def risk_status(prob: float) -> str:
        """Štítek rizika podle stejných hranic jako semafor (risk_light)."""
        light = risk_light(prob)
        tone = {"🟢": "good", "🟡": "warn", "🔴": "bad"}[light[0]]
        return ui.status_chip(light[2:], tone)

    # --- A. Hlavní doporučení: stroj + jedna krátká věta ---
    if both_loss:
        ui.hero(
            "Žádný stroj není ziskový",
            f"Relativně lépe: {better_machine}",
            "Při zadaných předpokladech by oba stroje v průměru prodělávaly.",
            tone="bad",
        )
    elif rec_profit == other_profit:
        ui.hero(
            "Stroje vycházejí stejně",
            f"{machine_1} / {machine_2}",
            "Oba stroje mají stejný očekávaný zisk – rozhodněte podle rizika a ceny.",
        )
    else:
        if rec_prob > other_prob:
            reason = "Vyšší očekávaný zisk, ale také vyšší riziko ztráty."
        else:
            parts = ["vyšší očekávaný zisk"]
            if rec_prob < other_prob:
                parts.append("nižší riziko ztráty")
            if rec_be < other_be:
                parts.append("méně dní potřebných k pokrytí fixních nákladů")
            text = parts[0] if len(parts) == 1 else ", ".join(parts[:-1]) + " a " + parts[-1]
            reason = text[0].upper() + text[1:] + "."
        ui.hero("Doporučený stroj", better_machine, reason)

    # --- B. Tři hlavní ukazatele doporučeného stroje ---
    ui.kpi_cards(
        [
            dict(
                label="Očekávaný měsíční zisk",
                value=fmt_czk(rec_profit),
                negative=rec_profit < 0,
                sub="Odhad průměrného měsíčního zisku.",
            ),
            dict(
                label="Riziko ztráty",
                value=fmt_pct(rec_prob * 100),
                chip=risk_status(rec_prob),
                sub="Pravděpodobnost, že měsíc skončí ztrátou.",
            ),
            dict(
                label="Bod zvratu",
                value=fmt_days(rec_be),
                sub="Dní pronájmu měsíčně potřebných k pokrytí fixních nákladů.",
            ),
        ]
    )

    # --- C. Jednoduché srovnání obou strojů ---
    st.markdown('<div class="de-section">Srovnání strojů</div>', unsafe_allow_html=True)

    def best_index(a, b, higher_is_better: bool):
        if a == b:
            return None
        return (0 if a > b else 1) if higher_is_better else (0 if a < b else 1)

    ui.compare_table(
        [machine_1, machine_2],
        [
            ("Měsíční zisk",
             [fmt_czk(expected_profit_mini), fmt_czk(expected_profit_dumper)],
             best_index(expected_profit_mini, expected_profit_dumper, True)),
            ("Riziko ztráty",
             [fmt_pct(risk_loss_mini), fmt_pct(risk_loss_dumper)],
             best_index(prob_loss_mini, prob_loss_dumper, False)),
            ("Bod zvratu",
             [fmt_days(break_even_mini), fmt_days(break_even_dumper)],
             best_index(break_even_mini, break_even_dumper, False)),
            ("Pořizovací cena",
             [fmt_czk(row_1["purchase_price"]), fmt_czk(row_2["purchase_price"])],
             None),
        ],
        rec_index=1 if rec_is_m2 else 0,
        tag="Relativně lepší" if both_loss else "Doporučeno",
    )
    st.markdown(
        '<div class="de-cmp-note">Zeleně ✓ = výhodnější hodnota.</div>',
        unsafe_allow_html=True,
    )

    # --- D. Co doporučujeme udělat dál (max. tři krátké kroky) ---
    if both_loss:
        steps = [
            "Za těchto předpokladů stroj nekupovat.",
            f"Prověřit vyšší cenu nebo vytížení – minimum pro {better_machine} je "
            f"{fmt_czk(rec_required)}/den.",
            "Doplnit skutečné náklady a výpočet zopakovat.",
        ]
    else:
        steps = [
            f"Ověřit, že {better_machine} bude pronajatý alespoň {fmt_days(rec_be)} měsíčně.",
            f"Zkontrolovat ceník: minimum {fmt_czk(rec_required)}/den, nyní {fmt_czk(rec_price)}/den.",
            "Doplnit skutečné náklady a výpočet zopakovat.",
        ]
    st.markdown(
        '<div class="de-section">Co udělat dál</div><ol class="de-steps">'
        + "".join(f"<li>{s}</li>" for s in steps)
        + "</ol>",
        unsafe_allow_html=True,
    )

    # --- E. Podrobnosti (ve výchozím stavu sbalené) ---
    with st.expander("Jak aplikace došla k doporučení"):
        st.markdown(
            f"Doporučený je stroj s **vyšším průměrným měsíčním ziskem** v simulaci "
            f"{fmt_num(n)} možných měsíců. {better_machine} vydělá více než {other_machine} "
            f"přibližně v **{fmt_pct(rec_better_share * 100, 0)}** simulovaných měsíců."
        )
        st.markdown(
            "Stejnou variantu doporučují i lineární a diskrétní optimalizace."
            if models_agree
            else "Lineární a diskrétní optimalizace vycházejí jinak – posuďte i záložky "
            "Riziko a Co ovlivňuje zisk."
        )
        col_s1, col_s2 = st.columns(2)
        col_s1.metric(f"Bodové hodnocení – {machine_1}", f"{scorecard_mini} / 4")
        col_s2.metric(f"Bodové hodnocení – {machine_2}", f"{scorecard_dumper} / 4")
        st.caption(
            "Body za vyšší zisk, nižší riziko ztráty, lepší nejhorší měsíc z 20 a rychlejší "
            "návratnost (jen když jsou oba stroje v zisku). Text níže sestavují pevná pravidla."
        )
        st.markdown(ai_text)

    with st.expander("Podrobné ukazatele obou strojů"):
        ui.compare_table(
            [machine_1, machine_2],
            [
                ("Ceník – denní pronájem", [fmt_czk(avg_daily_mini), fmt_czk(avg_daily_dumper)], None),
                (f"Minimální cena při {expected_days} dnech",
                 [fmt_czk(required_price_1), fmt_czk(required_price_2)], None),
                ("Aktuální ceník",
                 [
                     "Dostačuje" if avg_daily_mini >= required_price_1 else "Pod minimem",
                     "Dostačuje" if avg_daily_dumper >= required_price_2 else "Pod minimem",
                 ],
                 None),
            ],
            rec_index=1 if rec_is_m2 else 0,
            tag="Relativně lepší" if both_loss else "Doporučeno",
        )
        st.caption(
            "Minimální cena pokryje fixní a průměrné servisní náklady při zadaném odhadu poptávky. "
            "Bod zvratu na hlavní stránce počítá jen s fixními náklady (fixní náklad ÷ denní cena)."
        )

    with st.expander(f"Rychlý odhad bez simulace ({expected_days} dní pronájmu)"):
        col_q1, col_q2 = st.columns(2)
        col_q1.metric(f"{machine_1} – měsíční výsledek", fmt_czk(expected_profit_mini_real))
        col_q2.metric(f"{machine_2} – měsíční výsledek", fmt_czk(expected_profit_dumper_real))
        st.caption(
            "Denní cena × počet dní − fixní náklad − průměrný servis. "
            "Počet dní změníte v postranním panelu v „Pokročilém nastavení“."
        )
        if expected_profit_mini_real < 0 and expected_profit_dumper_real < 0:
            st.warning("Při tomto počtu dní jsou oba stroje ztrátové.")
        elif expected_profit_mini_real < 0 or expected_profit_dumper_real < 0:
            st.info("Jeden ze strojů je při tomto počtu dní ztrátový.")

    # --- F. Export reportu ---
    st.write("")
    with st.container(border=True):
        col_r1, col_r2 = st.columns([2, 1], vertical_alignment="center")
        col_r1.markdown(
            '<div class="de-section" style="margin:0">Investiční report</div>'
            f'<div style="color:{ui.UI_MUTED};font-size:.88rem">Shrnutí výsledků ve formátu Word. '
            "Fiktivní data – podklad pro rozhodnutí, ne investiční doporučení.</div>",
            unsafe_allow_html=True,
        )
        col_r2.download_button(
            label="Stáhnout report (DOCX)",
            data=build_docx_report(),
            file_name="decision_report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            width="stretch",
        )


# ==================================================
# ZÁLOŽKA 2 – MOŽNÉ VÝSLEDKY (MONTE CARLO)
# ==================================================

with tab2:
    st.markdown("### Jak mohou vypadat jednotlivé měsíce")
    st.markdown(
        f'<div class="de-intro">Simulace {fmt_num(n)} možných měsíců ukazuje, jak se může měnit '
        "zisk nebo ztráta při různém vytížení a nákladech.</div>",
        unsafe_allow_html=True,
    )

    ui.results_cards(
        [
            dict(
                name=name,
                color=MACHINE_COLOR[name],
                rows=[
                    ("Slabý měsíc (P10)", "1 z 10 měsíců dopadne hůř", fmt_czk(p10), "", p10 < 0),
                    ("Typický měsíc (medián)", "polovina měsíců dopadne lépe", fmt_czk(p50), "", p50 < 0),
                    ("Dobrý měsíc (P90)", "1 z 10 měsíců dopadne lépe", fmt_czk(p90), "", p90 < 0),
                    ("Průměr všech měsíců", "", fmt_czk(mean), "", mean < 0),
                    ("Podíl ztrátových měsíců", "", fmt_pct(prob * 100), loss_share_text(prob), False),
                ],
            )
            for name, p10, p50, p90, mean, prob in (
                (machine_1, p10_mini, p50_mini, p90_mini, expected_profit_mini, prob_loss_mini),
                (machine_2, p10_dumper, p50_dumper, p90_dumper, expected_profit_dumper, prob_loss_dumper),
            )
        ],
        rec_name=better_machine,
        tag="Relativně lepší" if both_loss else "Doporučeno",
    )

    st.markdown("### Rozložení měsíčního výsledku")
    fig_mc = go.Figure()
    for name, arr, med in ((machine_1, profit_mini, p50_mini), (machine_2, profit_dumper, p50_dumper)):
        fig_mc.add_trace(
            go.Histogram(
                x=arr, name=name, nbinsx=25, opacity=0.6,
                marker=dict(color=MACHINE_COLOR[name], line=dict(color="#fff", width=1)),
                hovertemplate="Zisk %{x} Kč<br>Počet měsíců: %{y}<extra>" + name + "</extra>",
            )
        )
        fig_mc.add_vline(x=med, line=dict(color=MACHINE_COLOR[name], dash="dot", width=2))
    fig_mc.add_vline(
        x=0, line=dict(color=ui.MUTED, dash="dash", width=1.5),
        annotation_text="0 Kč", annotation_position="top right",
        annotation=dict(bgcolor="rgba(255,255,255,0.9)", font=dict(color=ui.INK, size=12)),
    )
    fig_mc.update_layout(barmode="overlay")
    fig_mc.update_xaxes(title_text="Měsíční zisk / ztráta (Kč)")
    fig_mc.update_yaxes(title_text="Počet simulovaných měsíců")
    chart(ui.readable_axes(ui.style_fig(fig_mc, height=380)))
    ui.hint(
        "Čím více výsledků je napravo od 0 Kč, tím častěji stroj dosahuje zisku. "
        "Vlevo od 0 Kč jsou ztrátové měsíce. Tečkované čáry označují typický měsíc."
    )

    ui.insight(
        f"Pravděpodobnost, že <b>{machine_2}</b> dosáhne vyššího výsledku než "
        f"<b>{machine_1}</b>, je <b>{fmt_pct(prob_dumper_better * 100)}</b>."
    )

    with st.expander("Podrobnější analýza výsledků"):
        st.markdown('<div class="de-section">Kumulativní pravděpodobnost (CDF)</div>', unsafe_allow_html=True)
        fig_cdf = go.Figure()
        for name, arr in ((machine_1, profit_mini), (machine_2, profit_dumper)):
            sorted_arr = np.sort(arr)
            cdf = np.arange(1, len(sorted_arr) + 1) / len(sorted_arr)
            fig_cdf.add_trace(
                go.Scatter(
                    x=sorted_arr, y=cdf, mode="lines", name=name,
                    line=dict(color=MACHINE_COLOR[name], width=2),
                    hovertemplate="Výsledek ≤ %{x:,.0f} Kč<br>Pravděpodobnost %{y:.0%}<extra>"
                    + name + "</extra>",
                )
            )
        fig_cdf.add_vline(x=0, line=dict(color=ui.MUTED, dash="dash", width=1.5))
        fig_cdf.update_xaxes(title_text="Měsíční zisk / ztráta (Kč)")
        fig_cdf.update_yaxes(title_text="Pravděpodobnost", tickformat=".0%")
        chart(ui.readable_axes(ui.style_fig(fig_cdf, height=320)))
        ui.hint(
            "Graf ukazuje pravděpodobnost, že měsíční výsledek bude nižší nebo roven vybrané "
            "částce. Na hranici 0 Kč vidíte pravděpodobnost ztráty."
        )

        ui.interval_boxes(
            [
                (machine_1, MACHINE_COLOR[machine_1], f"{fmt_czk(ci_mini[0])} až {fmt_czk(ci_mini[1])}"),
                (machine_2, MACHINE_COLOR[machine_2], f"{fmt_czk(ci_dumper[0])} až {fmt_czk(ci_dumper[1])}"),
            ],
            label="Rozmezí prostředních 80 % simulovaných měsíců",
        )
        ui.hint("Dolní hranice odpovídá slabému měsíci (P10), horní dobrému měsíci (P90).")


# ==================================================
# ZÁLOŽKA 3 – RIZIKO
# ==================================================

with tab3:
    st.markdown("### Jak velké je riziko")

    def risk_chip(prob: float) -> str:
        """Štítek kategorie rizika podle stejných hranic jako semafor (risk_light)."""
        light = risk_light(prob)
        tone = {"🟢": "good", "🟡": "warn", "🔴": "bad"}[light[0]]
        return ui.status_chip(light[2:], tone)

    ui.results_cards(
        [
            dict(
                name=name,
                color=MACHINE_COLOR[name],
                rows=[
                    ("Kategorie rizika", "podle podílu ztrátových měsíců", risk_chip(prob), "", False),
                    ("Šance ztrátového měsíce", "", fmt_pct(prob * 100), loss_share_text(prob), False),
                    ("Velmi nepříznivý měsíc", "hranice 5 % nejhorších měsíců (VaR 5 %)",
                     fmt_czk(var95), "", var95 < 0),
                    ("Pravděpodobnost poruchy v měsíci", "údaj o stroji",
                     fmt_pct(row_m["failure_prob"] * 100), "", False),
                ],
            )
            for name, prob, var95, row_m in (
                (machine_1, prob_loss_mini, var95_mini, row_1),
                (machine_2, prob_loss_dumper, var95_dumper, row_2),
            )
        ],
        rec_name=better_machine,
        tag="Relativně lepší" if both_loss else "Doporučeno",
    )

    m1_dominates = prob_loss_mini < prob_loss_dumper and expected_profit_mini > expected_profit_dumper
    m2_dominates = prob_loss_dumper < prob_loss_mini and expected_profit_dumper > expected_profit_mini
    if m1_dominates or m2_dominates:
        winner = machine_1 if m1_dominates else machine_2
        ui.insight(
            f"Z pohledu rizika vychází lépe <b>{winner}</b>: má nižší šanci ztrátového měsíce "
            "a zároveň vyšší očekávaný výsledek."
        )
    else:
        ui.insight(
            "Ani jeden stroj není lepší v obou ohledech zároveň – jeden má vyšší očekávaný "
            "výsledek, druhý nižší riziko. Rozhodnutí záleží na tom, kolik rizika jste "
            "ochotni nést."
        )
    if prob_loss_mini >= 0.15 and prob_loss_dumper >= 0.15:
        ui.insight(
            "U obou strojů je riziko ztrátového měsíce při těchto předpokladech <b>vysoké</b>.",
            tone="bad",
        )

    st.markdown("### Riziko vs. výnos")
    fig_rr = go.Figure()
    # Bod s vyšším rizikem (vpravo) má název vlevo, aby se na úzkém displeji vešel do grafu.
    right_name = machine_1 if rr_x1 > rr_x2 else machine_2
    for name, x, y in ((machine_1, rr_x1, rr_y1), (machine_2, rr_x2, rr_y2)):
        fig_rr.add_trace(
            go.Scatter(
                x=[x], y=[y], mode="markers+text", name=name, text=[f"<b>{name}</b>"],
                textposition="middle left" if name == right_name else "middle right",
                textfont=dict(size=14, color=ui.INK),
                marker=dict(size=24, color=MACHINE_COLOR[name], line=dict(color="#fff", width=3)),
                hovertemplate="Šance ztráty %{x:.1f} %<br>Očekávaný zisk %{y:,.0f} Kč<extra>"
                + name + "</extra>",
            )
        )
    fig_rr.add_hline(y=0, line=dict(color=ui.MUTED, dash="dash", width=1.5))
    rr_pad = max(abs(rr_y1), abs(rr_y2)) * 0.25 + 1000
    fig_rr.update_xaxes(
        title_text="Šance ztrátového měsíce (%)", range=[-2, max(rr_x1, rr_x2) * 1.35 + 5],
        ticksuffix=" %",
    )
    fig_rr.update_yaxes(
        title_text="Očekávaný měsíční zisk (Kč)",
        range=[min(0, rr_y1, rr_y2) - rr_pad, max(0, rr_y1, rr_y2) + rr_pad],
        tickformat=",.0f",
    )
    chart(ui.readable_axes(ui.style_fig(fig_rr, height=380, legend=False)))
    ui.hint(
        "Čím více vlevo, tím menší riziko ztráty. Čím výše, tím vyšší očekávaný zisk. "
        "Nejvýhodnější kombinace se proto nachází vlevo nahoře."
    )

    with st.expander("Jak číst ukazatele rizika"):
        ui.explain(
            [
                ("Kategorie rizika (semafor)",
                 "Vychází z podílu ztrátových měsíců v simulaci: <b>Nízké</b> = méně než 5 %, "
                 "<b>Střední</b> = 5 % až méně než 15 %, <b>Vysoké</b> = 15 % a více."),
                ("Šance ztrátového měsíce",
                 "Podíl simulovaných měsíců, ve kterých stroj skončil ve ztrátě "
                 "(tržby nepokryly fixní a servisní náklady)."),
                ("Velmi nepříznivý měsíc (VaR 5 %)",
                 "Jde o odhad výsledku na hranici 5 % nejhorších simulovaných měsíců. "
                 "Skutečný výsledek může být i horší. Kladné číslo znamená, že i takto "
                 "nepříznivý měsíc skončí v zisku."),
                ("Pravděpodobnost poruchy v měsíci",
                 "Údaj o stroji z demonstračních dat. Při poruše stroj v simulaci přijde "
                 "o část dní pronájmu."),
            ]
        )


# ==================================================
# ZÁLOŽKA 4 – CO OVLIVŇUJE ZISK (CITLIVOST)
# ==================================================

with tab4:
    # Tato záložka jen zobrazuje již vypočtené hodnoty (citlivost, minimální cena,
    # bod zvratu, mapa ziskovosti) – žádný výpočet se zde nemění.
    days_text = fmt_days(expected_days, 0)

    # --- Co nejvíc mění měsíční zisk ---
    st.markdown("### Co nejvíc mění měsíční zisk")
    st.markdown(
        '<div class="de-intro">Graf ukazuje, jak moc se změní měsíční zisk, když se jeden údaj '
        "změní o 10 % nahoru nebo dolů a ostatní zůstanou stejné. Délka pruhu je rozdíl zisku "
        "mezi variantou −10 % a +10 %: čím delší pruh, tím větší vliv.</div>",
        unsafe_allow_html=True,
    )

    driver_short = {
        "Denní cena": "Denní cena",
        "Vytížení (dny/měsíc)": "Dny pronájmu",
        "Fixní náklad": "Fixní náklad",
        "Servisní náklad": "Servisní náklad",
    }
    driver_phrase = {
        "Denní cena": "denní cena",
        "Dny pronájmu": "počet dní pronájmu",
        "Fixní náklad": "fixní náklad",
        "Servisní náklad": "servisní náklad",
    }
    effect = {driver_short[d]: {} for d, _ in drivers}
    for label, value in sensitivity_data:
        drv, name = label.split(" – ", 1)
        effect[driver_short[drv]][name] = value
    # Pořadí podle největšího vlivu (při shodě zůstává pořadí z výpočtu).
    driver_order = sorted(effect, key=lambda d: max(effect[d].values()), reverse=True)
    top_effect = max(values) if values else 0.0
    top_drivers = [
        d for d in driver_order
        if top_effect > 0 and np.isclose(max(effect[d].values()), top_effect)
    ]

    fig_tornado = go.Figure()
    for name in (machine_1, machine_2):
        xs = [effect[d][name] for d in driver_order]
        fig_tornado.add_trace(
            go.Bar(
                x=xs, y=driver_order, orientation="h", name=name,
                marker=dict(color=MACHINE_COLOR[name], line=dict(color="#fff", width=2)),
                text=[fmt_czk(v) for v in xs], textposition="outside", cliponaxis=False,
                textfont=dict(color=ui.INK, size=13),
                customdata=[fmt_czk(v / 2) for v in xs],
                hovertemplate="<b>%{y}</b> · " + name
                + "<br>Změna o 10 % posune zisk o ±%{customdata}"
                + "<br>Rozdíl mezi −10 % a +10 %: %{x:,.0f} Kč<extra></extra>",
            )
        )
    fig_tornado.update_layout(barmode="group", bargap=0.28, bargroupgap=0.06)
    fig_tornado.update_yaxes(
        autorange="reversed", tickvals=driver_order,
        ticktext=[f"<b>{d}</b>" if d in top_drivers else d for d in driver_order],
    )
    fig_tornado.update_xaxes(
        range=[0, (top_effect or 1) * 1.55], tickformat=",.0f", nticks=5,
        title_text="Rozdíl měsíčního zisku (Kč)",
    )
    ui.readable_axes(ui.style_fig(fig_tornado, height=400))
    fig_tornado.update_yaxes(tickfont=dict(size=14, color=ui.INK))
    chart(fig_tornado)
    ui.hint(
        f"Počítáno pro odhad poptávky {days_text} pronájmu měsíčně (Pokročilé nastavení). "
        "Scénář trhu tuto část neovlivňuje. Tučně je zvýrazněn údaj s největším vlivem."
    )

    if top_drivers:
        top_machines = [
            name for name in (machine_1, machine_2)
            if any(np.isclose(effect[d][name], top_effect) for d in top_drivers)
        ]
        names = " a ".join(f"<b>{driver_phrase[d]}</b>" for d in top_drivers)
        where = (
            "u obou strojů" if len(top_machines) == 2
            else f"u stroje <b>{top_machines[0]}</b>"
        )
        verb = "ovlivňují" if len(top_drivers) > 1 else "ovlivňuje"
        which = "kteréhokoli z nich " if len(top_drivers) > 1 else ""
        text = (
            f"Nejvíce měsíční zisk {verb} {names} {where}: změna {which}o 10 % posune zisk "
            f"přibližně o <b>±{fmt_czk(top_effect / 2)}</b>."
        )
        if {"Denní cena", "Dny pronájmu"} <= set(top_drivers):
            text += (
                " Oba údaje mají stejný vliv, protože tržba je součinem ceny a počtu dní."
            )
        weakest = driver_order[-1]
        if weakest not in top_drivers:
            text += f" Nejmenší vliv má {driver_phrase[weakest]}."
        ui.insight(text)
    else:
        ui.insight("Při zadaných hodnotách změna žádného údaje o 10 % zisk nemění.")

    # --- Minimální denní cena pronájmu ---
    st.markdown("### Minimální denní cena pronájmu")

    def price_rows(required: float, actual: float) -> list:
        """Řádky karty: stav, minimum, aktuální cena a rozdíl (jen zobrazení)."""
        diff = actual - required
        ok = actual >= required
        chip = ui.status_chip(
            "Cena dostačuje" if ok else "Cena je pod minimem", "good" if ok else "bad"
        )
        if required > 0 and diff != 0:
            share = fmt_pct(abs(diff) / required * 100)
            sub = f"o {share} {'více' if diff > 0 else 'méně'} než minimum"
        else:
            sub = "přesně na minimu" if diff == 0 else ""
        if diff > 0:
            diff_html = f'<span class="de-pos">+{fmt_czk(diff)}</span>'
        else:
            diff_html = fmt_czk(diff)
        return [
            ("Stav", "aktuální cena oproti minimu", chip, "", False),
            ("Minimální denní cena", "pokryje fixní a servisní náklady",
             fmt_czk(required), "", False),
            ("Aktuální denní cena", "ceník stroje", fmt_czk(actual), "", False),
            ("Rezerva nad minimem" if ok else "Chybí do minima", "", diff_html, sub, diff < 0),
        ]

    ui.results_cards(
        [
            dict(name=name, color=MACHINE_COLOR[name], rows=price_rows(required, actual))
            for name, required, actual in (
                (machine_1, required_price_1, avg_daily_mini),
                (machine_2, required_price_2, avg_daily_dumper),
            )
        ],
        rec_name="",
        tag="",
    )
    if expected_days == 0:
        ui.insight(
            "Při odhadu 0 dní pronájmu nelze náklady pokrýt žádnou cenou. "
            "Minimální cena je v tomto případě počítána pro 1 den.",
            tone="bad",
        )
    ui.hint(
        "Minimální denní cena je cena, při které stroj právě pokryje své měsíční náklady "
        f"(zisk 0 Kč). Výpočet: (fixní náklad {fmt_czk(fixed_cost)} + průměrný servisní "
        f"náklad stroje) ÷ {days_text} pronájmu. Nepočítá s výpadky kvůli poruchám "
        "ani s kolísáním poptávky."
    )

    # --- Zisk podle počtu dní pronájmu ---
    st.markdown("### Zisk podle počtu dní pronájmu")
    be_values = {machine_1: break_even_mini, machine_2: break_even_dumper}
    be_finite = {k: v for k, v in be_values.items() if np.isfinite(v)}
    # Popisky míří od bodů ven (levý bod doleva nahoru, pravý doprava dolů), kde nevedou čáry.
    be_left = min(be_finite, key=be_finite.get) if be_finite else None

    fig_be = go.Figure()
    for name, curve in ((machine_1, profit_curve_mini), (machine_2, profit_curve_dumper)):
        fig_be.add_trace(
            go.Scatter(
                x=days_grid, y=curve, mode="lines", name=name,
                line=dict(color=MACHINE_COLOR[name], width=3),
                hovertemplate="%{x} dní → %{y:,.0f} Kč<extra>" + name + "</extra>",
            )
        )
    for name, be in be_finite.items():
        fig_be.add_vline(x=be, line=dict(color=MACHINE_COLOR[name], dash="dot", width=1.5))
        fig_be.add_trace(
            go.Scatter(
                x=[be], y=[0], mode="markers+text", showlegend=False,
                marker=dict(size=15, color=MACHINE_COLOR[name], line=dict(color="#fff", width=3)),
                text=[f"<b>{fmt_days(be, 1)}</b>"],
                textposition=(
                    ("top left" if be >= 7 else "top right") if name == be_left
                    else ("bottom right" if be <= 25 else "bottom left")
                ),
                textfont=dict(size=13, color=ui.INK), cliponaxis=False,
                hovertemplate=f"Bod zvratu: {fmt_days(be, 1)}<extra>{name}</extra>",
            )
        )
    fig_be.add_hline(y=0, line=dict(color=ui.INK, dash="dash", width=1.5))
    fig_be.update_xaxes(title_text="Počet dní pronájmu za měsíc", dtick=5)
    fig_be.update_yaxes(title_text="Měsíční zisk / ztráta (Kč)", tickformat=",.0f")
    chart(ui.readable_axes(ui.style_fig(fig_be, height=400)))
    ui.hint(
        "Čárkovaná vodorovná čára je nulový zisk. Body na ní označují bod zvratu – od tohoto "
        "počtu dní stroj pokryje fixní náklad. <b>Zjednodušený výpočet:</b> denní cena × počet dní − fixní náklad, "
        "bez servisních nákladů a poruch. Skutečný bod zvratu je proto o něco vyšší."
    )

    def be_days(be: float):
        """Celý počet dní, od kterého stroj pokryje fixní náklad (None = zvláštní případ)."""
        return int(np.ceil(be - 1e-9)) if np.isfinite(be) and 0 < be <= days_grid[-1] else None

    def be_sentence(name: str, be: float) -> str:
        if not np.isfinite(be):
            return f"U stroje <b>{name}</b> nelze bod zvratu určit."
        if be <= 0:
            return f"<b>{name}</b> pokrývá fixní náklad už od prvního dne pronájmu."
        if be > days_grid[-1]:
            return (
                f"<b>{name}</b> fixní náklad nepokryje ani při pronájmu po celý měsíc "
                f"(bod zvratu {fmt_days(be, 1)})."
            )
        return (
            f"<b>{name}</b> začne pokrývat fixní náklad přibližně od "
            f"<b>{be_days(be)} dní</b> pronájmu měsíčně."
        )

    d1, d2 = be_days(break_even_mini), be_days(break_even_dumper)
    if d1 is not None and d2 is not None:
        ui.insight(
            f"<b>{machine_1}</b> začne podle tohoto modelu pokrývat fixní náklad přibližně od "
            f"<b>{d1} dní</b> pronájmu měsíčně, <b>{machine_2}</b> přibližně od <b>{d2} dní</b>."
        )
    else:
        ui.insight(
            "Podle tohoto modelu: "
            + " ".join(be_sentence(name, be) for name, be in be_values.items())
        )

    # --- Mapa ziskovosti ---
    with st.expander("Mapa ziskovosti – kombinace ceny a počtu dní"):
        st.markdown(
            f'<div class="de-intro">Měsíční zisk stroje <b>{machine_1}</b> pro různé denní ceny '
            "(60–140 % ceníku) a počty dní pronájmu.</div>",
            unsafe_allow_html=True,
        )
        fig_heat = go.Figure(
            go.Heatmap(
                x=days_range, y=price_range, z=profit_matrix, zmid=0,
                colorscale=[[0, ui.BAD], [0.5, "#F1F3F5"], [1, ui.BLUE]],
                colorbar=dict(
                    title=dict(text="Zisk / ztráta (Kč)", side="top"),
                    orientation="h", y=-0.3, yanchor="top", x=0, xanchor="left",
                    len=1, thickness=12, tickformat=",.0f", nticks=5,
                ),
                hovertemplate="%{x} dní, %{y:,.0f} Kč/den<br>Zisk %{z:,.0f} Kč<extra></extra>",
            )
        )
        fig_heat.add_trace(
            go.Contour(
                x=days_range, y=price_range, z=profit_matrix, showscale=False,
                contours=dict(start=0, end=0, size=1, coloring="lines"),
                colorscale=[[0, ui.NAVY], [1, ui.NAVY]],
                line=dict(color=ui.NAVY, width=2, dash="dash"), hoverinfo="skip",
            )
        )
        point_in_map = (
            days_range[0] <= expected_days <= days_range[-1]
            and price_range[0] <= avg_daily_mini <= price_range[-1]
        )
        if point_in_map:
            fig_heat.add_trace(
                go.Scatter(
                    x=[expected_days], y=[avg_daily_mini], mode="markers+text",
                    marker=dict(size=16, color="#fff", line=dict(color=ui.INK, width=3)),
                    text=["<b>Vaše nastavení</b>"], textposition="top center",
                    textfont=dict(size=13, color=ui.INK), cliponaxis=False,
                    hovertemplate=(
                        f"Vaše nastavení: {days_text}, {fmt_czk(avg_daily_mini)}/den"
                        f"<br>Zisk {fmt_czk(profit_curve_mini[expected_days])}<extra></extra>"
                    ),
                )
            )
        fig_heat.update_xaxes(title_text="Počet dní pronájmu za měsíc", dtick=5)
        fig_heat.update_yaxes(title_text="Denní cena (Kč)", tickformat=",.0f")
        chart(ui.readable_axes(ui.style_fig(fig_heat, height=470, legend=False)))
        if profit_matrix.max() <= 0:
            ui.insight(
                "V zobrazeném rozsahu ceny a počtu dní není žádná zisková kombinace.",
                tone="bad",
            )
        ui.explain(
            [
                ("Jak mapu číst",
                 "Každé políčko je jedna kombinace: vodorovně počet dní pronájmu za měsíc, "
                 "svisle cena za den. <b>Modrá</b> znamená zisk, <b>červená</b> ztrátu. "
                 "Čím sytější barva, tím větší částka."),
                ("Přerušovaná čára",
                 "Hranice nulového zisku – tržby přesně pokryjí fixní náklad. Vpravo a nad "
                 "čárou stroj vydělává, vlevo a pod ní prodělává."),
                ("Bílý kroužek",
                 "Vaše současné nastavení: ceník stroje a odhad poptávky z Pokročilého "
                 "nastavení. Mění se spolu se vstupy v levém panelu."
                 if point_in_map else
                 f"Vaše nastavení ({days_text}) leží mimo zobrazený rozsah mapy "
                 f"({days_range[0]}–{days_range[-1]} dní), proto není vyznačeno."),
                ("Zjednodušení",
                 "Stejně jako graf zisku podle dní počítá mapa jen s fixním nákladem, "
                 "bez servisních nákladů a poruch."),
            ]
        )

    with st.expander("Jak se počítají údaje na této záložce"):
        ui.explain(
            [
                ("Vliv údajů (±10 %)",
                 f"Rychlý výpočet pro odhad poptávky {days_text}: zisk = denní cena × dny "
                 "− fixní náklad − průměrný servisní náklad. Vždy se změní jen jeden údaj. "
                 "Nejde o simulaci, proto ho neovlivní scénář trhu."),
                ("Minimální denní cena",
                 "(fixní náklad + průměrný servisní náklad) ÷ odhad poptávky ve dnech. "
                 "Zahrnuje servisní náklady."),
                ("Bod zvratu a mapa ziskovosti",
                 "Zjednodušeně: denní cena × dny − fixní náklad. Servisní náklady "
                 "nezahrnují, proto se mohou mírně lišit od minimální ceny."),
            ]
        )


# ==================================================
# ZÁLOŽKA 5 – NÁKUP STROJŮ (OPTIMALIZACE)
# ==================================================

with tab5:
    st.markdown(f"### Jaké stroje koupit za {fmt_czk(budget)}")
    ui.note(
        "Model zkouší kombinace 0–2 kusů u pěti strojů s nejvyšším očekávaným ziskem a vybere "
        "tu s nejvyšším měsíčním ziskem v rámci rozpočtu. Počítá s fixními náklady a průměrnou "
        "poptávkou jednotlivých strojů z demonstračních dat – nastavení fixního nákladu "
        "a scénáře trhu v postranním panelu se zde neuplatní. Rozpočet změníte v panelu."
    )
    st.write("")

    if best_combo is not None:
        active_machines = [
            (machines[i]["machine_type"], best_combo[i])
            for i in range(n_machines)
            if best_combo[i] > 0
        ]
        total_cost = sum(
            best_combo[i] * machines[i]["purchase_price"] for i in range(n_machines)
        )
        cols = st.columns(max(len(active_machines), 1))
        for col, (machine_name, count) in zip(cols, active_machines):
            col.metric(label=f"🚜 {machine_name}", value=f"{count} ks")

        col_p1, col_p2, col_p3 = st.columns(3)
        col_p1.metric("Celková investice", fmt_czk(total_cost))
        col_p2.metric("Očekávaný měsíční zisk sestavy", fmt_czk(best_profit))
        col_p3.metric("Zbývá z rozpočtu", fmt_czk(budget - total_cost))
        if total_cost > 0:
            st.caption(
                f"Měsíční zisk odpovídá přibližně {fmt_pct(best_profit / total_cost * 100, 2)} "
                "hodnoty investice."
            )
    else:
        st.warning("Při daném rozpočtu se nepodařilo najít vhodnou kombinaci strojů.")

    st.markdown("### Jak roste zisk s rozpočtem")
    current_idx = int(np.argmin(np.abs(np.array(budget_values) - budget)))
    fig_budget = go.Figure()
    fig_budget.add_trace(
        go.Scatter(
            x=budget_values, y=max_profits, mode="lines+markers", name="Maximální zisk",
            line=dict(color=ui.BLUE, width=2.5), marker=dict(size=9, color=ui.BLUE),
            hovertemplate="Rozpočet %{x:,.0f} Kč<br>Max. zisk %{y:,.0f} Kč/měs.<extra></extra>",
        )
    )
    fig_budget.add_trace(
        go.Scatter(
            x=[budget_values[current_idx]], y=[max_profits[current_idx]], mode="markers+text",
            name="Nejbližší zadanému rozpočtu", text=["Váš rozpočet"], textposition="bottom left",
            marker=dict(size=16, color=ui.ORANGE, line=dict(color="#fff", width=2)),
            hoverinfo="skip",
        )
    )
    fig_budget.update_xaxes(title_text="Investiční rozpočet (Kč)", tickformat=",.0f")
    fig_budget.update_yaxes(title_text="Maximální očekávaný měsíční zisk (Kč)")
    chart(ui.style_fig(fig_budget, height=380))
    st.caption(
        "Křivka je počítána pro rozpočty 0,5–3 mil. Kč. Oranžový bod je hodnota "
        "nejbližší vašemu rozpočtu."
    )

    with st.expander("🧮 Kontrola jinými výpočetními modely (LP, DP)"):
        if models_agree:
            st.success(
                f"Všechny tři rozhodovací modely se shodují: preferovanou variantou je **{mc_better}**."
            )
        else:
            st.warning(
                "Rozhodovací modely se neshodují. "
                "Výsledek je vhodné posoudit také podle rizika a citlivosti vstupů."
            )
        col_model1, col_model2, col_model3 = st.columns(3)
        col_model1.metric("Simulace (Monte Carlo)", mc_better)
        col_model2.metric("Lineární optimalizace (LP)", lp_better)
        col_model3.metric("Diskrétní optimalizace (DP)", dp_better)

        st.markdown(
            "**Jak modely fungují:** simulace pracuje s nejistotou a ukazuje rozpětí výsledků. "
            "LP a DP rozdělují 30 dní v měsíci mezi oba stroje při pevných předpokladech "
            "(plné vytížení, průměrné náklady) – jejich částky jsou proto teoretickým stropem, "
            "ne odhadem skutečného zisku."
        )
        col_lp1, col_lp2, col_lp3 = st.columns(3)
        col_lp1.metric(f"LP – dny pro {machine_1}", fmt_num(x1_opt, 1))
        col_lp2.metric(f"LP – dny pro {machine_2}", fmt_num(x2_opt, 1))
        col_lp3.metric("LP – teoretický zisk", fmt_czk(lp_profit))
        col_dp1, col_dp2, col_dp3 = st.columns(3)
        col_dp1.metric(f"DP – dny pro {machine_1}", fmt_num(best_d1))
        col_dp2.metric(f"DP – dny pro {machine_2}", fmt_num(best_d2))
        col_dp3.metric("DP – teoretický zisk", fmt_czk(best_dp_profit))


# ==================================================
# ZÁLOŽKA 6 – PLÁN ZAKÁZEK
# ==================================================

with tab6:
    st.markdown("### Plán zakázek a vytížení strojů")
    st.caption(
        "Zakázky ze zásobníku (fiktivní) jsou seřazeny podle očekávaného zisku, rizika "
        "a priority a naplánovány tak, aby se potřebné stroje nepřekrývaly."
    )

    plan = scheduled_jobs.copy()
    planned_mask = plan["schedule_status"].str.contains("Naplánováno")
    planned = plan[planned_mask].copy()

    # Zobrazení termínů a penále z existujících dat – plánovací algoritmus se nemění.
    planned["days_late"] = (planned["end_day"] - planned["deadline_days"]).clip(lower=0).astype(int)
    planned["penalty"] = planned["days_late"] * planned["delay_penalty_per_day"]
    planned["deadline_status"] = np.where(planned["days_late"] > 0, "Po termínu", "Včas")

    planned_count = int(planned_mask.sum())
    waiting_count = int(plan["schedule_status"].str.contains("Čeká").sum())
    total_plan_days = int(planned["end_day"].max()) if planned_count else 0
    late_count = int((planned["days_late"] > 0).sum())
    total_penalty = float(planned["penalty"].sum())

    col_j1, col_j2, col_j3, col_j4, col_j5 = st.columns(5)
    col_j1.metric("Naplánované zakázky", planned_count)
    col_j2.metric("Čekající zakázky", waiting_count)
    col_j3.metric("Délka plánu", f"{total_plan_days} dní")
    col_j4.metric("Po termínu", f"{late_count} z {planned_count}")
    col_j5.metric("Penále za zpoždění", fmt_czk(total_penalty))

    if late_count:
        st.warning(
            f"**{late_count} z {planned_count} zakázek skončí po smluvním termínu.** "
            f"Penále za zpoždění celkem {fmt_czk(total_penalty)} (fiktivní hodnoty). "
            "Současný plánovací algoritmus termíny zatím nezohledňuje – jeho úprava je "
            "součástí další etapy vývoje."
        )
    else:
        st.success("Všechny naplánované zakázky skončí do smluvního termínu.")

    # Doporučená prioritní zakázka (stejné pravidlo jako dosud)
    available_jobs = planned.copy()
    available_jobs["priority_score"] = available_jobs["priority"].map(priority_score).fillna(0)
    available_jobs["risk_score"] = available_jobs["risk_level"].map(risk_score).fillna(0)
    available_jobs["decision_score"] = (
        available_jobs["priority_score"] * 100000
        + available_jobs["expected_revenue"]
        - available_jobs["risk_score"] * 50000
    )
    ranked_jobs = available_jobs.sort_values(
        by=["decision_score", "risk_score", "expected_revenue"],
        ascending=[False, True, False],
    )
    best_job = ranked_jobs.iloc[0] if not ranked_jobs.empty else None

    st.markdown("### Harmonogram")
    gantt = planned.sort_values(["start_day", "end_day"])
    gantt["hover"] = [
        f"<b>{r.job_name}</b><br>Den {int(r.start_day)}–{int(r.end_day)} · termín den {int(r.deadline_days)}"
        f"<br>{r.deadline_status}"
        + (f" ({int(r.days_late)} dní, penále {fmt_czk(r.penalty)})" if r.days_late else "")
        + f"<br>Stroje: {r.required_machines.replace(';', ', ')}"
        f"<br>Priorita: {ui.PRIORITY_CZ.get(r.priority, r.priority)} · "
        f"Riziko: {ui.RISK_CZ.get(r.risk_level, r.risk_level)}"
        f"<br>Očekávaný zisk: {fmt_czk(r.expected_profit)}"
        for r in gantt.itertuples()
    ]
    fig_gantt = go.Figure()
    for status, color in (("Včas", ui.BLUE), ("Po termínu", ui.BAD)):
        part = gantt[gantt["deadline_status"] == status]
        if part.empty:
            continue
        fig_gantt.add_trace(
            go.Bar(
                y=part["job_name"], x=part["duration_days"], base=part["start_day"] - 1,
                orientation="h", name=status,
                marker=dict(color=color, line=dict(color="#fff", width=2)),
                hovertext=part["hover"], hoverinfo="text",
            )
        )
    # pořadí řádků podle začátku zakázky (jedna osa pro obě řady)
    fig_gantt.update_yaxes(categoryorder="array", categoryarray=list(gantt["job_name"]))
    fig_gantt.add_trace(
        go.Scatter(
            y=gantt["job_name"], x=gantt["deadline_days"], mode="markers", name="Smluvní termín",
            marker=dict(symbol="line-ns", size=18, line=dict(color=ui.NAVY, width=3)),
            hovertemplate="Termín: den %{x}<extra></extra>",
        )
    )
    if best_job is not None:
        fig_gantt.add_annotation(
            x=int(best_job["end_day"]), y=best_job["job_name"], text="⭐ Prioritní zakázka",
            showarrow=False, xanchor="left", xshift=8, font=dict(color=ui.NAVY, size=12),
        )
    fig_gantt.update_layout(barmode="overlay")
    fig_gantt.update_yaxes(autorange="reversed", title_text="")
    fig_gantt.update_xaxes(title_text="Den plánu", dtick=5, range=[0, max(total_plan_days, int(planned["deadline_days"].max())) + 4])
    chart(ui.style_fig(fig_gantt, height=max(360, 34 * len(gantt) + 80)))
    st.caption(
        "Modrá = zakázka skončí včas, červená = po termínu. Svislá čárka označuje smluvní termín."
    )

    st.markdown("### Detail zakázek")
    detail = gantt.assign(
        **{
            "Zakázka": gantt["job_name"],
            "Priorita": gantt["priority"].map(ui.PRIORITY_CZ),
            "Riziko": gantt["risk_level"].map(ui.RISK_CZ),
            "Potřebné stroje": gantt["required_machines"].str.replace(";", ", "),
            "Začátek (den)": gantt["start_day"].astype(int),
            "Konec (den)": gantt["end_day"].astype(int),
            "Termín (den)": gantt["deadline_days"].astype(int),
            "Zpoždění (dny)": gantt["days_late"],
            "Penále": gantt["penalty"].apply(fmt_czk),
            "Očekávaný výnos": gantt["expected_revenue"].apply(fmt_czk),
        }
    )[["Zakázka", "Priorita", "Riziko", "Potřebné stroje", "Začátek (den)", "Konec (den)",
       "Termín (den)", "Zpoždění (dny)", "Penále", "Očekávaný výnos"]]
    st.dataframe(detail, width="stretch", hide_index=True)

    st.markdown("### Vytížení strojů")
    machine_usage = []
    for machine in machine_capacity.keys():
        used_days = 0
        for _, job in scheduled_jobs.iterrows():
            required = [m.strip() for m in job["required_machines"].split(";")]
            if machine in required:
                used_days += int(job["duration_days"])
        utilization = used_days / total_plan_days * 100 if total_plan_days > 0 else 0
        machine_usage.append(
            {"Stroj": machine, "Využité dny": used_days, "Vytížení (%)": round(utilization, 1)}
        )
    machine_usage_df = pd.DataFrame(machine_usage).sort_values("Vytížení (%)", ascending=False)

    most_used_machine = machine_usage_df.iloc[0]
    bottleneck_machine = most_used_machine["Stroj"]
    bottleneck_utilization = most_used_machine["Vytížení (%)"]

    fig_usage = go.Figure(
        go.Bar(
            x=machine_usage_df["Vytížení (%)"], y=machine_usage_df["Stroj"], orientation="h",
            marker=dict(
                color=[ui.ORANGE if v >= 70 else ui.BLUE for v in machine_usage_df["Vytížení (%)"]],
                line=dict(color="#fff", width=2),
            ),
            text=[fmt_pct(v) for v in machine_usage_df["Vytížení (%)"]], textposition="outside",
            cliponaxis=False,
            customdata=machine_usage_df["Využité dny"],
            hovertemplate="%{y}<br>Vytížení %{x:.1f} %<br>Využité dny: %{customdata}<extra></extra>",
        )
    )
    fig_usage.add_vline(
        x=70, line=dict(color=ui.MUTED, dash="dash", width=1.5),
        annotation_text="70 % – hranice úzkého hrdla", annotation_position="top",
    )
    fig_usage.update_yaxes(autorange="reversed")
    fig_usage.update_xaxes(title_text="Vytížení v plánu (%)", range=[0, 110])
    chart(ui.style_fig(fig_usage, height=360, legend=False))

    if bottleneck_utilization >= 70:
        st.warning(
            f"🚨 Úzké hrdlo: **{bottleneck_machine}** je vytížen na "
            f"{fmt_pct(bottleneck_utilization, 0)} délky plánu. Tento stroj brzdí další zakázky."
        )
    else:
        st.success(
            f"✅ Kapacity jsou vyvážené. Nejvíce vytížený stroj je **{bottleneck_machine}** "
            f"({fmt_pct(bottleneck_utilization, 0)})."
        )

    if best_job is not None:
        st.markdown("### Doporučená prioritní zakázka")
        st.markdown(
            ui.machine_card(
                best_job["job_name"],
                ui.ORANGE,
                [
                    ("Očekávaný výnos", fmt_czk(best_job["expected_revenue"])),
                    ("Priorita", ui.PRIORITY_CZ.get(best_job["priority"], best_job["priority"])),
                    ("Riziko", ui.RISK_CZ.get(best_job["risk_level"], best_job["risk_level"])),
                    ("Termín v plánu", f"den {int(best_job['start_day'])}–{int(best_job['end_day'])}"),
                    ("Rozhodovací skóre", f"{fmt_num(best_job['decision_score'])} bodů"),
                ],
                recommended=True,
                tag="Priorita",
            ),
            unsafe_allow_html=True,
        )

        with st.expander("📊 Pořadí zakázek podle rozhodovacího skóre"):
            score_table = ranked_jobs[
                ["job_name", "priority", "risk_level", "expected_revenue", "decision_score"]
            ].copy()
            score_table["priority"] = score_table["priority"].map(ui.PRIORITY_CZ)
            score_table["risk_level"] = score_table["risk_level"].map(ui.RISK_CZ)
            score_table["expected_revenue"] = score_table["expected_revenue"].apply(fmt_czk)
            score_table["decision_score"] = score_table["decision_score"].apply(fmt_num)
            score_table.columns = ["Zakázka", "Priorita", "Riziko", "Očekávaný výnos", "Rozhodovací skóre"]
            st.dataframe(score_table, width="stretch", hide_index=True)
            st.caption(
                "Skóre = priorita × 100 000 + očekávaný výnos − riziko × 50 000. "
                "Pořadí v plánu používá jiné skóre (zisk upravený o riziko, priorita, "
                "dostupnost strojů, penále) – sjednocení je navrženo do další etapy."
            )

    with st.expander("⏱️ Ilustrativní úspora času"):
        baseline_days = 30
        saved_days = baseline_days - total_plan_days
        daily_profit = 15000
        col_i1, col_i2 = st.columns(2)
        col_i1.metric("Úspora proti referenčnímu plánu 30 dní", f"{saved_days} dní")
        col_i2.metric("Ilustrativní hodnota úspory", fmt_czk(saved_days * daily_profit))
        st.caption(
            "Ilustrativní výpočet s pevnými předpoklady: referenční plán 30 dní a hodnota "
            "15 000 Kč za každý ušetřený den. Nezohledňuje penále za zpoždění."
        )
