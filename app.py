import hmac
import importlib
import os
from io import BytesIO

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

import portfolio as pf
import scheduling as sch
import ui_theme as ui

# Streamlit při novém nasazení spustí app.py znovu, ale už načtené moduly (ui_theme,
# portfolio, scheduling) ponechá v paměti ve staré verzi. Explicitní reload zajistí, že se vždy
# použije aktuální soubor (jinak hrozí AttributeError u nově přidaných funkcí).
ui = importlib.reload(ui)
pf = importlib.reload(pf)
sch = importlib.reload(sch)
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
    '<div class="de-soon">Plán zakázek používá pevnou ukázkovou flotilu – vstupy v panelu '
    "ho zatím neovlivňují.</div>", unsafe_allow_html=True
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

# Pevný náhodný základ: stejné vstupy dají vždy stejné výsledky. Pro řízený přepočet
# s jiným základem stačí změnit tuto konstantu (případně ji později nabídnout v panelu).
SIMULATION_SEED = 42
rng = np.random.RandomState(SIMULATION_SEED)

# --- Stroj 1 ---
demand_mean_1 = row_1["demand_days_mean"]
mode_days_1 = np.clip(demand_mean_1, min_days, max_days)
util_days_1 = rng.triangular(min_days, mode_days_1, max_days, n)

service_cost_1 = rng.normal(
    row_1["service_cost_mean"],
    row_1["service_cost_std"],
    n,
)
service_cost_1 = np.clip(service_cost_1, 0, None)

failure_happens_1 = rng.rand(n) < row_1["failure_prob"]
failure_days_1 = np.where(
    failure_happens_1,
    rng.poisson(row_1["failure_days_mean"], n),
    0,
)
real_days_1 = np.maximum(util_days_1 - failure_days_1, 0)

price_std_1 = avg_daily_mini * 0.08
daily_price_1 = rng.normal(avg_daily_mini, price_std_1, n)
daily_price_1 = np.clip(daily_price_1, avg_daily_mini * 0.8, avg_daily_mini * 1.2)

rev_mini = real_days_1 * daily_price_1
profit_mini = rev_mini - fixed_cost - service_cost_1

# --- Stroj 2 ---
demand_mean_2 = row_2["demand_days_mean"]
mode_days_2 = np.clip(demand_mean_2, min_days, max_days)
util_days_2 = rng.triangular(min_days, mode_days_2, max_days, n)

service_cost_2 = rng.normal(
    row_2["service_cost_mean"],
    row_2["service_cost_std"],
    n,
)
service_cost_2 = np.clip(service_cost_2, 0, None)

failure_happens_2 = rng.rand(n) < row_2["failure_prob"]
failure_days_2 = np.where(
    failure_happens_2,
    rng.poisson(row_2["failure_days_mean"], n),
    0,
)
real_days_2 = np.maximum(util_days_2 - failure_days_2, 0)

price_std_2 = avg_daily_dumper * 0.08
daily_price_2 = rng.normal(avg_daily_dumper, price_std_2, n)
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

# --- Investiční scorecard (doplňkové bodové hodnocení) ---
# Bod získá jen stroj, který je v ukazateli skutečně lepší; při shodě nezíská bod nikdo.
# Hlavní doporučení (better_machine) určuje vyšší očekávaný zisk – body ho jen doplňují.


def _better_index(a, b, higher_is_better: bool):
    if a == b:
        return None
    return (0 if a > b else 1) if higher_is_better else (0 if a < b else 1)


def _payback_text(months) -> str:
    return f"{fmt_num(months, 1)} měsíce" if months is not None else "nelze určit (ztráta)"


score_items = [
    ("Očekávaný měsíční zisk", fmt_czk(expected_profit_mini), fmt_czk(expected_profit_dumper),
     _better_index(expected_profit_mini, expected_profit_dumper, True)),
    ("Šance ztrátového měsíce", fmt_pct(prob_loss_mini * 100), fmt_pct(prob_loss_dumper * 100),
     _better_index(prob_loss_mini, prob_loss_dumper, False)),
    ("Velmi nepříznivý měsíc (VaR 5 %)", fmt_czk(var95_mini), fmt_czk(var95_dumper),
     _better_index(var95_mini, var95_dumper, True)),
    ("Návratnost investice", _payback_text(roi_mini), _payback_text(roi_dumper),
     _better_index(roi_mini, roi_dumper, False)
     if roi_mini is not None and roi_dumper is not None else None),
]
scorecard_mini = sum(1 for *_, best in score_items if best == 0)
scorecard_dumper = sum(1 for *_, best in score_items if best == 1)

# --- Vysvětlení doporučení (stejné pravidlo jako hlavní karta v záložce Doporučení) ---
_names = (machine_1, machine_2)
_rec_i = 1 if better_machine == machine_2 else 0
_other = _names[1 - _rec_i]
_profits = (expected_profit_mini, expected_profit_dumper)
_points = (scorecard_mini, scorecard_dumper)

if expected_profit_mini < 0 and expected_profit_dumper < 0:
    _head = (
        f"**Proč ani jeden stroj není ziskový a {better_machine} vychází jen relativně lépe:**\n"
        f"- Oba stroje mají záporný očekávaný měsíční zisk; {better_machine} prodělává méně "
        f"({fmt_czk(_profits[_rec_i])} oproti {fmt_czk(_profits[1 - _rec_i])})."
    )
elif expected_profit_mini == expected_profit_dumper:
    _head = (
        "**Proč nelze dát jednoznačné doporučení:**\n"
        f"- Oba stroje mají stejný očekávaný měsíční zisk ({fmt_czk(expected_profit_mini)}). "
        "Rozhodněte podle rizika a pořizovací ceny."
    )
else:
    _head = (
        f"**Proč doporučujeme {better_machine}:**\n"
        f"- Má vyšší očekávaný měsíční zisk: {fmt_czk(_profits[_rec_i])} oproti "
        f"{fmt_czk(_profits[1 - _rec_i])} u stroje {_other}."
    )
if _points[1 - _rec_i] > _points[_rec_i]:
    _head += (
        f"\n- V doplňkovém bodovém hodnocení vychází lépe {_other} "
        f"({_points[1 - _rec_i]} : {_points[_rec_i]}) – je výhodnější v ukazatelích rizika "
        "nebo návratnosti. Zvažte, kolik rizika jste ochotni nést."
    )

_rows = "\n".join(
    f"- {label}: {machine_1} {v1} · {machine_2} {v2} – "
    + (f"lépe {_names[best]}" if best is not None else "shodné nebo nelze porovnat")
    for label, v1, v2, best in score_items
)
ai_text = f"""
{_head}

**Srovnání ukazatelů:**
{_rows}

**Rozpětí měsíčního výsledku (simulace):**
- Slabý měsíc (P10): {machine_1} {fmt_czk(p10_mini)} · {machine_2} {fmt_czk(p10_dumper)}
- Typický měsíc (medián): {machine_1} {fmt_czk(p50_mini)} · {machine_2} {fmt_czk(p50_dumper)}
- Dobrý měsíc (P90): {machine_1} {fmt_czk(p90_mini)} · {machine_2} {fmt_czk(p90_dumper)}

**Co doplnit před nákupem:**
- skutečnou poptávku po stroji {better_machine},
- servisní náklady, odstávky a pojištění,
- sezónnost (léto a zima).
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

# --- Nákup strojů: nejlepší sestava v rámci rozpočtu (výpočet v portfolio.py) ---
# Prohledají se všechny kombinace 0–2 kusů u všech typů strojů včetně „nic nekupovat“.
# Používají se individuální náklady strojů z dat, ne fixní náklad z postranního panelu.
purchase_options = pf.all_portfolios(df)
purchase_best = pf.best_portfolio(purchase_options, budget)
purchase_frontier = pf.budget_frontier(purchase_options)
purchase_min_profitable = pf.cheapest_profitable(purchase_options)

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

    doc.add_paragraph(f"Porovnání strojů: {machine_1} a {machine_2}")
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
        ("Bod zvratu (jen fixní náklady)", fmt_days(break_even_mini, 2), fmt_days(break_even_dumper, 2)),
        ("Šance ztrátového měsíce", fmt_pct(risk_loss_mini), fmt_pct(risk_loss_dumper)),
        ("Očekávaný měsíční zisk (simulace)", fmt_czk(expected_profit_mini), fmt_czk(expected_profit_dumper)),
    ]
    for metric, v1, v2 in rows:
        cells = table.add_row().cells
        cells[0].text = metric
        cells[1].text = v1
        cells[2].text = v2

    # Stejné tři případy jako hlavní karta v záložce Doporučení.
    doc.add_heading("Doporučení", level=1)
    verdict_p = doc.add_paragraph()
    if expected_profit_mini < 0 and expected_profit_dumper < 0:
        verdict_p.add_run("Žádný stroj není ziskový. ").bold = True
        verdict_p.add_run(
            f"Relativně lépe vychází {better_machine}, ale při zadaném nastavení jsou obě "
            "varianty v průměru ztrátové. Investici nedoporučujeme realizovat bez zvýšení ceny "
            "pronájmu, vyššího vytížení nebo snížení nákladů."
        )
    elif expected_profit_mini == expected_profit_dumper:
        verdict_p.add_run("Stroje vycházejí stejně. ").bold = True
        verdict_p.add_run("Oba mají stejný očekávaný zisk – rozhodněte podle rizika a ceny.")
    else:
        verdict_p.add_run(f"Doporučený stroj: {better_machine}. ").bold = True
        verdict_p.add_run(
            "Má vyšší očekávaný měsíční zisk v simulaci. Podrobné srovnání rizika a bodu "
            "zvratu je v tabulce výše."
        )
    doc.add_paragraph(
        f"Doplňkové bodové hodnocení (zisk, šance ztráty, velmi nepříznivý měsíc, návratnost; "
        f"při shodě bez bodu): {machine_1} {scorecard_mini} / 4, {machine_2} {scorecard_dumper} / 4."
    )

    doc.add_heading("Rozpětí možných výsledků (simulace)", level=1)
    doc.add_paragraph(
        f"Simulace vytváří {fmt_num(n)} možných měsíců s různou poptávkou, servisními náklady, "
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
        "Výběr nákupní sestavy (záložka Nákup strojů) a plán zakázek (záložka Plán zakázek) "
        "jsou v aplikaci samostatné analýzy na ukázkových datech. Tento report je nezahrnuje "
        "a navzájem zatím propojené nejsou."
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
        # Stejné tři případy jako hlavní karta výše (oba ztrátové / shoda / doporučení).
        if both_loss:
            st.markdown(
                f"Oba stroje jsou v simulaci {fmt_num(n)} možných měsíců **v průměru ztrátové**. "
                f"{better_machine} vychází jen relativně lépe – vydělá více než {other_machine} "
                f"přibližně v **{fmt_pct(rec_better_share * 100, 0)}** simulovaných měsíců."
            )
        elif rec_profit == other_profit:
            st.markdown(
                f"Oba stroje mají v simulaci {fmt_num(n)} možných měsíců **stejný průměrný "
                "měsíční zisk**."
            )
        else:
            st.markdown(
                f"Doporučený je stroj s **vyšším průměrným měsíčním ziskem** v simulaci "
                f"{fmt_num(n)} možných měsíců. {better_machine} vydělá více než {other_machine} "
                f"přibližně v **{fmt_pct(rec_better_share * 100, 0)}** simulovaných měsíců."
            )
        col_s1, col_s2 = st.columns(2)
        col_s1.metric(f"Bodové hodnocení – {machine_1}", f"{scorecard_mini} / 4")
        col_s2.metric(f"Bodové hodnocení – {machine_2}", f"{scorecard_dumper} / 4")
        ui.hint(
            "Doplňkové body za vyšší očekávaný zisk, nižší šanci ztrátového měsíce, lepší velmi "
            "nepříznivý měsíc (VaR 5 %) a rychlejší návratnost (jen když jsou oba stroje "
            "v zisku). Při shodě bod nezíská žádný stroj. Doporučení určuje očekávaný zisk; "
            "text níže sestavují pevná pravidla ze stejných výsledků."
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
        ui.hint(
            "Minimální cena pokryje fixní a průměrné servisní náklady při zadaném odhadu poptávky. "
            "Bod zvratu na hlavní stránce počítá jen s fixními náklady (fixní náklad ÷ denní cena)."
        )

    with st.expander(f"Rychlý odhad bez simulace ({expected_days} dní pronájmu)"):
        col_q1, col_q2 = st.columns(2)
        col_q1.metric(f"{machine_1} – měsíční výsledek", fmt_czk(expected_profit_mini_real))
        col_q2.metric(f"{machine_2} – měsíční výsledek", fmt_czk(expected_profit_dumper_real))
        ui.hint(
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
    ui.readable_axes(ui.style_fig(fig_tornado, height=450))
    fig_tornado.update_yaxes(tickfont=dict(size=14, color=ui.INK))
    # Vyšší horní okraj: ikony grafu (zoom, stažení…) mají vlastní řádek nad legendou
    # a nepřekrývají ji ani na mobilu, kde se legenda zalomí do dvou řádků.
    fig_tornado.update_layout(margin=dict(t=96))
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
    ui.readable_axes(ui.style_fig(fig_be, height=450))
    fig_be.update_layout(margin=dict(t=96))  # místo pro ikony grafu nad legendou
    chart(fig_be)
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
    st.markdown(
        '<div class="de-intro">Aplikace projde všechny kombinace strojů, které se vejdou do '
        "rozpočtu, a vybere tu s nejvyšším očekávaným měsíčním provozním ziskem. Pokud se "
        "žádný nákup nevyplatí, doporučí nic nekupovat.</div>",
        unsafe_allow_html=True,
    )

    def months_text(m) -> str:
        if m is None:
            return "–"
        return f"{fmt_num(m, 1)} měsíce"

    def portfolio_title(p: dict) -> str:
        return " · ".join(f"{k}× {name}" if k > 1 else name for name, k in p["counts"].items())

    pb = purchase_best
    if pb["units"] > 0:
        n_units = pb["units"]
        unit_word = "stroj" if n_units == 1 else "stroje" if n_units < 5 else "strojů"
        ui.hero(
            "Doporučená sestava",
            portfolio_title(pb),
            f"{n_units} {unit_word} za {fmt_czk(pb['investment'])} z rozpočtu {fmt_czk(budget)}.",
        )
        payback = pf.payback_months(pb)
        ui.kpi_cards(
            [
                dict(label="Celková investice", value=fmt_czk(pb["investment"]),
                     sub=f"zbývá {fmt_czk(budget - pb['investment'])} z rozpočtu"),
                dict(label="Očekávaný měsíční provozní zisk", value=fmt_czk(pb["profit"]),
                     sub=f"tržby {fmt_czk(pb['revenue'])} − náklady {fmt_czk(pb['costs'])}"),
                dict(label="Prostá návratnost", value=months_text(payback),
                     sub="investice ÷ měsíční provozní zisk"),
            ]
        )
        stress = pf.profit_with_demand(df, pb["counts"], 0.8)
        ui.insight(
            "Jde o <b>průměrný měsíc</b> při očekávané poptávce, ne o zaručený výsledek. "
            f"Pokud by poptávka byla o 20&nbsp;% nižší, zisk této sestavy by klesl na "
            f"<b>{fmt_czk(stress)}</b>.",
            tone="bad" if stress < 0 else "",
        )

        st.markdown("### Stroje v sestavě")
        ui.results_cards(
            [
                dict(
                    name=f"{name} – {line['units']} ks",
                    color=ui.NAVY,
                    rows=[
                        ("Investice", "pořizovací cena", fmt_czk(line["investment"]), "", False),
                        ("Dny pronájmu za měsíc", "poptávka po odečtení poruch",
                         fmt_num(line["days"], 1), "", False),
                        ("Tržby", "", fmt_czk(line["revenue"]), "", False),
                        ("Náklady", "fixní náklad a servis stroje", fmt_czk(line["costs"]), "", False),
                        ("Provozní zisk", "", fmt_czk(line["profit"]), "", line["profit"] < 0),
                    ],
                )
                for name, line in pb["lines"]
            ],
            rec_name="",
            tag="",
        )
    else:
        mp = purchase_min_profitable
        if mp is None:
            ui.hero(
                "Doporučení",
                "Nekupovat žádný stroj",
                "Žádná sestava strojů nemá při těchto datech kladný očekávaný provozní zisk.",
                tone="bad",
            )
        else:
            ui.hero(
                "Doporučení",
                "Zatím nekupovat",
                f"Rozpočet nestačí na žádnou ziskovou sestavu. Nejlevnější zisková volba "
                f"({portfolio_title(mp)}) stojí {fmt_czk(mp['investment'])}.",
                tone="bad",
            )

    # --- Graf: nejlepší zisk podle rozpočtu (stejný výpočet jako doporučení) ---
    st.markdown("### Jak roste zisk s rozpočtem")
    x_max = max(3_000_000, budget * 1.25)
    steps = [(x, p) for x, p in purchase_frontier if x <= x_max]
    xs = [x for x, _ in steps] + [x_max]
    ys = [p["profit"] for _, p in steps] + [steps[-1][1]["profit"]]
    step_labels = [portfolio_title(p) or "nic nekupovat" for _, p in steps]
    step_labels.append(step_labels[-1])
    fig_budget = go.Figure()
    fig_budget.add_trace(
        go.Scatter(
            x=xs, y=ys, mode="lines", line_shape="hv", name="Nejlepší sestava",
            line=dict(color=ui.BLUE, width=2.5), customdata=step_labels,
            hovertemplate="Rozpočet od %{x:,.0f} Kč<br>Zisk %{y:,.0f} Kč/měs."
            "<br>%{customdata}<extra></extra>",
        )
    )
    fig_budget.add_trace(
        go.Scatter(
            x=[budget], y=[pb["profit"]], mode="markers+text", name="Váš rozpočet",
            text=["<b>Váš rozpočet</b>"],
            textposition="top left" if budget > x_max * 0.5 else "top right",
            textfont=dict(size=13, color=ui.INK), cliponaxis=False,
            marker=dict(size=15, color=ui.ORANGE, line=dict(color="#fff", width=2)),
            hovertemplate=f"Váš rozpočet {fmt_czk(budget)}<br>Zisk {fmt_czk(pb['profit'])}/měs."
            "<extra></extra>",
        )
    )
    fig_budget.update_xaxes(title_text="Investiční rozpočet (Kč)", tickformat=",.0f",
                            range=[0, x_max], nticks=6)
    fig_budget.update_yaxes(title_text="Očekávaný měsíční provozní zisk (Kč)", tickformat=",.0f",
                            rangemode="tozero")
    chart(ui.readable_axes(ui.style_fig(fig_budget, height=380, legend=False)))
    ui.hint(
        "Každý schod znamená, že vyšší rozpočet umožní lepší sestavu. Oranžový bod je přesně "
        "váš rozpočet. Najetím na křivku zobrazíte, jaká sestava za daným ziskem stojí."
    )

    with st.expander("Předpoklady a metodika"):
        ui.explain(
            [
                ("Jak se sestava vybírá",
                 f"Prohledají se všechny kombinace 0–{pf.MAX_UNITS} kusů u všech "
                 f"{len(df)} typů strojů ({fmt_num(len(purchase_options))} možností, včetně "
                 "varianty nic nekupovat). Vybere se ta s nejvyšším měsíčním provozním ziskem, "
                 "která se vejde do rozpočtu; při shodě ta levnější. Ztrátová sestava se "
                 "nikdy nedoporučí."),
                ("Poptávka",
                 "Průměrná měsíční poptávka z dat je poptávka po celém typu stroje. Druhý kus "
                 "stejného typu ji nenásobí: pokryje jen dny, na které první kus nestačí "
                 "(nejvýše 30 dní na kus). Od dní se odečtou očekávané výpadky kvůli poruchám "
                 "(pravděpodobnost × průměrná délka)."),
                ("Náklady",
                 "Individuální měsíční fixní náklad a průměrný servis každého stroje z dat. "
                 "Fixní náklad z postranního panelu se zde nepoužívá – ostatní záložky s ním "
                 "počítají, proto se jejich částky liší. Data neuvádějí, zda fixní náklad "
                 "zahrnuje splátky."),
                ("Co model nezahrnuje",
                 "Společné provozní náklady půjčovny (prostory, mzdy, administrativa – v datech "
                 "nejsou), úroky, odpisy, daně, sezónnost a kolísání poptávky. Skutečný zisk "
                 "proto může být nižší."),
                ("Prostá návratnost",
                 "Za kolik měsíců by se investice vrátila z provozního zisku, bez úroků a "
                 "společných nákladů. Pokud fixní náklad obsahuje splátky, je návratnost "
                 "uvedena opatrně (spíše delší)."),
                ("Proč zde nejsou LP a DP",
                 "Dřívější srovnání lineární a diskrétní optimalizace se týkalo jen dvou strojů "
                 "z panelu a nemohlo potvrdit nákupní sestavu. Úplné prohledání všech kombinací "
                 "je samo přesnou diskrétní optimalizací."),
            ]
        )
        st.markdown('<div class="de-section">Ekonomika jednoho kusu podle dat</div>',
                    unsafe_allow_html=True)
        unit_rows = []
        for _, row in df.iterrows():
            one = pf.type_month(row, 1)
            unit_rows.append(
                {
                    "Stroj": row["machine_type"],
                    "Cena (Kč)": fmt_num(row["purchase_price"]),
                    "Poptávka (dny)": fmt_num(row["demand_days_mean"], 0),
                    "Tržby (Kč/měs.)": fmt_num(one["revenue"]),
                    "Náklady (Kč/měs.)": fmt_num(one["costs"]),
                    "Zisk (Kč/měs.)": fmt_num(one["profit"]),
                }
            )
        st.dataframe(pd.DataFrame(unit_rows), hide_index=True, width="stretch")


# ==================================================
# ZÁLOŽKA 6 – PLÁN ZAKÁZEK
# ==================================================

@st.cache_data(show_spinner="Hledám plán zakázek…")
def plan_jobs(jobs: pd.DataFrame, capacity: dict):
    """Nový plán (celočíselná optimalizace) a přínos jednotlivých zakázek pro plán."""
    result = sch.optimize_schedule(jobs, capacity)
    new_plan = sch.evaluate(jobs, result["starts"])
    contributions = sch.job_contributions(jobs, capacity, sch.summary(new_plan)["penalty"])
    return result, new_plan, contributions


with tab6:
    st.markdown("### Plán zakázek a vytížení strojů")
    st.markdown(
        '<div class="de-intro">Aplikace rozvrhne 12 zakázek ze zásobníku (fiktivní data) tak, '
        "aby žádný stroj nebyl ve stejný den využit víckrát, než kolik je kusů, a aby byl "
        "<b>zisk po penále co nejvyšší</b>.</div>",
        unsafe_allow_html=True,
    )

    plan_result, job_plan, job_contrib = plan_jobs(jobs_df, machine_capacity)
    plan_sum = sch.summary(job_plan)
    plan_conflicts = sch.conflicts(job_plan, machine_capacity)

    # Původní plán (dřívější pravidlo) – jen pro srovnání na stejných datech.
    old_starts = {
        idx: int(r["start_day"])
        for idx, r in scheduled_jobs.iterrows()
        if str(r["schedule_status"]).endswith("Naplánováno")
    }
    old_plan = sch.evaluate(jobs_df, old_starts)
    old_sum = sch.summary(old_plan)
    old_conflicts = sch.conflicts(old_plan, machine_capacity)

    ui.kpi_cards(
        [
            dict(label="Zisk po penále", value=fmt_czk(plan_sum["profit_after_penalty"]),
                 sub=f"zisk {fmt_czk(plan_sum['profit'])} − penále {fmt_czk(plan_sum['penalty'])}",
                 negative=plan_sum["profit_after_penalty"] < 0),
            dict(label="Zakázky po termínu", value=f"{plan_sum['late']} z {plan_sum['scheduled']}",
                 sub=f"penále celkem {fmt_czk(plan_sum['penalty'])}"),
            dict(label="Délka plánu", value=fmt_days(plan_sum["length"], 0),
                 sub=f"{plan_sum['scheduled']} naplánovaných zakázek"),
        ]
    )
    if plan_result["optimal"]:
        ui.insight(
            "Řešič <b>prokázal</b>, že při zadaných délkách zakázek, termínech a počtech strojů "
            "nemá žádný jiný plán nižší celkové penále. Nejde o záruku – skutečné zakázky se "
            "mohou zdržet nebo změnit."
        )
    else:
        ui.insight(
            "Zobrazený plán je <b>nejlepší nalezený</b>, optimalitu se v časovém limitu "
            "nepodařilo prokázat.",
            tone="bad",
        )
    if plan_sum["waiting"]:
        waiting = job_plan[~job_plan["scheduled"]]["job_name"].tolist()
        ui.insight(
            f"Nelze naplánovat (chybí potřebný stroj): <b>{', '.join(waiting)}</b>.", tone="bad"
        )

    st.markdown("### Srovnání s původním plánem")
    ui.compare_table(
        ["Původní plán", "Nový plán"],
        [
            ("Zakázky po termínu", [f"{old_sum['late']} z {old_sum['scheduled']}",
                                    f"{plan_sum['late']} z {plan_sum['scheduled']}"],
             1 if plan_sum["late"] < old_sum["late"] else None),
            ("Penále celkem", [fmt_czk(old_sum["penalty"]), fmt_czk(plan_sum["penalty"])],
             1 if plan_sum["penalty"] < old_sum["penalty"] else None),
            ("Délka plánu", [fmt_days(old_sum["length"], 0), fmt_days(plan_sum["length"], 0)],
             1 if plan_sum["length"] < old_sum["length"] else None),
            ("Zisk po penále", [fmt_czk(old_sum["profit_after_penalty"]),
                                fmt_czk(plan_sum["profit_after_penalty"])],
             1 if plan_sum["profit_after_penalty"] > old_sum["profit_after_penalty"] else None),
            ("Překročení kapacity strojů", [str(len(old_conflicts)), str(len(plan_conflicts))], None),
        ],
        rec_index=1,
        tag="Nový",
    )
    gain = plan_sum["profit_after_penalty"] - old_sum["profit_after_penalty"]
    ui.hint(
        f"Obě varianty plánují stejných {plan_sum['scheduled']} zakázek na stejných ukázkových "
        f"datech. Nový plán je o <b>{fmt_czk(gain)}</b> výhodnější jen díky jinému pořadí a "
        "využití volných dní strojů. Původní plán řadil zakázky podle skóre bez ohledu na "
        "termíny a stroj nevyužil, dokud nebyl volný pro další zakázku v pořadí."
    )

    # --- Harmonogram ---
    st.markdown("### Harmonogram")
    planned = job_plan[job_plan["scheduled"]].copy()
    planned["status"] = np.where(planned["days_late"] > 0, "Po termínu", "Včas")
    best_job = job_plan.loc[job_contrib.iloc[0]["job_index"]] if not job_contrib.empty else None
    gantt = planned.sort_values(["start_day", "end_day"])
    gantt["hover"] = [
        f"<b>{r.job_name}</b><br>Den {int(r.start_day)}–{int(r.end_day)} · termín den {r.deadline}"
        f"<br>{r.status}"
        + (f" ({r.days_late} dní, penále {fmt_czk(r.penalty)})" if r.days_late else "")
        + f"<br>Stroje: {', '.join(r.machines)}"
        f"<br>Zisk {fmt_czk(r.profit)} · po penále {fmt_czk(r.profit_after_penalty)}"
        for r in gantt.itertuples()
    ]
    fig_gantt = go.Figure()
    for status, color in (("Včas", ui.BLUE), ("Po termínu", ui.BAD)):
        part = gantt[gantt["status"] == status]
        if part.empty:
            continue
        fig_gantt.add_trace(
            go.Bar(
                y=part["job_name"], x=part["duration"], base=part["start_day"] - 1,
                orientation="h", name=status,
                marker=dict(color=color, line=dict(color="#fff", width=2)),
                hovertext=part["hover"], hoverinfo="text",
            )
        )
    fig_gantt.add_trace(
        go.Scatter(
            y=gantt["job_name"], x=gantt["deadline"], mode="markers", name="Termín",
            marker=dict(symbol="line-ns", size=18, line=dict(color=ui.NAVY, width=3)),
            hovertemplate="Termín: den %{x}<extra></extra>",
        )
    )
    fig_gantt.update_yaxes(categoryorder="array", categoryarray=list(gantt["job_name"]),
                           autorange="reversed", title_text="")
    x_end = max(plan_sum["length"], int(planned["deadline"].max()) if len(planned) else 0) + 1
    fig_gantt.update_xaxes(title_text="Den plánu", dtick=5, range=[0, x_end])
    fig_gantt.update_layout(barmode="overlay")
    ui.readable_axes(ui.style_fig(fig_gantt, height=max(420, 34 * len(gantt) + 140)))
    # Legenda pod grafem – na mobilu se tak nepřekrývá s ikonami grafu.
    fig_gantt.update_layout(legend=dict(orientation="h", yanchor="top", y=-0.1, x=0))
    chart(fig_gantt)
    ui.hint(
        "Modrá = zakázka skončí včas, červená = po termínu. Svislá čárka je smluvní termín "
        "(poslední den, kdy má být zakázka hotová)."
    )

    # --- Prioritní zakázka ---
    if best_job is not None:
        st.markdown("### Zakázka s největším přínosem pro plán")
        top = job_contrib.iloc[0]
        st.markdown(
            ui.machine_card(
                best_job["job_name"],
                ui.ORANGE,
                [
                    ("Zisk zakázky", fmt_czk(top["profit"])),
                    ("Penále, které zakázka přidá plánu", fmt_czk(top["added_penalty"])),
                    ("Čistý přínos pro plán", f"<b>{fmt_czk(top['contribution'])}</b>"),
                    ("Termín v plánu", f"den {int(best_job['start_day'])}–{int(best_job['end_day'])}"
                     f" (smluvní termín den {int(best_job['deadline'])})"),
                ],
                recommended=True,
                tag="Priorita",
            ),
            unsafe_allow_html=True,
        )
        ui.hint(
            "Čistý přínos = zisk zakázky − o kolik se kvůli ní zvýší nejnižší možné penále "
            "celého plánu (její vlastní zpoždění i zdržení ostatních zakázek). Tuto zakázku "
            "je proto nejméně výhodné ztratit nebo odložit."
        )
        with st.expander("Přínos všech zakázek pro plán"):
            ctab = job_contrib.assign(
                **{
                    "Zakázka": job_contrib["job_name"],
                    "Zisk": job_contrib["profit"].apply(fmt_czk),
                    "Přidané penále": job_contrib["added_penalty"].apply(fmt_czk),
                    "Čistý přínos": job_contrib["contribution"].apply(fmt_czk),
                }
            )[["Zakázka", "Zisk", "Přidané penále", "Čistý přínos"]]
            st.dataframe(ctab, width="stretch", hide_index=True)

    # --- Detail ---
    st.markdown("### Detail zakázek")
    detail = gantt.assign(
        **{
            "Zakázka": gantt["job_name"],
            "Stroje": gantt["machines"].apply(", ".join),
            "Začátek (den)": gantt["start_day"].astype(int),
            "Konec (den)": gantt["end_day"].astype(int),
            "Termín (den)": gantt["deadline"].astype(int),
            "Zpoždění (dny)": gantt["days_late"].astype(int),
            "Tržba": gantt["revenue"].apply(fmt_czk),
            "Náklady": gantt["costs"].apply(fmt_czk),
            "Zisk": gantt["profit"].apply(fmt_czk),
            "Penále": gantt["penalty"].apply(fmt_czk),
            "Zisk po penále": gantt["profit_after_penalty"].apply(fmt_czk),
            "Priorita": gantt["priority"].map(ui.PRIORITY_CZ),
            "Riziko": gantt["risk_level"].map(ui.RISK_CZ),
        }
    )[["Zakázka", "Stroje", "Začátek (den)", "Konec (den)", "Termín (den)", "Zpoždění (dny)",
       "Tržba", "Náklady", "Zisk", "Penále", "Zisk po penále", "Priorita", "Riziko"]]
    st.dataframe(detail, width="stretch", hide_index=True)

    # --- Vytížení strojů ---
    st.markdown("### Vytížení strojů")
    machine_usage_df = sch.utilization(job_plan, machine_capacity)
    bottleneck_machine = machine_usage_df.iloc[0]["machine"]
    bottleneck_utilization = float(machine_usage_df.iloc[0]["utilization"])
    fig_usage = go.Figure(
        go.Bar(
            x=machine_usage_df["utilization"],
            y=[f"{m} ({u} ks)" if u > 1 else m
               for m, u in zip(machine_usage_df["machine"], machine_usage_df["units"])],
            orientation="h",
            marker=dict(color=ui.BLUE, line=dict(color="#fff", width=2)),
            text=[fmt_pct(v) for v in machine_usage_df["utilization"]], textposition="outside",
            cliponaxis=False, customdata=machine_usage_df["used_days"],
            hovertemplate="%{y}<br>Vytížení %{x:.1f} %<br>Obsazené dny: %{customdata}<extra></extra>",
        )
    )
    fig_usage.update_yaxes(autorange="reversed")
    fig_usage.update_xaxes(title_text="Vytížení v plánu (%)", range=[0, 130])
    chart(ui.readable_axes(ui.style_fig(fig_usage, height=360, legend=False)))
    ui.hint(
        f"Vytížení = obsazené dny stroje ÷ (počet kusů × délka plánu {fmt_days(plan_sum['length'], 0)}). "
        f"Nejvytíženější je <b>{bottleneck_machine}</b> ({fmt_pct(bottleneck_utilization, 0)})."
    )

    with st.expander("Předpoklady a metodika"):
        ui.explain(
            [
                ("Cíl plánu",
                 "Co nejvyšší zisk po penále. Tržba, náklady a zisk zakázky na pořadí nezávisí, "
                 "takže při plánování všech zakázek je to totéž jako co nejnižší celkové penále. "
                 "Při stejném penále se zakázky dokončí co nejdříve."),
                ("Pojmy",
                 "<b>Tržba</b> = očekávaná tržba zakázky, <b>náklady</b> = přímé náklady zakázky "
                 "(z dat), <b>zisk</b> = tržba − náklady, <b>penále</b> = dny po termínu × denní "
                 "penále, <b>zisk po penále</b> = zisk − penále."),
                ("Stroje",
                 "Každá zakázka potřebuje po celou dobu všechny své stroje. V jednom dni lze stroj "
                 "použít tolikrát, kolik je kusů (vibrační deska 2 ks, ostatní 1 ks). Volné dny "
                 "stroje může využít kterákoli zakázka."),
                ("Jak se plán hledá",
                 "Celočíselná optimalizace (SciPy, řešič HiGHS) projde všechny přípustné začátky "
                 "zakázek. Pokud řešič optimum v časovém limitu neprokáže, aplikace to uvede."),
                ("Co model nezahrnuje",
                 "Víkendy a svátky, pozdější možný začátek zakázky, přejezdy a obsluhu strojů, "
                 "pronájmy strojů mimo zakázky, nejistotu délky zakázek a riziko (priorita "
                 "a riziko jsou jen informativní)."),
            ]
        )
