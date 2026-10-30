import contextlib
import hashlib
import html
import importlib
import os

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

import auth
import capacity as cp
import capacity_view as cv
import datamodel as dm
import econ_formula as ef
import fleet_economics as fe
import fleet_usage as fu
import fleet_view as fv
import formatting
import investice as inv
import navigace as navreg
import planner as pl
import planner_view as pv
import rental_view as rv
import portfolio as pf
import report_docx as rd
import scheduling as sch
import ui_theme as ui

# Streamlit při novém nasazení spustí app.py znovu, ale už načtené moduly (ui_theme, auth,
# datamodel, portfolio, scheduling, investice, report_docx…) ponechá v paměti ve staré verzi.
# Explicitní reload zajistí, že se vždy použije aktuální soubor (jinak hrozí AttributeError
# u nově přidaných funkcí).
formatting = importlib.reload(formatting)   # dřív než ui_theme a report_docx, které z něj importují
ui = importlib.reload(ui)
auth = importlib.reload(auth)
dm = importlib.reload(dm)
pf = importlib.reload(pf)
sch = importlib.reload(sch)
cp = importlib.reload(cp)
cv = importlib.reload(cv)
pl = importlib.reload(pl)
ef = importlib.reload(ef)
fu = importlib.reload(fu)                   # pořadí: využití → ekonomika → zobrazení (závisí na sobě)
fe = importlib.reload(fe)
fv = importlib.reload(fv)
pv = importlib.reload(pv)
rv = importlib.reload(rv)
inv = importlib.reload(inv)
rd = importlib.reload(rd)
navreg = importlib.reload(navreg)
from ui_theme import fmt_czk, fmt_days, fmt_num, fmt_pct

# ==================================================
# POMOCNÉ FUNKCE
# ==================================================


def format_czk(value) -> str:
    """Naformátuje částku v českém formátu s jednotkou Kč."""
    return fmt_czk(value)


# Semafor rizika (hranice 5 % a 15 %) – výpočetní pravidlo je v modulu investice.
risk_light = inv.risk_light


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


@st.cache_data(max_entries=4)
def load_data_model(data_version: str, code_version: str):
    """Společná datová vrstva (data/model/, fiktivní ukázka) a výsledek její kontroly.

    data_version = otisk obsahu datových souborů (dm.data_fingerprint). Je součástí klíče
    mezipaměti: Streamlit Cloud při novém nasazení neukončí běžící proces, takže bez otisku by
    se vracela data načtená před nasazením. Při změně dat se tak načtou znovu bez restartu.
    code_version = otisk zdrojového kódu modulů, které uložené funkce volají (CODE_MODULES) –
    totéž pro nasazení nového kódu (mezipaměť jinak pozná jen změnu kódu této funkce).

    Do mezipaměti se ukládají jen obyčejné tabulky a n-tice – ne objekty tříd z modulu
    datamodel. Ten se při každém běhu znovu načítá (importlib.reload); kdyby se to stalo
    souběžně v jiné relaci, uložení objektu staré třídy do mezipaměti by selhalo.
    """
    model = dm.load_model()
    issues = [(i.severity, i.table, i.message) for i in dm.validate(model)]
    return {name: getattr(model, name) for name in dm.SCHEMA}, issues


def load_data() -> pd.DataFrame:
    """Data o strojích ve stejném tvaru jako dříve (převod ze společné datové vrstvy)."""
    return dm.legacy_machines(data_model)


# ==================================================
# KONFIGURACE APLIKACE A PŘÍSTUP
# ==================================================

st.set_page_config(
    page_title="Decision Engine – Investice a provoz půjčovny",
    page_icon="📊",
    layout="wide",
)
ui.inject_css()

APP_TITLE = "Decision Engine"
APP_SUBTITLE = "Investice a provoz půjčovny"
APP_FOCUS = "Půjčovna malostrojů"
DEMO_BADGE = "Ukázková data"

# Přihlášení (modul auth.py). Údaje jsou jen ve Streamlit secrets:
# - jen klíč `password` → dosavadní přihlášení jedním heslem (beze změny),
# - sekce [users] → samostatné účty (jméno + heslo, v secrets jen hash hesla).
try:
    AUTH = auth.load_config(st.secrets)
except Exception:
    AUTH = auth.load_config({})


INV_KEEP = "_keep_"   # předpona uložených investičních vstupů (viz _inv_input)


def _log_out() -> None:
    """Odhlásí uživatele. Přihlašovací pole dostanou nový klíč, aby prohlížeč znovu
    neodeslal dříve vyplněné heslo (jinak by uživatele hned přihlásil zpět)."""
    for key in ("authenticated", "user_id", "user_name", "user_role", "admin_pw1", "admin_pw2"):
        st.session_state.pop(key, None)
    nonce = st.session_state.get("login_nonce", 0)
    for key in (f"login_user_{nonce}", f"login_password_{nonce}"):
        st.session_state.pop(key, None)
    st.session_state["login_nonce"] = nonce + 1
    for key in [k for k in st.session_state if str(k).startswith((INV_KEEP, "inv_"))]:   # vstupy předch. uživatele
        del st.session_state[key]


if not auth.can_log_in(AUTH):
    if AUTH["mode"] == "missing":
        st.error(
            "Přístupové heslo není nastaveno. Doplňte klíč `password` "
            "do Streamlit secrets (Settings → Secrets)."
        )
    else:
        st.error(
            "Přihlašovací účty nejsou správně nastavené – žádný aktivní účet nemá platné heslo. "
            "Zkontrolujte sekci `[users]` ve Streamlit secrets (Settings → Secrets)."
        )
    st.stop()

# Změna secrets (např. deaktivace účtu) se projeví při další akci přihlášeného uživatele.
if st.session_state.get("authenticated") and not auth.still_allowed(
    AUTH, st.session_state.get("user_id", auth.SINGLE_USER.username)
):
    _log_out()

# Po úspěšném přihlášení se přihlašovací pole skryje (stav je uložen v relaci).
if not st.session_state.get("authenticated", False):
    users_mode = AUTH["mode"] == "users"
    _, login_col, _ = st.columns([1, 1.2, 1])
    with login_col:
        ui.login_intro(
            APP_TITLE, APP_SUBTITLE,
            "Přihlaste se svým uživatelským jménem a heslem." if users_mode
            else "Pro pokračování zadejte přístupové heslo.",
        )
        nonce = st.session_state.get("login_nonce", 0)
        username = ""
        if users_mode:
            username = st.text_input("Uživatelské jméno", key=f"login_user_{nonce}")
        password = st.text_input(
            "Heslo" if users_mode else "Přístupové heslo",
            type="password",
            key=f"login_password_{nonce}",
            help="Po zadání stiskněte Enter." if users_mode
            else "Heslo je nastaveno ve Streamlit secrets. Po zadání stiskněte Enter.",
        )
        # Tlačítko jen potvrdí zadané údaje (stejně jako Enter).
        st.button("Přihlásit se", type="primary", width="stretch")
        if password and (username or not users_mode):
            user = auth.authenticate(AUTH, username, password)
            if user is not None:
                _log_out()
                st.session_state.update(
                    authenticated=True, user_id=user.username, user_name=user.name,
                    user_role=user.role,
                )
                st.rerun()
            st.error(auth.GENERIC_ERROR if users_mode else "Nesprávné heslo.")
        st.caption(
            "Pilotní prototyp. Všechny stroje, ceny, náklady a zakázky jsou fiktivní ukázková data."
        )
    st.stop()


# ==================================================
# HLAVIČKA APLIKACE
# ==================================================

ui.header(APP_TITLE, APP_SUBTITLE, DEMO_BADGE)


# ==================================================
# NAVIGACE (etapa M2) – levé menu se sekcemi
# ==================================================
# Stránky registruje navigace.py; soubory stranky/stranka_*.py jsou jen značky. Obsah aktivní
# stránky vykresluje tento skript (níže podle PAGE) – všechny stránky tak sdílejí jedno přihlášení,
# načtení dat i výpočty. Bez přihlášení se sem běh nedostane (st.stop výše), takže žádná stránka
# není dostupná ani přímou adresou.
_PAGE_OBJ = {
    key: st.Page(navreg.page_file(key), title=title, icon=icon, url_path=url,
                 default=(key == navreg.DEFAULT_PAGE))
    for key, _, title, icon, _, url in navreg.PAGES
}
nav = st.navigation({group: [_PAGE_OBJ[k] for k in keys] for group, keys in navreg.groups().items()},
                    position="sidebar")
PAGE = next(k for k, page in _PAGE_OBJ.items() if page is nav)
SHOW_INV_INPUTS = PAGE == "investice"          # investiční vstupy jen v sekci Investice


def page_link(key: str, label: str = None) -> None:
    """Odkaz na jinou sekci aplikace."""
    title = next(p[2] for p in navreg.PAGES if p[0] == key)
    st.page_link(_PAGE_OBJ[key], label=label or title, icon=_PAGE_OBJ[key].icon)


def _inv_input(container, widget: str, key: str, label: str, default_value, default_kwargs: dict, **kwargs):
    """Investiční vstup, jehož hodnota přežije přepnutí stránky (etapa M2).

    Vstupy jsou vidět jen v sekci Investice. Streamlit stav widgetu, který se na stránce nevykreslí,
    smaže – proto se hodnota ukládá i pod vlastní klíč INV_KEEP + key a po návratu se z něj obnoví.
    Mimo Investice se jen čte uložená hodnota (nebo původní výchozí hodnota). Výchozí hodnoty polí
    se nemění; default_kwargs = původní parametr výchozí hodnoty (index / value).
    """
    store = INV_KEEP + key
    if not SHOW_INV_INPUTS:
        return st.session_state.get(store, default_value)
    options = kwargs.get("options")
    if store in st.session_state and options is not None and st.session_state[store] not in options:
        del st.session_state[store]                  # např. po změně dat – vrátit se k výchozí hodnotě
    if key not in st.session_state and store in st.session_state:
        st.session_state[key] = st.session_state[store]
    if key not in st.session_state:
        kwargs.update(default_kwargs)
    value = getattr(container, widget)(label, key=key, **kwargs)
    st.session_state[store] = value
    return value


# ==================================================
# DATA A SIDEBAR VSTUPY
# ==================================================

# Postranní panel: jen označení aplikace, ukázková data, navigace a přihlášený uživatel.
# Investiční vstupy jsou na stránce Investice v panelu „Nastavení investičního scénáře“ nad výsledky.
# Jeho části se vytvoří dopředu, aby šly plnit na místech, kde je kód potřebuje – pořadí výpočtů
# tak zůstává beze změny. Každý vstup existuje jen jednou (stejné klíče jako dosud).
ui.sidebar_brand(APP_TITLE, APP_FOCUS)
st.sidebar.markdown(
    '<div class="de-demo"><b>Ukázková data</b> · všechny hodnoty jsou fiktivní</div>',
    unsafe_allow_html=True,
)
INV_PANEL_TITLE = "Nastavení investičního scénáře"
INV_TAB = navreg.INVESTICE_TABS[0]       # aktivní záložka sekce Investice
if SHOW_INV_INPUTS:
    # Pořadí na stránce: lišta záložek → panel nastavení → obsah aktivní záložky.
    # Lišta (st.tabs) slouží jen k přepínání; obsah všech záložek se počítá a vykresluje jako dřív
    # do kontejnerů pod panelem a viditelný je jen ten aktivní (ostatní skryje CSS – jako u st.tabs).
    # Při dvou stejných strojích se výsledky ani záložky nezobrazují (viz kontrola níže).
    _m1, _m2 = (st.session_state.get(k, st.session_state.get(INV_KEEP + k)) for k in ("inv_machine_1", "inv_machine_2"))
    if _m1 is None or _m1 != _m2:
        _inv_tab_bar = st.tabs(navreg.INVESTICE_TABS, key="inv_tab", on_change="rerun")
        INV_TAB = next((label for label, t in zip(navreg.INVESTICE_TABS, _inv_tab_bar) if t.open), INV_TAB)
    with st.expander(INV_PANEL_TITLE, icon=":material/tune:"):
        with st.container(key="de_inv_inputs"):
            _row_1 = st.columns(2, gap="medium")
            sb_machines, sb_costs = _row_1[0].container(), _row_1[1].container()
            _row_2 = st.columns(2, gap="medium")
            sb_market, sb_budget = _row_2[0].container(), _row_2[1].container()
            sb_advanced = st.container()
            sb_jobs = st.container()
    inv_summary_slot = st.empty()        # shrnutí zvolených hodnot – vidět i při sbaleném panelu
    inv_tab_area = st.container(key="de_inv_obsah")      # obsah záložek (pod panelem)
else:
    sb_machines = sb_costs = sb_market = sb_budget = sb_jobs = sb_advanced = inv_summary_slot = None
sb_user = st.sidebar.container()

# --- Přihlášený uživatel, odhlášení a nástroj správce (jen role admin) ---
with sb_user:
    st.markdown(
        f'<div class="de-user">Přihlášen: <b>{html.escape(st.session_state.get("user_name", ""))}</b>'
        f' · {html.escape(auth.ROLES.get(st.session_state.get("user_role", ""), ""))}</div>',
        unsafe_allow_html=True,
    )
    st.button("Odhlásit se", key="logout", on_click=_log_out, width="stretch")

# Verze dat a kódu jako součást klíčů mezipaměti (≈ 1 ms na běh) – viz load_data_model.
CODE_MODULES = (dm, sch, cp, cv, pl, ef, pv, fu, fe, fv)
data_version = dm.data_fingerprint()
code_version = dm.files_fingerprint([m.__file__ for m in CODE_MODULES])
_tables, _issues = load_data_model(data_version, code_version)
data_model = dm.DataModel(**_tables)
data_issues = [dm.Issue(*issue) for issue in _issues]
if dm.errors(data_issues):
    st.error(
        "Data aplikace neprošla kontrolou, výpočty proto nelze spustit:\n\n"
        + "\n".join(f"- {i}" for i in dm.errors(data_issues))
    )
    st.stop()

df = load_data()

jobs_df = dm.legacy_jobs(data_model)

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
if SHOW_INV_INPUTS:
    sb_jobs.markdown(
        '<div class="de-soon">Plán zakázek používá pevnou ukázkovou flotilu – tato nastavení '
        "ho zatím neovlivňují.</div>", unsafe_allow_html=True
    )

# Počet kusů každého typu stroje ze společné datové vrstvy (tabulka fleet_units).
machine_capacity = dm.fleet_capacity(data_model)

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

if SHOW_INV_INPUTS:
    sb_machines.markdown("### 1 · Porovnávané stroje")
machine_1 = _inv_input(
    sb_machines, "selectbox", "inv_machine_1", "Stroj 1", machine_options[0], dict(index=0),
    options=machine_options, help="První stroj, o jehož nákupu uvažujete.",
)
machine_2 = _inv_input(
    sb_machines, "selectbox", "inv_machine_2", "Stroj 2", machine_options[1], dict(index=1),
    options=machine_options, help="Druhý stroj pro srovnání. Musí se lišit od stroje 1.",
)

if machine_1 == machine_2 and PAGE in ("investice", "dashboard"):
    # Ostatní sekce investiční výsledky nezobrazují – nemají se kvůli tomu zablokovat.
    st.warning(
        f"Pro porovnání vyberte v panelu „{INV_PANEL_TITLE}“ dva **různé** stroje."
    )
    if PAGE != "investice":
        page_link("investice", "Otevřít sekci Investice a změnit stroje")
    st.stop()

row_1 = df[df["machine_type"] == machine_1].iloc[0]
row_2 = df[df["machine_type"] == machine_2].iloc[0]

avg_daily_mini = float(row_1["daily_price"])
avg_daily_dumper = float(row_2["daily_price"])

if SHOW_INV_INPUTS:
    sb_costs.markdown("### 2 · Náklady")
fixed_cost = _inv_input(
    sb_costs, "number_input", "inv_fixed_cost",
    "Měsíční fixní náklady na jeden stroj", 100000, dict(value=100000),
    min_value=0,
    step=5000,
    help="Částka v Kč, kterou stroj stojí každý měsíc bez ohledu na počet dní pronájmu "
    "(např. splátka, pojištění). Odečítá se od výsledku každého z obou porovnávaných strojů. "
    "Servisní náklady se počítají zvlášť z údajů o stroji. Výchozí hodnota je ukázková.",
)
if SHOW_INV_INPUTS:
    ui.sidebar_amount(sb_costs, fmt_czk(fixed_cost), "měsíčně na jeden stroj")

if SHOW_INV_INPUTS:
    sb_market.markdown("### 3 · Trh")
scenario = _inv_input(
    sb_market, "selectbox", "inv_scenario",
    "Scénář trhu", "Realistický", dict(index=1),
    options=["Pesimistický", "Realistický", "Optimistický"],
    help="Určuje, kolik dní v měsíci může být stroj pronajatý. "
    "Pesimistický 8–15 dní, realistický 15–25 dní, optimistický 22–30 dní.",
)

min_days, max_days = inv.scenario_days(scenario)          # rozsah dní pronájmu měsíčně
if scenario == "Pesimistický":
    scenario_note = "Slabší poptávka, opatrný scénář."
elif scenario == "Realistický":
    scenario_note = "Běžný provoz firmy."
else:
    scenario_note = "Silná poptávka a vysoké vytížení."

if SHOW_INV_INPUTS:
    sb_market.caption(
        f"{scenario_note} Simulace počítá s {min_days}–{max_days} dny pronájmu měsíčně."
    )

if SHOW_INV_INPUTS:
    sb_advanced.markdown("### Pokročilé nastavení")
expected_days = _inv_input(
    sb_advanced, "slider", "inv_expected_days",
    "Odhad poptávky pro rychlý výpočet (dny/měsíc)", 18, dict(value=18), min_value=0, max_value=31,
    help="Používá se pro rychlý výpočet bez simulace (karta „Rychlý odhad“), "
    "pro citlivost a pro minimální cenu pronájmu. Simulace pracuje s rozsahem "
    "dní podle zvoleného scénáře trhu.",
)
n = _inv_input(
    sb_advanced, "slider", "inv_n",
    "Počet simulací (měsíců)", 1000, dict(value=1000), min_value=100, max_value=10000,
    help="Kolik možných měsíců aplikace nasimuluje. Více simulací = stabilnější "
    "výsledky, ale pomalejší přepočet.",
)

if SHOW_INV_INPUTS:
    sb_budget.markdown("### 4 · Investiční rozpočet")
budget = _inv_input(
    sb_budget, "number_input", "inv_budget",
    "Investiční rozpočet", 3000000, dict(value=3000000),
    min_value=100000,
    step=100000,
    help="Kolik peněz je k dispozici na nákup strojů. Využívá záložka „Nákup strojů“.",
)
if SHOW_INV_INPUTS:
    ui.sidebar_amount(sb_budget, fmt_czk(budget), "k dispozici na nákup strojů")
    inv_summary_slot.markdown(
        f'<div class="de-inv-summary"><b>{html.escape(machine_1)}</b> × <b>{html.escape(machine_2)}</b> · '
        f"trh {html.escape(scenario.lower())} · fixní náklad {fmt_czk(fixed_cost)} měsíčně · "
        f"rozpočet {fmt_czk(budget)}</div>",
        unsafe_allow_html=True,
    )


# ==================================================
# INVESTIČNÍ VÝPOČTY (záložky 1–4) – modul investice.py
# ==================================================
# Simulace Monte Carlo, metriky, LP/DP, bodové hodnocení, citlivost, mapa ziskovosti a minimální
# cena se počítají v modulu investice (beze změny vzorců i pořadí losování). Zde se výsledky jen
# převezmou pod stejnými názvy, s jakými pracují záložky a export.
SIMULATION_SEED = inv.SIMULATION_SEED
_inv = inv.evaluate(row_1, row_2, machine_1, machine_2, fixed_cost, min_days, max_days, expected_days, n,
                    seed=SIMULATION_SEED)

# Simulace a základní metriky
profit_mini, profit_dumper = _inv["profit_mini"], _inv["profit_dumper"]
expected_profit_mini_real = _inv["expected_profit_mini_real"]
expected_profit_dumper_real = _inv["expected_profit_dumper_real"]
prob_loss_mini, prob_loss_dumper = _inv["prob_loss_mini"], _inv["prob_loss_dumper"]
risk_loss_mini, risk_loss_dumper = _inv["risk_loss_mini"], _inv["risk_loss_dumper"]
break_even_mini, break_even_dumper = _inv["break_even_mini"], _inv["break_even_dumper"]
expected_profit_mini, expected_profit_dumper = _inv["expected_profit_mini"], _inv["expected_profit_dumper"]
p10_mini, p10_dumper = _inv["p10_mini"], _inv["p10_dumper"]
p50_mini, p50_dumper = _inv["p50_mini"], _inv["p50_dumper"]
p90_mini, p90_dumper = _inv["p90_mini"], _inv["p90_dumper"]
var95_mini, var95_dumper = _inv["var95_mini"], _inv["var95_dumper"]
ci_mini, ci_dumper = _inv["ci_mini"], _inv["ci_dumper"]
prob_dumper_better = _inv["prob_dumper_better"]
roi_mini, roi_dumper = _inv["roi_mini"], _inv["roi_dumper"]
profit_diff, risk_diff = _inv["profit_diff"], _inv["risk_diff"]
better_machine = _inv["better_machine"]

# LP a DP, shoda modelů
x1_opt, x2_opt, lp_profit = _inv["x1_opt"], _inv["x2_opt"], _inv["lp_profit"]
best_d1, best_d2, best_dp_profit = _inv["best_d1"], _inv["best_d2"], _inv["best_dp_profit"]
mc_better, lp_better, dp_better = _inv["mc_better"], _inv["lp_better"], _inv["dp_better"]

# Doplňkové bodové hodnocení
scorecard_mini, scorecard_dumper = _inv["scorecard_mini"], _inv["scorecard_dumper"]

# Citlivost, bod zvratu, riziko × výnos, mapa ziskovosti, minimální cena
days_grid = _inv["days_grid"]
profit_curve_mini, profit_curve_dumper = _inv["profit_curve_mini"], _inv["profit_curve_dumper"]
drivers = _inv["drivers"]
sensitivity_data, labels, values = _inv["sensitivity_data"], _inv["labels"], _inv["values"]
rr_x1, rr_x2, rr_y1, rr_y2 = _inv["rr_x1"], _inv["rr_x2"], _inv["rr_y1"], _inv["rr_y2"]
days_range, price_range, profit_matrix = _inv["days_range"], _inv["price_range"], _inv["profit_matrix"]
required_price_1, required_price_2 = _inv["required_price_1"], _inv["required_price_2"]

# --- Texty shrnutí (jen prezentace výsledků výše) ---
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


# --- Bodové hodnocení: popisky a formát hodnot (body počítá investice.score_rows) ---
def _payback_text(months) -> str:
    return f"{fmt_num(months, 1)} měsíce" if months is not None else "nelze určit (ztráta)"


_SCORE_LABEL = {
    "profit": ("Očekávaný měsíční zisk", fmt_czk),
    "loss": ("Šance ztrátového měsíce", lambda p: fmt_pct(p * 100)),
    "var95": ("Velmi nepříznivý měsíc (VaR 5 %)", fmt_czk),
    "payback": ("Prostá návratnost investice", _payback_text),
}
score_items = [
    (_SCORE_LABEL[code][0], _SCORE_LABEL[code][1](v1), _SCORE_LABEL[code][1](v2), best)
    for code, v1, v2, best in _inv["score_rows"]
]

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
        f"\n- {_other} je ale lepší ve více doplňkových ukazatelích "
        f"({_points[1 - _rec_i]} ze 4 oproti {_points[_rec_i]} ze 4) – v riziku "
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

**Rozpětí měsíčního zisku stroje (simulace):**
- Slabý měsíc (P10): {machine_1} {fmt_czk(p10_mini)} · {machine_2} {fmt_czk(p10_dumper)}
- Typický měsíc (medián): {machine_1} {fmt_czk(p50_mini)} · {machine_2} {fmt_czk(p50_dumper)}
- Dobrý měsíc (P90): {machine_1} {fmt_czk(p90_mini)} · {machine_2} {fmt_czk(p90_dumper)}

**Co doplnit před nákupem:**
- skutečnou poptávku po stroji {better_machine},
- skutečné fixní a servisní náklady, odstávky a pojištění (výpočet zatím používá ukázkové hodnoty),
- sezónnost (léto a zima).
"""


# --- Nákup strojů: nejlepší sestava v rámci rozpočtu (výpočet v portfolio.py) ---
# Prohledají se všechny kombinace 0–2 kusů u všech typů strojů včetně „nic nekupovat“.
# Používají se individuální náklady strojů z dat, ne fixní náklad z nastavení investičního scénáře.
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
    """Investiční report DOCX – obsah sestavuje modul report_docx (beze změny dokumentu)."""
    return rd.build_report(
        avg_daily_dumper=avg_daily_dumper,
        avg_daily_mini=avg_daily_mini,
        better_machine=better_machine,
        break_even_dumper=break_even_dumper,
        break_even_mini=break_even_mini,
        expected_days=expected_days,
        expected_profit_dumper=expected_profit_dumper,
        expected_profit_mini=expected_profit_mini,
        fixed_cost=fixed_cost,
        machine_1=machine_1,
        machine_2=machine_2,
        max_days=max_days,
        min_days=min_days,
        n=n,
        p10_dumper=p10_dumper,
        p10_mini=p10_mini,
        p50_dumper=p50_dumper,
        p50_mini=p50_mini,
        p90_dumper=p90_dumper,
        p90_mini=p90_mini,
        required_price_1=required_price_1,
        required_price_2=required_price_2,
        risk_loss_dumper=risk_loss_dumper,
        risk_loss_mini=risk_loss_mini,
        scenario=scenario,
        scorecard_dumper=scorecard_dumper,
        scorecard_mini=scorecard_mini,
    )


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


UNIT_ECON_COLUMNS = ["Stroj", "Cena (Kč)", "Poptávka (dny)", "Tržby (Kč/měs.)", "Náklady (Kč/měs.)",
                     "Provozní zisk (Kč/měs.)"]


def unit_economics_rows() -> list:
    """Ekonomika jednoho kusu podle dat (portfolio.type_month pro 1 kus) – zdroj tabulky v Investicích
    (Nákup strojů) i ve Strojovém parku. Vrací nezaokrouhlené hodnoty; formátují se až při zobrazení."""
    rows = []
    for _, row in df.iterrows():
        one = pf.type_month(row, 1)
        rows.append(dict(machine=row["machine_type"], price=row["purchase_price"], demand=row["demand_days_mean"],
                         revenue=one["revenue"], costs=one["costs"], profit=one["profit"]))
    return rows


def unit_economics_texts(r: dict) -> list:
    """Hodnoty řádku tabulky v pořadí UNIT_ECON_COLUMNS (stejné formátování jako dosud)."""
    return [r["machine"], fmt_num(r["price"]), fmt_num(r["demand"], 0), fmt_num(r["revenue"]),
            fmt_num(r["costs"]), fmt_num(r["profit"])]


def show_unit_economics() -> None:
    """Tabulka „Ekonomika jednoho kusu podle dat“ (výpočet portfolio.type_month) – sekce Investice
    (Nákup strojů)."""
    st.markdown('<div class="de-section">Ekonomika jednoho kusu podle dat</div>',
                unsafe_allow_html=True)
    unit_rows = [dict(zip(UNIT_ECON_COLUMNS, unit_economics_texts(r))) for r in unit_economics_rows()]
    st.dataframe(pd.DataFrame(unit_rows), hide_index=True, width="stretch")


def chart(fig) -> None:
    st.plotly_chart(fig, width="stretch", config={"displaylogo": False, "locale": "cs"})


# ==================================================
# SEKCE INVESTICE – původní záložky 1–5 (beze změny obsahu)
# ==================================================

if PAGE == "investice":
    with inv_tab_area:                   # jeden kontejner na záložku; viditelný jen aktivní („_on“)
        tab1, tab2, tab3, tab4, tab5 = [
            st.container(key=f"de_inv_obsah_{i}_{'on' if label == INV_TAB else 'off'}")
            for i, label in enumerate(navreg.INVESTICE_TABS, start=1)
        ]


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
                    label="Očekávaný měsíční zisk stroje",
                    value=fmt_czk(rec_profit),
                    negative=rec_profit < 0,
                    sub=f"Průměr {fmt_num(n)} simulovaných měsíců: tržby − fixní náklad − servis.",
                ),
                dict(
                    label="Šance ztrátového měsíce",
                    value=fmt_pct(rec_prob * 100),
                    chip=risk_status(rec_prob),
                    sub="Podíl simulovaných měsíců, které skončí ztrátou.",
                ),
                dict(
                    label="Bod zvratu",
                    value=fmt_days(rec_be),
                    sub="Dní pronájmu měsíčně, které pokryjí fixní náklad (bez servisu a poruch).",
                ),
            ]
        )
        ui.hint(
            "Zisk stroje nezahrnuje společné náklady půjčovny (prostory, mzdy, administrativa) ani daně. "
            "Náklady jsou <b>ukázkové</b> – fixní náklad z nastavení investičního scénáře, servis z ukázkových dat; "
            "nejde o ověřené náklady firmy."
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
                ("Očekávaný měsíční zisk",
                 [fmt_czk(expected_profit_mini), fmt_czk(expected_profit_dumper)],
                 best_index(expected_profit_mini, expected_profit_dumper, True)),
                ("Šance ztrátového měsíce",
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
                f"Prověřit vyšší cenu nebo vytížení – minimum pro {better_machine} při "
                f"{expected_days} dnech pronájmu měsíčně je {fmt_czk(rec_required)}/den.",
                "Nahradit ukázkové náklady skutečnými a výpočet zopakovat.",
            ]
        else:
            steps = [
                f"Ověřit, že {better_machine} bude pronajatý víc než {fmt_days(rec_be)} měsíčně "
                "(bod zvratu – pokryje fixní náklad, servis už ne).",
                f"Zkontrolovat ceník: minimum při {expected_days} dnech pronájmu měsíčně je "
                f"{fmt_czk(rec_required)}/den (pokryje fixní náklad i servis), nyní {fmt_czk(rec_price)}/den.",
                "Nahradit ukázkové náklady skutečnými a výpočet zopakovat.",
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
            col_s1.metric(f"{machine_1} – lepší ukazatele", f"{scorecard_mini} ze 4")
            col_s2.metric(f"{machine_2} – lepší ukazatele", f"{scorecard_dumper} ze 4")
            ui.hint(
                "Doplňkové srovnání ve 4 ukazatelích: vyšší očekávaný zisk, nižší šance ztrátového měsíce, "
                "lepší velmi nepříznivý měsíc (VaR 5 %) a kratší návratnost (jen když oba stroje vydělávají). "
                "Při shodě nezíská bod nikdo. O doporučení rozhoduje jen očekávaný zisk; text níže "
                "sestavují pevná pravidla ze stejných výsledků."
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

        with st.expander(f"Rychlý odhad bez simulace – přesně {expected_days} dní pronájmu"):
            col_q1, col_q2 = st.columns(2)
            col_q1.metric(f"{machine_1} – zisk při {expected_days} dnech", fmt_czk(expected_profit_mini_real))
            col_q2.metric(f"{machine_2} – zisk při {expected_days} dnech", fmt_czk(expected_profit_dumper_real))
            ui.hint(
                f"Jiný výpočet než očekávaný zisk výše: ceníková cena × {expected_days} dní − fixní náklad "
                "− průměrný servis. Nepočítá s poruchami, kolísáním ceny a servisu ani s rozpětím dní podle "
                f"scénáře trhu (simulace: {min_days}–{max_days} dní), proto vychází jinak. "
                f"Počet dní změníte v panelu „{INV_PANEL_TITLE}“ v části „Pokročilé nastavení“."
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
            "měsíční zisk stroje (tržby − fixní náklad − servis) při různém počtu dní pronájmu, poruchách, "
            "ceně a servisních nákladech. Náklady jsou ukázkové.</div>",
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
                        ("Průměr všech měsíců", "= očekávaný měsíční zisk", fmt_czk(mean), "", mean < 0),
                        ("Šance ztrátového měsíce", "", fmt_pct(prob * 100), loss_share_text(prob), False),
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
            f"Pravděpodobnost, že <b>{machine_2}</b> dosáhne v jednom měsíci vyššího zisku než "
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

        st.markdown("### Riziko a očekávaný zisk")
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
            '<div class="de-intro">Graf ukazuje, jak moc se změní měsíční zisk stroje, když se jeden údaj '
            "změní o 10 % nahoru nebo dolů a ostatní zůstanou stejné. Délka pruhu je rozdíl zisku "
            "mezi variantou −10 % a +10 %: čím delší pruh, tím větší vliv. Počítá se rychlým odhadem "
            "pro pevný počet dní, ne simulací.</div>",
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
        st.markdown("### Kdy stroj pokryje fixní náklad")
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
                    hovertemplate="%{x} dní → po fixním nákladu %{y:,.0f} Kč<extra>" + name + "</extra>",
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
        fig_be.update_yaxes(title_text="Tržby − fixní náklad (Kč/měsíc)", tickformat=",.0f")
        ui.readable_axes(ui.style_fig(fig_be, height=450))
        fig_be.update_layout(margin=dict(t=96))  # místo pro ikony grafu nad legendou
        chart(fig_be)
        ui.hint(
            "Graf neukazuje zisk, ale tržby po odečtení fixního nákladu: denní cena × počet dní − fixní náklad, "
            "<b>bez servisních nákladů a poruch</b>. Body na čárkované nule jsou bod zvratu – od tohoto počtu "
            "dní stroj pokryje fixní náklad. Po odečtení servisu je potřeba dní o něco víc."
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
                f'<div class="de-intro">Tržby po odečtení fixního nákladu (bez servisu) stroje <b>{machine_1}</b> '
                "pro různé denní ceny (60–140 % ceníku) a počty dní pronájmu.</div>",
                unsafe_allow_html=True,
            )
            fig_heat = go.Figure(
                go.Heatmap(
                    x=days_range, y=price_range, z=profit_matrix, zmid=0,
                    colorscale=[[0, ui.BAD], [0.5, "#F1F3F5"], [1, ui.BLUE]],
                    colorbar=dict(
                        title=dict(text="Tržby − fixní náklad (Kč)", side="top"),
                        orientation="h", y=-0.3, yanchor="top", x=0, xanchor="left",
                        len=1, thickness=12, tickformat=",.0f", nticks=5,
                    ),
                    hovertemplate="%{x} dní, %{y:,.0f} Kč/den<br>Po fixním nákladu %{z:,.0f} Kč<extra></extra>",
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
                            f"<br>Po fixním nákladu {fmt_czk(profit_curve_mini[expected_days])}<extra></extra>"
                        ),
                    )
                )
            fig_heat.update_xaxes(title_text="Počet dní pronájmu za měsíc", dtick=5)
            fig_heat.update_yaxes(title_text="Denní cena (Kč)", tickformat=",.0f")
            chart(ui.readable_axes(ui.style_fig(fig_heat, height=470, legend=False)))
            if profit_matrix.max() <= 0:
                ui.insight(
                    "V zobrazeném rozsahu ceny a počtu dní tržby nikde nepokryjí ani fixní náklad.",
                    tone="bad",
                )
            ui.explain(
                [
                    ("Jak mapu číst",
                     "Každé políčko je jedna kombinace: vodorovně počet dní pronájmu za měsíc, "
                     "svisle cena za den. <b>Modrá</b> = tržby pokryjí fixní náklad, <b>červená</b> = nepokryjí. "
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
            "žádný nákup nevyplatí, doporučí nic nekupovat. <b>Jiný výpočet než záložky 1–4:</b> "
            "počítá s fixními náklady jednotlivých strojů z ukázkových dat (ne s fixním nákladem "
            "z nastavení investičního scénáře) a s průměrnou poptávkou bez simulace – částky za stejný stroj "
            "se proto liší.</div>",
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
                         sub=f"tržby {fmt_czk(pb['revenue'])} − fixní náklady a servis {fmt_czk(pb['costs'])}"),
                    dict(label="Prostá návratnost", value=months_text(payback),
                         sub="investice ÷ měsíční provozní zisk, bez úroků a společných nákladů"),
                ]
            )
            stress = pf.profit_with_demand(df, pb["counts"], 0.8)
            ui.insight(
                "Jde o <b>průměrný měsíc</b> při očekávané poptávce, ne o zaručený výsledek, a o provozní "
                "zisk strojů bez společných nákladů půjčovny (prostory, mzdy, administrativa), úroků a daní. "
                f"Pokud by poptávka byla o 20&nbsp;% nižší, provozní zisk této sestavy by klesl na "
                f"<b>{fmt_czk(stress)}</b>. Další krok: ověřit skutečnou poptávku a náklady strojů.",
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
                     "Fixní náklad z nastavení investičního scénáře se zde nepoužívá – ostatní záložky s ním "
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
            show_unit_economics()


# ==================================================
# ZÁLOŽKA 6 – PLÁN ZAKÁZEK
# ==================================================

@st.cache_data(show_spinner="Hledám plán zakázek…")
def plan_jobs(jobs: pd.DataFrame, capacity: dict, code_version: str):
    """Nový plán (celočíselná optimalizace) a přínos jednotlivých zakázek pro plán.
    code_version jen obnoví mezipaměť po nasazení nového kódu (viz load_data_model)."""
    result = sch.optimize_schedule(jobs, capacity)
    new_plan = sch.evaluate(jobs, result["starts"])
    contributions = sch.job_contributions(jobs, capacity, sch.summary(new_plan)["penalty"])
    return result, new_plan, contributions


if PAGE == "zakazky":                                     # sekce Vlastní zakázky (dříve záložka 6)
    # Modernizace UI: karty a sekce jako Dashboard / Pronájem strojů. Výpočty plánu (plan_jobs,
    # scheduling.summary/evaluate/conflicts/utilization) a přínos zakázek jsou beze změny.
    st.markdown(ui.JOBS_CSS, unsafe_allow_html=True)
    with ui.tab_container("zakazky"):
        st.markdown("### Plán zakázek a vytížení strojů")
        st.markdown(
            '<div class="de-intro">Aplikace rozvrhne 12 zakázek ze zásobníku (fiktivní data) tak, '
            "aby žádný stroj nebyl ve stejný den využit víckrát, než kolik je kusů, a aby byl "
            "<b>zisk zakázek po penále co nejvyšší</b>.</div>",
            unsafe_allow_html=True,
        )
        ui.note(
            "<b>Plán zatím nezohledňuje pronájmy a servis</b> – hlídá kapacitu jen mezi vlastními zakázkami. "
            "Konflikty s pronájmy a servisem ukazuje sekce <b>Plánování kapacity</b>."
        )

        plan_result, job_plan, job_contrib = plan_jobs(jobs_df, machine_capacity, code_version)
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
        gain = plan_sum["profit_after_penalty"] - old_sum["profit_after_penalty"]
        on_time = plan_sum["scheduled"] - plan_sum["late"]

        def jobs_delta(new, old, unit, better_lower, money=False):
            """Štítek změny nový − původní plán (symbol + text, barva jen doplňuje)."""
            diff = new - old
            if diff == 0:
                return ui.dash_chip("beze změny", "info")
            text = (("+" if diff > 0 else "−") + (fmt_czk(abs(diff)) if money else f"{fmt_num(abs(diff))} {unit}"))
            good = diff < 0 if better_lower else diff > 0
            return ui.dash_chip(text, "good" if good else "bad")

        # ---- KPI ----------------------------------------------------------------------------------
        ui.dash_kpis([
            dict(icon="coins", label="Zisk zakázek po penále", tone="accent",
                 value=fmt_czk(plan_sum["profit_after_penalty"]),
                 chip=ui.dash_chip("jen vlastní zakázky", "info"),
                 sub=f"zisk zakázek {fmt_czk(plan_sum['profit'])} − penále {fmt_czk(plan_sum['penalty'])} · "
                     "nejde o čistý zisk půjčovny"),
            dict(icon="spark", label="Zakázky v termínu", tone="good" if not plan_sum["late"] else "neutral",
                 value=f"{on_time}<small>z {plan_sum['scheduled']}</small>",
                 chip=(ui.dash_chip(f"{plan_sum['late']} po termínu", "bad") if plan_sum["late"]
                       else ui.dash_chip("vše v termínu", "good")),
                 sub=f"penále celkem {fmt_czk(plan_sum['penalty'])}"),
            dict(icon="route", label="Délka plánu", tone="neutral", value=fmt_days(plan_sum["length"], 0),
                 sub=f"{plan_sum['scheduled']} naplánovaných zakázek"),
            dict(icon="gauge", label="Zlepšení proti původnímu plánu", tone="good" if gain > 0 else "neutral",
                 value=("+" if gain > 0 else "") + fmt_czk(gain),
                 sub=f"zisk po penále: původní {fmt_czk(old_sum['profit_after_penalty'])} → nový "
                     f"{fmt_czk(plan_sum['profit_after_penalty'])}"),
        ])
        if plan_result["optimal"]:
            ui.insight(
                "Výpočet <b>ověřil</b>, že při zadaných délkách zakázek, termínech a počtech strojů "
                "neexistuje plán s nižším celkovým penále. Nejde o záruku – skutečné zakázky se "
                "mohou zdržet nebo změnit."
            )
        else:
            ui.insight(
                "Zobrazený plán je <b>nejlepší nalezený</b> – v časovém limitu se nepodařilo ověřit, "
                "že lepší plán neexistuje.",
                tone="bad",
            )
        if plan_sum["waiting"]:
            waiting = job_plan[~job_plan["scheduled"]]["job_name"].tolist()
            ui.insight(
                f"Nelze naplánovat (chybí potřebný stroj): <b>{', '.join(waiting)}</b>.", tone="bad"
            )

        # ---- Srovnání s původním plánem -----------------------------------------------------------
        st.markdown('<div class="dd-section-title">Srovnání s původním plánem<small>stejné zakázky a data, '
                    'jiné pořadí</small></div>', unsafe_allow_html=True)
        with ui.dash_section("cmp"):
            ui.dash_compare(
                ["Ukazatel", "Původní plán", "Nový plán", "Změna"],
                [
                    ("Zakázky po termínu", [f"{old_sum['late']} z {old_sum['scheduled']}",
                                            f"{plan_sum['late']} z {plan_sum['scheduled']}"],
                     jobs_delta(plan_sum["late"], old_sum["late"], "zakázky", better_lower=True)),
                    ("Penále celkem", [fmt_czk(old_sum["penalty"]), fmt_czk(plan_sum["penalty"])],
                     jobs_delta(plan_sum["penalty"], old_sum["penalty"], "", better_lower=True, money=True)),
                    ("Délka plánu", [fmt_days(old_sum["length"], 0), fmt_days(plan_sum["length"], 0)],
                     jobs_delta(plan_sum["length"], old_sum["length"], "dní", better_lower=True)),
                    ("Zisk po penále", [fmt_czk(old_sum["profit_after_penalty"]),
                                        fmt_czk(plan_sum["profit_after_penalty"])],
                     jobs_delta(plan_sum["profit_after_penalty"], old_sum["profit_after_penalty"], "",
                                better_lower=False, money=True)),
                    ("Kolize mezi zakázkami plánu", [str(len(old_conflicts)), str(len(plan_conflicts))],
                     jobs_delta(len(plan_conflicts), len(old_conflicts), "", better_lower=True)),
                ],
            )
            ui.dash_text(
                f"Obě varianty plánují stejných {plan_sum['scheduled']} zakázek na stejných ukázkových "
                f"datech. Nový plán je o <b>{fmt_czk(gain)}</b> výhodnější jen díky jinému pořadí a "
                "využití volných dní strojů. Původní plán řadil zakázky podle skóre bez ohledu na "
                "termíny a stroj nevyužil, dokud nebyl volný pro další zakázku v pořadí.", muted=True)
        ui.insight(
            "<b>Kolize mezi zakázkami plánu</b> počítá jen překrytí vlastních zakázek navzájem. Pronájmy "
            "strojů a servis plán nezná, takže s nimi se zakázky překrývat mohou – takové konflikty ukazuje "
            "sekce <b>Plánování kapacity</b>. Další krok: zkontrolovat tam konfliktní dny a vyřešit je ručně."
        )

        # ---- Harmonogram --------------------------------------------------------------------------
        st.markdown('<div class="dd-section-title">Harmonogram<small>pruh = doba zakázky, svislá čárka = '
                    'smluvní termín</small></div>', unsafe_allow_html=True)
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
        # Popisek osy: nejpřínosnější zakázka označená ★ (jen zobrazení, pořadí řádků beze změny)
        best_name = best_job["job_name"] if best_job is not None else None
        gantt_labels = {n: (f"★ {n}" if n == best_name else n) for n in gantt["job_name"]}
        fig_gantt = go.Figure()
        for status, color, label in (("Včas", ui.BLUE, "✓ Včas"), ("Po termínu", ui.BAD, "✕ Po termínu")):
            part = gantt[gantt["status"] == status]
            if part.empty:
                continue
            fig_gantt.add_trace(
                go.Bar(
                    y=part["job_name"].map(gantt_labels), x=part["duration"], base=part["start_day"] - 1,
                    orientation="h", name=label,
                    marker=dict(color=color, line=dict(color="#fff", width=2)),
                    text=["" if d <= 0 else f"+{int(d)} d" for d in part["days_late"]],
                    textposition="outside", cliponaxis=False,
                    textfont=dict(color=ui.BAD_TEXT, size=12),
                    hovertext=part["hover"], hoverinfo="text",
                )
            )
        fig_gantt.add_trace(
            go.Scatter(
                y=gantt["job_name"].map(gantt_labels), x=gantt["deadline"], mode="markers", name="Termín",
                marker=dict(symbol="line-ns", size=18, line=dict(color=ui.NAVY, width=3)),
                hovertemplate="Termín: den %{x}<extra></extra>",
            )
        )
        fig_gantt.update_yaxes(categoryorder="array", categoryarray=list(gantt["job_name"].map(gantt_labels)),
                               autorange="reversed", title_text="")
        x_end = max(plan_sum["length"], int(planned["deadline"].max()) if len(planned) else 0) + 5
        fig_gantt.update_xaxes(title_text="Den plánu", dtick=5, range=[0, x_end], showgrid=True,
                               gridcolor="#EEF0F3")
        fig_gantt.update_layout(barmode="overlay", bargap=0.35)
        ui.readable_axes(ui.style_fig(fig_gantt, height=max(420, 34 * len(gantt) + 140)))
        # Legenda pod grafem – na mobilu se tak nepřekrývá s ikonami grafu.
        fig_gantt.update_layout(legend=dict(orientation="h", yanchor="top", y=-0.16, x=0),
                                margin=dict(b=70))
        with ui.dash_section("gantt"):
            chart(fig_gantt)
            ui.dash_text(
                "<b>Modrá</b> = zakázka skončí včas, <b>červená</b> = po termínu (vpravo počet dní zpoždění). "
                "Svislá čárka je smluvní termín (poslední den, kdy má být zakázka hotová). "
                "★ = zakázka s největším přínosem pro plán.", muted=True)

        # ---- Prioritní zakázka --------------------------------------------------------------------
        if best_job is not None:
            top = job_contrib.iloc[0]
            with ui.dash_section("best_accent"):
                ui.dash_section_head("spark", "Zakázka s největším přínosem pro plán",
                                     "Nejméně výhodné ji ztratit nebo odložit", tone="accent")
                ui.dash_kv("Zakázka", "★ " + html.escape(str(best_job["job_name"])),
                           f"termín v plánu den {int(best_job['start_day'])}–{int(best_job['end_day'])} "
                           f"(smluvní termín den {int(best_job['deadline'])})")
                ui.dash_rows([
                    ("Zisk zakázky", fmt_czk(top["profit"]), ""),
                    ("Penále, které zakázka přidá plánu", fmt_czk(top["added_penalty"]), ""),
                    ("Čistý přínos pro plán", fmt_czk(top["contribution"]),
                     "zisk zakázky − přidané penále celého plánu", True),
                ])
                ui.dash_text(
                    "Čistý přínos = zisk zakázky − o kolik se kvůli ní zvýší nejnižší možné penále "
                    "celého plánu (její vlastní zpoždění i zdržení ostatních zakázek). Tuto zakázku "
                    "je proto nejméně výhodné ztratit nebo odložit.", muted=True)
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

        # ---- Detail -------------------------------------------------------------------------------
        st.markdown('<div class="dd-section-title">Detail zakázek<small>v pořadí podle začátku · '
                    'tabulku lze posouvat do strany</small></div>', unsafe_allow_html=True)
        detail = gantt.assign(
            **{
                "Zakázka": gantt["job_name"].map(gantt_labels),
                "Stav": np.where(gantt["days_late"] > 0, "✕ Po termínu", "✓ Včas"),
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
        )[["Zakázka", "Stav", "Začátek (den)", "Konec (den)", "Termín (den)", "Zpoždění (dny)", "Zisk po penále",
           "Zisk", "Penále", "Tržba", "Náklady", "Stroje", "Priorita", "Riziko"]]
        st.dataframe(
            detail, width="stretch", hide_index=True, height=35 * (len(detail) + 1) + 3,   # všechny řádky
            column_config={
                "Zakázka": st.column_config.TextColumn(pinned=True, width="medium"),
                "Stav": st.column_config.TextColumn(width=110),
                "Stroje": st.column_config.TextColumn(width="medium"),
            },
        )

        # ---- Vytížení strojů ----------------------------------------------------------------------
        st.markdown('<div class="dd-section-title">Vytížení strojů zakázkami<small>jen vlastní zakázky '
                    'v plánu</small></div>', unsafe_allow_html=True)
        machine_usage_df = sch.utilization(job_plan, machine_capacity)
        bottleneck_machine = machine_usage_df.iloc[0]["machine"]
        bottleneck_utilization = float(machine_usage_df.iloc[0]["utilization"])
        usage_labels = [f"{m} ({u} ks)" if u > 1 else m
                        for m, u in zip(machine_usage_df["machine"], machine_usage_df["units"])]
        fig_usage = go.Figure(
            go.Bar(
                x=machine_usage_df["utilization"],
                y=usage_labels,
                orientation="h",
                marker=dict(color=[ui.ACCENT if i == 0 else ui.BLUE for i in range(len(machine_usage_df))],
                            line=dict(color="#fff", width=2)),
                text=[fmt_pct(v) for v in machine_usage_df["utilization"]], textposition="outside",
                cliponaxis=False, customdata=machine_usage_df["used_days"],
                hovertemplate="%{y}<br>Vytížení %{x:.1f} %<br>Obsazené dny: %{customdata}<extra></extra>",
            )
        )
        fig_usage.update_yaxes(autorange="reversed")
        fig_usage.update_xaxes(title_text="Vytížení zakázkami v plánu (%)", range=[0, 130], showgrid=True,
                               gridcolor="#EEF0F3")
        fig_usage.add_vline(x=100, line=dict(color=ui.NAVY, width=1, dash="dot"))
        fig_usage.update_layout(bargap=0.35)
        with ui.dash_section("usage"):
            chart(ui.readable_axes(ui.style_fig(fig_usage, height=max(320, 34 * len(machine_usage_df) + 110),
                                                legend=False)))
            ui.dash_text(
                f"Vytížení = dny stroje na zakázkách ÷ (počet kusů × délka plánu {fmt_days(plan_sum['length'], 0)}). "
                "Pronájmy a servis nezahrnuje – celkovou obsazenost ukazuje sekce Plánování kapacity. "
                f"Nejvytíženější je <b>{bottleneck_machine}</b> ({fmt_pct(bottleneck_utilization, 0)}, oranžově). "
                "Tečkovaná čára = 100 % kapacity.", muted=True)

        with st.expander("Předpoklady a metodika"):
            ui.explain(
                [
                    ("Cíl plánu",
                     "Co nejvyšší zisk zakázek po penále. Tržba, náklady a zisk zakázky na pořadí nezávisí, "
                     "takže při plánování všech zakázek je to totéž jako co nejnižší celkové penále. "
                     "Při stejném penále se zakázky dokončí co nejdříve."),
                    ("Pojmy",
                     "<b>Tržba</b> = očekávaná tržba zakázky, <b>náklady</b> = přímé náklady zakázky "
                     "(z dat), <b>zisk</b> = tržba − náklady, <b>penále</b> = dny po termínu × denní "
                     "penále, <b>zisk po penále</b> = zisk − penále. Nejde o čistý zisk firmy: fixní náklady "
                     "strojů a společné náklady půjčovny se neodečítají a data neuvádějí, zda přímé náklady "
                     "zakázek obsahují náklady vlastních strojů."),
                    ("Stroje",
                     "Každá zakázka potřebuje po celou dobu všechny své stroje. V jednom dni lze stroj "
                     "použít tolikrát, kolik je kusů (vibrační deska 2 ks, ostatní 1 ks). Volné dny "
                     "stroje může využít kterákoli zakázka."),
                    ("Jak se plán hledá",
                     "Celočíselná optimalizace (SciPy, řešič HiGHS) projde všechny přípustné začátky "
                     "zakázek. Pokud v časovém limitu neověří, že lepší plán neexistuje, aplikace to uvede."),
                    ("Co model nezahrnuje",
                     "Víkendy a svátky, pozdější možný začátek zakázky, přejezdy a obsluhu strojů, "
                     "pronájmy strojů mimo zakázky, servisní odstávky, nejistotu délky zakázek a riziko "
                     "(priorita a riziko jsou jen informativní)."),
                ]
            )

# ==================================================
# ZÁLOŽKA 7 – KAPACITA STROJŮ
# ==================================================
# Kalendář jen zobrazuje výsledky modulu capacity.py (capacity_view = převod na tabulky a mřížku).
# Termíny zakázek přebírá z plánu záložky 6 (stejná funkce plan_jobs, výsledek z mezipaměti).

@st.cache_data(show_spinner=False, max_entries=64)
def capacity_grid(data_version: str, code_version: str, job_starts: tuple, start: int, end: int, types: tuple, plan_start):
    """Mřížka, její HTML a konflikty za období. Streamlit spouští kód všech záložek při každém
    kliknutí – mezipaměť zajistí, že se kalendář přepočítá jen při změně jeho vstupů (včetně
    změny dat: data_version). Ukládají se jen obyčejné tabulky a text (viz load_data_model)."""
    tables, _ = load_data_model(data_version, code_version)
    model = dm.DataModel(**tables)
    cal = cv.build(model, dict(job_starts))
    names = cv.type_names(model)
    grid = cv.grid(cal, start, end, list(types))
    return (grid, cv.grid_html(grid, names, cv.fleet_units(cal), plan_start),
            cv.conflicts_table(cal, start, end, list(types), names, cv.job_info(model), plan_start))


CAP_VIEWS = ["Kalendář dostupnosti", "Přehled využití"]


def show_in_calendar(type_id: str) -> None:
    """Z přehledu využití do kalendáře: filtr i detail dne na vybraný typ stroje.
    Volá se jako callback tlačítka – před vykreslením widgetů, proto lze měnit jejich stav."""
    st.session_state["cap_types"] = [type_id]
    st.session_state["cap_detail_type"] = type_id
    st.session_state["cap_view"] = CAP_VIEWS[0]


if PAGE == "kapacita":                                    # sekce Plánování kapacity (dříve záložka 7)
    with ui.tab_container("kapacita"):
        # UX/UI 2: společný standard záložek (ui_theme) + pravidla kalendáře (cv.CALENDAR_TAB_CSS).
        # Souhrnné ukazatele se vypisují do místa NAD filtry (summary_slot), ale počítají se až PO přečtení
        # filtrů – vždy tedy odpovídají aktuálnímu výběru. Widgety, jejich klíče, pořadí a výchozí hodnoty se nemění.
        st.markdown(cv.TAB_CSS + cv.CALENDAR_TAB_CSS, unsafe_allow_html=True)
        ui.tab_header(
            "Kapacita strojů po dnech",
            "Kolik kusů každého stroje je v jednotlivých dnech volných a co je drží: "
            "<b>potvrzené</b> pronájmy, vlastní zakázky podle plánu ze záložky 6 a servis. <b>Předběžné</b> "
            "rezervace kapacitu neblokují – jsou vyznačené zvlášť jako možné riziko. Pronájmy a servis jsou "
            "<b>fiktivní ukázková data</b>.",
        )

        _, cap_plan, _ = plan_jobs(jobs_df, machine_capacity, code_version)
        cap_jobs = cv.job_info(data_model)
        cap_names = cv.type_names(data_model)
        cap_starts = cv.job_starts_from_plan(jobs_df, cap_plan)
        cap_cal = cv.build(data_model, cap_starts)
        cap_default_end, cap_max_day = cv.period_bounds(cap_cal)
        cap_types_all = [t for t in cap_names if t in cap_cal.type_ids()]

        cap_summary = ui.summary_slot("kapacita")              # souhrn – naplní se až po filtrech níže

        with ui.filter_box("kapacita", "Období, stroje a zobrazení"):
            f1, f2 = st.columns([3, 2])
            with f1:
                cap_pick = st.multiselect(
                    "Typy strojů", cap_types_all, default=[], format_func=lambda t: cap_names[t],
                    key="cap_types", placeholder="Všechny typy strojů",
                )
            cap_types = [t for t in cap_types_all if t in cap_pick] or cap_types_all
            with f2:
                cap_start = st.date_input(
                    "Datum prvního dne plánu (nepovinné)", value=None, key="cap_plan_start", format="DD.MM.YYYY",
                    help="Jen pro zobrazení kalendářních dat a dnů v týdnu. Plán počítá ve dnech od začátku.",
                )
            cap_from, cap_to = st.slider(
                "Období (dny plánu)", min_value=1, max_value=cap_max_day, value=(1, cap_default_end), key="cap_period",
            )
            cap_view = st.radio(
                "Zobrazení", CAP_VIEWS, horizontal=True, key="cap_view", label_visibility="collapsed",
            )

        with cap_summary:
            st.markdown(ui.scope_html(*cv.scope_texts(cap_from, cap_to, cap_types, cap_types_all, cap_names, cap_start)),
                        unsafe_allow_html=True)
            if cap_cal.issues:
                st.warning(
                    "Některé údaje jsou neplatné nebo neúplné – dotčené stroje jsou označené „nelze určit“:\n\n"
                    + "\n".join(f"- {text}" for _, text in cap_cal.issues)
                )

        if cap_view == CAP_VIEWS[0]:
            cap_grid, cap_html, cap_conf = capacity_grid(
                data_version, code_version, tuple(sorted(cap_starts.items())), cap_from, cap_to, tuple(cap_types), cap_start)
            cap_sum = cv.summary(cap_grid, cap_names)
            n_conf = len(cap_sum["conflict_days"])
            with cap_summary:
                ui.kpi_cards(
                    [
                        dict(label="Dny s konfliktem", value=str(n_conf),
                             sub=("bez konfliktu v období" if not n_conf else
                                  "potvrzené nároky převyšují počet kusů · dny: "
                                  + cv.days_text(cap_sum["conflict_days"], cap_start)),
                             negative=n_conf > 0),
                        dict(label="Nelze určit", value=str(cap_sum["unknown_cells"]),
                             sub=("dní × strojů s neúplnými údaji o kapacitě" if cap_sum["unknown_cells"]
                                  else "údaje o kapacitě jsou úplné")),
                        dict(label="Nejvytíženější stroj", value=html.escape(cap_sum["busiest"] or "–"),
                             sub=(f"obsazeno {fmt_pct(cap_sum['busiest_share'], 0)} strojodnů v období (vč. servisu)"
                                  if cap_sum["busiest"] else "")),
                    ]
                )
            st.markdown(cap_html, unsafe_allow_html=True)
            ui.hint(
                "Číslo v buňce = <b>volné kusy / kusy k dispozici</b> ten den; pruh ukazuje jednotlivé kusy. "
                "Konec rezervace i zakázky je včetně posledního dne. Na mobilu kalendář posuňte do strany; "
                "po podržení buňky (na počítači najetí myší) se zobrazí podrobnosti."
            )
            ui.hint(cv.data_info(data_model, data_version))
            cap_late = cv.late_jobs(cap_cal, cap_jobs, cap_types, cap_from, cap_to)
            if cap_late:
                ui.insight(
                    "Plán záložky 6 dokončí některé zakázky <b>po termínu</b> – kalendář je ukazuje až do "
                    "skutečného konce: "
                    + "; ".join(
                        f"{html.escape(name)} do {cv.day_label(end, cap_start)} (termín {cv.day_label(dl, cap_start)})"
                        for _, name, end, dl in cap_late
                    )
                    + "."
                )

            st.markdown("### Konflikty v období")
            if cap_conf.empty:
                st.markdown(ui.status_chip(ui.status_text("Bez konfliktu", "good"), "good")
                            + " Ve zvoleném období a u vybraných strojů "
                            "nepřesahuje obsazení počet kusů.", unsafe_allow_html=True)
            else:
                n_rows = len(cap_conf)
                st.markdown(ui.status_chip(ui.status_text(
                    f"{n_rows} {'konflikt' if n_rows == 1 else 'konflikty' if n_rows < 5 else 'konfliktů'}", "bad"), "bad"),
                    unsafe_allow_html=True)
                st.dataframe(cap_conf, width="stretch", hide_index=True)
                ui.hint("Konflikt = potvrzené nároky (pronájmy, zakázky z plánu, servis) převyšují počet kusů. "
                        "Vzniká hlavně tím, že plán záložky 6 zatím o pronájmech a servisu neví. Aplikace "
                        "konflikty jen ukazuje – nic nepřesouvá a neurčuje, kdo má přednost. "
                        "<b>Další krok:</b> vyřešit ručně, např. posunem zakázky, pronájmu nebo servisu.")

            st.markdown("### Detail dne")
            d1, d2 = st.columns([3, 2])
            with d1:
                cap_dt = st.selectbox("Stroj", cap_types, format_func=lambda t: cap_names[t], key="cap_detail_type")
            if "cap_detail_day" in st.session_state:          # po změně období držet den v rozsahu
                st.session_state["cap_detail_day"] = min(max(int(st.session_state["cap_detail_day"]), cap_from), cap_to)
            with d2:
                cap_dd = st.number_input("Den plánu", min_value=cap_from, max_value=cap_to, value=cap_from, step=1,
                                         key="cap_detail_day")
            cap_row = cap_grid[(cap_grid["type_id"] == cap_dt) & (cap_grid["day"] == cap_dd)].iloc[0]
            st.markdown(
                f"**{html.escape(cap_names[cap_dt])}, {cv.day_label(int(cap_dd), cap_start)}:** "
                + cv.day_status_text(cap_row)
            )
            cap_detail = cv.day_detail(cap_cal, cap_dt, int(cap_dd), cap_jobs, cap_start)
            if cap_detail.empty:
                ui.hint("V tento den stroj nedrží žádný pronájem, zakázka ani servis.")
            else:
                st.dataframe(cap_detail, width="stretch", hide_index=True)
            st.dataframe(cv.unit_states(cap_cal, cap_dt, int(cap_dd)), width="stretch", hide_index=True)
            ui.hint("Pronájmy a zakázky jsou rezervované na typ stroje, ne na konkrétní kus – o volnu rozhodují "
                    "počty kusů výše.")

        if cap_view == CAP_VIEWS[0]:
            st.markdown("### Ověřit nový požadavek")
            ui.hint("Rychlá kontrola kapacity: vejde se požadavek na stroj v daném termínu? Předběžné rezervace "
                    "kapacitu neblokují, dokud je nezapočtete volbou níže. Nic se nerezervuje ani nepřesouvá. "
                    "Ekonomiku a porovnání více požadavků řeší záložka Rozhodovací plánovač.")
            q1, q2, q3, q4 = st.columns([3, 1, 1, 1])
            with q1:
                req_type = st.selectbox("Typ stroje", cap_types_all, format_func=lambda t: cap_names[t], key="cap_req_type")
            with q2:
                req_from = st.number_input("Od dne", min_value=1, max_value=365, value=13, step=1, key="cap_req_from")
            with q3:
                req_to = st.number_input("Do dne (včetně)", min_value=1, max_value=365, value=16, step=1,
                                         key="cap_req_to")
            with q4:
                req_qty = st.number_input("Počet kusů", min_value=1, max_value=10, value=1, step=1, key="cap_req_qty")
            req_tent = st.checkbox("Počítat i s předběžnými rezervacemi (opatrnější pohled)", key="cap_req_tent")
            req = cv.check_request(cap_cal, req_type, int(req_from), int(req_to), int(req_qty), req_tent,
                                   cap_jobs, cap_start)
            st.markdown(
                ui.status_chip(ui.status_text(req["status"], req["tone"]), req["tone"]) + " "
                + html.escape(req["headline"]),
                unsafe_allow_html=True,
            )
            if not req["problems"].empty:
                st.dataframe(req["problems"], width="stretch", hide_index=True)
            for text in req["reasons"]:
                st.markdown(f"- {html.escape(text)}")
            for day_text, text in req["warnings"]:
                st.markdown(f"- ⚠️ {html.escape(day_text)}: {html.escape(text)}")

            with st.expander("Jak kalendář počítá"):
                ui.explain(
                    [
                        ("Zdroj výpočtu",
                         "Všechny počty, konflikty i ověření požadavku počítá kapacitní výpočet aplikace; "
                         "kalendář je jen zobrazuje."),
                        ("Dny",
                         "Celé dny, začátek i konec včetně. Rezervace může navázat den po skončení předchozí; "
                         "předání ve stejný den je konflikt."),
                        ("Kapacita",
                         "Kapacita dne = kusy ve flotile k dispozici ten den minus kusy v servisu."),
                        ("Stavy",
                         "Potvrzené pronájmy, zakázky a servis kapacitu blokují. Předběžné rezervace (přerušovaná linka) ji "
                         "neblokují, jen varují; při ověření nového požadavku je lze volbou započítat. V Rozhodovacím "
                         "plánovači soutěží předběžné rezervace o volnou kapacitu stejně jako nové požadavky. "
                         "Zakázka bez termínu v plánu (nebo neplatný údaj) znamená „nelze určit“ – nikdy volno."),
                        ("Zakázky",
                         "Termíny zakázek jsou z plánu záložky 6 včetně zpoždění. Plán zatím nebere ohled na "
                         "pronájmy a servis, proto tu mohou být konflikty."),
                        ("Data",
                         "Pronájmy (FIKT-…) a servisní odstávky jsou fiktivní ukázka; dva konflikty jsou v nich "
                         "záměrně. Ostatní záložky je nepoužívají."),
                    ]
                )

        # ------------------------------------------------------------------
        # E2.4 – Přehled využití (stejné filtry typu a období jako kalendář)
        # ------------------------------------------------------------------
        if cap_view == CAP_VIEWS[1]:
            use_grid, _, _ = capacity_grid(
                data_version, code_version, tuple(sorted(cap_starts.items())), cap_from, cap_to, tuple(cap_types), cap_start)
            use_table = cv.utilization_table(use_grid, cap_names, cap_cal.unknown_capacity)
            use_fleet = cv.fleet_utilization(use_table, use_grid)
            n_days = cap_to - cap_from + 1
            cap_total = use_fleet["capacity"]
            pct = lambda v: (fmt_pct(100.0 * v / cap_total, 0) if cap_total else "nelze určit")

            ui.note(
                "<b>Ukázková data</b> – přehled vychází z fiktivních pronájmů, servisu a plánu zakázek "
                "ze záložky 6. Výsledky jsou ilustrativní."
            )
            with cap_summary:
                st.markdown(
                    f"**Kapacita flotily ve zvoleném období:** {fmt_num(cap_total)} strojodnů "
                    f"({len(use_table)} typů strojů, {n_days} {'den' if n_days == 1 else 'dny' if n_days < 5 else 'dní'})"
                    + (f" – bez typů s neúplnými údaji o flotile: {html.escape(', '.join(use_fleet['excluded']))}"
                       if use_fleet["excluded"] else "")
                )
                use_conf = use_fleet["conflict_days"]
                ui.kpi_cards(
                    [
                        dict(label="Využití flotily", value=cv.share_text(use_fleet["util_min"], use_fleet["util_max"]),
                             sub=f"pronájmy a vlastní zakázky z {fmt_num(cap_total)} strojodnů (servis se nepočítá)"),
                        dict(label="Volná kapacita", value=f"{fmt_num(use_fleet['free'])} strojodnů",
                             sub=f"{pct(use_fleet['free'])} kapacity"
                             + (f" · nelze určit {fmt_num(use_fleet['unknown'])}" if use_fleet["unknown"] else "")),
                        dict(label="Konfliktní dny", value=str(len(use_conf)),
                             sub=("bez konfliktu v období" if not use_conf else
                                  "dny: " + cv.days_text(use_conf, cap_start)
                                  + f" · poptávka nad kapacitu {fmt_num(use_fleet['excess_demand'])} strojodnů"),
                             negative=bool(use_conf)),
                    ]
                )
            ui.kpi_cards(
                [
                    dict(label="Pronájmy", value=f"{fmt_num(use_fleet['rental'])} strojodnů",
                         sub=f"{pct(use_fleet['rental'])} kapacity · potvrzené rezervace"),
                    dict(label="Vlastní zakázky", value=f"{fmt_num(use_fleet['job'])} strojodnů",
                         sub=f"{pct(use_fleet['job'])} kapacity · podle plánu záložky 6"),
                    dict(label="Servis", value=f"{fmt_num(use_fleet['maintenance'])} strojodnů",
                         sub=f"{pct(use_fleet['maintenance'])} kapacity · stroj nedostupný"),
                ]
            )

            if use_fleet["conflict"]:
                ui.insight(
                    f"V konfliktních dnech je <b>{fmt_num(use_fleet['conflict'])} strojodnů</b> obsazených, ale bez "
                    "rozhodnutí o přednosti nelze určit, zda pronájmem, zakázkou, nebo servisem – v grafu "
                    "i tabulce jsou vedené zvlášť jako „konflikt“. Proto může být využití uvedené rozpětím.",
                    tone="bad",
                )

            use_fig, use_height = cv.utilization_figure(use_table, use_fleet)
            use_fig = ui.readable_axes(ui.style_fig(use_fig, height=use_height))
            use_fig.update_layout(legend=cv.LEGEND_BELOW)
            chart(use_fig)
            ui.hint(
                "Každý řádek = 100 % kapacity typu stroje (kusy × dny). Číslo za názvem = využití "
                "(pronájmy + vlastní zakázky). Servis je nedostupná kapacita, ne využití."
            )

            high = use_table[use_table["rating"] == "vysoké"]["name"].tolist()
            low = use_table[use_table["rating"] == "nízké"]["name"].tolist()
            ui.insight(
                f"<b>Vysoké využití</b> (aspoň {fmt_num(cv.HIGH_USE)} %): "
                f"{html.escape(', '.join(high)) if high else 'žádný stroj'} · "
                f"<b>Nízké využití</b> (pod {fmt_num(cv.LOW_USE)} %): "
                f"{html.escape(', '.join(low)) if low else 'žádný stroj'}. Hranice jsou orientační."
            )
            if use_fleet["tentative"]:
                ui.hint(
                    f"Předběžné rezervace: {fmt_num(use_fleet['tentative'])} strojodnů – kapacitu neblokují "
                    "a do využití se nepočítají."
                )

            st.markdown("### Využití podle typu stroje")
            st.dataframe(cv.utilization_display(use_table), width="stretch", hide_index=True)

            n1, n2 = st.columns([3, 2])
            with n1:
                use_pick = st.selectbox("Detail stroje v kalendáři", use_table["type_id"].tolist(),
                                        format_func=lambda t: cap_names.get(t, t), key="use_detail_type")
            with n2:
                st.button("Zobrazit v kalendáři", key="use_to_calendar", on_click=show_in_calendar,
                          args=(use_pick,), width="stretch")

            with st.expander("Jak se přehled počítá"):
                ui.explain(
                    [
                        ("Zdroj",
                         "Denní kapacita z kapacitního modulu – stejná jako v kalendáři (stejné filtry, data "
                         "a plán zakázek ze záložky 6)."),
                        ("Kapacita",
                         "Součet kusů k dispozici v jednotlivých dnech ve strojodnech (1 strojoden = 1 kus "
                         "stroje na 1 den). Objednaný stroj se počítá až od dodání."),
                        ("Rozdělení",
                         "Každý den a stroj patří do jedné skupiny – nic se nepočítá dvakrát. Běžný den: "
                         "pronájem, vlastní zakázky, servis a volno. Konfliktní den (víc nároků než kusů): "
                         "celá kapacita je „konflikt“, protože bez rozhodnutí o přednosti nelze určit, kdo "
                         "stroj skutečně má. Nároky nad počet kusů jsou „poptávka nad kapacitu“ a do "
                         "kapacity se nepočítají."),
                        ("Využití",
                         "(pronájmy + vlastní zakázky) ÷ kapacita. Servis se nepočítá. V konfliktním dni je "
                         "jisté jen to, co servis nemůže zabrat; proto může vyjít rozpětí. Nikdy nepřesáhne "
                         "100 %."),
                        ("Obsazenost × využití",
                         "Obsazenost v kalendáři („Nejvytíženější stroj“) zahrnuje i servis – stroj je "
                         "nedostupný. Využití v tomto přehledu servis nezahrnuje – jen produktivní obsazení "
                         "pronájmy a vlastními zakázkami."),
                        ("Předběžné rezervace",
                         "Kapacitu neblokují a do využití se nepočítají – jsou uvedené jen pro informaci."),
                        ("Hodnocení",
                         f"Vysoké využití od {fmt_num(cv.HIGH_USE)} %, nízké pod {fmt_num(cv.LOW_USE)} %. Hranice jsou "
                         "orientační – nejsou podložené skutečnými firemními daty. Zasahuje-li rozpětí využití "
                         "do dvou kategorií, je hodnocení „nelze určit“."),
                        ("Nelze určit",
                         "Kde chybí údaje (např. zakázka bez termínu v plánu), zbylá kapacita není volno, "
                         "ale „nelze určit“."),
                    ]
                )


# ==================================================
# ZÁLOŽKA 8 – ROZHODOVACÍ PLÁNOVAČ (E2.5b)
# ==================================================
# Jen zobrazuje výsledek planner.py (nad capacity.py a economics.py); planner_view převádí
# formulář na požadavek a výsledek na karty. Požadavky existují jen v relaci uživatele –
# nic se neukládá, nepotvrzuje ani nemění v datech.

DP_KINDS = ["Pronájem stroje", "Vlastní zakázka"]
DP_NONE = "—"
DP_JOB_ROWS = 3

# Vysvětlení plánovače („Jak plánovač rozhoduje“). Rozhodovací centrum je ukazuje v jednom seznamu,
# Pronájem strojů ve skupinách – texty jsou stejné (doplněné o dvě věty z dřívějšího úvodu stránky).
DP_EXPLAIN = [
    ("Zdroj", "Rozhodovací pravidla aplikace nad kapacitním a ekonomickým výpočtem – stejná data a plán "
              "zakázek jako záložky 6 a 7."),
    ("Provoz", "Potvrzené pronájmy, zakázky v plánu a servis blokují kapacitu. Požadavek je proveditelný, "
               "jen když se vejde do zbytku; žádná doporučená varianta nepřekročí kapacitu."),
    ("Ekonomika", "Odděleně od provozu: příspěvek = tržba − variabilní náklady (u zakázky také přímé "
                  "náklady a penále). Fixní náklady strojů a společné náklady se neodečítají, nejde tedy "
                  "o čistý zisk. Chybějící údaj se nepočítá jako 0 – známá částka je jen horní mez. "
                  "Strojoden = 1 kus stroje na 1 den."),
    ("Štítky ekonomiky", "★ Doporučeno – prokazatelně nejlepší varianta. Méně výhodné – lepší je jiná "
                         "varianta. Nevyplatí se – každá varianta, která se vejde, je ztrátová. "
                         "= Rovnocenné – varianty vycházejí stejně. ⇄ Záleží na nákladech – rozhoduje "
                         "výše nezadaného nákladu (viz hranice zlomu). ? Nelze rozhodnout – chybí údaje. "
                         "Bez porovnání – požadavek nemá soupeře."),
    ("Doporučeno", "Jen když se ze soupeřících požadavků vejde část a jedna varianta je ekonomicky lepší "
                   "i v nejhorším případě. Samostatný požadavek bez soupeře je „Proveditelné“."),
    ("Nelze rozhodnout", "Chybí údaje o požadavku nebo kapacitě, nebo se vejde jen část soupeřících "
                         "požadavků a ekonomika chybí."),
    ("Konflikty", "Konflikt dvou potvrzených závazků se neřeší automaticky – jde k ručnímu řešení."),
    ("Alternativy", "Jen posun zakázky v jejím okně bez penále. Jiný stroj ani externí zapůjčení – "
                    "chybí údaje."),
    ("Nic se nemění", "Požadavky ani výsledky se neukládají; rezervace a zakázky v datech zůstávají beze "
                      "změny."),
]
_DP_EX = dict(DP_EXPLAIN)
DP_EXPLAIN_GROUPS = [
    ("Provoz a kapacita", [("Zdroj", _DP_EX["Zdroj"]), ("Provoz", _DP_EX["Provoz"]),
                           ("Předběžné rezervace", "Předběžné rezervace z dat soutěží o volnou kapacitu stejně jako "
                                                   "nové požadavky a nikdy se automaticky neruší."),
                           ("Nic se nemění", _DP_EX["Nic se nemění"])]),
    ("Ekonomika", [("Ekonomika", _DP_EX["Ekonomika"]),
                   ("Provoz × ekonomika", "Plánovač odpovídá zvlášť na dvě otázky – vejde se požadavek do kapacity "
                                          "a která varianta přinese nejvyšší příspěvek. Požadavek se může vejít, "
                                          "a přesto být ekonomicky méně výhodný.")]),
    ("Význam doporučení", [("Štítky ekonomiky", _DP_EX["Štítky ekonomiky"]), ("Doporučeno", _DP_EX["Doporučeno"]),
                           ("Nelze rozhodnout", _DP_EX["Nelze rozhodnout"])]),
    ("Konflikty a alternativy", [("Konflikty", _DP_EX["Konflikty"]), ("Alternativy", _DP_EX["Alternativy"])]),
]
assert sorted(k for _, items in DP_EXPLAIN_GROUPS for k, _ in items if k in _DP_EX) == sorted(_DP_EX)   # nic nechybí




@st.cache_data(show_spinner=False, max_entries=32)
def planner_result(data_version: str, code_version: str, job_starts: tuple, requests: tuple):
    """Výsledek plánovače jen z obyčejných typů (viz load_data_model)."""
    tables, _ = load_data_model(data_version, code_version)
    return pv.evaluate(dm.DataModel(**tables), dict(job_starts), pv.unfreeze(requests))


def dp_remove(ref: str) -> None:
    st.session_state["dp_requests"] = [r for r in st.session_state.get("dp_requests", []) if r["ref"] != ref]


def dp_clear() -> None:
    st.session_state["dp_requests"] = []


@st.fragment
def dp_economic_comparison(panels: list, sig: str) -> None:
    """Panely „Ekonomické porovnání“ s volitelnou simulací „Co kdyby“ (E2.5c 4D-B).

    Fragment: změna hypotetické hodnoty přepočítá jen tuto část (graf se značkou a rámeček simulace).
    Jedna nezadaná sazba (stejné pole i typ stroje, jediný scénář) = jedna společná hodnota pro všechny
    skupiny. Skutečné doporučení, karty, data ani rezervace se nemění.
    """
    shown = set()
    for panel in panels:
        spec = panel.get("sim")
        key = pv.sim_key(spec, sig) if spec and spec["available"] else None
        x = st.session_state.get(key) if key else None
        result = pv.simulate(spec, x) if key and x is not None else None
        st.markdown(pv.panel_html(panel, result), unsafe_allow_html=True)
        if key is None:
            continue
        if key not in shown:
            shown.add(key)
            st.number_input(
                f"Co kdyby – zkuste předpoklad: {spec['unknown']}", min_value=0.0, value=None, step=100.0,
                format="%.2f", key=key,
                help="Prázdné pole = žádná simulace. Zadaná 0 je platná hypotetická hodnota. Počítá se s hodnotou "
                     "na haléře, jak ji pole ukazuje. Nic se neukládá a skutečné doporučení se nemění.")
        else:
            ui.hint("Předpoklad pro tento náklad je společný se simulací výše – jde o stejný ekonomický parametr.")
        if result is not None:
            st.markdown(pv.sim_html(spec, result), unsafe_allow_html=True)


if PAGE in ("pronajem", "rozhodovaci_centrum"):              # dříve záložka 8 (Pronájem = stejný nástroj, jen pronájem)
    # Pronájem strojů: modernizovaný vzhled (karty jako Dashboard); Rozhodovací centrum beze změny vzhledu.
    # Formulář, validace, požadavky a výsledek plánovače jsou pro obě stránky stejné.
    DP_RENT = PAGE == "pronajem"
    if DP_RENT:
        st.markdown(ui.RENTAL_CSS, unsafe_allow_html=True)
    with (ui.tab_container("pronajem") if DP_RENT else contextlib.nullcontext()):
        if DP_RENT:
            ui.tab_header("Pronájem strojů", "Ověřte, zda se nový pronájem vejde do kapacity a zda se vyplatí.")
            st.markdown(pv.CSS, unsafe_allow_html=True)
            ui.note("<b>Ukázková data</b> – požadavky se nikam neukládají ani nic nemění v rezervacích či "
                    "zakázkách; po obnovení stránky zmizí. Podrobnosti v části „Jak plánovač rozhoduje“.")
        else:
            st.markdown("### Rozhodovací plánovač" if PAGE == "rozhodovaci_centrum" else "### Pronájem strojů – nový požadavek")
            st.markdown(pv.CSS, unsafe_allow_html=True)
            ui.note(
                "<b>Ukázková data</b> – pronájmy, servis a zakázky jsou fiktivní. Požadavky zadané zde se "
                "<b>nikam neukládají</b>, nic se nepotvrzuje ani nemění v rezervacích či zakázkách; po obnovení "
                "stránky zmizí."
            )
            st.markdown(
                '<div class="de-intro">Zadejte požadavek na pronájem nebo vlastní zakázku. Plánovač odpovídá zvlášť na '
                "dvě otázky: <b>Provoz</b> – vejde se požadavek do kapacity, kterou nechávají potvrzené pronájmy, "
                "zakázky z plánu záložky 6 a servis? <b>Ekonomika</b> – která varianta přinese nejvyšší příspěvek "
                "(tržba − variabilní náklady; nejde o čistý zisk)? Požadavek se může vejít, a přesto být ekonomicky "
                "méně výhodný. Předběžné rezervace z dat soutěží o volnou kapacitu stejně jako nové požadavky "
                "a nikdy se automaticky neruší.</div>",
                unsafe_allow_html=True,
            )

        dp_names = cv.type_names(data_model)
        dp_types = [t for t in dp_names if t in set(data_model.fleet_units["type_id"])]
        _, dp_plan, _ = plan_jobs(jobs_df, machine_capacity, code_version)
        dp_starts = cv.job_starts_from_plan(jobs_df, dp_plan)
        dp_requests = st.session_state.setdefault("dp_requests", [])

        if not DP_RENT:
            st.markdown("#### Nový požadavek")
        if PAGE == "pronajem":
            dp_kind = DP_KINDS[0]
        else:
            dp_kind = st.radio("Typ požadavku", DP_KINDS, horizontal=True, key="dp_kind")
        if dp_kind == DP_KINDS[1]:
            # výchozí hodnoty polí zakázky jen přes stav relace (šablona je pak může přepsat bez varování)
            for i in range(DP_JOB_ROWS):
                st.session_state.setdefault(f"dp_j_type{i}", DP_NONE)
                st.session_state.setdefault(f"dp_j_qty{i}", 1)
            st.session_state.setdefault("dp_j_start", 1)
            st.session_state.setdefault("dp_j_dur", 1)
            dp_jobs = data_model.jobs["job_id"].tolist()
            dp_tpl = st.selectbox(
                "Vyplnit podle existující zakázky (nepovinné – jen šablona, původní zakázka se nemění)",
                [DP_NONE] + dp_jobs, key="dp_template",
                format_func=lambda j: j if j == DP_NONE else f"{j} – {data_model.jobs.set_index('job_id').loc[j, 'name']}",
            )
            if dp_tpl != st.session_state.get("dp_template_used"):
                st.session_state["dp_template_used"] = dp_tpl
                if dp_tpl != DP_NONE:                      # předvyplnit formulář (před vykreslením polí)
                    tpl = pv.job_template(data_model, dp_tpl)
                    st.session_state["dp_j_name"] = tpl["name"]
                    for i in range(DP_JOB_ROWS):
                        t, q = tpl["needs"][i] if i < len(tpl["needs"]) else (DP_NONE, 1)
                        st.session_state[f"dp_j_type{i}"] = t if t in dp_types else DP_NONE
                        st.session_state[f"dp_j_qty{i}"] = q
                    st.session_state["dp_j_dur"] = tpl["duration"] or 1
                    st.session_state["dp_j_start"] = tpl["earliest"] or 1
                    for key, val in (("dp_j_earliest", tpl["earliest"]), ("dp_j_deadline", tpl["deadline"]),
                                     ("dp_j_revenue", tpl["revenue"]), ("dp_j_direct", tpl["direct_cost"]),
                                     ("dp_j_penalty", tpl["penalty_per_day"])):
                        st.session_state[key] = val      # chybějící hodnota zůstane prázdná, ne 0

        with (ui.dash_section("rent_form") if DP_RENT else contextlib.nullcontext()):
            if DP_RENT:
                ui.dash_section_head("truck", "Nový požadavek na pronájem",
                                     "Dny plánu (den 1 = první den) · prázdná cena = ceník stroje · požadavky jsou "
                                     "společné s Rozhodovacím centrem, kde zadáte i vlastní zakázku", tone="accent")
            with st.form("dp_form", border=not DP_RENT):
                if dp_kind == DP_KINDS[0]:
                    r1, r2 = st.columns([3, 1])
                    with r1:
                        dp_r_type = st.selectbox("Stroj", dp_types, format_func=lambda t: dp_names[t], key="dp_r_type")
                    with r2:
                        dp_r_qty = st.number_input("Počet kusů", min_value=1, max_value=10, value=1, step=1, key="dp_r_qty")
                    r3, r4, r5 = st.columns(3)
                    with r3:
                        dp_r_from = st.number_input("Od dne", min_value=1, max_value=365, value=21, step=1, key="dp_r_from")
                    with r4:
                        dp_r_to = st.number_input("Do dne (včetně)", min_value=1, max_value=365, value=24, step=1,
                                                  key="dp_r_to")
                    with r5:
                        dp_r_price = st.number_input(
                            "Cena za den (Kč, nepovinné)", min_value=0.0, value=None, step=500.0, key="dp_r_price",
                            help="Nevyplněno = ceník typu stroje z ukázkových dat.")
                else:
                    dp_j_name = st.text_input("Název zakázky (nepovinné)", key="dp_j_name", max_chars=80)
                    st.caption("Stroje, které zakázka potřebuje současně po celou dobu:")
                    dp_rows = []
                    for i in range(DP_JOB_ROWS):
                        c1, c2 = st.columns([3, 1])
                        with c1:
                            t = st.selectbox(f"Stroj {i + 1}", [DP_NONE] + dp_types, key=f"dp_j_type{i}",
                                             format_func=lambda t: t if t == DP_NONE else dp_names[t])
                        with c2:
                            q = st.number_input(f"Kusů {i + 1}", min_value=1, max_value=10, step=1, key=f"dp_j_qty{i}")
                        dp_rows.append((None if t == DP_NONE else t, q))
                    j1, j2, j3, j4 = st.columns(4)
                    with j1:
                        dp_j_start = st.number_input("Začátek (den)", min_value=1, max_value=365, step=1, key="dp_j_start")
                    with j2:
                        dp_j_dur = st.number_input("Délka (dny)", min_value=1, max_value=120, step=1, key="dp_j_dur")
                    with j3:
                        dp_j_earliest = st.number_input("Nejdřívější začátek", min_value=1, max_value=365, value=None,
                                                        step=1, key="dp_j_earliest",
                                                        help="Okno pro nabídku posunu bez penále (nepovinné).")
                    with j4:
                        dp_j_deadline = st.number_input("Termín dokončení", min_value=1, max_value=365, value=None,
                                                        step=1, key="dp_j_deadline")
                    e1, e2, e3 = st.columns(3)
                    with e1:
                        dp_j_revenue = st.number_input("Tržba (Kč)", min_value=0.0, value=None, step=1000.0,
                                                       key="dp_j_revenue")
                    with e2:
                        dp_j_direct = st.number_input("Přímé náklady (Kč)", min_value=0.0, value=None, step=1000.0,
                                                      key="dp_j_direct")
                    with e3:
                        dp_j_penalty = st.number_input("Penále za den (Kč)", min_value=0.0, value=None, step=500.0,
                                                       key="dp_j_penalty")
                    st.caption("Nevyplněný ekonomický údaj se nepočítá jako 0 – výsledek bude označený jako neúplný.")
                b1, b2 = st.columns(2)
                with b1:
                    dp_alone = st.form_submit_button("Vyhodnotit požadavek" if DP_RENT else "Vyhodnotit samostatně",
                                                     width="stretch")
                with b2:
                    dp_add = st.form_submit_button("Přidat k porovnání", width="stretch")

        if dp_alone or dp_add:
            existing = [] if dp_alone else [r["ref"] for r in dp_requests]
            ref = pv.next_ref(data_model, existing)
            if dp_kind == DP_KINDS[0]:
                new = pv.rental_request(ref, dp_r_type, dp_r_qty, dp_r_from, dp_r_to, dp_r_price)
            else:
                new = pv.job_request(ref, dp_j_name, dp_rows, dp_j_start, dp_j_dur, dp_j_earliest, dp_j_deadline,
                                     dp_j_revenue, dp_j_direct, dp_j_penalty)
            problems = pv.form_problems(new)
            if problems:
                st.error(" ".join(problems))
            elif dp_add and len(dp_requests) >= pv.MAX_REQUESTS:
                st.warning(f"K porovnání lze mít nejvýš {pv.MAX_REQUESTS} požadavků – nejdřív některý odeberte.")
            else:
                st.session_state["dp_requests"] = ([new] if dp_alone else dp_requests + [new])
                dp_requests = st.session_state["dp_requests"]

        if not dp_requests:
            ui.hint("Zatím není zadaný žádný požadavek.")
        else:
            with (ui.dash_section("rent_list") if DP_RENT else contextlib.nullcontext()):
                if DP_RENT:
                    ui.dash_section_head("route", "Porovnávané požadavky",
                                         "Vyhodnocují se společně – soupeří o stejnou volnou kapacitu", tone="neutral")
                else:
                    st.markdown("#### Porovnávané požadavky")
                for r in dp_requests:
                    c1, c2 = st.columns([5, 1])
                    what = "pronájem" if r["kind"] == "rental" else "zakázka"
                    machines = ", ".join(f"{dp_names.get(t, t)} × {q}" for t, q in r["needs"])
                    with c1:
                        st.markdown(f"**{pv.md_escape(r['ref'])}** – {what}: {pv.md_escape(machines)}, "
                                    f"dny {r['start']}–{r['end']}")
                    with c2:
                        st.button("Odebrat", key=f"dp_rm_{r['ref']}", on_click=dp_remove, args=(r["ref"],), width="stretch")
                st.button("Vymazat vše", key="dp_clear", on_click=dp_clear)

            res = planner_result(data_version, code_version, tuple(sorted(dp_starts.items())), pv.freeze(dp_requests))
            if DP_RENT:
                dp_related = [m for m in res["manual"] if m["related"]]
                dp_card = next(c for c in res["cards"] if c["ref"] == dp_requests[-1]["ref"])
                st.markdown(f'<div class="dd-section-title">Výsledek vyhodnocení<small>požadavek '
                            f'{html.escape(dp_card["ref"])}</small></div>', unsafe_allow_html=True)
                ui.dash_kpis(rv.rent_kpis(dp_card, dp_related))
                if len(res["cards"]) > 1:
                    st.markdown('<div class="dd-section-title">Porovnání variant<small>proveditelnost, přínos '
                                'a doporučení každé varianty</small></div>', unsafe_allow_html=True)
                    ui.dash_variants(rv.variant_items(res["cards"], dp_names))
                st.markdown('<div class="dd-section-title">Podrobné zdůvodnění<small>ekonomické porovnání, karty '
                            'požadavků, kalendář a konflikty</small></div>', unsafe_allow_html=True)
            else:
                st.markdown("#### Výsledek")
            # E2.5c 4D-B: „Co kdyby“ – hypotetická hodnota se drží jen v relaci; klíč obsahuje otisk dat
            # a požadavků, takže se po jejich změně hodnota sama vymaže (staré klíče se odstraní)
            dp_sig = hashlib.sha256(repr((data_version, sorted(dp_starts.items()), pv.freeze(dp_requests)))
                                    .encode("utf-8")).hexdigest()[:12]
            dp_sim_keys = {pv.sim_key(p["sim"], dp_sig) for p in res["panels"] if p.get("sim") and p["sim"]["available"]}
            for k in [k for k in st.session_state if str(k).startswith(pv.SIM_KEY_PREFIX) and k not in dp_sim_keys]:
                del st.session_state[k]
            dp_economic_comparison(res["panels"], dp_sig)
            for card in res["cards"]:
                st.markdown(pv.card_html(card, dp_names), unsafe_allow_html=True)
                if card["blocking"]:
                    with st.expander(f"Co blokuje {card['ref']} ({len(card['blocking'])} dní × stroj)"):
                        st.dataframe(pv.blocking_table(card, dp_names), width="stretch", hide_index=True)
                for alt in card["alternatives"]:
                    st.markdown(f"- Alternativa (jen návrh, nic se nemění): {pv.md_escape(alt)}")
                if card["note"]:
                    ui.hint(html.escape(card["note"].capitalize()))
            if res["tentative"]:
                ui.hint(
                    "Předběžné rezervace v datech (rovnocenné novým požadavkům, nikdy se automaticky neruší): "
                    + html.escape(", ".join(f"{t['ref']} – {dp_names.get(t['type_id'], t['type_id'])} dny "
                                            f"{t['start']}–{t['end']}" for t in res["tentative"]))
                )

            dp_types_used = tuple(t for t in dp_names if any(t == x for r in dp_requests for x, _ in r["needs"]))
            dp_lo = max(1, min(r["start"] for r in dp_requests) - 2)
            dp_hi = max(r["end"] for r in dp_requests) + 2
            if not DP_RENT:
                st.markdown("#### Dostupnost dotčených strojů")
            _, dp_html, _ = capacity_grid(data_version, code_version, tuple(sorted(dp_starts.items())), dp_lo, dp_hi,
                                          dp_types_used, None)
            with (st.expander("Dostupnost dotčených strojů – kalendář") if DP_RENT else contextlib.nullcontext()):
                st.markdown(dp_html, unsafe_allow_html=True)
                ui.hint("Kalendář ukazuje kapacitu po potvrzených závazcích a servisu – nové požadavky v něm nejsou.")

            related = [m for m in res["manual"] if m["related"]]
            if DP_RENT:
                st.markdown('<div class="dd-section-title">Konflikty potvrzených závazků</div>', unsafe_allow_html=True)
            else:
                st.markdown("#### Konflikty potvrzených závazků")
            if related:
                for m in related:
                    ui.insight(f"<b>K ručnímu řešení:</b> {html.escape(m['text'])}"
                               + "".join(f"<br>Posun zakázky: {html.escape(a)}" for a in m["alternatives"]), tone="bad")
            else:
                ui.hint("U dotčených strojů a dnů nejsou konflikty potvrzených závazků.")
            if len(res["manual"]) > len(related):
                with st.expander(f"Všechny konflikty potvrzených závazků ({len(res['manual'])})"):
                    for m in res["manual"]:
                        st.markdown(f"- {pv.md_escape(m['text'])}")

            with st.expander("Chybějící údaje" + (f" ({len(res['missing_relevant'])})" if res["missing_relevant"] else "")):
                if res["missing_relevant"]:
                    for x in res["missing_relevant"]:
                        st.markdown(f"- {pv.md_escape(x)}")          # může obsahovat text od uživatele
                else:
                    st.markdown("Pro vyhodnocení zadaných požadavků žádné údaje nechybí.")
                for n in res["notes"]:
                    st.markdown(f"- {pv.md_escape(n)}")

        with st.expander("Jak plánovač rozhoduje"):
            if DP_RENT:                                      # stejné texty, jen rozdělené do skupin
                for dp_group, dp_items in DP_EXPLAIN_GROUPS:
                    st.markdown(f'<div class="dd-group">{dp_group}</div>', unsafe_allow_html=True)
                    ui.explain(dp_items)
            else:
                ui.explain(DP_EXPLAIN)


if PAGE == "zakazky":
    page_link("rozhodovaci_centrum", "Ověřit novou zakázku v Rozhodovacím centru")


# ==================================================
# DASHBOARD (etapa M2) – jen souhrn existujících výsledků, žádné nové výpočty
# ==================================================

if PAGE == "dashboard":
    st.markdown(ui.DASHBOARD_CSS, unsafe_allow_html=True)
    with ui.tab_container("dashboard"):
        ui.tab_header(
            "Přehled půjčovny",
            "Souhrn výsledků sekcí na <b>ukázkových datech</b> – Dashboard nic nepřepočítává, čísla odpovídají "
            "sekcím, kam vedou odkazy.",
        )
        # Plán zakázek a kapacita: stejné funkce a výchozí nastavení jako při otevření sekcí
        _, dash_plan, _ = plan_jobs(jobs_df, machine_capacity, code_version)
        dash_sum = sch.summary(dash_plan)
        dash_starts = cv.job_starts_from_plan(jobs_df, dash_plan)
        dash_cal = cv.build(data_model, dash_starts)
        dash_end, _ = cv.period_bounds(dash_cal)
        dash_names = cv.type_names(data_model)
        dash_types = [t for t in dash_names if t in dash_cal.type_ids()]
        dash_grid, _, dash_conflicts = capacity_grid(data_version, code_version, tuple(sorted(dash_starts.items())),
                                                     1, dash_end, tuple(dash_types), None)
        dash_cap = cv.summary(dash_grid, dash_names)
        n_dash_conf = len(dash_cap["conflict_days"])
        # Využití flotily: totéž jako Plánování kapacity → Přehled využití ve výchozím nastavení
        dash_use = cv.fleet_utilization(cv.utilization_table(dash_grid, dash_names, dash_cal.unknown_capacity),
                                        dash_grid)
        dash_capacity = dash_use["capacity"]
        dash_pct = lambda v: fmt_pct(100.0 * v / dash_capacity, 0) if dash_capacity else "nelze určit"
        dash_bookings = data_model.rental_bookings
        dash_confirmed = int((dash_bookings["status"] == cp.BLOCKING).sum())
        dash_tentative = len(dash_bookings) - dash_confirmed
        dash_requests = st.session_state.get("dp_requests", [])
        dash_rec_profit = expected_profit_dumper if better_machine == machine_2 else expected_profit_mini
        dash_rec_loss = prob_loss_dumper if better_machine == machine_2 else prob_loss_mini
        dash_period = f"dny 1–{dash_end}"
        dash_word = lambda n, one, few, many: cv.plural(n, one, few, many).split(" ", 1)[1]   # jen slovo

        # ---- 1. Hlavní KPI ------------------------------------------------------------------------
        ui.dash_kpis([
            dict(icon="gauge", label="Vytížení strojového parku", tone="neutral",
                 value=cv.share_text(dash_use["util_min"], dash_use["util_max"]),
                 sub=f"pronájmy a zakázky z {fmt_num(dash_capacity)} strojodnů · {dash_period}"
                     + (f" · bez: {html.escape(', '.join(dash_use['excluded']))}" if dash_use["excluded"] else "")),
            dict(icon="coins", label="Ekonomický výsledek", tone="accent",
                 value=fmt_czk(dash_sum["profit_after_penalty"]),
                 chip=ui.dash_chip("dílčí údaj – jen vlastní zakázky", "info"),
                 sub="zisk vlastních zakázek po penále · celkový výsledek půjčovny zatím není k dispozici"),
            dict(icon="truck", label="Pronájmy a zakázky", tone="neutral",
                 pair=True,
                 value=(f"<span>{dash_confirmed}<small>{dash_word(dash_confirmed, 'pronájem', 'pronájmy', 'pronájmů')}</small></span> "
                        f"<span>· {dash_sum['scheduled']}<small>{dash_word(dash_sum['scheduled'], 'zakázka', 'zakázky', 'zakázek')}</small></span>"),
                 sub=f"potvrzené pronájmy a zakázky v plánu · +{cv.plural(dash_tentative, 'předběžná rezervace', 'předběžné rezervace', 'předběžných rezervací')} · "
                     "stav k dnešku nelze určit"),
            dict(icon="alert", label="Kapacitní konflikty", tone="bad" if n_dash_conf else "good",
                 value=f"{n_dash_conf}<small>{dash_word(n_dash_conf, 'den', 'dny', 'dní')}</small>",
                 chip=ui.dash_chip("vyžaduje řešení", "bad") if n_dash_conf else ui.dash_chip("bez konfliktu", "good"),
                 sub=("bez konfliktu" if not n_dash_conf else "dny " + cv.days_text(dash_cap["conflict_days"]))
                     + f" · všechny stroje, {dash_period}"),
        ])

        # ---- 2. Provoz a ekonomika ----------------------------------------------------------------
        col_op, col_ec = st.columns(2, gap="medium")
        with col_op, ui.dash_section("provoz"):
            ui.dash_section_head("route", "Provoz půjčovny a vlastních zakázek",
                                 f"Obsazenost strojů a plán zakázek · {dash_period}", tone="neutral")
            ui.dash_rows([
                ("Pronájmy", f"{fmt_num(dash_use['rental'])} strojodnů",
                 f"{dash_pct(dash_use['rental'])} kapacity · potvrzené rezervace"),
                ("Vlastní zakázky", f"{fmt_num(dash_use['job'])} strojodnů",
                 f"{dash_pct(dash_use['job'])} kapacity · podle plánu zakázek"),
                ("Servis", f"{fmt_num(dash_use['maintenance'])} strojodnů",
                 f"{dash_pct(dash_use['maintenance'])} kapacity · stroj nedostupný"),
                ("Volná kapacita", f"{fmt_num(dash_use['free'])} strojodnů", f"{dash_pct(dash_use['free'])} kapacity",
                 True),
                ("Nejvytíženější stroj", html.escape(dash_cap["busiest"] or "–"),
                 (fmt_pct(dash_cap["busiest_share"], 0) + " obsazenost vč. servisu")
                 if dash_cap["busiest_share"] is not None else ""),
                ("Zakázky po termínu", f"{dash_sum['late']} z {dash_sum['scheduled']}",
                 f"délka plánu {dash_sum['length']} dní"),
            ])
            with ui.dash_links("provoz"):
                page_link("kapacita")
                page_link("zakazky")
                page_link("pronajem")
        with col_ec, ui.dash_section("ekonomika"):
            ui.dash_section_head("coins", "Ekonomika a hospodaření",
                                 "Plán zakázek za celé období · investice měsíčně")
            ui.dash_rows([
                ("Tržby zakázek", fmt_czk(dash_sum["revenue"]), "plán vlastních zakázek"),
                ("Náklady zakázek", fmt_czk(dash_sum["costs"]), ""),
                ("Penále za zpoždění", fmt_czk(dash_sum["penalty"]), f"{dash_sum['late']} zakázky po termínu"
                 if dash_sum["late"] else "bez zpoždění"),
                ("Zisk zakázek po penále",
                 ('<span class="v neg">' if dash_sum["profit_after_penalty"] < 0 else "")
                 + fmt_czk(dash_sum["profit_after_penalty"])
                 + ("</span>" if dash_sum["profit_after_penalty"] < 0 else ""), "dílčí výsledek – jen vlastní zakázky",
                 True),
                (("Relativně lepší stroj" if both_loss else "Doporučený stroj") + " – očekávaný zisk",
                 fmt_czk(dash_rec_profit),
                 f"{html.escape(better_machine)} · simulace" + (" · oba stroje prodělávají" if both_loss else "")),
                ("Doporučená sestava – provozní zisk",
                 fmt_czk(purchase_best["profit"]) if purchase_best["units"] > 0 else "–",
                 f"investice {fmt_czk(purchase_best['investment'])}" if purchase_best["units"] > 0
                 else "žádná zisková sestava v rozpočtu"),
            ])
            ui.dash_text("Ukázková data. Zisk zakázek nezahrnuje pronájmy ani fixní náklady flotily – nejde o celkový "
                         "zisk půjčovny.", muted=True)
            with ui.dash_links("ekonomika"):
                page_link("investice")
                page_link("strojovy_park")

        # ---- 3. Doporučení a upozornění -------------------------------------------------------------
        st.markdown('<div class="dd-group">Doporučení a upozornění</div>', unsafe_allow_html=True)
        with st.container(key="dd_alerts"):        # na užších obrazovkách karty pod sebou (CSS)
            col_inv, col_cap, col_dc = st.columns(3, gap="medium")
        with col_inv, ui.dash_section("investice_accent"):
            ui.dash_section_head("spark", "Investiční doporučení", tone="accent")
            if purchase_best["units"] > 0:
                dash_set = " · ".join(f"{k}× {name}" if k > 1 else name for name, k in purchase_best["counts"].items())
                dash_set_sub = (f"{fmt_czk(purchase_best['profit'])} měsíčně · investice "
                                f"{fmt_czk(purchase_best['investment'])}")
            else:
                dash_set, dash_set_sub = "Nekupovat", "žádná zisková sestava v rozpočtu"
            ui.dash_kv("Relativně lepší stroj" if both_loss else "Doporučený stroj", html.escape(better_machine),
                       f"porovnání {html.escape(machine_1)} × {html.escape(machine_2)} · očekávaný zisk {fmt_czk(dash_rec_profit)}"
                       + (" – oba prodělávají" if both_loss else "")
                       + f" · šance ztrátového měsíce {fmt_pct(dash_rec_loss * 100)}")
            ui.dash_kv("Doporučená sestava (Nákup strojů)", html.escape(dash_set), dash_set_sub)
            with ui.dash_links("investice"):
                page_link("investice", "Otevřít Investice")
        with col_cap, ui.dash_section("konflikty_" + ("bad" if len(dash_conflicts) else "good")):
            ui.dash_section_head("alert", "Kapacitní konflikty", tone="bad" if len(dash_conflicts) else "good")
            if len(dash_conflicts):
                items = "".join(
                    f'<li><b>{html.escape(str(r["Stroj"]))}</b> – {html.escape(str(r["Dny"]))}'
                    f'<small>chybí {html.escape(str(r["Chybí kusů"]))} ks · {html.escape(str(r["Jistota"]))}</small></li>'
                    for _, r in dash_conflicts.head(4).iterrows()
                )
                more = len(dash_conflicts) - 4
                st.markdown(ui.dash_chip(cv.plural(len(dash_conflicts), "konflikt", "konflikty", "konfliktů")
                                         + " k řešení", "bad")
                            + f'<ul class="dd-list">{items}</ul>'
                            + (f'<div class="dd-cap">a další ({more})</div>' if more > 0 else ""),
                            unsafe_allow_html=True)
            else:
                st.markdown(ui.dash_chip("bez konfliktu", "good"), unsafe_allow_html=True)
                ui.dash_text(f"V období {dash_period} nemá žádný stroj víc nároků než kusů.", muted=True)
            with ui.dash_links("konflikty"):
                page_link("kapacita", "Otevřít kalendář kapacity")
        with col_dc, ui.dash_section("rozhodovaci_centrum_neutral"):
            ui.dash_section_head("route", "Rozhodovací centrum", tone="neutral")
            if dash_requests:
                dash_res = planner_result(data_version, code_version, tuple(sorted(dash_starts.items())),
                                          pv.freeze(dash_requests))
                dash_refs = {r["ref"] for r in dash_requests}
                items = "".join(
                    f'<li><b>{html.escape(c["ref"])}</b> – {html.escape(c["title"])}'
                    f'<small>Provoz: {html.escape(c["op_badge"])}'
                    + (f' · Ekonomika: {html.escape(c["econ_badge"])}' if c["econ_badge"] else "")
                    + "</small></li>"
                    for c in dash_res["cards"] if c["ref"] in dash_refs
                )
                st.markdown(ui.dash_chip(cv.plural(len(dash_requests), "rozpracovaný požadavek",
                                                   "rozpracované požadavky", "rozpracovaných požadavků"), "info")
                            + f'<ul class="dd-list">{items}</ul>', unsafe_allow_html=True)
            else:
                st.markdown(ui.dash_chip("0 rozpracovaných požadavků", "info"), unsafe_allow_html=True)
                ui.dash_text("Zatím žádný požadavek. Nový pronájem nebo zakázku ověříte v Rozhodovacím centru "
                             "(jen v této relaci, nic se neukládá).", muted=True)
            with ui.dash_links("rozhodovaci_centrum"):
                page_link("rozhodovaci_centrum")
                page_link("pronajem", "Nový pronájem")
        ui.hint(cv.data_info(data_model, data_version))


# ==================================================
# EKONOMIKA FLOTILY (E2.2 – K3.4c) – výsledek flotily za období
# ==================================================
# Jen zobrazuje výsledek fleet_economics.period_result převedený zobrazovací vrstvou fleet_view.build.
# Termíny zakázek z plánu Vlastních zakázek (stejná funkce plan_jobs, výsledek z mezipaměti), výchozí
# období jako Plánování kapacity. Nic se nevybírá, neruší ani nepřesouvá; fixní náklady a režie se
# odečtou jen při výslovně zvolené metodě přepočtu (výchozí: nezadávat).
FLEET_ECON_TITLE = "Ekonomika flotily"
FLEET_ECON_INTRO = ("Společný ekonomický výsledek flotily za zvolené období: jisté příspěvky pronájmů "
                    "a vlastních zakázek, fixní náklady, režie a konflikty. Vychází z <b>ukázkových dat</b>.")
FLEET_ECON_NOT_NET_PROFIT = fv.NOT_NET_PROFIT
FLEET_ECON_METHODS = {
    "Nezadávat – fixní náklady a režie se neodečtou (výsledek je horní mez)": None,
    "Zadat počet měsíců": "months",
    "Podle kalendáře (vyžaduje datum prvního dne plánu)": "calendar_days",
}
FE_CHIP_TONE = {"good": "good", "warn": "warn", "info": "info", "bad": "bad", "neutral": "info"}
FE_KPI_ICONS = {"pronajem": "truck", "zakazky": "route", "fixni": "gauge", "celkem": "coins"}


@st.cache_data(show_spinner="Počítám ekonomiku flotily…", max_entries=32)
def fleet_economics_view(data_version: str, code_version: str, job_starts: tuple, start: int, end: int,
                         method, months, plan_start):
    """Obsah stránky Ekonomika flotily (fleet_view.build nad fleet_economics.period_result).

    Výpočet trvá jednotky sekund – mezipaměť ho přepočítá jen při změně vstupů nebo dat / kódu.
    Ukládají se jen obyčejné hodnoty (texty, čísla, seznamy), viz load_data_model.
    """
    tables, _ = load_data_model(data_version, code_version)
    model = dm.DataModel(**tables)
    pr = fe.period_result(model, start, end, dict(job_starts), fixed_cost_method=method, months=months,
                          plan_start=plan_start)
    return fv.build(pr, cv.type_names(model), plan_start)


def fe_chip(status: str, tone: str) -> str:
    return ui.dash_chip(status, FE_CHIP_TONE.get(tone, "info"))


def fe_open_calendar(type_ids: list) -> None:
    """Plánování kapacity s filtrem na stroje konfliktu (stejné klíče jako show_in_calendar).

    Widgety kalendáře na této stránce nejsou, proto lze jejich stav nastavit před přepnutím stránky.
    Nic jiného se nemění – jen filtr zobrazení."""
    st.session_state["cap_types"] = list(type_ids)
    st.session_state["cap_detail_type"] = type_ids[0]
    st.session_state["cap_view"] = CAP_VIEWS[0]
    st.switch_page(_PAGE_OBJ["kapacita"])


def fe_open_usage(start: int, end: int, plan_start) -> None:
    """Plánování kapacity → Přehled využití se stejným obdobím (všechny stroje). Jen filtry zobrazení."""
    st.session_state["cap_types"] = []
    st.session_state["cap_period"] = (int(start), int(end))
    st.session_state["cap_plan_start"] = plan_start
    st.session_state["cap_view"] = CAP_VIEWS[1]
    st.switch_page(_PAGE_OBJ["kapacita"])


def fe_entry_row(r: dict, detail: bool = False) -> tuple:
    """Řádek závazku pro ui.dash_rows: (popisek, částka, poznámka se štítkem stavu) – na mobilu pod sebou."""
    label = f"{r['ref']} – {r['kind']}" + (f" · {r['name']}" if r["name"] else "")
    facts = [r["period"], r["machines_text"], r["reason"]]
    if detail:
        facts += [f"termín: {r['date_source']}", f"kus: {r['unit_assignment']}"]
    note = html.escape(" · ".join(x for x in facts if x)) + " · " + fe_chip(r["status"], r["tone"])
    return label, html.escape(r["text"]), note


def fe_banner(b: dict) -> None:
    text = html.escape(b["text"])
    if b["text"].startswith(FLEET_ECON_NOT_NET_PROFIT):
        text = f"<b>{html.escape(FLEET_ECON_NOT_NET_PROFIT)}</b>" + html.escape(b["text"][len(FLEET_ECON_NOT_NET_PROFIT):])
    ui.insight(text, "bad" if b["tone"] in ("bad", "warn") else "")


if PAGE == "ekonomika":
    st.markdown(ui.ECON_CSS, unsafe_allow_html=True)
    with ui.tab_container("ekonomika"):
        ui.tab_header(FLEET_ECON_TITLE, FLEET_ECON_INTRO)
        fe_badge = st.empty()
        _, fe_plan, _ = plan_jobs(jobs_df, machine_capacity, code_version)
        fe_starts = cv.job_starts_from_plan(jobs_df, fe_plan)
        fe_default_end, fe_max_day = cv.period_bounds(cv.build(data_model, fe_starts))
        fe_top = ui.summary_slot("ekonomika")                 # upozornění a ukazatele – naplní se po filtrech

        with ui.filter_box("ekonomika", "Období a fixní náklady"):
            f1, f2 = st.columns([3, 2])
            with f1:
                fe_from, fe_to = st.slider("Období (dny plánu)", min_value=1, max_value=fe_max_day,
                                           value=(1, fe_default_end), key="fe_period")
            with f2:
                fe_start = st.date_input("Datum prvního dne plánu (nepovinné)", value=None, key="fe_plan_start",
                                         format="DD.MM.YYYY")
            fe_label = st.radio("Fixní náklady a režie za období", list(FLEET_ECON_METHODS), key="fe_method")
            fe_method = FLEET_ECON_METHODS[fe_label]
            fe_months = None
            if fe_method == "months":
                fe_months = st.number_input("Počet měsíců", min_value=0.01, max_value=24.0, value=1.0, step=0.25,
                                            format="%.2f", key="fe_months")
            ui.hint("Metodu přepočtu měsíčních fixních nákladů a režie na období volíte vy – aplikace žádnou "
                    "nepředpokládá. Režie se odečte jen tehdy, když je v datech zadaná.")

        fe_view = fleet_economics_view(data_version, code_version, tuple(sorted(fe_starts.items())),
                                       int(fe_from), int(fe_to), fe_method, fe_months, fe_start)
        fe_badge.markdown(ui.status_chip("Ukázková data" + (" / neúplné" if fe_view["status"] != fv.COMPLETE else ""),
                                         "warn"), unsafe_allow_html=True)

        with fe_top:
            for b in fe_view["banners"]:
                fe_banner(b)
            if fe_view["computable"]:
                st.markdown(f'<div class="dd-section-title">Hlavní ukazatele <small>{html.escape(fe_view["period"])}'
                            "</small></div>", unsafe_allow_html=True)
                ui.dash_kpis([dict(
                    icon=FE_KPI_ICONS[h["key"]], label=h["label"], value=html.escape(h["text"]), cls="word",
                    chip=fe_chip(h["status"], h["tone"]),
                    sub=html.escape(h["meaning"]) + f'<span class="dd-period">Období: {html.escape(h["period"])}</span>',
                    tone="neutral") for h in fe_view["headline"]])

        if fe_view["computable"]:
            st.markdown('<div class="dd-section-title">Rozpad výsledku <small>jen podložené částky – neznámé '
                        "položky jsou označené</small></div>", unsafe_allow_html=True)
            fe_cols = st.columns(len(fe_view["components"]))
            for col, g in zip(fe_cols, fe_view["components"]):
                with col, ui.dash_section("ekon_" + {"Pronájem strojů": "pronajem", "Vlastní zakázky": "zakazky"}
                                          .get(g["group"], "flotila")):
                    ui.dash_section_head({"Pronájem strojů": "truck", "Vlastní zakázky": "route"}.get(g["group"], "coins"),
                                         g["group"], html.escape(g["note"]))
                    ui.dash_rows([(r["item"], html.escape(r["text"]), fe_chip(r["status"], r["tone"]),
                                   r.get("total", False)) for r in g["rows"]])

            st.markdown('<div class="dd-section-title">Co chybí k přesnému výsledku</div>', unsafe_allow_html=True)
            with ui.dash_section("ekon_chybi_" + ("bad" if fe_view["missing"] else "good")):
                if fe_view["missing"]:
                    for grp in fe_view["missing"]:
                        st.markdown(f'<div class="dd-group">{html.escape(grp["group"])}</div><ul class="dd-list">'
                                    + "".join(f"<li>{html.escape(item)}</li>" for item in grp["items"]) + "</ul>",
                                    unsafe_allow_html=True)
                else:
                    ui.dash_text("Pro zvolené období a metodu nechybí žádný údaj ani rozhodnutí.")

            # K3.4e2: stručný souhrn využití strojů (Usage.fleet – stejná čísla jako Plánování kapacity)
            fe_use = fe_view["usage_summary"]
            if fe_use:
                st.markdown('<div class="dd-section-title">Využití strojů za období <small>'
                            + html.escape(f"{fe_view['period']} · kapacita flotily {fe_use['capacity_text']}")
                            + "</small></div>", unsafe_allow_html=True)
                with ui.dash_section("ekon_vyuziti_neutral"):
                    ui.dash_rows([(r["label"], html.escape(r["text"]), html.escape(r["meaning"]))
                                  for r in fe_use["rows"]])
                    ui.dash_text(html.escape(fe_use["composition"]), muted=True)
                    for note in fe_use["notes"]:
                        ui.dash_text(html.escape(note), muted=True)
                    with ui.dash_links("ekon_vyuziti"):
                        if st.button("Podrobné využití v Plánování kapacity", key="fe_usage_detail",
                                     icon=":material/calendar_month:", type="tertiary"):
                            fe_open_usage(fe_from, fe_to, fe_start)

            fe_cases = fe_view["conflicts"]
            st.markdown('<div class="dd-section-title">Kapacitní konflikty a možné varianty <small>'
                        + html.escape(f"{len(fe_cases)} {fv._plural(len(fe_cases), 'případ', 'případy', 'případů')}")
                        + "</small></div>", unsafe_allow_html=True)
            if not fe_cases:
                ui.dash_text("V období nejsou konflikty potvrzených pronájmů, zakázek a servisu.", muted=True)
            for case in fe_cases:
                with ui.dash_section(f"ekon_konflikt_{case['id']}_bad"):
                    ui.dash_section_head("alert", case["title"], html.escape(case["note"]), tone="bad")
                    ui.dash_rows([(f"{c['ref']} – {c['kind']}", html.escape(c["text"]),
                                   html.escape(c["period"]) + " · " + fe_chip(c["status"], c["tone"]))
                                  for c in case["commitments"]]
                                 + ([("Servis", "", html.escape(", ".join(case["services"])))] if case["services"] else []))
                    ui.dash_variants([dict(
                        title=v["title"], sub=f"{v['feasibility']} · vyřeší konflikt: {v['resolves']}",
                        tag=v["amount"]["status"],
                        rows=[("Příspěvek dotčených závazků", html.escape(v["amount"]["text"]))],
                        why="<ul>" + "".join(f"<li>{html.escape(x)}</li>" for x in
                                             ([v["change"]] if v["change"] else [])
                                             + ([f"Bližší termíny blokují: {', '.join(v['blockers'])}"]
                                                if v["blockers"] else [])
                                             + [f"Varování: {w}" for w in v["warnings"]]
                                             + ([v["amount"]["explanation"]] if v["amount"]["explanation"] else []))
                        + "</ul>") for v in case["variants"]])
                    # K3.4d: porovnání variant – jen informace (econ_formula), nic se nevybírá ani nedoporučuje
                    with st.expander(f"Porovnání variant – jen informace ({len(case['comparisons'])})"):
                        if case["comparison_note"]:
                            ui.dash_text(html.escape(case["comparison_note"]))
                        if case["comparisons"]:
                            st.markdown('<ul class="dd-list">' + "".join(
                                f"<li><b>{html.escape(c['a'])} × {html.escape(c['b'])}</b> · {html.escape(c['kind'])}"
                                f"<small>{html.escape(c['text'])}</small></li>" for c in case["comparisons"])
                                + "</ul>", unsafe_allow_html=True)
                        ui.dash_text(html.escape(fv.COMPARISON_DISCLAIMER), muted=True)
                    with ui.dash_links(f"ekon_{case['id']}"):
                        if st.button("Zobrazit v kalendáři", key=f"fe_cal_{case['id']}",
                                     icon=":material/calendar_month:", type="tertiary"):
                            fe_open_calendar(case["types"])
                        if case["jobs"]:
                            page_link("zakazky", "Vlastní zakázky")
            if fe_cases:
                sc = fe_view["scenarios"]
                with ui.dash_section("ekon_rozpeti_neutral"):
                    ui.dash_kv("Rozpětí celkového výsledku přes ověřené varianty", html.escape(sc["cell"]["text"]),
                               fe_chip(sc["cell"]["status"], sc["cell"]["tone"]) + " "
                               + html.escape(sc["cell"]["explanation"] or sc.get("text", "")))
                    ui.dash_text(html.escape(sc.get("text", "")) + " Rozhodnutí je na majiteli půjčovny.", muted=True)

            # K3.4e1: předběžné, přesahující a nejisté závazky – nezapočítané částky (jen informace)
            fe_pending = fe_view["pending"]
            st.markdown('<div class="dd-section-title">Předběžné a nejisté závazky <small>'
                        + html.escape(f"nezapočítáno – {len(fe_pending)} "
                                      + fv._plural(len(fe_pending), "závazek", "závazky", "závazků"))
                        + "</small></div>", unsafe_allow_html=True)
            with ui.dash_section("ekon_nezapocitano_neutral"):
                ui.dash_text(html.escape(fv.PENDING_NOTE), muted=True)
                if fe_pending:
                    ui.dash_rows([fe_entry_row(r) for r in fe_pending])
                else:
                    ui.dash_text("V období nejsou předběžné, přesahující ani nejisté závazky.", muted=True)

            # K3.4e1: evidence všech závazků z dat (každý právě jednou, ve skupinách podle stavu)
            with st.expander(f"Evidence závazků ({len(fe_view['ledger'])})"):
                ui.dash_text("Každý pronájem a každá vlastní zakázka z dat právě jednou. Do výsledku se započítává "
                             "jen skupina „Jisté využití“; částky jsou příspěvky celých závazků.", muted=True)
                for grp in fe_view["ledger_groups"]:
                    st.markdown(f'<div class="dd-group">{html.escape(grp["title"])} ({len(grp["rows"])})</div>',
                                unsafe_allow_html=True)
                    ui.dash_rows([fe_entry_row(r, detail=True) for r in grp["rows"]])

        # K3.4e2: metodika – jen popis skutečné implementace, bez závěrů a rad
        with st.expander("Jak se počítá"):
            ui.explain([(title, html.escape(text)) for title, text in fe_view["methodology"]])


# ==================================================
# STROJOVÝ PARK (etapa M2) – existující údaje o strojích
# ==================================================

if PAGE == "strojovy_park":
    # Modernizace UI: karty a sekce jako Dashboard / Vlastní zakázky. Počty kusů = data flotily
    # (machine_capacity), ekonomika = jeden kus podle portfolio.type_month – nic nového se nepočítá.
    st.markdown(ui.FLEET_CSS, unsafe_allow_html=True)
    with ui.tab_container("strojovy_park"):
        ui.tab_header(
            "Strojový park",
            "Typy strojů, počty kusů ve flotile a modelová ekonomika <b>jednoho kusu</b> podle "
            "<b>ukázkových dat</b>. Údaje se zde jen zobrazují.",
        )
        fleet_econ = unit_economics_rows()
        fleet_total = sum(machine_capacity.values())
        fleet_types = len(machine_capacity)
        fleet_by_profit = sorted(fleet_econ, key=lambda r: r["profit"], reverse=True)
        fleet_by_revenue = sorted(fleet_econ, key=lambda r: r["revenue"], reverse=True)
        fleet_loss = [r for r in fleet_econ if r["profit"] < 0]
        fleet_word = lambda n, one, few, many: cv.plural(n, one, few, many).split(" ", 1)[1]

        # ---- KPI (souhrn za flotilu jen tam, kde je jednoznačný: počty kusů a typů) -------------------
        ui.dash_kpis([
            dict(icon="truck", label="Počet strojů ve flotile", tone="neutral",
                 value=f"{fleet_total}<small>{fleet_word(fleet_total, 'kus', 'kusy', 'kusů')}</small>",
                 sub="fyzické kusy z dat flotily · ukázková data"),
            dict(icon="route", label="Počet typů strojů", tone="neutral",
                 value=f"{fleet_types}<small>{fleet_word(fleet_types, 'typ', 'typy', 'typů')}</small>",
                 sub="různé kategorie strojů ve flotile"),
            dict(icon="coins", label="Měsíční tržby jednoho kusu", tone="accent",
                 value=f"{fmt_num(fleet_by_revenue[-1]['revenue'])}–{fmt_czk(fleet_by_revenue[0]['revenue'])}",
                 chip=ui.dash_chip("model · 1 kus", "info"),
                 sub=f"rozpětí podle typu · nejvíc {html.escape(fleet_by_revenue[0]['machine'])}"),
            dict(icon="spark", label="Měsíční provozní zisk jednoho kusu", tone="accent",
                 value=f"{fmt_num(fleet_by_profit[-1]['profit'])}–{fmt_czk(fleet_by_profit[0]['profit'])}",
                 chip=(ui.dash_chip(f"{len(fleet_loss)} z {fleet_types} typů ve ztrátě", "bad") if fleet_loss
                       else ui.dash_chip(f"všech {fleet_types} typů v zisku", "good")),
                 sub=f"rozpětí podle typu · nejvíc {html.escape(fleet_by_profit[0]['machine'])}"),
        ])
        ui.dash_text("Souhrn tržeb a zisku za celou flotilu se záměrně nezobrazuje – ekonomika je model "
                     "<b>jednoho kusu</b> pro rozhodování o nákupu, ne výsledek stávající flotily "
                     "(podrobnosti v části Metodika).", muted=True)

        # ---- Stroje ve flotile ----------------------------------------------------------------------
        st.markdown(f'<div class="dd-section-title">Počet kusů ve flotile<small>{fleet_total} '
                    f'{fleet_word(fleet_total, "kus", "kusy", "kusů")} · {fleet_types} '
                    f'{fleet_word(fleet_types, "typ", "typy", "typů")}</small></div>', unsafe_allow_html=True)
        with ui.dash_section("flotila"):
            ui.dash_units(list(machine_capacity.items()),
                          f"Celkem <b>{fleet_total}</b> {fleet_word(fleet_total, 'kus', 'kusy', 'kusů')} "
                          "ve flotile (ukázková data).")

        # ---- Ekonomika jednoho kusu -----------------------------------------------------------------
        st.markdown('<div class="dd-section-title">Ekonomika jednoho kusu podle dat<small>model na 1 kus '
                    'a průměrný měsíc · částky v Kč</small></div>', unsafe_allow_html=True)
        with ui.dash_section("ekonomika"):
            ui.dash_compare(
                UNIT_ECON_COLUMNS,
                [(t[0], t[1:5], f'<span class="neg">{t[5]}</span>' if r["profit"] < 0 else t[5])
                 for r, t in ((r, unit_economics_texts(r)) for r in fleet_econ)],
            )
            ui.dash_text("Ekonomika jednoho kusu používá individuální fixní náklady a servis strojů z dat – "
                         "stejně jako Investice → Nákup strojů.", muted=True)

        # ---- Graf provozního zisku (stejné hodnoty jako tabulka) ------------------------------------
        st.markdown('<div class="dd-section-title">Provozní zisk jednoho kusu podle typu<small>seřazeno '
                    'od nejvyššího · Kč za měsíc</small></div>', unsafe_allow_html=True)
        fig_fleet = go.Figure(
            go.Bar(
                x=[r["profit"] for r in fleet_by_profit],
                y=[r["machine"] for r in fleet_by_profit],
                orientation="h",
                marker=dict(color=[ui.BAD if r["profit"] < 0 else ui.BLUE for r in fleet_by_profit],
                            line=dict(color="#fff", width=2)),
                text=[fmt_czk(r["profit"]) for r in fleet_by_profit], textposition="outside", cliponaxis=False,
                customdata=[[fmt_czk(r["revenue"]), fmt_czk(r["costs"])] for r in fleet_by_profit],
                hovertemplate="%{y}<br>Provozní zisk 1 kusu: %{text}<br>Tržby %{customdata[0]} · "
                              "náklady %{customdata[1]}<extra></extra>",
            )
        )
        fleet_max = max(abs(r["profit"]) for r in fleet_econ) or 1
        fig_fleet.update_yaxes(autorange="reversed", title_text="")
        fig_fleet.update_xaxes(title_text="Kč za měsíc · jeden kus", showgrid=True, gridcolor="#EEF0F3",
                               range=[min(0, min(r["profit"] for r in fleet_econ)) * 1.6, fleet_max * 1.6],
                               nticks=5,
                               tickformat=",.0f")
        fig_fleet.update_layout(bargap=0.35, separators=", ")
        if fleet_loss:
            fig_fleet.add_vline(x=0, line=dict(color=ui.NAVY, width=1))
        with ui.dash_section("graf"):
            chart(ui.readable_axes(ui.style_fig(fig_fleet, height=max(320, 34 * len(fleet_econ) + 110),
                                                legend=False)))
            ui.dash_text("Graf ukazuje přesně hodnoty sloupce „Provozní zisk“ z tabulky výše – ekonomiku "
                         "<b>jednoho kusu</b> každého typu, ne celé flotily."
                         + (" Červeně = ztráta." if fleet_loss else ""), muted=True)

        # ---- Metodika a data ------------------------------------------------------------------------
        with st.expander("Metodika a data"):
            ui.explain([
                ("Zdroj", "Stejný výpočet jako Investice → Nákup strojů (modul portfolio, jeden kus): "
                          "průměrný měsíc podle ukázkových dat o strojích."),
                ("Tržby", "Ceníková denní cena × dny pronájmu. Dny = průměrná poptávka typu (sloupec Poptávka, "
                          "nejvýš počet dní v měsíci) bez dní ztracených poruchami (pravděpodobnost × průměrná "
                          "délka poruchy)."),
                ("Náklady", "Individuální měsíční fixní náklad stroje a průměrný servis z dat (ne fixní náklad "
                            "z nastavení investičního scénáře)."),
                ("Provozní zisk", "Tržby − náklady jednoho kusu za průměrný měsíc. Nejde o čistý zisk: "
                                  "společné náklady půjčovny a daně se neodečítají."),
                ("Proč ne souhrn za flotilu",
                 "Model popisuje očekávaný měsíc jednoho kusu pro rozhodování o nákupu. Součet za stávající "
                 "flotilu by předpokládal, že stejná poptávka platí pro každý kus (u více kusů stejného typu "
                 "se poptávka dělí) a nezohlednil by skutečné pronájmy, zakázky a servis z kapacitního "
                 "kalendáře. Celkový výsledek půjčovny proto zatím není k dispozici – stejně jako na Dashboardu."),
                ("Počty kusů", "Kusy z dat flotily – stejný údaj, se kterým počítá plán zakázek a kapacita."),
            ])
        ui.hint(cv.data_info(data_model, data_version))
        with ui.dash_links("strojovy_park"):
            page_link("kapacita", "Obsazenost strojů po dnech – Plánování kapacity")
            page_link("investice", "Nákup strojů – Investice")


# ==================================================
# NASTAVENÍ (etapa M2) – uživatel, správa účtů, data
# ==================================================

if PAGE == "nastaveni":
    with ui.tab_container("nastaveni"):
        ui.tab_header("Nastavení", "Přihlášený uživatel, správa účtů a informace o datech.")
        st.markdown("### Uživatel")
        st.markdown(
            f'Přihlášen: <b>{html.escape(st.session_state.get("user_name", ""))}</b>'
            f' · {html.escape(auth.ROLES.get(st.session_state.get("user_role", ""), ""))}',
            unsafe_allow_html=True,
        )
        st.button("Odhlásit se", key="logout_nastaveni", on_click=_log_out)
        st.markdown("### Správa účtů")
        if st.session_state.get("user_role") != "admin":
            ui.hint("Správu účtů vidí jen administrátor.")
        if st.session_state.get("user_role") == "admin":
            with st.expander("Správa účtů (jen administrátor)"):
                if AUTH["mode"] == "users":
                    for u in AUTH["users"].values():
                        st.markdown(f"- **{html.escape(u.name)}** (`{u.username}`) – {u.role_label}, "
                                    + ("aktivní" if u.active else "neaktivní"))
                    for problem in AUTH["problems"]:
                        st.warning(problem)
                else:
                    st.markdown("Aplikace používá jedno společné heslo (účty zatím nejsou nastavené).")
                st.markdown("**Vytvořit hash hesla pro Streamlit secrets**")
                pw1 = st.text_input("Nové heslo", type="password", key="admin_pw1")
                pw2 = st.text_input("Heslo znovu", type="password", key="admin_pw2")
                if pw1 or pw2:
                    problems = auth.password_problems(pw1)
                    if pw1 != pw2:
                        st.error("Hesla se neshodují.")
                    elif problems:
                        st.error("Heslo nevyhovuje: " + ", ".join(problems) + ".")
                    else:
                        st.code(auth.hash_password(pw1), language=None)
                        st.caption("Hash zkopírujte do password_hash v secrets. Heslo se nikam neukládá.")
        st.markdown("### Data")
        ui.hint(cv.data_info(data_model, data_version))
        ui.hint("Všechny údaje jsou fiktivní ukázková data (složka data/model). Aplikace je jen čte – nic se neukládá.")
        for issue in data_issues:
            st.markdown(f"- {html.escape(str(issue))}")


# Spuštění aktivní stránky (soubor stránky je jen značka – obsah vykreslil kód výše)
nav.run()
