"""Vzhled aplikace Decision Engine: barvy, formátování čísel, CSS a společné prvky UI.

Modul neobsahuje žádné výpočty – jen prezentaci výsledků.
"""

import base64
import html
from contextlib import contextmanager
from pathlib import Path

import streamlit as st

# ==================================================
# BARVY
# ==================================================

NAVY = "#14304F"        # prvky grafů (hranice, termíny)
ORANGE = "#D97706"      # Stroj 2 v grafech
BLUE = "#1F5FA8"        # Stroj 1
TEXT = "#1B2A3A"
MUTED = "#5B6675"
BORDER = "#DDE3EA"
SURFACE = "#F3F5F8"
GOOD = "#1E7B4A"
WARN = "#B7791F"
BAD = "#B42318"

# Barvy řad v grafech – ověřeno validátorem (kontrast i rozlišitelnost pro barvoslepé).
SERIES = [BLUE, ORANGE]

ASSETS = Path(__file__).parent / "assets"


# ==================================================
# ČESKÉ FORMÁTOVÁNÍ ČÍSEL
# ==================================================

# Funkce jsou v modulu formatting (bez Streamlitu, používá je i export DOCX); zde se znovu
# vystavují, aby zůstalo ui.fmt_czk i „from ui_theme import fmt_czk“.
from formatting import NBSP, fmt_czk, fmt_days, fmt_num, fmt_pct  # noqa: E402,F401


PRIORITY_CZ = {"High": "Vysoká", "Medium": "Střední", "Low": "Nízká"}
RISK_CZ = {"High": "Vysoké", "Medium": "Střední", "Low": "Nízké"}


# ==================================================
# VIZUÁLNÍ IDENTITA APLIKACE
# ==================================================
# Barvy rozhraní (lišta, nadpisy, akcent, popisky). Barvy grafů výše zůstávají beze změny.

INK = "#1F2933"          # hlavní text a nadpisy
BAR = "#1F2329"          # antracitová horní lišta
ACCENT = "#B45309"       # tlumená oranžová – akcent (kontrast 5,0:1 na bílé i s bílým textem)
ACCENT_SOFT = "#FDF4EA"  # jemné podbarvení doporučené varianty
UI_MUTED = "#4B5563"     # popisky a vysvětlivky (kontrast 7,6:1)
UI_LINE = "#E3E6EA"      # rámečky a oddělovače
UI_BG = "#F6F7F9"        # světlé plochy (sidebar, záhlaví tabulek)
OK_TEXT, OK_BG = "#166534", "#E8F5EC"
WARN_TEXT, WARN_BG = "#92400E", "#FEF3C7"
BAD_TEXT, BAD_BG = "#B42318", "#FDECEA"


def mark_src() -> str:
    """Vlastní neutrální symbol aplikace Decision Engine (graf + symbol rozhodnutí) jako data URI."""
    svg = (ASSETS / "decision_engine_mark.svg").read_bytes()
    return "data:image/svg+xml;base64," + base64.b64encode(svg).decode()


# ==================================================
# CSS
# ==================================================


# ==================================================
# SPOLEČNÝ STANDARD ZÁLOŽEK (UX/UI 2, etapa 0)
# ==================================================
# Pravidla platí JEN uvnitř obalu záložky vytvořeného funkcí tab_container() – tedy v prvku s třídou
# „st-key-de_tab_<název>“. Záložky bez obalu se nemění. Hodnoty převzaté z pilotu záložky 7 (PR #38).
# Používá jen vlastní třídy (de-…) a třídy st-key-*, žádné automaticky generované třídy Streamlitu.

TAB_KEY_PREFIX = "de_tab_"
_TAB = '[class*="st-key-de_tab_"]'
TAB_STANDARD_CSS = f"""
{_TAB} .de-tab-title{{font-size:1.25rem;font-weight:700;color:#1F2933;margin:.2rem 0 .35rem 0;line-height:1.3;padding:0}}
{_TAB} .de-tab-intro{{color:#1F2933;font-size:.97rem;line-height:1.5;margin:0 0 .6rem 0}}
{_TAB} h3{{font-size:1.05rem !important;font-weight:700;margin:1.4rem 0 .5rem 0 !important;padding:0 !important}}
{_TAB} .de-scope{{display:flex;flex-wrap:wrap;align-items:baseline;gap:.25rem .6rem;font-size:.875rem;color:#374151;
  background:#F6F7F9;border:1px solid #E3E6EA;border-radius:8px;padding:.45rem .75rem;margin:.1rem 0 .6rem 0}}
{_TAB} .de-scope b{{color:#1F2933}}
{_TAB} .de-scope .de-scope-hint{{color:#4B5563}}
{_TAB} .de-filter-label{{font-size:.8125rem;font-weight:700;letter-spacing:.05em;text-transform:uppercase;color:#4B5563;margin:0 0 .1rem 0}}
{_TAB} .de-kpis{{margin-bottom:.9rem}}
{_TAB} .de-kpi .lbl{{font-size:.8125rem}}
{_TAB} .de-kpi .sub{{font-size:.875rem}}
{_TAB} .de-status{{font-size:.8125rem}}
@media (min-width:641px) and (max-width:900px){{
  {_TAB} .de-kpis{{grid-template-columns:repeat(3,minmax(0,1fr));gap:.6rem}}
  {_TAB} .de-kpi{{padding:.75rem .85rem}}
  {_TAB} .de-kpi .val{{font-size:1.45rem;overflow-wrap:anywhere}}
}}
@media (max-width:640px){{
  {_TAB} .de-tab-title{{font-size:1.15rem}}
  {_TAB} .de-kpis{{gap:.5rem}}
  {_TAB} .de-kpi{{padding:.65rem .85rem}}
  {_TAB} .de-kpi .val{{font-size:1.45rem;margin:.2rem 0}}
}}
"""


def inject_css() -> None:
    st.markdown(
        f"""
        <style>
        /* 3,5rem: obsah začíná pod systémovou lištou Streamlitu (60 px), která je při sbaleném panelu bílá */
        .block-container {{ padding-top: 3.5rem; padding-bottom: 3rem; max-width: 1200px; }}
        h1, h2, h3, h4 {{ color: {INK}; }}
        h2 {{ font-size: 1.4rem !important; }}
        h3 {{ font-size: 1.12rem !important; }}
        [data-testid="stCaptionContainer"], [data-testid="stCaptionContainer"] p {{ color: {UI_MUTED} !important; }}
        [data-testid="stWidgetLabel"] p {{ color: {INK}; font-weight: 500; }}

        /* Postranní panel */
        [data-testid="stSidebar"] {{ background: {UI_BG}; border-right: 1px solid {UI_LINE}; }}
        [data-testid="stSidebar"] h3 {{
            font-size: .74rem !important; text-transform: uppercase; letter-spacing: .06em;
            color: {UI_MUTED} !important; border-top: 1px solid {UI_LINE};
            padding-top: .9rem !important; margin-top: .55rem !important;
        }}
        .de-side-brand {{ display: flex; align-items: center; gap: .55rem; margin: -.4rem 0 .7rem 0; }}
        .de-side-brand img {{ width: 30px; height: 30px; }}
        .de-side-brand b {{ color: {INK}; font-size: .98rem; }}
        .de-side-brand span {{ display: block; color: {UI_MUTED}; font-size: .74rem; }}
        .de-demo {{
            background: #fff; border: 1px solid {UI_LINE}; border-left: 3px solid {ACCENT};
            border-radius: 8px; padding: .35rem .6rem; font-size: .74rem; color: {UI_MUTED}; margin-bottom: .2rem;
        }}
        .de-demo b {{ color: {INK}; }}
        /* Čitelnost: Streamlit dává popiskům průhlednost 0,6 a krajním hodnotám posuvníku světlou barvu */
        [data-testid="stSidebar"] [data-testid="stCaptionContainer"] {{ opacity: 1; color: {UI_MUTED}; }}
        [data-testid="stSidebar"] [data-testid="stSliderTickBar"],
        [data-testid="stSidebar"] [data-testid="stSliderTickBar"] p {{ color: {UI_MUTED} !important; }}
        /* Panel „Nastavení investičního scénáře“ na stránce Investice (stejné nadpisy skupin jako dřív v panelu) */
        .st-key-de_inv_inputs h3 {{
            font-size: .74rem !important; text-transform: uppercase; letter-spacing: .06em;
            color: {UI_MUTED} !important; padding-top: .3rem !important; margin-top: .2rem !important;
        }}
        .st-key-de_inv_inputs [data-testid="stCaptionContainer"] {{ opacity: 1; color: {UI_MUTED}; }}
        .st-key-de_inv_inputs [data-testid="stSliderTickBar"],
        .st-key-de_inv_inputs [data-testid="stSliderTickBar"] p {{ color: {UI_MUTED} !important; }}
        /* Investice: obsah neaktivních záložek je skrytý (lišta záložek je nad panelem nastavení) */
        [class*="st-key-de_inv_obsah_"][class*="_off"] {{ display: none; }}
        .de-inv-summary {{ font-size: .85rem; color: {UI_MUTED}; margin: -.4rem 0 .4rem .1rem; }}
        .de-inv-summary b {{ color: {INK}; }}
        .de-amount {{ font-size: .85rem; color: {UI_MUTED}; margin: -.55rem 0 .3rem .1rem; }}
        .de-amount b {{ color: {INK}; font-variant-numeric: tabular-nums; }}
        .de-soon {{
            font-size: .82rem; color: {UI_MUTED}; border-top: 1px solid {UI_LINE};
            padding-top: .8rem; margin-top: .5rem; margin-bottom: .8rem;
        }}
        .de-user {{
            font-size: .85rem; color: {UI_MUTED}; border-top: 1px solid {UI_LINE};
            padding-top: .8rem; margin-top: .6rem; margin-bottom: .4rem;
        }}
        .de-user b {{ color: {INK}; }}

        /* Záložky */
        .stTabs [data-baseweb="tab-list"] {{ gap: .15rem; border-bottom: 1px solid {UI_LINE}; overflow-x: auto; }}
        .stTabs [data-baseweb="tab"] {{ padding: .55rem .85rem; font-weight: 600; color: {UI_MUTED}; }}
        .stTabs [aria-selected="true"] {{ color: {INK}; }}
        .stTabs [data-baseweb="tab-highlight"] {{ background-color: {ACCENT}; height: 3px; }}

        [data-testid="stMetric"] {{
            background: #fff; border: 1px solid {UI_LINE}; border-radius: 10px; padding: .75rem 1rem;
        }}
        [data-testid="stMetricLabel"] p {{ white-space: normal; color: {UI_MUTED}; font-size: .85rem; }}
        [data-testid="stMetricValue"] {{ color: {INK}; font-size: 1.45rem; }}
        [data-testid="stExpander"] details {{ border-radius: 10px; border-color: {UI_LINE}; background: #fff; }}
        [data-testid="stExpander"] summary p {{ color: {INK}; font-weight: 600; }}

        /* Horní lišta */
        .de-header {{
            display: flex; align-items: center; gap: .85rem; background: {BAR}; color: #fff;
            border-radius: 12px; padding: .65rem 1.1rem; margin-bottom: 1rem;
        }}
        .de-header img {{ width: 40px; height: 40px; flex: 0 0 auto; }}
        .de-header .de-names {{ min-width: 0; flex: 1 1 auto; }}
        .de-title {{ font-size: 1.25rem; font-weight: 700; line-height: 1.2; color: #fff; }}
        .de-sub {{ font-size: .86rem; color: #D1D5DB; line-height: 1.3; }}
        .de-badge {{
            flex: 0 0 auto; border: 1px solid #6B7280; color: #F3F4F6; font-weight: 600; font-size: .72rem;
            letter-spacing: .04em; padding: .22rem .6rem; border-radius: 999px; white-space: nowrap;
        }}

        /* Přihlášení */
        .de-login {{ text-align: center; margin: 3rem 0 1.2rem 0; }}
        .de-login img {{ width: 64px; height: 64px; }}
        .de-login .t {{ font-size: 1.7rem; font-weight: 800; color: {INK}; margin-top: .6rem; }}
        .de-login .s {{ color: {UI_MUTED}; font-size: 1rem; }}
        .de-login .p {{ color: {INK}; font-size: .95rem; margin-top: 1.4rem; }}

        /* Karty používané v záložkách */
        .de-card {{
            background: #fff; border: 1px solid {UI_LINE}; border-radius: 12px; padding: 1rem 1.15rem;
            height: 100%; margin-bottom: .6rem;
        }}
        .de-card.rec {{ border: 2px solid {ACCENT}; }}
        .de-card-title {{
            margin: 0 0 .5rem 0; font-size: 1.08rem; font-weight: 700; color: {INK};
            display: flex; align-items: center; gap: .55rem;
        }}
        .de-dot {{ width: .75rem; height: .75rem; border-radius: 50%; display: inline-block; }}
        .de-tag {{
            font-size: .7rem; font-weight: 700; color: #fff; background: {ACCENT};
            padding: .15rem .5rem; border-radius: 999px; margin-left: auto;
        }}
        .de-row {{
            display: flex; justify-content: space-between; gap: 1rem; padding: .4rem 0;
            border-bottom: 1px solid {UI_LINE}; font-size: .95rem;
        }}
        .de-row:last-child {{ border-bottom: none; }}
        .de-row span:first-child {{ color: {UI_MUTED}; }}
        .de-row span:last-child {{ font-weight: 600; color: {INK}; text-align: right; }}
        .de-note {{
            background: #fff; border: 1px solid {UI_LINE}; border-left: 4px solid {INK};
            border-radius: 8px; padding: .6rem .8rem; font-size: .88rem; color: {UI_MUTED};
        }}

        /* Záložka Doporučení */
        .de-hero {{
            background: #fff; border: 1px solid {UI_LINE}; border-left: 6px solid {ACCENT};
            border-radius: 12px; padding: 1rem 1.4rem; margin: .2rem 0 1rem 0;
        }}
        .de-hero.bad {{ border-left-color: {BAD_TEXT}; }}
        .de-hero .lbl {{ font-size: .76rem; font-weight: 700; letter-spacing: .08em; text-transform: uppercase; color: {ACCENT}; }}
        .de-hero.bad .lbl {{ color: {BAD_TEXT}; }}
        .de-hero .ttl {{ font-size: 1.9rem; font-weight: 800; color: {INK}; line-height: 1.15; margin: .15rem 0 .3rem 0; }}
        .de-hero .txt {{ font-size: 1.05rem; color: {INK}; }}

        .de-kpis {{ display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: .9rem; margin: 0 0 1.4rem 0; }}
        .de-kpi {{ background: #fff; border: 1px solid {UI_LINE}; border-radius: 12px; padding: .9rem 1.15rem; }}
        .de-kpi .lbl {{ font-size: .76rem; font-weight: 700; letter-spacing: .05em; text-transform: uppercase; color: {UI_MUTED}; }}
        .de-kpi .val {{ font-size: 1.95rem; font-weight: 800; color: {INK}; line-height: 1.15; margin: .35rem 0 .3rem 0;
                        display: flex; align-items: center; gap: .5rem; flex-wrap: wrap; }}
        .de-kpi .val.neg {{ color: {BAD_TEXT}; }}
        .de-kpi .sub {{ font-size: .88rem; color: {UI_MUTED}; }}
        .de-status {{
            display: inline-block; font-size: .74rem; font-weight: 700; padding: .14rem .55rem;
            border-radius: 999px; letter-spacing: .01em;
        }}
        .de-status.good {{ color: {OK_TEXT}; background: {OK_BG}; }}
        .de-status.warn {{ color: {WARN_TEXT}; background: {WARN_BG}; }}
        .de-status.bad {{ color: {BAD_TEXT}; background: {BAD_BG}; }}

        .de-section {{ font-size: 1.02rem; font-weight: 700; color: {INK}; margin: .2rem 0 .55rem 0; }}
        .de-cmp {{ width: 100%; border-collapse: separate; border-spacing: 0; border: 1px solid {UI_LINE};
                   border-radius: 12px; overflow: hidden; margin-bottom: .4rem; background: #fff; }}
        .de-cmp th, .de-cmp td {{ padding: .65rem 1rem; border-bottom: 1px solid {UI_LINE}; text-align: right; }}
        .de-cmp tr:last-child td {{ border-bottom: none; }}
        .de-cmp th {{ font-size: .88rem; color: {INK}; background: {UI_BG}; font-weight: 700; }}
        .de-cmp th:first-child, .de-cmp td:first-child {{ text-align: left; color: {UI_MUTED}; font-weight: 500; }}
        .de-cmp td {{ font-size: .98rem; color: {INK}; font-variant-numeric: tabular-nums; }}
        .de-cmp .rec {{ background: {ACCENT_SOFT}; }}
        .de-cmp th.rec {{ background: #FAE6D2; }}
        .de-cmp .best {{ font-weight: 700; color: {OK_TEXT}; }}
        .de-cmp .best::after {{ content: " ✓"; }}
        .de-cmp .tag {{
            display: inline-block; margin-left: .4rem; font-size: .68rem; font-weight: 700; color: #fff;
            background: {ACCENT}; padding: .1rem .45rem; border-radius: 999px; vertical-align: middle;
        }}
        .de-cmp-note {{ font-size: .8rem; color: {UI_MUTED}; margin: 0 0 1.4rem 0; }}
        .de-steps {{ margin: 0 0 1.2rem 0; padding-left: 1.3rem; }}
        .de-steps li {{ margin: .25rem 0; color: {INK}; }}

        /* Záložka Možné výsledky */
        .de-res-grid {{ display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 1rem; margin: .4rem 0 1.4rem 0; }}
        .de-res-card {{ background: #fff; border: 1px solid {UI_LINE}; border-radius: 12px; padding: .9rem 1.15rem .6rem 1.15rem; }}
        .de-res-card.rec {{ border: 2px solid {ACCENT}; }}
        .de-res-row {{
            display: grid; grid-template-columns: minmax(0, 1fr) auto; gap: .2rem 1rem; align-items: baseline;
            padding: .55rem 0; border-top: 1px solid {UI_LINE};
        }}
        .de-res-row .l {{ color: {INK}; font-size: .93rem; }}
        .de-res-row .h {{ display: block; color: {UI_MUTED}; font-size: .78rem; margin-top: .05rem; }}
        .de-res-row .v {{ color: {INK}; font-size: 1.02rem; font-weight: 700; text-align: right;
                          font-variant-numeric: tabular-nums; white-space: nowrap; }}
        .de-res-row .v.neg {{ color: {BAD_TEXT}; }}
        .de-res-row .v .de-pos {{ color: {OK_TEXT}; }}
        .de-res-row .s {{ display: block; color: {UI_MUTED}; font-size: .78rem; font-weight: 500; }}
        .de-intro {{ color: {INK}; font-size: 1rem; margin: -.2rem 0 .2rem 0; }}
        .de-hint {{ color: #374151; font-size: .94rem; line-height: 1.5; margin: -.3rem 0 1rem 0; }}
        .de-insight {{
            border-left: 3px solid {ACCENT}; padding: .15rem 0 .15rem .85rem;
            color: {INK}; font-size: 1rem; margin: 0 0 1.4rem 0;
        }}
        .de-insight b {{ color: {INK}; }}
        .de-insight.bad {{ border-left-color: {BAD_TEXT}; }}
        .de-explain {{ margin: .2rem 0 .3rem 0; }}
        .de-explain .it {{ padding: .55rem 0; border-top: 1px solid {UI_LINE}; }}
        .de-explain .it:first-child {{ border-top: none; padding-top: .1rem; }}
        .de-explain .t {{ color: {INK}; font-weight: 700; font-size: .95rem; }}
        .de-explain .d {{ color: #374151; font-size: .94rem; line-height: 1.5; margin-top: .1rem; }}
        .de-ci-grid {{ display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: .7rem; margin: .4rem 0 .4rem 0; }}
        .de-ci {{ background: {UI_BG}; border: 1px solid {UI_LINE}; border-radius: 8px; padding: .55rem .8rem; }}
        .de-ci .n {{ color: {INK}; font-size: .88rem; font-weight: 600; display: flex; align-items: center; gap: .45rem; }}
        .de-ci .l {{ color: {UI_MUTED}; font-size: .8rem; }}
        .de-ci .v {{ color: {INK}; font-size: .95rem; font-weight: 600; font-variant-numeric: tabular-nums; }}

        @media (max-width: 768px) {{
            .de-kpis {{ grid-template-columns: 1fr; gap: .6rem; }}
            .de-res-grid, .de-ci-grid {{ grid-template-columns: 1fr; gap: .7rem; }}
            .de-hero {{ padding: .9rem 1rem; }}
            .de-hero .ttl {{ font-size: 1.5rem; }}
            .de-kpi .val {{ font-size: 1.6rem; }}
            .de-cmp th, .de-cmp td {{ padding: .5rem .45rem; font-size: .85rem; }}
        }}
        @media (max-width: 640px) {{
            .block-container {{ padding-left: 1rem; padding-right: 1rem; }}
            .de-header {{ padding: .55rem .8rem; gap: .6rem; }}
            .de-header img {{ width: 32px; height: 32px; }}
            .de-title {{ font-size: 1.08rem; }}
            .de-sub {{ font-size: .78rem; }}
            .de-badge {{ font-size: .64rem; padding: .18rem .45rem; }}
            [data-testid="stMetricValue"] {{ font-size: 1.25rem; }}
        }}
        @media (max-width: 420px) {{
            .de-sub {{ display: none; }}
        }}
        {TAB_STANDARD_CSS}
        </style>
        """,
        unsafe_allow_html=True,
    )


# ==================================================
# KOMPONENTY
# ==================================================


def header(title: str, subtitle: str, badge: str) -> None:
    """Výškově úsporná horní lišta: symbol, název, podtitul, štítek. Text se zalamuje, neořezává."""
    st.markdown(
        f"""
        <div class="de-header">
            <img src="{mark_src()}" alt="">
            <div class="de-names">
                <div class="de-title">{html.escape(title)}</div>
                <div class="de-sub">{html.escape(subtitle)}</div>
            </div>
            <div class="de-badge">{html.escape(badge)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def login_intro(title: str, subtitle: str, prompt: str) -> None:
    st.markdown(
        f'<div class="de-login"><img src="{mark_src()}" alt=""><div class="t">{html.escape(title)}</div>'
        f'<div class="s">{html.escape(subtitle)}</div><div class="p">{html.escape(prompt)}</div></div>',
        unsafe_allow_html=True,
    )


def sidebar_amount(container, amount: str, note: str) -> None:
    """Zadaná částka pod číselným polem v českém formátu (např. „100 000 Kč měsíčně“).

    Nativní st.number_input neumí oddělovat tisíce, proto se hodnota zobrazí ještě takto.
    """
    container.markdown(
        f'<div class="de-amount"><b>{html.escape(amount)}</b> {html.escape(note)}</div>',
        unsafe_allow_html=True,
    )


def sidebar_brand(title: str, subtitle: str) -> None:
    st.sidebar.markdown(
        f'<div class="de-side-brand"><img src="{mark_src()}" alt=""><div><b>{html.escape(title)}</b>'
        f"<span>{html.escape(subtitle)}</span></div></div>",
        unsafe_allow_html=True,
    )


def machine_card(name: str, color: str, rows: list, recommended: bool = False, tag: str = "Doporučeno") -> str:
    """HTML karta stroje: rows = [(popisek, hodnota), ...]."""
    tag_html = f'<span class="de-tag">{html.escape(tag)}</span>' if recommended else ""
    rows_html = "".join(
        f"<div class='de-row'><span>{html.escape(label)}</span><span>{value}</span></div>"
        for label, value in rows
    )
    return (
        f'<div class="de-card {"rec" if recommended else ""}">'
        f'<div class="de-card-title"><span class="de-dot" style="background:{color}"></span>{html.escape(name)}{tag_html}</div>'
        f"{rows_html}</div>"
    )


def note(text: str) -> None:
    st.markdown(f'<div class="de-note">{text}</div>', unsafe_allow_html=True)


def hero(label: str, title: str, text: str, tone: str = "") -> None:
    """Dominantní karta s doporučením (záložka Doporučení)."""
    st.markdown(
        f'<div class="de-hero {tone}"><div class="lbl">{html.escape(label)}</div>'
        f'<div class="ttl">{html.escape(title)}</div><div class="txt">{text}</div></div>',
        unsafe_allow_html=True,
    )


def status_chip(text: str, tone: str) -> str:
    """Barevný štítek stavu – barva vždy doplněná slovním popisem (good / warn / bad)."""
    return f'<span class="de-status {tone}">{html.escape(text)}</span>'


# --- Společné komponenty standardu záložek (UX/UI 2, etapa 0) -----------------------------------
# Jen zobrazení: nic nepočítají, nečtou ani nezapisují stav widgetů. Klíče obalů mají vlastní
# předpony (de_tab_, de_summary_, de_filters_), aby se nemohly potkat s klíči widgetů.

STATUS_SYMBOL = {"good": "✓", "bad": "✕", "warn": "?"}


def status_text(text: str, tone: str) -> str:
    """Text štítku stavu se symbolem, aby význam nenesla jen barva (✓ / ✕ / ?)."""
    sym = STATUS_SYMBOL.get(tone, "")
    return f"{sym} {text}" if sym else text


def tab_container(slug: str):
    """Obal záložky – uvnitř platí společný standard (TAB_STANDARD_CSS). Použití: with tab, tab_container("…"):"""
    return st.container(key=TAB_KEY_PREFIX + slug)


def tab_header(title: str, intro_html: str = "") -> None:
    """Nadpis záložky a krátký úvod (nejvýš 1–2 věty). intro_html je pevný text aplikace (smí obsahovat <b>)."""
    st.markdown(f'<div class="de-tab-title" role="heading" aria-level="2">{html.escape(title)}</div>',
                unsafe_allow_html=True)
    if intro_html:
        st.markdown(f'<div class="de-tab-intro">{intro_html}</div>', unsafe_allow_html=True)


def summary_slot(slug: str):
    """Místo pro souhrnné ukazatele NAD ovládacími prvky. Plní se až PO přečtení filtrů, takže vždy
    odpovídá aktuálnímu výběru (vzor z pilotu záložky 7)."""
    return st.container(key=f"de_summary_{slug}")


@contextmanager
def filter_box(slug: str, label: str):
    """Rámeček s filtry a přepínači záložky. Widgety uvnitř si ponechávají své klíče i výchozí hodnoty."""
    with st.container(border=True, key=f"de_filters_{slug}"):
        st.markdown(f'<div class="de-filter-label">{html.escape(label)}</div>', unsafe_allow_html=True)
        yield


def scope_html(scope: str, subject: str, hint: str = "mění se podle filtrů níže") -> str:
    """Řádek „Platí pro“ nad souhrnem: pro jaké období a co platí ukazatele (texty se escapují)."""
    return (f'<div class="de-scope"><span>Platí pro: <b>{html.escape(scope)}</b> · '
            f'<b>{html.escape(subject)}</b></span><span class="de-scope-hint">{html.escape(hint)}</span></div>')


def kpi_cards(cards: list) -> None:
    """Řada velkých KPI karet. cards = [dict(label, value, sub, negative, chip)], na mobilu pod sebou."""
    items = "".join(
        f'<div class="de-kpi"><div class="lbl">{html.escape(c["label"])}</div>'
        f'<div class="val {"neg" if c.get("negative") else ""}">{c["value"]}{c.get("chip", "")}</div>'
        f'<div class="sub">{c.get("sub", "")}</div></div>'
        for c in cards
    )
    st.markdown(f'<div class="de-kpis">{items}</div>', unsafe_allow_html=True)


def compare_table(names: list, rows: list, rec_index: int, tag: str) -> None:
    """Srovnávací tabulka dvou strojů.

    rows = [(popisek, [hodnota_1, hodnota_2], index_lepší_hodnoty nebo None)].
    Sloupec doporučeného stroje je jemně zvýrazněn, lepší hodnota v řádku označena ✓.
    """
    head = "".join(
        f'<th class="{"rec" if i == rec_index else ""}">{html.escape(n)}'
        + (f'<span class="tag">{html.escape(tag)}</span>' if i == rec_index else "")
        + "</th>"
        for i, n in enumerate(names)
    )
    body = ""
    for label, values, best in rows:
        cells = "".join(
            f'<td class="{"rec" if i == rec_index else ""} {"best" if best == i else ""}">{v}</td>'
            for i, v in enumerate(values)
        )
        body += f"<tr><td>{html.escape(label)}</td>{cells}</tr>"
    st.markdown(
        f'<table class="de-cmp"><thead><tr><th>Ukazatel</th>{head}</tr></thead>'
        f"<tbody>{body}</tbody></table>",
        unsafe_allow_html=True,
    )


def results_cards(cards: list, rec_name: str, tag: str) -> None:
    """Dvě karty výsledků vedle sebe (na mobilu pod sebou).

    cards = [dict(name, color, rows=[(popisek, nápověda, hodnota, doplněk, záporná)])].
    """
    out = ""
    for c in cards:
        rec = c["name"] == rec_name
        tag_html = f'<span class="de-tag">{html.escape(tag)}</span>' if rec else ""
        rows = "".join(
            f'<div class="de-res-row"><div class="l">{html.escape(label)}'
            + (f'<span class="h">{html.escape(hint)}</span>' if hint else "")
            + f'</div><div class="v {"neg" if neg else ""}">{value}'
            + (f'<span class="s">{html.escape(sub)}</span>' if sub else "")
            + "</div></div>"
            for label, hint, value, sub, neg in c["rows"]
        )
        out += (
            f'<div class="de-res-card {"rec" if rec else ""}"><div class="de-card-title">'
            f'<span class="de-dot" style="background:{c["color"]}"></span>{html.escape(c["name"])}{tag_html}</div>'
            f"{rows}</div>"
        )
    st.markdown(f'<div class="de-res-grid">{out}</div>', unsafe_allow_html=True)


def hint(text: str) -> None:
    """Nenápadné, ale čitelné vysvětlení pod grafem."""
    st.markdown(f'<div class="de-hint">{text}</div>', unsafe_allow_html=True)


def insight(text: str, tone: str = "") -> None:
    """Krátký zvýrazněný závěr (bez velké KPI karty). tone="bad" = červená linka pro varování."""
    cls = f"de-insight {tone}" if tone else "de-insight"
    st.markdown(f'<div class="{cls}">{text}</div>', unsafe_allow_html=True)


def explain(items: list) -> None:
    """Čitelný seznam vysvětlivek: items = [(název, popis v HTML)]."""
    rows = "".join(
        f'<div class="it"><div class="t">{html.escape(t)}</div><div class="d">{d}</div></div>'
        for t, d in items
    )
    st.markdown(f'<div class="de-explain">{rows}</div>', unsafe_allow_html=True)


def interval_boxes(items: list, label: str) -> None:
    """Méně výrazné boxy s intervalem: items = [(název, barva, text hodnoty)]."""
    boxes = "".join(
        f'<div class="de-ci"><div class="n"><span class="de-dot" style="background:{color}"></span>'
        f'{html.escape(name)}</div><div class="l">{html.escape(label)}</div><div class="v">{value}</div></div>'
        for name, color, value in items
    )
    st.markdown(f'<div class="de-ci-grid">{boxes}</div>', unsafe_allow_html=True)


def readable_axes(fig):
    """Tmavší popisky os a legendy pro lepší čitelnost (doplňuje style_fig, nemění data)."""
    axis = dict(
        tickfont=dict(color="#374151", size=12),
        title_font=dict(color=INK, size=13),
        linecolor="#C9CED6",
    )
    fig.update_xaxes(**axis)
    fig.update_yaxes(**axis)
    fig.update_layout(legend=dict(font=dict(color=INK, size=13)))
    return fig


# ==================================================
# GRAFY (Plotly)
# ==================================================


def style_fig(fig, height: int = 380, legend: bool = True):
    """Jednotný vzhled grafů: písmo, barvy os, česká čísla, legenda nahoře."""
    fig.update_layout(
        height=height,
        template="plotly_white",
        separators=", ",
        font=dict(family="Source Sans Pro, sans-serif", size=13, color=TEXT),
        margin=dict(l=10, r=10, t=36, b=10),
        paper_bgcolor="#fff",
        plot_bgcolor="#fff",
        hoverlabel=dict(bgcolor="#fff", font_size=13, bordercolor=BORDER),
        showlegend=legend,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0, title_text=""),
    )
    fig.update_xaxes(gridcolor="#EEF1F5", linecolor=BORDER, zerolinecolor=BORDER)
    fig.update_yaxes(gridcolor="#EEF1F5", linecolor=BORDER, zerolinecolor=BORDER)
    return fig


# ---------------------------------------------------------------------------
# Dashboard (etapa M3 – první fáze): KPI karty, sekce a upozornění.
# Jen zobrazení – hodnoty dodává app.py z existujících výpočtů. Styl platí jen uvnitř Dashboardu.
# ---------------------------------------------------------------------------

DASH_KEY = TAB_KEY_PREFIX + "dashboard"
_DASH = f".st-key-{DASH_KEY}"

# Jednoduché obrysové ikony (24×24, barva podle textu) – bez externích zdrojů
DASH_ICONS = {
    "gauge": '<path d="M4 15a8 8 0 1 1 16 0"/><path d="M12 15l4-5"/><circle cx="12" cy="15" r="1.4"/>',
    "coins": '<ellipse cx="12" cy="6.5" rx="7" ry="2.8"/><path d="M5 6.5v5c0 1.6 3.1 2.8 7 2.8s7-1.2 7-2.8v-5"/>'
             '<path d="M5 11.5v5c0 1.6 3.1 2.8 7 2.8s7-1.2 7-2.8v-5"/>',
    "truck": '<path d="M3 7h11v9H3z"/><path d="M14 10h4l3 3v3h-7"/><circle cx="7" cy="17.5" r="1.8"/>'
             '<circle cx="17" cy="17.5" r="1.8"/>',
    "alert": '<path d="M12 4l9 16H3z"/><path d="M12 10v4.5"/><circle cx="12" cy="17.2" r=".6"/>',
    "spark": '<path d="M12 3v4M12 17v4M3 12h4M17 12h4"/><circle cx="12" cy="12" r="3"/>',
    "route": '<circle cx="6" cy="18" r="2"/><circle cx="18" cy="6" r="2"/><path d="M8 18h6a4 4 0 0 0 0-8H10a4 4 0 0 1 0-8h6"/>',
}


def dash_icon(name: str) -> str:
    return (f'<svg class="dd-ico" viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" '
            f'stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">{DASH_ICONS[name]}</svg>')


def dash_chip(text: str, tone: str) -> str:
    """Stavový štítek: barva vždy se symbolem a textem (význam nenese jen barva)."""
    return f'<span class="dd-chip {tone}">{html.escape(status_text(text, tone))}</span>'


def dash_kpis(cards: list) -> None:
    """Hlavní KPI karty. cards = [dict(icon, label, value, sub, chip, tone)] – value/sub/chip jsou hotové HTML.

    tone (barva ikony): "accent" (výchozí), "neutral", "good", "bad" – barva jen tam, kde nese význam.
    """
    items = "".join(
        f'<div class="dd-kpi"><div class="dd-kpi-top"><span class="dd-ico-box {c.get("tone", "accent")}">'
        f'{dash_icon(c["icon"])}</span><span class="dd-lbl">{html.escape(c["label"])}</span></div>'
        f'<div class="dd-val{" pair" if c.get("pair") else ""}{" " + c["cls"] if c.get("cls") else ""}">'
        f'{c["value"]}</div>{c.get("chip", "")}'
        f'<div class="dd-sub">{c.get("sub", "")}</div></div>'
        for c in cards
    )
    st.markdown(f'<div class="dd-kpis">{items}</div>', unsafe_allow_html=True)


def dash_section(slug: str):
    """Ohraničená sekce Dashboardu (uvnitř se smějí použít i odkazy st.page_link).

    Přípona klíče _good/_bad/_accent/_neutral určuje barvu proužku u karet upozornění (viz CSS).
    """
    return st.container(border=True, key=f"de_dash_{slug}")


def dash_links(slug: str):
    """Řádek odkazů do sekcí (zalamuje se; na mobilu pod sebou podle šířky)."""
    return st.container(key=f"dd_links_{slug}")


def dash_section_head(icon: str, title: str, caption: str = "", tone: str = "accent") -> None:
    st.markdown(
        f'<div class="dd-head"><span class="dd-ico-box {tone}">{dash_icon(icon)}</span><div>'
        f'<div class="dd-title" role="heading" aria-level="3">{html.escape(title)}</div>'
        + (f'<div class="dd-cap">{caption}</div>' if caption else "")
        + "</div></div>",
        unsafe_allow_html=True,
    )


def dash_rows(rows: list) -> None:
    """Řádky „popisek – hodnota“: rows = [(popisek, hodnota HTML, poznámka HTML nebo ""[, zvýraznit])]."""
    body = "".join(
        f'<div class="dd-row{" key" if len(r) > 3 and r[3] else ""}"><span class="l">{html.escape(r[0])}'
        + (f'<small>{r[2]}</small>' if r[2] else "")
        + f'</span><span class="v">{r[1]}</span></div>'
        for r in rows
    )
    st.markdown(f'<div class="dd-rows">{body}</div>', unsafe_allow_html=True)


def dash_kv(label: str, value_html: str, sub_html: str = "") -> None:
    """Popisek nad výraznou hodnotou (karty doporučení)."""
    st.markdown(
        f'<div class="dd-kv"><div class="k">{html.escape(label)}</div><div class="val">{value_html}</div>'
        + (f'<div class="s">{sub_html}</div>' if sub_html else "") + "</div>",
        unsafe_allow_html=True,
    )


def dash_text(text_html: str, muted: bool = False) -> None:
    st.markdown(f'<div class="dd-text{" muted" if muted else ""}">{text_html}</div>', unsafe_allow_html=True)


_TONES = {"accent": (ACCENT, ACCENT_SOFT), "good": ("#166534", "#E8F5EC"), "bad": (BAD_TEXT, BAD_BG),
          "neutral": ("#334155", "#EEF2F7")}
def _tone_css(scope: str) -> str:
    return "\n".join(
        f"{scope} .dd-ico-box.{name} {{ color: {fg}; background: {bg}; }}\n"
        f'{scope} [class*="st-key-de_dash_"][class*="_{name}"] {{ border-left: 4px solid {fg} !important; }}'
        for name, (fg, bg) in _TONES.items()
    )


def dash_css(key: str) -> str:
    """Styl karet Dashboardu (dd-*) omezený na kontejner s klíčem `key` – používá Dashboard i Pronájem strojů."""
    _DASH = f".st-key-{key}"
    _TONE_CSS = _tone_css(_DASH)
    return f"""
<style>
{_DASH} .de-tab-title {{ font-size: 1.4rem; margin-bottom: .1rem; }}
{_DASH} .de-tab-intro {{ color: {UI_MUTED}; font-size: .92rem; margin-bottom: .9rem; }}
/* KPI karty */
{_DASH} .dd-kpis {{ display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 1rem; margin: .1rem 0 1.4rem 0; }}
{_DASH} .dd-kpi {{ background: #fff; border: 1px solid {UI_LINE}; border-radius: 14px; padding: 1rem 1.15rem;
                 box-shadow: 0 1px 3px rgba(16,24,40,.06); display: flex; flex-direction: column; gap: .45rem; min-width: 0; }}
{_DASH} .dd-kpi-top {{ display: flex; align-items: center; gap: .6rem; }}
{_DASH} .dd-ico-box {{ display: inline-flex; align-items: center; justify-content: center; flex: 0 0 auto;
                     width: 2.1rem; height: 2.1rem; border-radius: 10px; }}
{_DASH} .dd-ico {{ width: 1.2rem; height: 1.2rem; }}
{_DASH} .dd-lbl {{ font-size: .76rem; font-weight: 700; letter-spacing: .05em; text-transform: uppercase;
                 color: {UI_MUTED}; line-height: 1.25; }}
{_DASH} .dd-val {{ font-size: 2.15rem; font-weight: 800; color: {INK}; line-height: 1.1; letter-spacing: -.01em;
                 font-variant-numeric: tabular-nums; overflow-wrap: anywhere; margin-top: .15rem; }}
{_DASH} .dd-val.pair {{ font-size: 1.75rem; display: flex; flex-wrap: wrap; column-gap: .45rem; row-gap: .1rem; }}
{_DASH} .dd-val.pair > span {{ white-space: nowrap; }}
{_DASH} .dd-val small {{ font-size: 1rem; font-weight: 600; color: {UI_MUTED}; margin-left: .2rem; letter-spacing: 0; }}
{_DASH} .dd-sub {{ font-size: .8rem; color: {UI_MUTED}; line-height: 1.45; margin-top: auto; padding-top: .55rem;
                 border-top: 1px solid {UI_LINE}; }}
{_DASH} .dd-chip {{ align-self: flex-start; font-size: .74rem; font-weight: 700; padding: .16rem .6rem; border-radius: 999px;
                  white-space: nowrap; }}
{_DASH} .dd-chip.good {{ color: #166534; background: #E8F5EC; }}
{_DASH} .dd-chip.warn {{ color: {WARN_TEXT}; background: {WARN_BG}; }}
{_DASH} .dd-chip.bad {{ color: {BAD_TEXT}; background: {BAD_BG}; }}
{_DASH} .dd-chip.info {{ color: #334155; background: #EEF2F7; }}
/* Sekce a karty */
{_DASH} [class*="st-key-de_dash_"] {{ border-radius: 14px !important; border-color: {UI_LINE} !important; background: #fff;
                                    box-shadow: 0 1px 3px rgba(16,24,40,.06); padding: 1.1rem 1.2rem !important; }}
{_DASH} [data-testid="stHorizontalBlock"] {{ align-items: stretch; }}
{_DASH} [data-testid="stColumn"] > [data-testid="stVerticalBlock"] {{ height: 100%; }}
{_DASH} [data-testid="stColumn"] [data-testid="stLayoutWrapper"]:has(> [class*="st-key-de_dash_"]),
{_DASH} [data-testid="stColumn"] [class*="st-key-de_dash_"] {{ height: 100%; }}
{_DASH} .dd-head {{ display: flex; gap: .7rem; align-items: center; padding-bottom: .75rem;
                  border-bottom: 1px solid {UI_LINE}; margin-bottom: .1rem; }}
{_DASH} .dd-title {{ font-size: 1.06rem; font-weight: 700; color: {INK}; line-height: 1.3; }}
{_DASH} .dd-cap {{ font-size: .82rem; color: {UI_MUTED}; line-height: 1.4; }}
{_DASH} .dd-rows {{ margin: 0; }}
{_DASH} .dd-row {{ display: flex; justify-content: space-between; align-items: baseline; gap: 1rem; padding: .55rem .1rem;
                 border-top: 1px solid {UI_LINE}; font-size: .94rem; }}
{_DASH} .dd-row:first-child {{ border-top: none; }}
{_DASH} .dd-row .l {{ color: {INK}; font-weight: 500; }}
{_DASH} .dd-row .l small {{ display: block; color: {UI_MUTED}; font-size: .78rem; font-weight: 400; margin-top: .05rem; }}
{_DASH} .dd-row .v {{ color: {INK}; font-weight: 700; text-align: right; white-space: nowrap; font-variant-numeric: tabular-nums; }}
{_DASH} .dd-row .v .neg, {_DASH} .dd-row .v.neg {{ color: {BAD_TEXT}; }}
{_DASH} .dd-row.key {{ background: {UI_BG}; border-radius: 8px; border-top-color: transparent; padding: .6rem .6rem; margin: .15rem 0; }}
{_DASH} .dd-row.key + .dd-row {{ border-top-color: transparent; }}
{_DASH} .dd-row.key .l, {_DASH} .dd-row.key .v {{ font-weight: 800; font-size: 1rem; }}
{_DASH} .dd-kv {{ margin: .1rem 0 .55rem 0; }}
{_DASH} .dd-kv .k {{ font-size: .74rem; font-weight: 700; letter-spacing: .05em; text-transform: uppercase; color: {UI_MUTED}; }}
{_DASH} .dd-kv .val {{ font-size: 1.12rem; font-weight: 800; color: {INK}; line-height: 1.3; margin: .1rem 0 .1rem 0; }}
{_DASH} .dd-kv .s {{ font-size: .82rem; color: {UI_MUTED}; line-height: 1.45; }}
{_DASH} .dd-text {{ font-size: .92rem; color: {INK}; line-height: 1.5; margin: .1rem 0 .2rem 0; }}
{_DASH} .dd-text.muted {{ color: {UI_MUTED}; font-size: .84rem; }}
{_DASH} .dd-text b {{ color: {INK}; }}
{_DASH} ul.dd-list {{ margin: .45rem 0 .1rem 0 !important; padding-left: 0 !important; list-style: none; }}
{_DASH} ul.dd-list li {{ margin-left: 0 !important; padding: .45rem 0; border-top: 1px solid {UI_LINE};
                       font-size: .92rem; color: {INK}; line-height: 1.4; }}
{_DASH} ul.dd-list li:first-child {{ border-top: none; }}
{_DASH} ul.dd-list li small {{ display: block; color: {UI_MUTED}; font-size: .78rem; }}
{_DASH} .dd-group {{ font-size: .78rem; font-weight: 700; letter-spacing: .07em; text-transform: uppercase; color: {UI_MUTED};
                   margin: .9rem 0 .2rem 0; }}
/* Odkazy do sekcí */
{_DASH} [class*="st-key-dd_links_"] {{ flex-direction: row !important; flex-wrap: wrap; gap: .1rem 1.2rem !important;
                                     margin-top: auto; padding-top: .55rem; border-top: 1px solid {UI_LINE}; }}
{_DASH} [class*="st-key-dd_links_"] > div {{ width: auto !important; flex: 0 0 auto; }}
{_DASH} [data-testid="stLayoutWrapper"]:has(> [class*="st-key-dd_links_"]) {{ margin-top: auto; }}
{_DASH} [data-testid="stPageLink"] a {{ padding: .15rem .3rem; border-radius: 8px; }}
{_DASH} [data-testid="stPageLink"] p {{ font-size: .88rem; font-weight: 600; color: {ACCENT}; }}
{_TONE_CSS}
@media (max-width: 1100px) {{
  {_DASH} .dd-kpis {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
}}
@media (max-width: 1000px) {{
  {_DASH} .st-key-dd_alerts [data-testid="stHorizontalBlock"] {{ flex-direction: column; gap: .9rem !important; }}
  {_DASH} .st-key-dd_alerts [data-testid="stColumn"] {{ width: 100% !important; flex: 1 1 auto !important; }}
}}
@media (max-width: 640px) {{
  {_DASH} .dd-kpis {{ gap: .6rem; margin-bottom: 1rem; }}
  {_DASH} .dd-kpi {{ padding: .8rem .8rem; border-radius: 12px; gap: .3rem; }}
  {_DASH} .dd-kpi-top {{ gap: .45rem; }}
  {_DASH} .dd-ico-box {{ width: 1.7rem; height: 1.7rem; border-radius: 8px; }}
  {_DASH} .dd-ico {{ width: 1rem; height: 1rem; }}
  {_DASH} .dd-lbl {{ font-size: .66rem; letter-spacing: .03em; }}
  {_DASH} .dd-val, {_DASH} .dd-val.pair {{ font-size: 1.45rem; }}
  {_DASH} .dd-val small {{ font-size: .8rem; }}
  {_DASH} .dd-sub {{ font-size: .74rem; padding-top: .4rem; }}
  {_DASH} .dd-chip {{ font-size: .68rem; white-space: normal; }}
  {_DASH} [class*="st-key-de_dash_"] {{ padding: .9rem .95rem !important; }}
}}
</style>
"""


DASHBOARD_CSS = dash_css(DASH_KEY)



# ---------------------------------------------------------------------------
# Pronájem strojů (modernizace UI): stejné komponenty jako Dashboard (dd-*) + karty variant.
# Jen zobrazení – hodnoty dodává app.py z výsledku plánovače (planner_view.evaluate).
# ---------------------------------------------------------------------------

RENT_KEY = TAB_KEY_PREFIX + "pronajem"
_RENT = f".st-key-{RENT_KEY}"


def dash_variants(items: list) -> None:
    """Karty porovnávaných variant: items = [dict(title, sub, tag, rows=[(popisek, hodnota HTML)], why)]."""
    cards = "".join(
        f'<div class="dd-var"><div class="dd-var-head"><span class="t">{html.escape(it["title"])}</span>'
        + (f'<span class="tag">{html.escape(it["tag"])}</span>' if it.get("tag") else "")
        + f'</div><div class="dd-var-sub">{html.escape(it["sub"])}</div>'
        + "".join(f'<div class="dd-var-row"><span>{html.escape(lbl)}</span><span>{val}</span></div>'
                  for lbl, val in it["rows"])
        + (f'<div class="dd-var-why">{it["why"]}</div>' if it.get("why") else "")
        + "</div>"
        for it in items
    )
    st.markdown(f'<div class="dd-vars">{cards}</div>', unsafe_allow_html=True)


RENTAL_CSS = dash_css(RENT_KEY).replace("</style>", f"""
{_RENT} .dd-section-title {{ font-size: 1.08rem; font-weight: 700; color: {INK}; margin: 1.3rem 0 .5rem 0; }}
{_RENT} .dd-section-title small {{ font-weight: 500; color: {UI_MUTED}; font-size: .84rem; margin-left: .4rem; }}
{_RENT} [class*="st-key-de_dash_rent_form"] [data-testid="stForm"] {{ border: none; padding: 0; }}
{_RENT} .dd-vars {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(260px, 1fr)); gap: .9rem;
                  margin: .2rem 0 1rem 0; }}
{_RENT} .dd-var {{ background: #fff; border: 1px solid {UI_LINE}; border-radius: 14px; padding: .9rem 1rem;
                 box-shadow: 0 1px 3px rgba(16,24,40,.06); display: flex; flex-direction: column; gap: .3rem; }}
{_RENT} .dd-var-head {{ display: flex; justify-content: space-between; gap: .5rem; align-items: baseline; }}
{_RENT} .dd-var-head .t {{ font-weight: 800; color: {INK}; font-size: 1rem; }}
{_RENT} .dd-var-head .tag {{ font-size: .72rem; color: {UI_MUTED}; background: {UI_BG}; border-radius: 999px;
                           padding: .1rem .5rem; white-space: nowrap; }}
{_RENT} .dd-var-sub {{ font-size: .82rem; color: {UI_MUTED}; margin-bottom: .2rem; }}
{_RENT} .dd-var-row {{ display: flex; justify-content: space-between; align-items: center; gap: .8rem; padding: .35rem 0;
                     border-top: 1px solid {UI_LINE}; font-size: .9rem; color: {INK}; }}
{_RENT} .dd-var-row span:last-child {{ text-align: right; font-weight: 700; }}
{_RENT} .dd-var-why {{ font-size: .8rem; color: {UI_MUTED}; line-height: 1.45; border-top: 1px solid {UI_LINE}; padding-top: .4rem; }}
{_RENT} .dd-kpis .dd-val.word {{ font-size: 1.55rem; }}
@media (max-width: 640px) {{
  {_RENT} .dd-kpis .dd-val.word {{ font-size: 1.15rem; }}
}}
</style>""")
