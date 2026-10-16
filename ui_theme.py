"""Vzhled aplikace Decision Engine: barvy, formátování čísel, CSS a společné prvky UI.

Modul neobsahuje žádné výpočty – jen prezentaci výsledků.
"""

import base64
import html
from pathlib import Path

import streamlit as st

# ==================================================
# BARVY
# ==================================================

NAVY = "#14304F"        # hlavička, nadpisy
ORANGE = "#D97706"      # zvýraznění, doporučená varianta, Stroj 2
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

NBSP = " "


def fmt_num(value, decimals: int = 0) -> str:
    """Číslo v českém formátu: mezera mezi tisíci, desetinná čárka, znaménko minus."""
    if value is None:
        return "–"
    text = f"{abs(float(value)):,.{decimals}f}"
    text = text.replace(",", NBSP).replace(".", ",")
    return ("−" if float(value) < 0 and text.strip("0,") else "") + text


def fmt_czk(value, decimals: int = 0) -> str:
    return f"{fmt_num(value, decimals)}{NBSP}Kč"


def fmt_pct(value_percent, decimals: int = 1) -> str:
    """Formátuje hodnotu, která už je v procentech (např. 24.6 → „24,6 %“)."""
    return f"{fmt_num(value_percent, decimals)}{NBSP}%"


def fmt_days(value, decimals: int = 1) -> str:
    return f"{fmt_num(value, decimals)}{NBSP}dní"


PRIORITY_CZ = {"High": "Vysoká", "Medium": "Střední", "Low": "Nízká"}
RISK_CZ = {"High": "Vysoké", "Medium": "Střední", "Low": "Nízké"}


# ==================================================
# LOGO
# ==================================================


def logo_src() -> str:
    """Vrátí zdroj loga pro <img>.

    Pořadí: 1) Streamlit secrets `logo_url` (URL nebo data URI – firemní logo mimo Git),
    2) lokální soubor assets/logo.png|jpg|svg (je v .gitignore), 3) neutrální zástupné logo.
    """
    try:
        url = str(st.secrets.get("logo_url", "")).strip()
    except Exception:
        url = ""
    if url:
        return url

    for name, mime in (("logo.png", "image/png"), ("logo.jpg", "image/jpeg"), ("logo.svg", "image/svg+xml")):
        path = ASSETS / name
        if path.exists():
            return f"data:{mime};base64," + base64.b64encode(path.read_bytes()).decode()

    placeholder = ASSETS / "logo_placeholder.svg"
    return "data:image/svg+xml;base64," + base64.b64encode(placeholder.read_bytes()).decode()


# ==================================================
# CSS
# ==================================================


def inject_css() -> None:
    st.markdown(
        f"""
        <style>
        .block-container {{ padding-top: 1.4rem; padding-bottom: 3rem; max-width: 1280px; }}
        h1, h2, h3, h4 {{ color: {NAVY}; }}
        h2 {{ font-size: 1.45rem !important; }}
        h3 {{ font-size: 1.15rem !important; }}

        [data-testid="stSidebar"] {{ background: {SURFACE}; border-right: 1px solid {BORDER}; }}
        [data-testid="stSidebar"] h3 {{ font-size: .95rem !important; margin: .4rem 0 0 0; }}

        .stTabs [data-baseweb="tab-list"] {{ gap: .2rem; border-bottom: 1px solid {BORDER}; overflow-x: auto; }}
        .stTabs [data-baseweb="tab"] {{ padding: .55rem .9rem; font-weight: 600; color: {MUTED}; }}
        .stTabs [aria-selected="true"] {{ color: {NAVY}; }}
        .stTabs [data-baseweb="tab-highlight"] {{ background-color: {ORANGE}; height: 3px; }}

        [data-testid="stMetric"] {{
            background: #fff; border: 1px solid {BORDER}; border-radius: 10px; padding: .75rem 1rem;
        }}
        [data-testid="stMetricLabel"] p {{ white-space: normal; color: {MUTED}; font-size: .85rem; }}
        [data-testid="stMetricValue"] {{ color: {NAVY}; font-size: 1.5rem; }}
        [data-testid="stExpander"] details {{ border-radius: 10px; border-color: {BORDER}; }}

        .de-header {{
            display: flex; align-items: center; gap: 1rem; background: {NAVY}; color: #fff;
            border-radius: 12px; padding: .85rem 1.2rem; margin-bottom: 1.1rem;
        }}
        .de-header img {{ height: 46px; max-width: 170px; background: #fff; border-radius: 6px; padding: 4px; }}
        .de-header.stacked {{ flex-direction: column; align-items: flex-start; gap: .6rem; }}
        .de-header.stacked .de-badge {{ margin-left: 0; }}
        .de-title {{ font-size: 1.4rem; font-weight: 700; line-height: 1.2; }}
        .de-sub {{ font-size: .9rem; opacity: .85; }}
        .de-badge {{
            margin-left: auto; background: {ORANGE}; color: #fff; font-weight: 700; font-size: .75rem;
            letter-spacing: .03em; padding: .3rem .7rem; border-radius: 999px; white-space: nowrap;
        }}

        .de-card {{
            background: #fff; border: 1px solid {BORDER}; border-radius: 12px; padding: 1rem 1.15rem;
            height: 100%; margin-bottom: .6rem;
        }}
        .de-card.rec {{ border: 2px solid {ORANGE}; box-shadow: 0 2px 12px rgba(217,119,6,.12); }}
        .de-card-title {{
            margin: 0 0 .5rem 0; font-size: 1.1rem; font-weight: 700; color: {NAVY};
            display: flex; align-items: center; gap: .55rem;
        }}
        .de-dot {{ width: .75rem; height: .75rem; border-radius: 50%; display: inline-block; }}
        .de-tag {{
            font-size: .7rem; font-weight: 700; color: #fff; background: {ORANGE};
            padding: .15rem .5rem; border-radius: 999px; margin-left: auto;
        }}
        .de-row {{
            display: flex; justify-content: space-between; gap: 1rem; padding: .4rem 0;
            border-bottom: 1px dashed {BORDER}; font-size: .95rem;
        }}
        .de-row:last-child {{ border-bottom: none; }}
        .de-row span:first-child {{ color: {MUTED}; }}
        .de-row span:last-child {{ font-weight: 600; color: {TEXT}; text-align: right; }}

        .de-verdict {{
            border-radius: 12px; padding: 1rem 1.2rem; margin: .2rem 0 1rem 0;
            background: #FFF7ED; border-left: 6px solid {ORANGE}; color: {TEXT};
        }}
        .de-verdict.bad {{ background: #FEF3F2; border-left-color: {BAD}; }}
        .de-verdict .big {{ font-size: 1.35rem; font-weight: 700; color: {NAVY}; margin-bottom: .25rem; }}

        .de-note {{
            background: #fff; border: 1px solid {BORDER}; border-left: 4px solid {NAVY};
            border-radius: 8px; padding: .6rem .8rem; font-size: .85rem; color: {MUTED};
        }}
        .de-demo {{
            background: #FFF7ED; border: 1px solid #F6D3A8; border-radius: 8px; padding: .55rem .7rem;
            font-size: .8rem; color: #7A4A06; margin-bottom: .6rem;
        }}

        @media (max-width: 640px) {{
            .block-container {{ padding-left: 1rem; padding-right: 1rem; }}
            .de-header {{ flex-wrap: wrap; }}
            .de-badge {{ margin-left: 0; }}
            .de-title {{ font-size: 1.15rem; }}
            [data-testid="stMetricValue"] {{ font-size: 1.25rem; }}
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


# ==================================================
# KOMPONENTY
# ==================================================


def header(title: str, subtitle: str, badge: str, stacked: bool = False) -> None:
    st.markdown(
        f"""
        <div class="de-header {'stacked' if stacked else ''}">
            <img src="{logo_src()}" alt="Logo společnosti">
            <div>
                <div class="de-title">{html.escape(title)}</div>
                <div class="de-sub">{html.escape(subtitle)}</div>
            </div>
            <div class="de-badge">{html.escape(badge)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def verdict(title: str, text: str, tone: str = "") -> None:
    st.markdown(
        f'<div class="de-verdict {tone}"><div class="big">{title}</div>{text}</div>',
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
