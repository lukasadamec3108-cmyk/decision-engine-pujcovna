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
            background: #FFF7ED; border: 1px solid #F6D3A8; border-radius: 8px; padding: .35rem .6rem;
            font-size: .75rem; color: #7A4A06; margin-bottom: .2rem;
        }}

        /* Postranní panel: jasně oddělené sekce */
        [data-testid="stSidebar"] h3 {{
            font-size: .78rem !important; text-transform: uppercase; letter-spacing: .05em;
            color: {MUTED} !important; border-top: 1px solid {BORDER};
            padding-top: .9rem !important; margin-top: .5rem !important;
        }}

        /* Záložka Doporučení – manažerský přehled */
        .de-hero {{
            background: #fff; border: 1px solid {BORDER}; border-left: 6px solid {ORANGE};
            border-radius: 14px; padding: 1.2rem 1.5rem; margin: .3rem 0 1rem 0;
            box-shadow: 0 2px 14px rgba(20,48,79,.06);
        }}
        .de-hero.bad {{ border-left-color: {BAD}; }}
        .de-hero .lbl {{
            font-size: .78rem; font-weight: 700; letter-spacing: .08em; text-transform: uppercase; color: {ORANGE};
        }}
        .de-hero.bad .lbl {{ color: {BAD}; }}
        .de-hero .ttl {{ font-size: 2rem; font-weight: 800; color: {NAVY}; line-height: 1.15; margin: .2rem 0 .45rem 0; }}
        .de-hero .txt {{ font-size: 1.05rem; color: {TEXT}; max-width: 60rem; }}

        .de-kpis {{ display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 1rem; margin: 0 0 1.6rem 0; }}
        .de-kpi {{ background: #fff; border: 1px solid {BORDER}; border-radius: 12px; padding: 1rem 1.2rem; }}
        .de-kpi .lbl {{ font-size: .78rem; font-weight: 700; letter-spacing: .04em; text-transform: uppercase; color: {MUTED}; }}
        .de-kpi .val {{ font-size: 2rem; font-weight: 800; color: {NAVY}; line-height: 1.15; margin: .4rem 0 .35rem 0; }}
        .de-kpi .val.neg {{ color: {BAD}; }}
        .de-kpi .sub {{ font-size: .88rem; color: {MUTED}; }}
        .de-status {{
            display: inline-block; font-size: .78rem; font-weight: 700; padding: .12rem .55rem;
            border-radius: 999px; margin-right: .35rem;
        }}
        .de-status.good {{ color: {GOOD}; background: #E7F5EC; }}
        .de-status.warn {{ color: #8A5A00; background: #FDF3DC; }}
        .de-status.bad {{ color: {BAD}; background: #FEE4E2; }}

        .de-section {{ font-size: 1.05rem; font-weight: 700; color: {NAVY}; margin: .2rem 0 .6rem 0; }}
        .de-cmp {{ width: 100%; border-collapse: separate; border-spacing: 0; border: 1px solid {BORDER};
                   border-radius: 12px; overflow: hidden; margin-bottom: 1.4rem; background: #fff; }}
        .de-cmp th, .de-cmp td {{ padding: .7rem 1rem; border-bottom: 1px solid {BORDER}; text-align: right; }}
        .de-cmp tr:last-child td {{ border-bottom: none; }}
        .de-cmp th {{ font-size: .85rem; color: {NAVY}; background: {SURFACE}; font-weight: 700; }}
        .de-cmp th:first-child, .de-cmp td:first-child {{ text-align: left; color: {MUTED}; font-weight: 500; }}
        .de-cmp td {{ font-size: .98rem; color: {TEXT}; font-variant-numeric: tabular-nums; }}
        .de-cmp .rec {{ background: #FFF7ED; }}
        .de-cmp th.rec {{ background: #FDEBD3; }}
        .de-cmp .best {{ font-weight: 700; color: {NAVY}; }}
        .de-cmp .best::after {{ content: " ✓"; color: {GOOD}; }}
        .de-cmp .tag {{
            display: inline-block; margin-left: .4rem; font-size: .68rem; font-weight: 700; color: #fff;
            background: {ORANGE}; padding: .1rem .45rem; border-radius: 999px; vertical-align: middle;
        }}
        .de-cmp-note {{ font-size: .8rem; color: {MUTED}; margin: -1rem 0 1.4rem 0; }}

        .de-why {{ border-left: 3px solid {ORANGE}; padding: .1rem 0 .1rem 1rem; margin: 0 0 1.6rem 0; }}
        .de-why p {{ margin: .25rem 0; font-size: 1rem; color: {TEXT}; }}

        .de-steps {{ margin: 0 0 1.4rem 0; padding-left: 1.3rem; }}
        .de-steps li {{ margin: .3rem 0; color: {TEXT}; }}

        @media (max-width: 768px) {{
            .de-kpis {{ grid-template-columns: 1fr; gap: .7rem; }}
            .de-hero {{ padding: 1rem 1.1rem; }}
            .de-hero .ttl {{ font-size: 1.5rem; }}
            .de-kpi .val {{ font-size: 1.6rem; }}
            .de-cmp th, .de-cmp td {{ padding: .55rem .5rem; font-size: .86rem; }}
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


def kpi_cards(cards: list) -> None:
    """Řada velkých KPI karet. cards = [dict(label, value, sub, negative)], na mobilu pod sebou."""
    items = "".join(
        f'<div class="de-kpi"><div class="lbl">{html.escape(c["label"])}</div>'
        f'<div class="val {"neg" if c.get("negative") else ""}">{c["value"]}</div>'
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
