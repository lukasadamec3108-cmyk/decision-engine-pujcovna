"""Registr stránek aplikace (etapa M2) – jediný zdroj pořadí, názvů, ikon a skupin levého menu.

Obsah stránek vykresluje app.py (podle aktivní stránky). Soubory stranky/stranka_*.py jsou jen
značky pro nativní navigaci Streamlitu (st.Page potřebuje soubor, aby šlo stránku otevřít
adresou i v automatických testech – AppTest.switch_page).

Modul nezávisí na Streamlitu – používají ho app.py i testy.
"""

from pathlib import Path

ROOT = Path(__file__).resolve().parent
PAGES_DIR = "stranky"

# (klíč, soubor, název v menu, ikona, skupina menu, adresa)
PAGES = [
    ("dashboard", "stranka_dashboard.py", "Dashboard", ":material/dashboard:", "Přehled", "dashboard"),
    ("ekonomika", "stranka_ekonomika.py", "Ekonomika flotily", ":material/account_balance_wallet:", "Přehled",
     "ekonomika-flotily"),
    ("pronajem", "stranka_pronajem.py", "Pronájem strojů", ":material/local_shipping:", "Provoz", "pronajem"),
    ("zakazky", "stranka_zakazky.py", "Vlastní zakázky", ":material/construction:", "Provoz", "zakazky"),
    ("strojovy_park", "stranka_strojovy_park.py", "Strojový park", ":material/precision_manufacturing:", "Provoz",
     "strojovy-park"),
    ("kapacita", "stranka_kapacita.py", "Plánování kapacity", ":material/calendar_month:", "Provoz", "kapacita"),
    ("rozhodovaci_centrum", "stranka_rozhodovaci_centrum.py", "Rozhodovací centrum", ":material/insights:", "Provoz",
     "rozhodovaci-centrum"),
    ("investice", "stranka_investice.py", "Investice", ":material/savings:", "Investice", "investice"),
    ("nastaveni", "stranka_nastaveni.py", "Nastavení", ":material/settings:", "Systém", "nastaveni"),
]
DEFAULT_PAGE = "dashboard"
KEYS = [p[0] for p in PAGES]

# Sekce Investice: původní záložky 1–5 (stejné popisky)
INVESTICE_TABS = ["🏆 Doporučení", "📈 Možné výsledky", "⚠️ Riziko", "🎯 Co ovlivňuje zisk", "🛒 Nákup strojů"]

# Kam se přesunuly původní záložky aplikace (číslo záložky → stránka, případně záložka uvnitř stránky)
ORIGINAL_TABS = {
    1: ("investice", "🏆 Doporučení"), 2: ("investice", "📈 Možné výsledky"), 3: ("investice", "⚠️ Riziko"),
    4: ("investice", "🎯 Co ovlivňuje zisk"), 5: ("investice", "🛒 Nákup strojů"),
    6: ("zakazky", None), 7: ("kapacita", None), 8: ("rozhodovaci_centrum", None),
}


def page_file(key: str) -> Path:
    """Absolutní cesta k souboru stránky (st.Page ji pak najde bez ohledu na spouštěný skript)."""
    name = next(p[1] for p in PAGES if p[0] == key)
    return ROOT / PAGES_DIR / name


def page_relpath(key: str) -> str:
    """Cesta stránky relativně ke kořeni projektu (pro AppTest.switch_page)."""
    return f"{PAGES_DIR}/{next(p[1] for p in PAGES if p[0] == key)}"


def groups() -> dict:
    """Skupiny menu v pořadí stránek: {skupina: [klíče]}."""
    out = {}
    for key, _, _, _, group, _ in PAGES:
        out.setdefault(group, []).append(key)
    return out
