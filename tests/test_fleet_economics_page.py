"""Testy kostry stránky „Ekonomika flotily“ (E2.2 – K3.4b). Hesla jsou náhodná, jen pro test.

Stránka zatím nic nepočítá: jen nadpis, štítek „Ukázková data / neúplné“, upozornění „Neúplný výsledek –
nejde o čistý zisk.“ a vysvětlení, že výsledky budou doplněny (K3.4c).

Spuštění: python -m pytest tests/test_fleet_economics_page.py
"""

import os
import re
import secrets
import sys
from pathlib import Path

from streamlit.testing.v1 import AppTest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))
import auth  # noqa: E402
import fleet_view  # noqa: E402
import navigace  # noqa: E402
from app_pages import open_page  # noqa: E402

KEY = "ekonomika"
LEGACY = secrets.token_urlsafe(12)
TESTER = secrets.token_urlsafe(12)
USERS = {
    "lukas": {"name": "Lukáš", "role": "admin", "active": True, "legacy_password": True},
    "honza": {"name": "Honza Test", "role": "tester", "active": True,
              "password_hash": auth.hash_password(TESTER, iterations=100_000)},
}


def app(users=False):
    os.chdir(ROOT)
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300)
    at.secrets["password"] = LEGACY
    if users:
        at.secrets["users"] = USERS
    at.run()
    return at


def login_single(at):
    at.text_input[0].set_value(LEGACY).run()
    assert not at.exception, [e.value for e in at.exception]
    return at


def login_user(at, user, password):
    [u] = [t for t in at.text_input if t.label == "Uživatelské jméno"]
    [p] = [t for t in at.text_input if t.label == "Heslo"]
    u.set_value(user)
    p.set_value(password).run()
    return at


def body(at):
    """Viditelný obsah stránky bez stylů a bez společné hlavičky aplikace."""
    return [m.value for m in at.main.markdown
            if not m.value.lstrip().startswith("<style") and 'class="de-header' not in m.value]


def plain(values):
    return re.sub(r"<[^>]+>", " ", " ".join(values)).replace("\xa0", " ")


# ---------------------------------------------------------------------------
# Navigace
# ---------------------------------------------------------------------------

def test_page_is_in_overview_group_directly_below_dashboard():
    assert navigace.KEYS[:2] == ["dashboard", KEY]
    assert navigace.groups()["Přehled"] == ["dashboard", KEY]
    entry = next(p for p in navigace.PAGES if p[0] == KEY)
    assert entry[2] == "Ekonomika flotily" and entry[5] == "ekonomika-flotily"
    assert navigace.page_file(KEY).is_file()
    assert navigace.DEFAULT_PAGE == "dashboard"                         # výchozí stránka se nemění
    assert KEY not in {k for k, _ in navigace.ORIGINAL_TABS.values()}    # původní záložky beze změny


# ---------------------------------------------------------------------------
# Obsah kostry
# ---------------------------------------------------------------------------

def test_skeleton_shows_only_required_texts():
    at = open_page(login_single(app()), KEY)
    assert not at.exception and not at.error
    text = plain(body(at))
    assert "Ekonomika flotily" in text
    assert "Ukázková data / neúplné" in text
    assert fleet_view.NOT_NET_PROFIT in text                            # stejný text jako zobrazovací vrstva
    assert "budou doplněny v následujícím kroku" in text
    assert any('class="de-tab-title"' in m for m in body(at))            # společný styl stránek
    assert any('class="de-insight bad"' in m for m in body(at))
    assert any(b.key == "logout" for b in at.sidebar.button)             # odhlášení jako na ostatních stránkách


def test_skeleton_shows_no_amounts_indicators_or_inputs():
    at = open_page(login_single(app()), KEY)
    text = plain(body(at))
    assert "Kč" not in text and not re.search(r"\d", text), text         # žádná částka ani číslo
    assert not any("de-kpi" in m or "dd-kpi" in m for m in body(at))      # žádné ukazatele
    assert len(at.main.dataframe) == 0 and len(at.tabs) == 0 and len(at.main.expander) == 0
    for kind in (at.main.button, at.main.selectbox, at.main.slider, at.main.number_input, at.main.radio,
                 at.main.date_input, at.main.multiselect, at.main.text_input):
        assert len(kind) == 0
    assert "zisk" not in text.replace(fleet_view.NOT_NET_PROFIT, "").lower()


def test_app_does_not_compute_fleet_economics_yet():
    src = (ROOT / "app.py").read_text(encoding="utf-8")
    for name in ("period_result", "fleet_view", "fleet_economics", "conflict_options", "period_ledger"):
        assert name not in src, name


# ---------------------------------------------------------------------------
# Přístup
# ---------------------------------------------------------------------------

def test_not_reachable_without_login():
    at = open_page(app(users=True), KEY)
    assert not at.exception
    assert [t.label for t in at.text_input] == ["Uživatelské jméno", "Heslo"]
    assert "Ekonomika flotily" not in plain(m.value for m in at.markdown)


def test_reachable_for_logged_in_roles_like_other_overviews():
    for user, password in (("lukas", LEGACY), ("honza", TESTER)):
        at = login_user(app(users=True), user, password)
        assert any(b.key == "logout" for b in at.sidebar.button), user
        open_page(at, KEY)
        assert not at.exception and "Ekonomika flotily" in plain(body(at)), user
        open_page(at, "dashboard")                                       # stejná oprávnění jako Dashboard
        assert not at.exception and "Souhrn výsledků" in plain(body(at)), user
