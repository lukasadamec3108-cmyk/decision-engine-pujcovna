"""Testy přihlášení v běžící aplikaci (Streamlit AppTest). Hesla jsou náhodná, jen pro test.

Spuštění: python -m pytest tests   nebo   python tests/test_login_app.py
"""

import os
import secrets
import sys
from pathlib import Path

from streamlit.testing.v1 import AppTest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))
import auth  # noqa: E402
import navigace  # noqa: E402
from app_pages import open_page  # noqa: E402

LEGACY = secrets.token_urlsafe(12)
HONZA = secrets.token_urlsafe(12)
MARTIN = secrets.token_urlsafe(12)
USERS = {
    "lukas": {"name": "Lukáš", "role": "admin", "active": True, "legacy_password": True},
    "honza": {"name": "Honza Podroužek", "role": "tester", "active": True,
              "password_hash": auth.hash_password(HONZA, iterations=100_000)},
    "martin": {"name": "Martin Štefan", "role": "owner", "active": False,
               "password_hash": auth.hash_password(MARTIN, iterations=100_000)},
}


def app(with_users=True, password=LEGACY):
    os.chdir(ROOT)
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300)
    if password is not None:
        at.secrets["password"] = password
    if with_users:
        at.secrets["users"] = USERS
    at.run()
    return at


def field(at, label):
    return [t for t in at.text_input if t.label == label][0]


def login(at, user, password):
    if user is not None:
        field(at, "Uživatelské jméno").set_value(user)
    field(at, "Heslo" if user is not None else "Přístupové heslo").set_value(password).run()
    return at


def logged_in(at) -> bool:
    """Přihlášený uživatel má v panelu odhlášení a vidí obsah aplikace (výchozí stránka Dashboard)."""
    return (not at.exception and any(b.key == "logout" for b in at.sidebar.button)
            and any('class="de-header' in m.value for m in at.main.markdown))


def content_hidden(at) -> bool:
    """Žádný chráněný obsah: bez odhlášení, hlavičky, záložek, ukazatelů a tabulek."""
    return (not any(b.key == "logout" for b in at.button) and len(at.tabs) == 0
            and not any(("de-header" in m.value or "de-kpi" in m.value) and not m.value.lstrip().startswith("<style")
                        for m in at.markdown)
            and len(at.dataframe) == 0)


def errors(at):
    return [e.value for e in at.error]


def test_protected_content_hidden_before_login():
    at = app()
    assert content_hidden(at) and not at.exception
    assert [t.label for t in at.text_input] == ["Uživatelské jméno", "Heslo"]


def test_no_page_is_reachable_without_login():
    at = app()
    for key in navigace.KEYS:                                   # i přímou adresou stránky
        open_page(at, key)
        assert content_hidden(at) and not at.exception, key
        assert [t.label for t in at.text_input] == ["Uživatelské jméno", "Heslo"], key


def test_lukas_logs_in_with_current_password():
    at = login(app(), "lukas", LEGACY)
    assert logged_in(at)
    assert any("Lukáš" in m.value for m in at.sidebar.markdown)
    open_page(at, "nastaveni")
    assert any("Správa účtů" in e.label for e in at.main.expander)          # nástroj správce (Nastavení)


def test_honza_logs_in_without_admin_tools():
    at = login(app(), "honza", HONZA)
    assert logged_in(at)
    assert any("Honza Podroužek" in m.value for m in at.sidebar.markdown)
    open_page(at, "nastaveni")
    assert not any("Správa účtů" in e.label for e in at.expander)          # nikde v aplikaci
    assert not [t for t in at.text_input if t.key in ("admin_pw1", "admin_pw2")]
    assert any("jen administrátor" in m.value for m in at.main.markdown)


def test_wrong_password_rejected():
    at = login(app(), "honza", "spatne-heslo-123")
    assert not logged_in(at) and auth.GENERIC_ERROR in errors(at)


def test_inactive_martin_rejected():
    at = login(app(), "martin", MARTIN)
    assert not logged_in(at) and auth.GENERIC_ERROR in errors(at)


def test_logout():
    at = login(app(), "honza", HONZA)
    assert logged_in(at)
    at.button(key="logout").click().run()
    assert content_hidden(at) and field(at, "Heslo").value == ""
    at.run()                                   # další překreslení nesmí uživatele přihlásit zpět
    assert content_hidden(at)
    login(at, "honza", HONZA)
    assert logged_in(at)                       # nové přihlášení po odhlášení funguje


def test_deactivated_account_is_logged_out_on_next_action():
    at = login(app(), "honza", HONZA)
    assert logged_in(at)
    at.secrets["users"] = dict(USERS, honza=dict(USERS["honza"], active=False))
    at.run()
    assert content_hidden(at)


def test_single_password_mode_works_as_before():
    at = app(with_users=False)
    assert [t.label for t in at.text_input] == ["Přístupové heslo"]
    at = login(at, None, "spatne")
    assert "Nesprávné heslo." in errors(at)
    at = login(app(with_users=False), None, LEGACY)
    assert logged_in(at)


def test_missing_configuration():
    at = app(with_users=False, password=None)
    assert content_hidden(at) and any("není nastaveno" in e for e in errors(at))


def test_admin_hash_tool_creates_valid_hash():
    at = login(app(), "lukas", LEGACY)
    open_page(at, "nastaveni")
    new = secrets.token_urlsafe(14)
    at.text_input(key="admin_pw1").set_value(new)
    at.text_input(key="admin_pw2").set_value(new).run()
    produced = [c.value for c in at.main.code]
    assert produced and auth.verify_password(new, produced[0])


def test_logout_from_settings_page():
    at = login(app(), "honza", HONZA)
    open_page(at, "nastaveni")
    at.button(key="logout_nastaveni").click().run()
    assert content_hidden(at) and not logged_in(at)


def test_logout_clears_kept_investment_inputs():
    at = login(app(), "honza", HONZA)
    open_page(at, "investice")
    at.number_input(key="inv_fixed_cost").set_value(150000).run()
    assert at.session_state["_keep_inv_fixed_cost"] == 150000
    at.button(key="logout").click().run()
    assert not [k for k in at.session_state.keys() if str(k).startswith("_keep_")]
    login(at, "honza", HONZA)
    open_page(at, "investice")
    assert at.number_input(key="inv_fixed_cost").value == 100000          # výchozí hodnota


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK", name)
