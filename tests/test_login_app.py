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
import auth  # noqa: E402

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
    return len(at.tabs) == 7 and not at.exception             # 6 původních záložek + Kapacita strojů


def errors(at):
    return [e.value for e in at.error]


def test_protected_content_hidden_before_login():
    at = app()
    assert len(at.tabs) == 0 and not at.exception
    assert [t.label for t in at.text_input] == ["Uživatelské jméno", "Heslo"]


def test_lukas_logs_in_with_current_password():
    at = login(app(), "lukas", LEGACY)
    assert logged_in(at)
    assert any("Lukáš" in m.value for m in at.sidebar.markdown)
    assert any("Správa účtů" in e.label for e in at.sidebar.expander)       # nástroj správce


def test_honza_logs_in_without_admin_tools():
    at = login(app(), "honza", HONZA)
    assert logged_in(at)
    assert any("Honza Podroužek" in m.value for m in at.sidebar.markdown)
    assert not any("Správa účtů" in e.label for e in at.sidebar.expander)


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
    assert len(at.tabs) == 0 and field(at, "Heslo").value == ""
    at.run()                                   # další překreslení nesmí uživatele přihlásit zpět
    assert len(at.tabs) == 0
    login(at, "honza", HONZA)
    assert logged_in(at)                       # nové přihlášení po odhlášení funguje


def test_deactivated_account_is_logged_out_on_next_action():
    at = login(app(), "honza", HONZA)
    assert logged_in(at)
    at.secrets["users"] = dict(USERS, honza=dict(USERS["honza"], active=False))
    at.run()
    assert len(at.tabs) == 0


def test_single_password_mode_works_as_before():
    at = app(with_users=False)
    assert [t.label for t in at.text_input] == ["Přístupové heslo"]
    at = login(at, None, "spatne")
    assert "Nesprávné heslo." in errors(at)
    at = login(app(with_users=False), None, LEGACY)
    assert logged_in(at)


def test_missing_configuration():
    at = app(with_users=False, password=None)
    assert len(at.tabs) == 0 and any("není nastaveno" in e for e in errors(at))


def test_admin_hash_tool_creates_valid_hash():
    at = login(app(), "lukas", LEGACY)
    new = secrets.token_urlsafe(14)
    at.text_input(key="admin_pw1").set_value(new)
    at.text_input(key="admin_pw2").set_value(new).run()
    produced = [c.value for c in at.sidebar.code]
    assert produced and auth.verify_password(new, produced[0])


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK", name)
