"""Testy přihlašování (auth.py). Hesla se generují náhodně při běhu testu – v repozitáři žádná nejsou.

Spuštění: python -m pytest tests   nebo   python tests/test_auth.py
"""

import secrets
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import auth  # noqa: E402

LEGACY = secrets.token_urlsafe(12)
HONZA = secrets.token_urlsafe(12)
MARTIN = secrets.token_urlsafe(12)


def config(**overrides):
    users = {
        "lukas": {"name": "Lukáš", "role": "admin", "active": True, "legacy_password": True},
        "honza": {"name": "Honza Podroužek", "role": "tester", "active": True,
                  "password_hash": auth.hash_password(HONZA, iterations=100_000)},
        "martin": {"name": "Martin Štefan", "role": "owner", "active": False,
                   "password_hash": auth.hash_password(MARTIN, iterations=100_000)},
    }
    users.update(overrides)
    return auth.load_config({"password": LEGACY, "users": users})


def test_hash_roundtrip_and_salt():
    h1, h2 = auth.hash_password("tajne-heslo-1"), auth.hash_password("tajne-heslo-1")
    assert h1 != h2                                   # jiná sůl → jiný hash
    assert auth.verify_password("tajne-heslo-1", h1) and auth.verify_password("tajne-heslo-1", h2)
    assert not auth.verify_password("tajne-heslo-2", h1)
    assert "tajne-heslo-1" not in h1 and h1.startswith("pbkdf2_sha256$600000$")


def test_invalid_hash_never_verifies():
    for bad in ("", "abc", "pbkdf2_sha256$10$AAAA$BBBB", "md5$1$x$y", "VLOZTE_HASH_Z_NASTROJE"):
        assert not auth.verify_password("cokoli", bad) and not auth.is_valid_hash(bad)


def test_single_password_mode_unchanged():
    c = auth.load_config({"password": LEGACY})
    assert c["mode"] == "single" and auth.can_log_in(c)
    assert auth.authenticate(c, "", LEGACY).is_admin
    assert auth.authenticate(c, "", LEGACY + " ") is not None      # mezera navíc nevadí (jako dříve)
    assert auth.authenticate(c, "", "spatne") is None
    assert auth.load_config({})["mode"] == "missing" and not auth.can_log_in(auth.load_config({}))


def test_lukas_keeps_current_password():
    u = auth.authenticate(config(), "lukas", LEGACY)
    assert u is not None and u.is_admin and u.name == "Lukáš"
    assert auth.authenticate(config(), "Lukas ", LEGACY) is not None     # jméno bez ohledu na velikost


def test_honza_own_account():
    u = auth.authenticate(config(), "honza", HONZA)
    assert u is not None and u.role == "tester" and not u.is_admin
    assert auth.authenticate(config(), "honza", LEGACY) is None          # cizí heslo nefunguje
    assert auth.authenticate(config(), "lukas", HONZA) is None


def test_wrong_password_and_unknown_user():
    c = config()
    assert auth.authenticate(c, "honza", "spatne-heslo") is None
    assert auth.authenticate(c, "neznamy", HONZA) is None
    assert auth.authenticate(c, "honza", "") is None


def test_inactive_martin_cannot_log_in():
    assert auth.authenticate(config(), "martin", MARTIN) is None
    on = config(martin={"name": "Martin Štefan", "role": "owner", "active": True,
                        "password_hash": auth.hash_password(MARTIN, iterations=100_000)})
    assert auth.authenticate(on, "martin", MARTIN).role == "owner"


def test_missing_active_flag_means_inactive():
    c = config(honza={"name": "H", "role": "tester",
                      "password_hash": auth.hash_password(HONZA, iterations=100_000)})
    assert auth.authenticate(c, "honza", HONZA) is None


def test_deactivation_and_password_change_are_per_user():
    c = config(honza={"name": "H", "role": "tester", "active": False,
                      "password_hash": auth.hash_password(HONZA, iterations=100_000)})
    assert not auth.still_allowed(c, "honza") and auth.still_allowed(c, "lukas")
    new = secrets.token_urlsafe(12)
    c2 = config(honza={"name": "H", "role": "tester", "active": True,
                       "password_hash": auth.hash_password(new, iterations=100_000)})
    assert auth.authenticate(c2, "honza", HONZA) is None and auth.authenticate(c2, "honza", new)
    assert auth.authenticate(c2, "lukas", LEGACY) is not None


def test_config_problems_are_reported():
    c = config(petr={"name": "P", "role": "boss", "active": True, "password_hash": "x"},
               honza={"name": "H", "role": "tester", "active": True, "password_hash": "VLOZTE"})
    assert any("neznámá role" in p for p in c["problems"])
    assert any("honza" in p and "password_hash" in p for p in c["problems"])
    assert auth.authenticate(c, "honza", HONZA) is None and auth.can_log_in(c)   # Lukáš dál může


def test_no_usable_account():
    c = auth.load_config({"users": {"honza": {"role": "tester", "active": True, "password_hash": "x"}}})
    assert c["mode"] == "users" and not auth.can_log_in(c)
    c2 = auth.load_config({"users": {"lukas": {"role": "admin", "active": True, "legacy_password": True}}})
    assert not auth.can_log_in(c2)          # legacy_password bez klíče password nefunguje


def test_password_rules():
    assert auth.password_problems("kratke")
    assert auth.password_problems(" mezera-na-zacatku")
    assert auth.password_problems("dostatecne-dlouhe") == []


def test_no_real_hashes_or_passwords_in_repository_files():
    # V repozitáři smí být jen zástupné texty – žádný platný hash ani heslo.
    import re
    pattern = re.compile(r"pbkdf2_sha256\$\d+\$[A-Za-z0-9+/=]+\$[A-Za-z0-9+/=]+")
    files = [p for p in ROOT.rglob("*") if p.is_file() and ".git" not in p.parts
             and p.suffix in (".py", ".md", ".toml", ".example", ".csv", ".txt", ".cfg")]
    files.append(ROOT / ".streamlit" / "secrets.toml.example")
    for path in files:
        if path.name == "secrets.toml":
            continue                      # lokální soubor mimo repozitář (.gitignore)
        for match in pattern.findall(path.read_text(encoding="utf-8", errors="ignore")):
            assert not auth.is_valid_hash(match), f"platný hash v {path}"
    example = (ROOT / ".streamlit" / "secrets.toml.example").read_text(encoding="utf-8")
    assert 'password = "ZMENTE_ME"' in example


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK", name)
