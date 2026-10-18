"""Přihlašování uživatelů Decision Engine. Čisté funkce bez Streamlitu (kvůli testům).

Dva režimy podle obsahu Streamlit secrets:

1. JEDNO HESLO (dosavadní stav): v secrets je jen `password = "…"`. Aplikace se chová přesně
   jako dříve – zadá se jen heslo a přihlášený má roli administrátora.
2. UŽIVATELSKÉ ÚČTY: v secrets je sekce `[users]`. Každý účet má vlastní jméno a heslo:

       [users.honza]
       name = "Honza Podroužek"
       role = "tester"                      # admin | tester | owner
       active = true
       password_hash = "pbkdf2_sha256$600000$…$…"

   Účet může místo `password_hash` použít `legacy_password = true` – pak se ověřuje stávajícím
   heslem `password` z nejvyšší úrovně secrets (pro bezpečný přechod správce bez změny hesla).

Hesla se v aplikaci ani v repozitáři neukládají. V secrets je jen jednosměrný hash
(PBKDF2-HMAC-SHA256 se solí), ze kterého heslo nelze zpětně získat.
"""

import base64
import hashlib
import hmac
import secrets as pysecrets
from dataclasses import dataclass

ALGORITHM = "pbkdf2_sha256"
ITERATIONS = 600_000            # doporučení OWASP pro PBKDF2-HMAC-SHA256
ROLES = {
    "admin": "administrátor",
    "tester": "mentor / tester",
    "owner": "majitel",
}
MIN_PASSWORD_LENGTH = 10
GENERIC_ERROR = "Nesprávné uživatelské jméno nebo heslo."


@dataclass(frozen=True)
class User:
    username: str
    name: str
    role: str
    active: bool
    password_hash: str = ""
    legacy_password: bool = False

    @property
    def role_label(self) -> str:
        return ROLES.get(self.role, self.role)

    @property
    def is_admin(self) -> bool:
        return self.role == "admin"


# ---------------------------------------------------------------------------
# Hash hesla
# ---------------------------------------------------------------------------

def _b64(data: bytes) -> str:
    return base64.b64encode(data).decode("ascii")


def hash_password(password: str, iterations: int = ITERATIONS, salt: bytes = None) -> str:
    """Vytvoří hash hesla ve tvaru pbkdf2_sha256$<iterace>$<sůl>$<hash>."""
    if not password:
        raise ValueError("Heslo nesmí být prázdné.")
    salt = salt or pysecrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, iterations)
    return f"{ALGORITHM}${iterations}${_b64(salt)}${_b64(digest)}"


def verify_password(password: str, stored: str) -> bool:
    """Ověří heslo proti uloženému hashi. Neplatný formát hashe = neověřeno."""
    try:
        algorithm, iterations, salt, digest = str(stored).strip().split("$")
        if algorithm != ALGORITHM:
            return False
        iterations = int(iterations)
        if iterations < 100_000:
            return False
        salt, digest = base64.b64decode(salt), base64.b64decode(digest)
    except (ValueError, TypeError):
        return False
    candidate = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, iterations)
    return hmac.compare_digest(candidate, digest)


def is_valid_hash(stored: str) -> bool:
    try:
        algorithm, iterations, salt, digest = str(stored).strip().split("$")
        return (algorithm == ALGORITHM and int(iterations) >= 100_000
                and len(base64.b64decode(salt)) >= 8 and len(base64.b64decode(digest)) == 32)
    except (ValueError, TypeError):
        return False


def password_problems(password: str) -> list:
    """Základní kontrola síly nového hesla (pro nástroj správce)."""
    problems = []
    if len(password) < MIN_PASSWORD_LENGTH:
        problems.append(f"heslo musí mít alespoň {MIN_PASSWORD_LENGTH} znaků")
    if password.strip() != password:
        problems.append("heslo nesmí začínat ani končit mezerou")
    return problems


def _dummy_check(password: str) -> None:
    """Stejně náročné ověření jako u existujícího účtu – neprozradí, zda jméno existuje.

    Počítá se až při přihlášení neznámým jménem (ne při každém načtení modulu).
    """
    verify_password(password, hash_password("decision-engine-dummy", salt=b"0123456789abcdef"))


# ---------------------------------------------------------------------------
# Konfigurace ze secrets
# ---------------------------------------------------------------------------

def _to_dict(value) -> dict:
    """Převede sekci secrets (AttrDict apod.) na obyčejný slovník."""
    try:
        return {k: (_to_dict(v) if hasattr(v, "keys") else v) for k, v in dict(value).items()}
    except (TypeError, ValueError):
        return {}


def _as_bool(value, default: bool) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return default
    return str(value).strip().lower() in ("true", "1", "ano", "yes")


def normalize_username(username: str) -> str:
    return str(username or "").strip().lower()


def load_config(secrets) -> dict:
    """Načte přihlašovací konfiguraci ze secrets (mapování).

    Vrací {"mode": "users" | "single" | "missing", "users": {jméno: User},
           "legacy_password": str, "problems": [popisy chyb konfigurace]}.
    """
    try:
        raw = _to_dict(secrets)
    except Exception:
        raw = {}
    legacy = str(raw.get("password", "") or "").strip()
    users_raw = raw.get("users")
    problems = []
    if not isinstance(users_raw, dict) or not users_raw:
        return {"mode": "single" if legacy else "missing", "users": {}, "legacy_password": legacy,
                "problems": problems}

    users = {}
    for key, cfg in users_raw.items():
        username = normalize_username(key)
        if not isinstance(cfg, dict):
            problems.append(f"účet {username}: chybná struktura")
            continue
        role = str(cfg.get("role", "")).strip().lower()
        if role not in ROLES:
            problems.append(f"účet {username}: neznámá role „{role}“")
            continue
        legacy_flag = _as_bool(cfg.get("legacy_password"), False)
        pw_hash = str(cfg.get("password_hash", "") or "").strip()
        if legacy_flag and not legacy:
            problems.append(f"účet {username}: legacy_password vyžaduje klíč password")
        if not legacy_flag and not is_valid_hash(pw_hash):
            problems.append(f"účet {username}: chybí platný password_hash")
        users[username] = User(
            username=username,
            name=str(cfg.get("name", username)).strip() or username,
            role=role,
            active=_as_bool(cfg.get("active"), False),     # bez výslovného active = true je účet vypnutý
            password_hash=pw_hash,
            legacy_password=legacy_flag,
        )
    return {"mode": "users", "users": users, "legacy_password": legacy, "problems": problems}


def can_log_in(config: dict) -> bool:
    if config["mode"] == "single":
        return True
    return any(_usable(u, config) for u in config["users"].values())


def _usable(user: User, config: dict) -> bool:
    if not user.active:
        return False
    if user.legacy_password:
        return bool(config["legacy_password"])
    return is_valid_hash(user.password_hash)


# ---------------------------------------------------------------------------
# Ověření přihlášení
# ---------------------------------------------------------------------------

SINGLE_USER = User(username="admin", name="Administrátor", role="admin", active=True)


def authenticate(config: dict, username: str, password: str):
    """Vrátí User při úspěšném přihlášení, jinak None. Neprozrazuje, proč se nepodařilo."""
    password = (password or "").strip()
    if not password:
        return None
    if config["mode"] == "single":
        ok = hmac.compare_digest(password.encode("utf-8"), config["legacy_password"].encode("utf-8"))
        return SINGLE_USER if ok else None
    if config["mode"] != "users":
        return None
    user = config["users"].get(normalize_username(username))
    if user is None:
        _dummy_check(password)                           # stejná doba odpovědi jako u existujícího
        return None
    if user.legacy_password:
        ok = bool(config["legacy_password"]) and hmac.compare_digest(
            password.encode("utf-8"), config["legacy_password"].encode("utf-8"))
    else:
        ok = verify_password(password, user.password_hash)
    return user if ok and _usable(user, config) else None


def still_allowed(config: dict, username: str) -> bool:
    """Platí přihlášení i po změně secrets? (deaktivace účtu se projeví při dalším kliknutí)"""
    if config["mode"] == "single":
        return username == SINGLE_USER.username
    user = config["users"].get(normalize_username(username))
    return user is not None and _usable(user, config)
