"""Společná datová vrstva Decision Engine (etapa E1).

Jeden datový základ pro pronájem strojů i vlastní zakázky: typy strojů, jednotlivé kusy
flotily, údržba, rezervace pronájmů, poptávka po pronájmu, zakázky a jejich požadavky,
obsluha a společná nastavení. Data ve složce `data/model/` jsou FIKTIVNÍ ukázka.

Modul data jen načte a zkontroluje. Výpočty aplikace dostávají data přes převodní
funkce `legacy_machines`, `legacy_jobs` a `fleet_capacity` ve stejném tvaru jako dříve,
takže se v této etapě žádný výsledek nemění.

Náklady jsou rozdělené do skupin, aby se později nezapočítaly dvakrát:
- vlastnické (ownership_cost_month): splátka, odpis, pojištění – platí se i při nečinnosti,
- provozní fixní (fixed_operating_cost_month, service_cost_month_*): pevné měsíční náklady kusu,
- variabilní (variable_cost_rental_day / variable_cost_job_day): jen za dny použití,
- režie půjčovny (settings: overhead_cost_month): nerozpočítává se na stroje.
Pokud rozdělení fixních nákladů zatím není známé, je celá částka ve sloupci
fixed_cost_month_unsplit a sloupce vlastnických a provozních fixních nákladů jsou prázdné.
"""

import hashlib
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

DEFAULT_DIR = Path(__file__).parent / "data" / "model"

# ---------------------------------------------------------------------------
# Schéma: sloupec → (typ, povinný). Typy: "id", "text", "int", "num", "prob", "bool", "enum:…"
# ---------------------------------------------------------------------------
STATUS_JOB = "enum:confirmed,tentative,opportunity"
STATUS_BOOKING = "enum:confirmed,tentative"
SCHEMA = {
    "machine_types": {
        "type_id": ("id", True), "name": ("text", True), "role": ("text", False),
        "rental_price_day": ("num", True), "purchase_price": ("num", True),
        "ownership_cost_month": ("num", False), "fixed_operating_cost_month": ("num", False),
        "fixed_cost_month_unsplit": ("num", False),
        "service_cost_month_mean": ("num", True), "service_cost_month_std": ("num", True),
        "variable_cost_rental_day": ("num", False), "variable_cost_job_day": ("num", False),
        "failure_prob": ("prob", True), "failure_days_mean": ("num", True),
        "rental_with_operator": ("bool", False),
    },
    "fleet_units": {
        "unit_id": ("id", True), "type_id": ("id", True), "status": ("enum:owned,leased,ordered", True),
        "available_from_day": ("int", False), "available_to_day": ("int", False),
    },
    "maintenance": {
        "unit_id": ("id", True), "start_day": ("int", True), "days": ("int", True), "note": ("text", False),
    },
    "rental_bookings": {
        "booking_id": ("id", True), "type_id": ("id", True), "start_day": ("int", True),
        "end_day": ("int", True), "units": ("int", True), "price_day": ("num", True),
        "with_operator": ("bool", False), "status": (STATUS_BOOKING, True),
    },
    "rental_demand": {
        "type_id": ("id", True), "period": ("text", True), "expected_days": ("num", True),
        "std_days": ("num", False), "source": ("text", True),
    },
    "jobs": {
        "job_id": ("id", True), "name": ("text", True), "job_type": ("text", False),
        "status": (STATUS_JOB, True), "p_win": ("prob", False),
        "priority": ("enum:High,Medium,Low", True), "risk_level": ("enum:High,Medium,Low", True),
        "revenue": ("num", True), "direct_cost": ("num", True), "duration_days": ("int", True),
        "earliest_start_day": ("int", True), "deadline_day": ("int", True),
        "penalty_per_day": ("num", True),
    },
    "job_requirements": {
        "job_id": ("id", True), "resource_kind": ("enum:machine,operator", True),
        "resource_id": ("id", True), "quantity": ("int", True), "position": ("int", True),
    },
    "operator_roles": {
        "role_id": ("id", True), "name": ("text", True), "headcount": ("int", False),
        "daily_cost": ("num", False), "machine_types": ("text", False),
    },
    "settings": {
        "key": ("id", True), "value": ("text", False), "unit": ("text", False), "note": ("text", False),
    },
}
PRIMARY_KEYS = {
    "machine_types": ["type_id"], "fleet_units": ["unit_id"], "rental_bookings": ["booking_id"],
    "rental_demand": ["type_id", "period"], "jobs": ["job_id"],
    "job_requirements": ["job_id", "resource_kind", "resource_id"], "operator_roles": ["role_id"],
    "settings": ["key"],
}


@dataclass
class DataModel:
    machine_types: pd.DataFrame
    fleet_units: pd.DataFrame
    maintenance: pd.DataFrame
    rental_bookings: pd.DataFrame
    rental_demand: pd.DataFrame
    jobs: pd.DataFrame
    job_requirements: pd.DataFrame
    operator_roles: pd.DataFrame
    settings: pd.DataFrame


@dataclass
class Issue:
    severity: str   # "chyba" nebo "upozornění"
    table: str
    message: str

    def __str__(self) -> str:
        return f"[{self.severity}] {self.table}: {self.message}"


# ---------------------------------------------------------------------------
# Načtení
# ---------------------------------------------------------------------------

def data_fingerprint(directory=DEFAULT_DIR) -> str:
    """Otisk obsahu všech CSV souborů dat (SHA-256). Změní se při jakékoli změně dat.

    Aplikace ho používá jako klíč mezipaměti: po nasazení nových dat se načtou znovu i bez
    restartu běžícího procesu. Soubory jsou malé, výpočet trvá zlomky milisekundy.
    """
    digest = hashlib.sha256()
    for path in sorted(Path(directory).glob("*.csv")):
        digest.update(path.name.encode("utf-8") + b"\0" + path.read_bytes() + b"\0")
    return digest.hexdigest()


def files_fingerprint(paths) -> str:
    """Otisk obsahu zadaných souborů (SHA-256) – např. zdrojového kódu modulů.

    Mezipaměť Streamlitu (st.cache_data) pozná změnu jen u kódu samotné uložené funkce, ne
    u modulů, které volá. Otisk jejich zdrojového kódu v klíči zajistí, že se po nasazení
    nového kódu výsledky spočítají znovu i bez restartu běžícího procesu.
    """
    digest = hashlib.sha256()
    for path in sorted(str(p) for p in paths):
        digest.update(Path(path).name.encode("utf-8") + b"\0" + Path(path).read_bytes() + b"\0")
    return digest.hexdigest()


def load_model(directory=DEFAULT_DIR) -> DataModel:
    """Načte všechny tabulky jako text (typy kontroluje až validate)."""
    directory = Path(directory)
    tables = {}
    for name in SCHEMA:
        path = directory / f"{name}.csv"
        tables[name] = pd.read_csv(path, dtype=str, keep_default_na=False) if path.exists() else None
    missing = [n for n, t in tables.items() if t is None]
    if missing:
        raise FileNotFoundError("Chybí datové soubory: " + ", ".join(f"{m}.csv" for m in missing))
    return DataModel(**tables)


def _num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series.replace("", None), errors="coerce")


def _filled(series: pd.Series) -> pd.Series:
    return series.astype(str).str.strip() != ""


# ---------------------------------------------------------------------------
# Kontrola
# ---------------------------------------------------------------------------

def validate(model: DataModel) -> list:
    issues = []

    def err(table, msg):
        issues.append(Issue("chyba", table, msg))

    def warn(table, msg):
        issues.append(Issue("upozornění", table, msg))

    # 1) Sloupce, povinné hodnoty, typy a rozsahy
    for table, cols in SCHEMA.items():
        df = getattr(model, table)
        missing = [c for c in cols if c not in df.columns]
        if missing:
            err(table, "chybí sloupce " + ", ".join(missing))
            continue
        extra = [c for c in df.columns if c not in cols]
        if extra:
            warn(table, "neznámé sloupce " + ", ".join(extra))
        for col, (kind, required) in cols.items():
            s = df[col].astype(str).str.strip()
            filled = s != ""
            if required and (~filled).any():
                err(table, f"sloupec {col}: prázdná povinná hodnota v řádcích {list(df.index[~filled] + 2)}")
            vals = s[filled]
            if kind in ("int", "num", "prob"):
                nums = pd.to_numeric(vals, errors="coerce")
                bad = nums.isna()
                if bad.any():
                    err(table, f"sloupec {col}: není číslo ({', '.join(vals[bad].head(3))})")
                if kind == "int" and (nums.dropna() % 1 != 0).any():
                    err(table, f"sloupec {col}: musí být celé číslo")
                if (nums.dropna() < 0).any():
                    err(table, f"sloupec {col}: záporná hodnota")
                if kind == "prob" and (nums.dropna() > 1).any():
                    err(table, f"sloupec {col}: pravděpodobnost musí být 0–1")
            elif kind == "bool":
                bad = ~vals.str.lower().isin(["true", "false", "1", "0", "ano", "ne"])
                if bad.any():
                    err(table, f"sloupec {col}: očekáváno ano/ne")
            elif kind.startswith("enum:"):
                allowed = kind[5:].split(",")
                bad = ~vals.isin(allowed)
                if bad.any():
                    err(table, f"sloupec {col}: nepovolená hodnota {sorted(set(vals[bad]))}, povoleno {allowed}")
        # Jedinečnost klíčů
        keys = PRIMARY_KEYS.get(table)
        if keys and len(df):
            dup = df.duplicated(subset=keys, keep=False)
            if dup.any():
                err(table, f"duplicitní klíč {keys}: {df.loc[dup, keys].drop_duplicates().values.tolist()}")
    if any(i.severity == "chyba" and "chybí sloupce" in i.message for i in issues):
        return issues

    mt, fleet, jobs, req = model.machine_types, model.fleet_units, model.jobs, model.job_requirements
    type_ids = set(mt["type_id"])
    unit_ids = set(fleet["unit_id"])
    role_ids = set(model.operator_roles["role_id"])

    # 2) Jednoznačné názvy strojů
    if mt["name"].str.strip().str.lower().duplicated().any():
        err("machine_types", "názvy strojů nejsou jedinečné")

    # 3) Vazby mezi tabulkami
    def fk(table, df, col, valid, target):
        bad = sorted(set(df[col]) - valid)
        if bad:
            err(table, f"{col} odkazuje na neexistující {target}: {bad}")

    fk("fleet_units", fleet, "type_id", type_ids, "typ stroje")
    fk("maintenance", model.maintenance, "unit_id", unit_ids, "kus flotily")
    fk("rental_bookings", model.rental_bookings, "type_id", type_ids, "typ stroje")
    fk("rental_demand", model.rental_demand, "type_id", type_ids, "typ stroje")
    fk("job_requirements", req, "job_id", set(jobs["job_id"]), "zakázku")
    fk("job_requirements", req[req["resource_kind"] == "machine"], "resource_id", type_ids, "typ stroje")
    fk("job_requirements", req[req["resource_kind"] == "operator"], "resource_id", role_ids, "roli obsluhy")
    for roles in model.operator_roles["machine_types"]:
        for t in [x.strip() for x in str(roles).split(";") if x.strip()]:
            if t not in type_ids:
                err("operator_roles", f"machine_types odkazuje na neexistující typ stroje: {t}")

    # 4) Úplnost: každá zakázka má aspoň jeden stroj, každý typ aspoň jednu poptávku
    no_req = sorted(set(jobs["job_id"]) - set(req.loc[req["resource_kind"] == "machine", "job_id"]))
    if no_req:
        err("jobs", f"zakázky bez požadovaného stroje: {no_req}")
    without_units = sorted(type_ids - set(fleet["type_id"]))
    if without_units:
        warn("fleet_units", f"typy strojů bez kusu ve flotile: {without_units}")

    # 5) Náklady – žádné dvojí započtení fixních nákladů
    unsplit = _filled(mt["fixed_cost_month_unsplit"])
    split = _filled(mt["ownership_cost_month"]) | _filled(mt["fixed_operating_cost_month"])
    both = mt.loc[unsplit & split, "type_id"].tolist()
    if both:
        err("machine_types", f"fixní náklady zadané nerozdělené i rozdělené zároveň (dvojí započtení): {both}")
    none = mt.loc[~unsplit & ~split, "type_id"].tolist()
    if none:
        err("machine_types", f"chybí fixní náklady: {none}")
    if unsplit.any():
        warn("machine_types", "rozdělení fixních nákladů na vlastnické a provozní zatím není zadané "
             f"({int(unsplit.sum())} typů) – doplní majitel")

    # 6) Časové údaje a kapacita
    b = model.rental_bookings
    if len(b):
        bad = _num(b["end_day"]) < _num(b["start_day"])
        if bad.any():
            err("rental_bookings", f"konec před začátkem: {b.loc[bad, 'booking_id'].tolist()}")
        units = fleet.groupby("type_id").size()
        conf = b[b["status"] == "confirmed"]
        for t, grp in conf.groupby("type_id"):
            days = {}
            for r in grp.itertuples():
                for d in range(int(r.start_day), int(r.end_day) + 1):
                    days[d] = days.get(d, 0) + int(r.units)
            over = [d for d, n in days.items() if n > units.get(t, 0)]
            if over:
                err("rental_bookings", f"potvrzené rezervace {t} překračují počet kusů ve dnech {over[:5]}")
    j = jobs
    dl, dur, es = _num(j["deadline_day"]), _num(j["duration_days"]), _num(j["earliest_start_day"])
    if (dur < 1).any():
        err("jobs", "délka zakázky musí být alespoň 1 den")
    if (es < 1).any():
        err("jobs", "nejdřívější začátek musí být den 1 nebo později")
    tight = j.loc[es + dur - 1 > dl, "job_id"].tolist()
    if tight:
        warn("jobs", f"zakázky nelze dokončit v termínu ani při okamžitém začátku: {tight}")
    tent = j[(j["status"] != "confirmed") & ~_filled(j["p_win"])]
    if len(tent):
        warn("jobs", f"nepotvrzené zakázky bez pravděpodobnosti získání: {tent['job_id'].tolist()}")
    m = model.maintenance
    if len(m) and (_num(m["days"]) < 1).any():
        err("maintenance", "údržba musí trvat alespoň 1 den")
    return issues


def errors(issues: list) -> list:
    return [i for i in issues if i.severity == "chyba"]


# ---------------------------------------------------------------------------
# Převod do dosavadního tvaru (zpětná kompatibilita aplikace)
# ---------------------------------------------------------------------------

def legacy_machines(model: DataModel) -> pd.DataFrame:
    """Tabulka strojů ve tvaru dřívějšího machines_demo.csv."""
    mt = model.machine_types
    demand = model.rental_demand.drop_duplicates("type_id").set_index("type_id")
    fixed = _num(mt["fixed_cost_month_unsplit"]).where(
        _filled(mt["fixed_cost_month_unsplit"]),
        _num(mt["ownership_cost_month"]).fillna(0) + _num(mt["fixed_operating_cost_month"]).fillna(0),
    )
    out = pd.DataFrame({
        "machine_type": mt["name"].values,
        "daily_price": _num(mt["rental_price_day"]).values,
        "purchase_price": _num(mt["purchase_price"]).values,
        "monthly_fixed_cost": fixed.values,
        "service_cost_mean": _num(mt["service_cost_month_mean"]).values,
        "service_cost_std": _num(mt["service_cost_month_std"]).values,
        "demand_days_mean": _num(demand.loc[mt["type_id"], "expected_days"]).values,
        "demand_days_std": _num(demand.loc[mt["type_id"], "std_days"]).values,
        "failure_prob": _num(mt["failure_prob"]).astype(float).values,
        "failure_days_mean": _num(mt["failure_days_mean"]).values,
        "role": mt["role"].values,
    })
    return _ints(out, exclude=("failure_prob",))


def legacy_jobs(model: DataModel) -> pd.DataFrame:
    """Tabulka zakázek ve tvaru dřívějšího jobs_demo.csv (stroje spojené středníkem)."""
    names = model.machine_types.set_index("type_id")["name"]
    req = model.job_requirements
    req = req[req["resource_kind"] == "machine"].assign(_pos=_num(req["position"]))
    machines = (req.sort_values(["job_id", "_pos"]).groupby("job_id")["resource_id"]
                .apply(lambda ids: ";".join(names[i] for i in ids)))
    j = model.jobs
    out = pd.DataFrame({
        "job_id": j["job_id"].values, "job_name": j["name"].values, "job_type": j["job_type"].values,
        "priority": j["priority"].values, "duration_days": _num(j["duration_days"]).values,
        "expected_revenue": _num(j["revenue"]).values, "cost_base": _num(j["direct_cost"]).values,
        "risk_level": j["risk_level"].values, "required_machines": machines.loc[j["job_id"]].values,
        "deadline_days": _num(j["deadline_day"]).values,
        "delay_penalty_per_day": _num(j["penalty_per_day"]).values,
    })
    return _ints(out)


def fleet_capacity(model: DataModel) -> dict:
    """Počet kusů každého typu stroje (název → počet) v pořadí tabulky typů strojů."""
    counts = model.fleet_units.groupby("type_id").size()
    return {r.name: int(counts.get(r.type_id, 0)) for r in model.machine_types.itertuples()}


def _ints(df: pd.DataFrame, exclude=()) -> pd.DataFrame:
    for c in df.columns:
        if c not in exclude and pd.api.types.is_numeric_dtype(df[c]) and (df[c].dropna() % 1 == 0).all():
            df[c] = df[c].astype("int64")
    return df
