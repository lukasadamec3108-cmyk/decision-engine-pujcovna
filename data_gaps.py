"""Přehled chybějících údajů (etapa E3.2) – vyhodnocení bez Streamlitu, jen pro čtení.

Projde datový model E1 a najde údaje, které chybějí nebo jsou neplatné a které skutečně používají
už implementované výpočty (economics.py, plánovač, kapacita, investice). Nic nedoplňuje, neodhaduje
ani neukládá. Zobrazení (texty, seskupení) je v data_gaps_view.py.

PRAVIDLA STAVŮ (stejná jako economics.py a capacity.py – nic nového)
  chybí     … prázdná buňka, chybějící řádek nastavení nebo chybějící sloupec / záznam,
  neplatný  … text místo čísla, záporná částka, pravděpodobnost mimo 0–1, celé číslo menší než 1,
              nepovolená hodnota nastavení, duplicitní záznam,
  (OK)      … skutečná nula „0“ je údaj, ne chybějící hodnota – do přehledu se nedostane.
Údaje, pro které datový model nemá pole nebo které žádný výpočet nepoužívá, jsou zvlášť v NOT_TRACKED_ITEMS
(„zatím není evidováno“) a do počtu chybějících údajů se nepočítají.

PRAVIDLA PRIORIT (PRIORITY_RULES – vysvětlitelná, ne libovolná)
  vysoká  … H1 neplatná hodnota nebo duplicitní záznam (výpočet nelze provést nebo je nejednoznačný),
            H2 chybí povinný údaj (bez něj výsledek nelze určit ani jako horní mez),
            H3 chybí nastavení, které brání úplnému výsledku u všech zakázek najednou;
  střední … M1 chybí náklad, kvůli kterému je výsledek jen horní mez nebo orientační;
  nízká   … L1 doplňující údaj, který dnešní výpočty nemění.

Každý problém je v seznamu právě jednou (klíč tabulka + sloupec + záznam); duplicitní řádek se hlásí
jednou jako „duplicitní záznam“ a jeho buňky se nevyhodnocují znovu.
"""

import dataclasses
from dataclasses import dataclass, field

import capacity as cp
import datamodel as dm
import economics as ec

MISSING, INVALID, NOT_TRACKED = "chybí", "neplatný", "zatím není evidováno"
HIGH, MEDIUM, LOW = "vysoká", "střední", "nízká"
PRIORITIES = (HIGH, MEDIUM, LOW)

PRIORITY_RULES = {
    "H1": (HIGH, "Neplatná hodnota nebo duplicitní záznam – výpočet nelze provést nebo je nejednoznačný."),
    "H2": (HIGH, "Chybí povinný údaj – bez něj nelze výsledek určit ani jako horní mez."),
    "H3": (HIGH, "Chybí nastavení, které brání úplnému výsledku u všech zakázek najednou."),
    "M1": (MEDIUM, "Chybí náklad – výsledek je bez něj jen horní mez nebo orientační."),
    "L1": (LOW, "Doplňující údaj – dnešní výpočty nemění, zpřesní evidenci."),
}

# Oblasti výpočtů: klíč → (název, kde v aplikaci)
AREAS = {
    "zakazky": ("Příspěvek vlastních zakázek", "Ekonomika flotily, Rozhodovací centrum"),
    "pronajmy": ("Příspěvek pronájmů", "Ekonomika flotily, Pronájem strojů, Rozhodovací centrum"),
    "flotila": ("Výsledek flotily za období", "Ekonomika flotily"),
    "srovnani": ("Zakázka, nebo pronájem?", "Ekonomika flotily"),
    "planovac": ("Doporučení rozhodovacího plánovače", "Pronájem strojů, Rozhodovací centrum"),
    "plan": ("Plán zakázek a kapacita", "Vlastní zakázky, Plánování kapacity"),
    "investice": ("Investiční model a nákup strojů", "Investice, Strojový park, Dashboard"),
}

RECORD_KIND = {"machine_types": "stroj", "jobs": "zakázka", "job_requirements": "zakázka",
               "rental_bookings": "pronájem", "maintenance": "servis", "rental_demand": "stroj",
               "settings": "nastavení", "fleet_units": "kus flotily", "operator_roles": "obsluha"}

ACCEPTED_BOOL = ("ano", "true", "1", "yes", "ne", "false", "0", "no")   # = economics.direct_cost_includes_machines


@dataclass(frozen=True)
class Field:
    table: str
    column: str
    kind: str            # money | int | prob
    label: str
    required: bool
    why: str
    areas: tuple
    rule_missing: str    # pravidlo priority, když údaj chybí (neplatný = vždy H1)


MACHINE_FIELDS = (
    Field("machine_types", "rental_price_day", "money", "Ceníková cena pronájmu za den", True,
          "Bez ceny nelze spočítat tržbu z pronájmu.", ("pronajmy", "planovac", "srovnani", "investice"), "H2"),
    Field("machine_types", "purchase_price", "money", "Pořizovací cena stroje", True,
          "Bez ní nelze spočítat návratnost a nákupní sestavu.", ("investice",), "H2"),
    Field("machine_types", "service_cost_month_mean", "money", "Průměrný měsíční servis", True,
          "Je součástí fixních nákladů stroje a flotily.", ("flotila", "investice"), "H2"),
    Field("machine_types", "service_cost_month_std", "money", "Rozptyl měsíčního servisu", True,
          "Simulace investic s ním počítá kolísání nákladů na servis.", ("investice",), "H2"),
    Field("machine_types", "failure_prob", "prob", "Pravděpodobnost poruchy", True,
          "Investiční model s ní odečítá dny výpadku.", ("investice",), "H2"),
    Field("machine_types", "failure_days_mean", "money", "Průměrná délka poruchy (dny)", True,
          "Investiční model s ní odečítá dny výpadku.", ("investice",), "H2"),
    Field("machine_types", "variable_cost_rental_day", "money", "Variabilní náklad na den pronájmu", False,
          "Bez něj se od tržby z pronájmu neodečte náklad za den použití (opotřebení, palivo…) – příspěvek "
          "pronájmu je jen horní mez.", ("pronajmy", "flotila", "planovac", "srovnani"), "M1"),
    Field("machine_types", "variable_cost_job_day", "money", "Variabilní náklad na den práce na vlastní zakázce",
          False, "Bez něj se od příspěvku zakázky neodečte náklad vlastních strojů za dny na zakázce – příspěvek "
          "zakázky je jen horní mez.", ("zakazky", "flotila", "planovac", "srovnani"), "M1"),
)
JOB_FIELDS = (
    Field("jobs", "revenue", "money", "Tržba zakázky", True, "Bez tržby nelze spočítat příspěvek zakázky.",
          ("zakazky", "flotila", "srovnani", "planovac", "plan"), "H2"),
    Field("jobs", "direct_cost", "money", "Přímé náklady zakázky", True,
          "Bez přímých nákladů nelze spočítat příspěvek zakázky.",
          ("zakazky", "flotila", "srovnani", "planovac", "plan"), "H2"),
    Field("jobs", "penalty_per_day", "money", "Penále za den zpoždění", True,
          "Bez něj nelze ocenit zpoždění v plánu zakázek.", ("zakazky", "planovac", "plan"), "H2"),
    Field("jobs", "duration_days", "int", "Délka zakázky (dny)", True,
          "Bez délky nelze zakázku naplánovat ani spočítat strojodny.", ("plan", "zakazky", "srovnani", "planovac"),
          "H2"),
    Field("jobs", "earliest_start_day", "int", "Nejdřívější začátek zakázky", True,
          "Bez něj nelze zakázku naplánovat.", ("plan", "planovac"), "H2"),
    Field("jobs", "deadline_day", "int", "Termín dokončení zakázky", True,
          "Bez termínu nelze určit zpoždění a penále.", ("plan", "zakazky", "planovac"), "H2"),
)
BOOKING_FIELDS = (
    Field("rental_bookings", "price_day", "money", "Cena pronájmu za den v rezervaci", True,
          "Bez ceny nelze spočítat tržbu z rezervace.", ("pronajmy", "flotila", "planovac"), "H2"),
    Field("rental_bookings", "start_day", "int", "Začátek pronájmu", True,
          "Bez termínu nelze rezervaci zapsat do kapacity.", ("plan", "pronajmy", "planovac"), "H2"),
    Field("rental_bookings", "end_day", "int", "Konec pronájmu", True,
          "Bez termínu nelze rezervaci zapsat do kapacity.", ("plan", "pronajmy", "planovac"), "H2"),
    Field("rental_bookings", "units", "int", "Počet pronajatých kusů", True,
          "Bez počtu kusů nelze spočítat kapacitu ani tržbu.", ("plan", "pronajmy", "planovac"), "H2"),
)
MAINTENANCE_FIELDS = (
    Field("maintenance", "start_day", "int", "Začátek servisu", True,
          "Bez začátku nelze servis zapsat do kapacity.", ("plan",), "H2"),
    Field("maintenance", "days", "int", "Délka servisu (dny)", True,
          "Bez délky nelze servis zapsat do kapacity.", ("plan",), "H2"),
)
REQUIREMENT_FIELD = Field("job_requirements", "quantity", "int", "Počet kusů stroje pro zakázku", True,
                          "Bez počtu kusů nelze spočítat strojodny ani kapacitu zakázky.",
                          ("plan", "zakazky", "srovnani", "planovac"), "H2")
DEMAND_FIELD = Field("rental_demand", "expected_days", "money", "Modelová poptávka po pronájmu (dny za měsíc)", True,
                     "Investiční model a nákupní sestava z ní odhadují počet pronajatých dní.", ("investice",), "H2")
SETTING_MACHINES = Field("settings", "direct_cost_includes_machines", "bool",
                         "Zda přímé náklady zakázek už obsahují náklady vlastních strojů", False,
                         "Bez tohoto údaje nelze rozhodnout, zda od přímých nákladů zakázky odečíst ještě náklady "
                         "vlastních strojů – hrozí dvojí započtení nebo vynechání. Příspěvek žádné zakázky proto "
                         "nemůže být úplný a ani doplněné variabilní náklady na zakázce se nepoužijí.",
                         ("zakazky", "flotila", "planovac", "srovnani"), "H3")
SETTING_OVERHEAD = Field("settings", "overhead_cost_month", "money", "Společná měsíční režie půjčovny", False,
                         "Bez režie je celkový výsledek flotily jen horní mez (režie se neodečte).", ("flotila",),
                         "M1")
FIXED_SPLIT = Field("machine_types", "ownership_cost_month", "money",
                    "Rozdělení fixních nákladů na vlastnické a provozní", False,
                    "Celkové fixní náklady stroje jsou zadané; rozdělení na vlastnické (splátka, odpis, pojištění) "
                    "a provozní dnes žádný výpočet nemění – zpřesní jen evidenci.", (), "L1")
FIXED_TOTAL = Field("machine_types", "fixed_cost_month_unsplit", "money", "Měsíční fixní náklady stroje", True,
                    "Bez fixních nákladů nelze spočítat náklady flotily ani investiční model.",
                    ("flotila", "investice"), "H2")

# Údaje, které datový model nemá nebo které žádný výpočet nepoužívá (ověřeno v kódu – viz testy).
NOT_TRACKED_ITEMS = (
    dict(key="doprava", label="Náklady na dopravu stroje", field_exists=False,
         detail="Datový model nemá pole pro dopravu stroje k zákazníkovi nebo na zakázku.",
         would_affect=("pronajmy", "zakazky")),
    dict(key="obsluha", label="Náklady na obsluhu (strojníka)", field_exists=True,
         detail="Tabulka rolí obsluhy (operator_roles, denní náklad) i příznak pronájmu s obsluhou v datech jsou, "
                "ale žádný výpočet je zatím nepoužívá.", would_affect=("pronajmy", "zakazky")),
    dict(key="poptavka", label="Skutečná poptávka po pronájmu", field_exists=False,
         detail="Evidovaná je jen modelová poptávka „typický měsíc“ pro investiční model; skutečné nebo odmítnuté "
                "poptávky datový model nemá. Srovnání „Zakázka, nebo pronájem?“ proto poptávku nepředpokládá.",
         would_affect=("srovnani", "investice")),
    dict(key="kusy", label="Přiřazení rezervací a zakázek ke konkrétním kusům", field_exists=False,
         detail="Rezervace a zakázky uvádějí jen typ stroje, ne konkrétní kus.", would_affect=("plan",)),
    dict(key="nahrada", label="Smluvní náhrada při zrušení nebo zkrácení pronájmu", field_exists=False,
         detail="Bez ní jsou varianty řešení konfliktů s uvolněním pronájmu jen podmíněné.",
         would_affect=("flotila",)),
    dict(key="p_win", label="Pravděpodobnost získání nepotvrzené zakázky", field_exists=True,
         detail="Pole v datech je, ale žádný výpočet ho zatím nepoužívá.", would_affect=()),
    dict(key="kalendar", label="Pracovní dny v týdnu a horizont plánování", field_exists=True,
         detail="Řádky v nastavení dat jsou, ale výpočty zatím počítají kalendářní dny a hodnoty nepoužívají.",
         would_affect=()),
)


@dataclass(frozen=True)
class Gap:
    """Jeden chybějící nebo neplatný údaj u jednoho záznamu."""
    key: tuple                # (tabulka, sloupec, záznam) – jedinečný
    label: str
    table: str
    column: str
    record_kind: str
    record_id: str
    record_name: str
    status: str               # MISSING | INVALID
    value: str                # původní text neplatné hodnoty ("" u chybějící)
    why: str
    areas: tuple
    rule: str
    priority: str
    where: str


@dataclass
class Report:
    gaps: list = field(default_factory=list)
    not_tracked: list = field(default_factory=list)
    fictitious: bool = False
    jobs_total: int = 0
    jobs_incomplete: list = field(default_factory=list)   # zakázky s neúplným příspěvkem (economics.py)
    notes: list = field(default_factory=list)

    @property
    def summary(self) -> dict:
        machines = sorted({g.record_id for g in self.gaps if g.record_kind == "stroj"})
        direct_jobs = {g.record_id for g in self.gaps if g.record_kind == "zakázka"}
        jobs = sorted(direct_jobs | set(self.jobs_incomplete))
        areas = [a for a in AREAS if any(a in g.areas for g in self.gaps)]
        return dict(
            total=len(self.gaps),
            missing=sum(g.status == MISSING for g in self.gaps),
            invalid=sum(g.status == INVALID for g in self.gaps),
            by_priority={p: sum(g.priority == p for g in self.gaps) for p in PRIORITIES},
            machines=machines, jobs=jobs, jobs_total=self.jobs_total, areas=areas,
            not_tracked=len(self.not_tracked),
        )


def _where(table, column, record=None) -> str:
    row = f", řádek {record}" if record else ""
    return f"datový soubor data/model/{table}.csv, sloupec {column}{row} (aplikace zatím nemá formulář pro úpravy)"


def _check(raw, kind) -> tuple:
    """(stav, text hodnoty) podle pravidel economics._amount / capacity._int; (None, "") = údaj je v pořádku."""
    text = "" if raw is None else str(raw).strip()
    if kind == "int":
        n = cp._int(raw)
        if n is None:
            return MISSING, ""
        return (None, "") if isinstance(n, int) and n >= 1 else (INVALID, text)
    res = ec._amount(raw, "", False)
    if res.missing:
        return MISSING, ""
    if res.invalid or (kind == "prob" and res.value > 1):
        return INVALID, text
    return None, ""


def _gap(f: Field, status, value, record_id, record_name, rule=None, label=None, why=None, where=None) -> Gap:
    rule = "H1" if status == INVALID else (rule or f.rule_missing)
    return Gap(key=(f.table, f.column, record_id), label=label or f.label, table=f.table, column=f.column,
               record_kind=RECORD_KIND[f.table], record_id=record_id, record_name=record_name, status=status,
               value=value, why=why or f.why, areas=f.areas, rule=rule, priority=PRIORITY_RULES[rule][0],
               where=where or _where(f.table, f.column))


class _Collector:
    def __init__(self):
        self.gaps, self.keys = [], set()

    def add(self, gap: Gap):
        if gap.key not in self.keys:                 # každý problém jednou
            self.keys.add(gap.key)
            self.gaps.append(gap)


def _missing_columns(model, out: _Collector) -> set:
    """Chybějící sloupce, které výpočty používají – jedna položka za sloupec (záznam = celá tabulka)."""
    used = {}
    for f in (*MACHINE_FIELDS, *JOB_FIELDS, *BOOKING_FIELDS, *MAINTENANCE_FIELDS, REQUIREMENT_FIELD, DEMAND_FIELD,
              FIXED_TOTAL):
        used.setdefault(f.table, []).append(f)
    gone = set()
    for table, fields in used.items():
        df = getattr(model, table)
        for f in fields:
            if f.column not in df.columns:
                gone.add((table, f.column))
                out.add(_gap(f, MISSING, "", "celá tabulka", f"sloupec {f.column} v datech chybí", rule="H2",
                             why=f.why + " Sloupec v datech chybí."))
    return gone


def _duplicates(model, out: _Collector) -> dict:
    """Duplicitní klíče (datamodel.PRIMARY_KEYS) – jedna položka za duplicitní klíč; vrací {tabulka: {klíč}}."""
    dup_keys = {}
    for table in ("machine_types", "jobs", "rental_bookings", "settings", "rental_demand"):
        df = getattr(model, table)
        keys = [k for k in dm.PRIMARY_KEYS[table] if k in df.columns]
        if not keys or df.empty:
            continue
        mask = df.duplicated(subset=keys, keep=False)
        for value in df.loc[mask, keys[0]].astype(str).unique():
            dup_keys.setdefault(table, set()).add(value)
            f = Field(table, keys[0], "id", "Duplicitní záznam", True,
                      "Stejný záznam je v datech vícekrát – není jasné, která hodnota platí.", (), "H1")
            out.add(Gap(key=(table, "duplicitní záznam", value), label="Duplicitní záznam", table=table,
                        column=keys[0], record_kind=RECORD_KIND[table], record_id=value, record_name=value,
                        status=INVALID, value=value, why=f.why, areas=_table_areas(table), rule="H1",
                        priority=HIGH, where=_where(table, keys[0])))
    return dup_keys


def _table_areas(table) -> tuple:
    fields = [f for f in (*MACHINE_FIELDS, *JOB_FIELDS, *BOOKING_FIELDS, *MAINTENANCE_FIELDS, DEMAND_FIELD,
                          SETTING_MACHINES, SETTING_OVERHEAD) if f.table == table]
    return tuple(a for a in AREAS if any(a in f.areas for f in fields))


def _first_rows(df, key) -> list:
    """Řádky tabulky – u duplicitního klíče jen první (duplicita se hlásí zvlášť, buňky jen jednou)."""
    if key not in df.columns:
        return []
    return [r for _, r in df.drop_duplicates(subset=[key]).iterrows()]


def _machine_gaps(model, out, gone, includes):
    mt = model.machine_types
    for r in _first_rows(mt, "type_id"):
        tid, name = str(r["type_id"]), str(r.get("name", "") or r["type_id"])
        for f in MACHINE_FIELDS:
            if (f.table, f.column) in gone:
                continue
            if f.column == "variable_cost_job_day" and includes is True:
                continue                                   # stroje jsou v přímých nákladech – údaj se nepoužije
            status, value = _check(r[f.column], f.kind)
            if status:
                out.add(_gap(f, status, value, tid, name))
        # Fixní náklady: celek, nebo rozdělení (economics.unit_fixed_costs)
        cols = ("fixed_cost_month_unsplit", "ownership_cost_month", "fixed_operating_cost_month")
        if not all(c in mt.columns for c in cols):
            continue
        states = {c: _check(r[c], "money") for c in cols}
        bad = [(c, v) for c, (s, v) in states.items() if s == INVALID]
        if bad:
            for c, v in bad:
                f = Field("machine_types", c, "money", FIXED_TOTAL.label, True, FIXED_TOTAL.why, FIXED_TOTAL.areas, "H2")
                out.add(_gap(f, INVALID, v, tid, name))
            continue
        unsplit = states["fixed_cost_month_unsplit"][0] is None
        split = [states[c][0] is None for c in cols[1:]]
        if not unsplit and not any(split):
            out.add(_gap(FIXED_TOTAL, MISSING, "", tid, name,
                         where=_where("machine_types", "fixed_cost_month_unsplit (nebo ownership_cost_month "
                                      "a fixed_operating_cost_month)")))
        elif unsplit and not all(split):
            out.add(_gap(FIXED_SPLIT, MISSING, "", tid, name,
                         where=_where("machine_types", "ownership_cost_month a fixed_operating_cost_month")))


def _job_gaps(model, out, gone):
    jobs, req = model.jobs, model.job_requirements
    for r in _first_rows(jobs, "job_id"):
        jid, name = str(r["job_id"]), str(r.get("name", "") or r["job_id"])
        for f in JOB_FIELDS:
            if (f.table, f.column) in gone:
                continue
            status, value = _check(r[f.column], f.kind)
            if status:
                out.add(_gap(f, status, value, jid, name))
        if {"job_id", "resource_kind", "quantity", "resource_id"} <= set(req.columns):
            machines = req[(req["job_id"] == r["job_id"]) & (req["resource_kind"] == "machine")]
            if machines.empty:
                f = Field("job_requirements", "resource_id", "id", "Stroje potřebné pro zakázku", True,
                          "Zakázka nemá v datech žádný stroj – nelze ji naplánovat ani porovnat s pronájmem.",
                          REQUIREMENT_FIELD.areas, "H2")
                out.add(_gap(f, MISSING, "", jid, name))
            for m in machines.itertuples():
                status, value = _check(m.quantity, "int")
                if status:
                    out.add(Gap(key=("job_requirements", "quantity", f"{jid}/{m.resource_id}"),
                                label=f"{REQUIREMENT_FIELD.label} ({m.resource_id})", table="job_requirements",
                                column="quantity", record_kind="zakázka", record_id=jid, record_name=name,
                                status=status, value=value, why=REQUIREMENT_FIELD.why, areas=REQUIREMENT_FIELD.areas,
                                rule="H1" if status == INVALID else "H2", priority=HIGH,
                                where=_where("job_requirements", "quantity")))


def _row_gaps(model, out, gone, table, key, fields, name_of):
    for r in _first_rows(getattr(model, table), key):
        rid = str(r[key])
        for f in fields:
            if (f.table, f.column) in gone:
                continue
            status, value = _check(r[f.column], f.kind)
            if status:
                out.add(_gap(f, status, value, rid, name_of(r)))


def _demand_gaps(model, out, gone):
    mt, demand = model.machine_types, model.rental_demand
    if "type_id" not in mt.columns or ("rental_demand", "expected_days") in gone or "type_id" not in demand.columns:
        return
    for r in _first_rows(mt, "type_id"):
        tid, name = str(r["type_id"]), str(r.get("name", "") or r["type_id"])
        rows = demand[demand["type_id"] == r["type_id"]]
        if rows.empty:
            out.add(_gap(DEMAND_FIELD, MISSING, "", tid, name,
                         why=DEMAND_FIELD.why + " Pro tento stroj v datech není žádný řádek poptávky."))
            continue
        status, value = _check(rows.iloc[0]["expected_days"], "money")
        if status:
            out.add(_gap(DEMAND_FIELD, status, value, tid, name))


def _setting_gaps(model, out, includes):
    s = model.settings
    present = s[s["key"] == SETTING_MACHINES.column] if "key" in s.columns else s.iloc[0:0]
    raw = ec.setting(model, SETTING_MACHINES.column) if "key" in s.columns else ""
    where = ("datový soubor data/model/settings.csv, řádek s klíčem direct_cost_includes_machines a hodnotou "
             "„ano“ nebo „ne“ (aplikace zatím nemá formulář pro úpravy)")
    if includes is None:
        status = INVALID if raw and raw.lower() not in ACCEPTED_BOOL else MISSING
        why = SETTING_MACHINES.why + ("" if len(present) else " V nastavení dat tento řádek zatím vůbec není.")
        out.add(_gap(SETTING_MACHINES, status, raw if status == INVALID else "", SETTING_MACHINES.column,
                     "nastavení dat", why=why, where=where))
    raw = ec.setting(model, SETTING_OVERHEAD.column) if "key" in s.columns else ""
    status, value = _check(raw, "money")
    if status:
        out.add(_gap(SETTING_OVERHEAD, status, value, SETTING_OVERHEAD.column, "nastavení dat",
                     where="datový soubor data/model/settings.csv, řádek overhead_cost_month (Kč za měsíc; "
                           "skutečná nula „0“ je platný údaj)"))


def _incomplete_jobs(model) -> tuple:
    """Zakázky, jejichž příspěvek economics.job_contribution není úplný (bez penále) – dotčené výsledky."""
    if "job_id" not in model.jobs.columns:
        return 0, []
    ids = list(dict.fromkeys(str(j) for j in model.jobs["job_id"]))
    out = []
    for jid in ids:
        try:
            res = ec.job_contribution(model, jid, 0)
            if not res.complete:
                out.append(jid)
        except (ValueError, TypeError, KeyError):
            out.append(jid)                                  # neplatný údaj – výsledek nelze spočítat
    return len(ids), out


def evaluate(model: dm.DataModel) -> Report:
    """Přehled chybějících a neplatných údajů (jen čtení). Pořadí: priorita, pak tabulka a pořadí v datech."""
    out = _Collector()
    gone = _missing_columns(model, out)
    _duplicates(model, out)
    includes = ec.direct_cost_includes_machines(model) if "key" in model.settings.columns else None
    _setting_gaps(model, out, includes)
    if "type_id" in model.machine_types.columns:
        _machine_gaps(model, out, gone, includes)
        _demand_gaps(model, out, gone)
    if "job_id" in model.jobs.columns:
        _job_gaps(model, out, gone)
    if "booking_id" in model.rental_bookings.columns:
        _row_gaps(model, out, gone, "rental_bookings", "booking_id", BOOKING_FIELDS, lambda r: str(r["booking_id"]))
    if "unit_id" in model.maintenance.columns and len(model.maintenance):
        m = model.maintenance.reset_index(drop=True)
        m = m.assign(_ref=[f"{u} #{i + 1}" for i, u in enumerate(m["unit_id"])])
        _row_gaps(dataclasses.replace(model, maintenance=m), out, gone, "maintenance", "_ref", MAINTENANCE_FIELDS,
                  lambda r: f"servis {r['unit_id']}")
    rank = {p: i for i, p in enumerate(PRIORITIES)}
    gaps = sorted(out.gaps, key=lambda g: rank[g.priority])          # stabilní: v rámci priority pořadí nálezu
    total, incomplete = _incomplete_jobs(model)
    notes = []
    if "type_id" in model.machine_types.columns and model.machine_types.empty:
        notes.append("V datech nejsou žádné stroje.")
    if "job_id" in model.jobs.columns and model.jobs.empty:
        notes.append("V datech nejsou žádné vlastní zakázky.")
    return Report(gaps=gaps, not_tracked=[dict(t, status=NOT_TRACKED) for t in NOT_TRACKED_ITEMS],
                  fictitious=ec.is_fictitious(model), jobs_total=total, jobs_incomplete=incomplete, notes=notes)
