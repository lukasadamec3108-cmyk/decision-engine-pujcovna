"""Jednotné využití fyzických strojů za zvolené období (etapa E2.2 – K2). Bez Streamlitu.

Staví na kapacitním kalendáři z capacity.py (stejné kusy, rezervace, zakázky, servis a stejná
pravidla stavů). Nic nepřesouvá, neruší ani neoptimalizuje a nikde se zatím nezobrazuje.

STROJODNY
---------
- 1 strojoden = 1 fyzický kus na 1 den. Období [start, end] zahrnuje oba krajní dny; dny jsou
  čísla (den 1 = první den plánu) nebo kalendářní data (přechod přes měsíc i rok).
- Dostupná kapacita typu v daný den = kusy ve flotile, které ten den k dispozici jsou
  (objednaný kus až od data dodání). Servis je jedna ze složek kapacity, ne její zmenšení.
- Každý dostupný kus se v každém dni započítá PRÁVĚ JEDNOU do jedné složky:
      pronájem + vlastní zakázka + servis + konflikt + nelze určit + volno = kapacita.
  Nároky, které se do kapacity nevejdou, se vykazují zvlášť jako „nárok nad kapacitu“ –
  do kapacity se nepřičítají a nic se kvůli nim tiše neruší.

STAVY ZÁZNAMŮ (jako capacity.py)
--------------------------------
confirmed … blokuje kapacitu; tentative / opportunity … neblokuje, vykazuje se zvlášť
(include_tentative=True je započte jako blokující); cancelled … ignoruje se; neznámý stav
nebo neplatný termín … chyba dat → volné strojodny typu jsou „nelze určit“.

KONFLIKTY
---------
Přednost pronájmu, zakázky ani servisu se NEURČUJE. Konflikt nastane, když
- konkrétní kus má ve stejný den víc různých využití (dva servisní záznamy téhož kusu se
  nepočítají – kus je v servisu jednou; jde jen o upozornění na data),
- nároky na typ stroje převýší dostupné kusy – pak jsou v konfliktu VŠECHNY dostupné kusy
  typu v tom dni (bez priority nelze říct, které využití ustoupí),
- konkrétní kus je naplánovaný na den, kdy není ve flotile.
Konfliktní strojoden se započítá jednou (složka „konflikt“), přebytek jako nárok nad kapacitu.

KONKRÉTNÍ KUS × TYP STROJE
--------------------------
Rezervace a zakázky v datech E1 uvádějí jen typ stroje. Takový nárok se konkrétnímu kusu
nepřiřazuje: počty strojodnů po typech jsou jisté, ale v tabulce kusů je stav
„nejednoznačné přiřazení“ – kromě případu, kdy jsou obsazené všechny zbývající kusy typu
stejným druhem využití (pak je druh jistý, jen ne konkrétní rezervace).
"""

from dataclasses import dataclass, field
from datetime import date

import pandas as pd

import capacity as cp

RENTAL, JOB, SERVICE, FREE = "pronájem", "vlastní zakázka", "servis", "volno"
CONFLICT, AMBIGUOUS, UNKNOWN, OUT = "konflikt", "nejednoznačné přiřazení", "nelze určit", "mimo flotilu"
KIND_STATE = {"rental": RENTAL, "job": JOB, "maintenance": SERVICE}

COMPLETE, INCOMPLETE, NOT_COMPUTABLE = "úplný", "neúplný", "nelze spočítat"
GROUPS = ["rental", "job", "maintenance", "conflict", "unknown", "free"]
DAY_COLUMNS = (["day", "type_id", "capacity"] + GROUPS
               + ["excess", "tentative", "ambiguous", "capacity_known", "status"])
UNIT_COLUMNS = ["day", "unit_id", "type_id", "in_capacity", "state", "sources", "note"]
TYPE_COLUMNS = (["type_id", "capacity"] + GROUPS
                + ["excess", "tentative", "ambiguous", "conflict_days", "capacity_known",
                   "unit_level", "balanced"])


@dataclass
class Usage:
    """Výsledek machine_days.

    status    … úplný | neúplný (část strojodnů nelze určit nebo chybí počet kusů) | nelze spočítat
    by_unit   … každý fyzický kus × den, právě jeden řádek (UNIT_COLUMNS)
    by_day    … typ stroje × den, strojodny po složkách (DAY_COLUMNS)
    by_type   … součty za období po typech (TYPE_COLUMNS)
    fleet     … součty za typy se známým počtem kusů + seznam vynechaných typů
    conflicts … [{scope: kus|typ, id, type_id, kind, days, sources, excess}]
    notes     … nejistoty a chybějící údaje (texty)
    """
    status: str
    start: object = None
    end: object = None
    by_unit: pd.DataFrame = field(default_factory=lambda: pd.DataFrame(columns=UNIT_COLUMNS))
    by_day: pd.DataFrame = field(default_factory=lambda: pd.DataFrame(columns=DAY_COLUMNS))
    by_type: pd.DataFrame = field(default_factory=lambda: pd.DataFrame(columns=TYPE_COLUMNS))
    fleet: dict = field(default_factory=dict)
    conflicts: list = field(default_factory=list)
    notes: list = field(default_factory=list)


def _period_problem(cal: cp.Calendar, start, end):
    """Text chyby období, nebo None. Nic nevyhazuje – neplatné období = „nelze spočítat“."""
    if start is None or end is None:
        return "chybí začátek nebo konec období"
    if not (cp._is_day(start) and cp._is_day(end)):
        return "neplatné datum období"
    if isinstance(start, date) != isinstance(end, date):
        return "období míchá čísla dní a kalendářní data"
    if cal.day_kind is not None and cp._kind(start) is not cal.day_kind:
        return "období používá jiný druh dnů (čísla × data) než evidence"
    if end < start:
        return "konec období je před začátkem"
    return None


def machine_days(cal: cp.Calendar, start, end, include_tentative: bool = False) -> Usage:
    """Využití fyzických strojů po dnech za období start–end (včetně) z kapacitního kalendáře."""
    problem = _period_problem(cal, start, end)
    if problem:
        return Usage(NOT_COMPUTABLE, start, end, notes=[problem])

    blocking = (cp.BLOCKING,) + (cp.NON_BLOCKING if include_tentative else ())
    firm = [a for a in cal.assignments if a.status in blocking]
    soft = [a for a in cal.assignments if a.status not in blocking]
    units_of = {}
    for u in cal.units:
        units_of.setdefault(u.type_id, []).append(u)

    unit_rows, day_rows, conflicts = [], [], []
    for day in cp.days(start, end):
        for t in cal.type_ids():
            row, units, conf = _type_day(cal, t, day, units_of.get(t, []), firm, soft)
            day_rows.append(row)
            unit_rows += units
            conflicts += conf

    by_unit = pd.DataFrame(unit_rows, columns=UNIT_COLUMNS).sort_values(["day", "unit_id"], ignore_index=True)
    by_day = pd.DataFrame(day_rows, columns=DAY_COLUMNS)
    by_type = _by_type(by_day, by_unit, cal)
    fleet = _fleet(by_type)
    notes = _notes(cal, start, end, by_type, include_tentative)
    status = COMPLETE if (by_day["unknown"].sum() == 0 and by_type["capacity_known"].astype(bool).all()) else INCOMPLETE
    return Usage(status, start, end, by_unit, by_day, by_type, fleet, _merge(conflicts), notes)


def from_model(model, start, end, job_starts=None, plan_start=None, include_tentative: bool = False) -> Usage:
    """Využití za období přímo z datového modelu E1 (kalendář přes capacity.from_model)."""
    return machine_days(cp.from_model(model, plan_start=plan_start, job_starts=job_starts),
                        start, end, include_tentative)


# ---------------------------------------------------------------------------
# Jeden typ stroje v jednom dni
# ---------------------------------------------------------------------------

def _type_day(cal, t, day, units, firm, soft):
    available = [u for u in units if cp._unit_available(u, day)]
    today = [a for a in firm if a.type_id == t and a.start <= day <= a.end]
    pooled = [a for a in today if a.unit_id is None]                  # známý jen typ stroje
    per_unit = {}
    for a in today:
        if a.unit_id is not None:
            per_unit.setdefault(a.unit_id, []).append(a)

    unit_rows, conflicts = [], []
    counts = {g: 0 for g in GROUPS}
    excess = 0
    ids = {u.unit_id for u in available}

    def uses(unit_id):
        """Různá využití kusu; překrývající se servisní záznamy téhož kusu = jeden servis."""
        items = per_unit.get(unit_id, [])
        service = [a for a in items if a.kind == "maintenance"]
        return [a for a in items if a.kind != "maintenance"] + service[:1], items

    # Kusy mimo flotilu: servis se nepočítá, pronájem / zakázka je konflikt mimo kapacitu.
    for u in units:
        if u.unit_id in ids:
            continue
        planned = [a for a in per_unit.get(u.unit_id, []) if a.kind != "maintenance"]
        note = ("přesto naplánováno: " + ", ".join(a.label for a in planned)) if planned else ""
        unit_rows.append(dict(day=day, unit_id=u.unit_id, type_id=t, in_capacity=False, state=OUT,
                              sources=_labels(planned), note=note))
        if planned:
            excess += len(planned)
            conflicts.append(dict(scope="kus", id=u.unit_id, type_id=t, day=day,
                                  kind="využití kusu mimo dobu, kdy je ve flotile",
                                  sources=_labels(planned), excess=len(planned)))

    pool = [u for u in available if not per_unit.get(u.unit_id)]
    need = sum(a.quantity for a in pooled)
    over = need > len(pool)
    risk = sum(j.quantity for j in cal.open_jobs if j.type_id == t and j.may_occupy(day))
    unsure = t in cal.unknown_occupancy

    if over:
        counts["conflict"] = len(available)
        over_by = need - len(pool) + sum(len(uses(u.unit_id)[0]) - 1 for u in available if per_unit.get(u.unit_id))
        excess += over_by
        # servis kusu, který ten den není ve flotile, kapacitu neubírá (jako capacity._counts)
        active = [a for a in today if not (a.kind == "maintenance" and a.unit_id not in ids)]
        sources = _labels(active)
        conflicts.append(dict(scope="typ", id=t, type_id=t, day=day,
                              kind="překročení kapacity typu – " + cp._conflict_kind({a.kind for a in active}),
                              sources=sources, excess=over_by))
        for u in available:
            unit_rows.append(dict(day=day, unit_id=u.unit_id, type_id=t, in_capacity=True, state=CONFLICT,
                                  sources=sources, note=f"nároky převyšují dostupné kusy typu o {over_by}"))
    else:
        for u in available:
            distinct, raw = uses(u.unit_id)
            if not raw:
                continue
            if len(distinct) == 1:
                counts[distinct[0].kind] += 1
                state, note = KIND_STATE[distinct[0].kind], ""
                if len(raw) > 1:
                    note = "překrývající se servisní záznamy – započteno jednou"
            else:
                counts["conflict"] += 1
                excess += len(distinct) - 1
                state, note = CONFLICT, ""
                conflicts.append(dict(scope="kus", id=u.unit_id, type_id=t, day=day,
                                      kind=cp._conflict_kind({a.kind for a in distinct}),
                                      sources=_labels(distinct), excess=len(distinct) - 1))
            unit_rows.append(dict(day=day, unit_id=u.unit_id, type_id=t, in_capacity=True, state=state,
                                  sources=_labels(raw), note=note))
        # Nároky jen na typ: počty jisté, konkrétní kus ne.
        kinds = {a.kind for a in pooled}
        for a in pooled:
            counts[a.kind] += a.quantity
        rest = len(pool) - need
        unknown = rest if unsure else min(rest, risk)
        counts["unknown"] += unknown
        counts["free"] += rest - unknown
        if need == 0:
            state = UNKNOWN if unknown else FREE
            note = _uncertainty(cal, t, unsure, risk) if unknown else ""
        elif rest == 0 and len(kinds) == 1:
            state = KIND_STATE[kinds.pop()]
            note = "konkrétní kus neevidován – obsazeny všechny zbývající kusy typu"
        else:
            state = AMBIGUOUS
            note = (f"obsazeno {need} z {len(pool)} kusů bez určení konkrétního kusu"
                    + (f"; {_uncertainty(cal, t, unsure, risk)}" if unknown else ""))
        for u in pool:
            unit_rows.append(dict(day=day, unit_id=u.unit_id, type_id=t, in_capacity=True, state=state,
                                  sources=_labels(pooled), note=note))

    tentative = sum(a.quantity for a in soft if a.type_id == t and a.start <= day <= a.end)
    ambiguous = sum(1 for r in unit_rows if r["state"] == AMBIGUOUS)
    known = t not in cal.unknown_capacity
    if over and known:
        status = CONFLICT                     # přetížení je jisté i bez chybějících údajů
    elif not known or counts["unknown"]:
        status = UNKNOWN
    elif counts["conflict"] or excess:
        status = CONFLICT
    else:
        status = "jisté"
    row = dict(day=day, type_id=t, capacity=len(available), **counts, excess=excess, tentative=tentative,
               ambiguous=ambiguous, capacity_known=known, status=status)
    return row, unit_rows, conflicts


def _labels(items) -> str:
    return ", ".join(sorted(dict.fromkeys(a.label for a in items)))


def _uncertainty(cal, t, unsure, risk) -> str:
    if unsure:
        return "neplatný nebo neúplný záznam: " + "; ".join(cal.issues_for(t))
    return f"může obsadit zakázka bez pevného termínu ({risk} ks)"


# ---------------------------------------------------------------------------
# Souhrny
# ---------------------------------------------------------------------------

def _by_type(by_day, by_unit, cal) -> pd.DataFrame:
    rows = []
    for t in cal.type_ids():
        part = by_day[by_day["type_id"] == t]
        sums = {k: int(part[k].sum()) for k in ["capacity"] + GROUPS + ["excess", "tentative", "ambiguous"]}
        units = by_unit[(by_unit["type_id"] == t) & by_unit["in_capacity"]]
        used = units[units["state"].isin([RENTAL, JOB, AMBIGUOUS])]
        derived = used["note"].str.startswith("konkrétní kus neevidován").any() or (used["state"] == AMBIGUOUS).any()
        if not len(used):
            level = "bez pronájmů a zakázek"
        elif derived:
            level = "jen typ stroje – konkrétní kus neurčen"
        else:
            level = "konkrétní kusy"
        known = t not in cal.unknown_capacity
        balanced = sum(sums[g] for g in GROUPS) == sums["capacity"] if known else None
        rows.append(dict(type_id=t, **sums, conflict_days=int((part["status"] == CONFLICT).sum()),
                         capacity_known=known, unit_level=level, balanced=balanced))
    return pd.DataFrame(rows, columns=TYPE_COLUMNS)


def _fleet(by_type) -> dict:
    known = by_type.loc[by_type["capacity_known"].astype(bool)]
    sums = {k: int(known[k].sum()) for k in ["capacity"] + GROUPS + ["excess", "tentative", "ambiguous"]}
    return dict(**sums, balanced=sum(sums[g] for g in GROUPS) == sums["capacity"],
                excluded=by_type.loc[~by_type["capacity_known"].astype(bool), "type_id"].tolist())


def _notes(cal, start, end, by_type, include_tentative) -> list:
    notes = [text for _, text in cal.issues]
    for o in cp.overlapping_maintenance(cal, start, end):
        notes.append(f"kus {o['unit_id']}: překrývající se servisní záznamy ({', '.join(o['sources'])}) – "
                     "v každém dni započten jednou")
    for j in cal.open_jobs:
        if j.may_overlap(start, end):
            notes.append(j.label + " – dotčené volné strojodny jsou „nelze určit“")
    for t in by_type.loc[by_type["unit_level"].astype(str).str.startswith("jen typ"), "type_id"]:
        notes.append(f"{t}: rezervace / zakázky uvádějí jen typ stroje – konkrétní kus nelze určit")
    for t in by_type.loc[~by_type["capacity_known"].astype(bool), "type_id"]:
        notes.append(f"{t}: počet kusů nelze spolehlivě určit – součet strojodnů je neurčitelný")
    if not include_tentative and int(by_type["tentative"].sum()):
        notes.append("předběžné rezervace a nabídky kapacitu neblokují (vykázány zvlášť)")
    return notes


def _merge(items) -> list:
    """Sloučí konflikty stejného kusu / typu se stejnými zdroji v po sobě jdoucích dnech."""
    out, last = [], {}
    for c in items:
        key = (c["scope"], c["id"], c["kind"], c["sources"])
        prev = last.get(key)
        if prev is not None and cp._shift(prev["days"][-1], 1) == c["day"]:
            prev["days"].append(c["day"])
            prev["excess"] = max(prev["excess"], c["excess"])
            continue
        entry = dict(scope=c["scope"], id=c["id"], type_id=c["type_id"], kind=c["kind"], days=[c["day"]],
                     sources=c["sources"], excess=c["excess"])
        out.append(entry)
        last[key] = entry
    return sorted(out, key=lambda c: (c["days"][0], c["scope"], c["id"]))
