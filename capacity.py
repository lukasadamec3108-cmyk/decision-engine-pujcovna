"""Kapacitní plánování strojů (etapa E2.3A). Jen výpočetní logika, bez Streamlitu.

Hlídá, aby se nenaplánovalo víc využití strojů, než kolik jich firma v daný den má.
Aplikace modul zatím nepoužívá (napojení na vizuální kalendář až v E2.3B).

PRAVIDLA
--------
- Čas po celých dnech. Interval [začátek, konec] zahrnuje OBA krajní dny.
  Pronájem 1.–5. a další od 6. na sebe navazují bez konfliktu; předání ve stejný den
  (konec 5., začátek 5.) je konflikt – na převoz a kontrolu stroje mezi nimi nezbývá čas.
- Dny jsou buď čísla (den 1 = první den plánu, jako v datech E1), nebo kalendářní data
  (datetime.date) – pak výpočet správně přechází přes konec měsíce i roku. V jednom
  výpočtu se obojí nesmí míchat.
- Kapacita typu v daný den = kusy ve flotile, které jsou ten den k dispozici (vlastněné
  a leasingové; objednané až od data dodání) − kusy v servisu / údržbě.
- Počet strojů ≠ strojodny: kapacita i obsazenost se počítají v kusech pro každý den zvlášť.
- Stavy a blokování kapacity (návrh):
    confirmed              … potvrzené → BLOKUJE kapacitu
    tentative, opportunity … předběžné rezervace a nabídky → neblokují, vykazují se zvlášť
                             jako riziko (lze je volitelně započítat: include_tentative=True)
    cancelled              … zrušené → ignorují se (datový model E1 tento stav zatím nemá)
- Přednost pronájmu ani zakázky modul NEurčuje, nic nepřesouvá, nemaže ani neoptimalizuje.
- Rezervace známá jen na úrovni typu stroje se k žádnému konkrétnímu kusu nepřiřazuje –
  počítá se jen počet kusů. Konflikt konkrétního kusu se hledá jen tam, kde je kus uveden.
- Chybějící údaj NENÍ nulová obsazenost: neplatný záznam nebo potvrzená zakázka bez
  pevného termínu udělá výsledek „nelze určit“, pokud by mohl rozhodnutí změnit.
"""

from dataclasses import dataclass, field
from datetime import date, timedelta
from numbers import Integral
from typing import Optional

import pandas as pd

BLOCKING = "confirmed"
NON_BLOCKING = ("tentative", "opportunity")
IGNORED = ("cancelled",)
KIND_LABEL = {"rental": "pronájem", "job": "vlastní zakázka", "maintenance": "servis / údržba"}
UNIT_STATUSES = ("owned", "leased", "ordered")

FEASIBLE, INFEASIBLE, UNKNOWN = "realizovatelné", "nerealizovatelné", "nelze určit"
FREE, FULL, OVER = "volno", "plno", "překročeno"


# ---------------------------------------------------------------------------
# Datové struktury
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class Unit:
    """Fyzický kus stroje."""
    unit_id: str
    type_id: str
    status: str = "owned"                 # owned | leased | ordered
    available_from: object = None         # den / datum, odkdy je kus k dispozici (None = od začátku)
    available_to: object = None           # den / datum, dokdy včetně (None = bez omezení)


@dataclass(frozen=True)
class Assignment:
    """Využití stroje: pronájem, vlastní zakázka nebo servis. start i end jsou včetně."""
    kind: str                             # rental | job | maintenance
    ref: str                              # kód rezervace / zakázky / servisu
    type_id: str
    start: object
    end: object
    quantity: int = 1                     # počet kusů současně (ne strojodny)
    status: str = BLOCKING
    unit_id: Optional[str] = None         # jen pokud je známý konkrétní kus

    @property
    def label(self) -> str:
        return f"{KIND_LABEL.get(self.kind, self.kind)} {self.ref}"


@dataclass(frozen=True)
class OpenJob:
    """Potvrzená zakázka bez pevného termínu: proběhne někdy od `earliest` do `latest`.

    latest = None … konec není zaručen (termín dokončení je jen smluvní, zakázka může
    se zpožděním a penále skončit později) → může obsadit stroj kdykoli od `earliest`.
    """
    ref: str
    type_id: str
    earliest: object
    latest: object
    duration: int
    quantity: int = 1

    def may_occupy(self, day) -> bool:
        return self.earliest <= day and (self.latest is None or day <= self.latest)

    def may_overlap(self, start, end) -> bool:
        return self.earliest <= end and (self.latest is None or start <= self.latest)

    @property
    def label(self) -> str:
        end = self.latest if self.latest is not None else "neznámý konec – může skončit po termínu"
        return f"vlastní zakázka {self.ref} (bez pevného termínu, možné obsazení {self.earliest} – {end})"


@dataclass
class Calendar:
    units: list                                            # platné kusy
    assignments: list = field(default_factory=list)        # platná, nezrušená využití
    open_jobs: list = field(default_factory=list)          # potvrzené zakázky bez termínu
    issues: list = field(default_factory=list)             # [(type_id, popis)] neplatné / neúplné údaje
    unknown_capacity: set = field(default_factory=set)     # typy, u nichž nevíme počet kusů
    unknown_occupancy: set = field(default_factory=set)    # typy s potvrzeným využitím bez platných údajů
    day_kind: Optional[type] = None                        # date nebo int podle vstupů (None = žádné dny)

    def type_ids(self) -> list:
        return sorted({u.type_id for u in self.units} | self.unknown_capacity)

    def issues_for(self, type_id: str) -> list:
        return [text for t, text in self.issues if t == type_id]


# ---------------------------------------------------------------------------
# Práce se dny
# ---------------------------------------------------------------------------

def _is_day(value) -> bool:
    return isinstance(value, date) or (isinstance(value, Integral) and not isinstance(value, bool))


def _is_count(value) -> bool:
    return isinstance(value, Integral) and not isinstance(value, bool) and value >= 1


def _valid_interval(start, end) -> bool:
    if not (_is_day(start) and _is_day(end)):
        return False
    if isinstance(start, date) != isinstance(end, date):
        return False
    return end >= start


def days(start, end) -> list:
    """Všechny dny intervalu včetně obou konců (čísla i data, přes konec měsíce / roku)."""
    if isinstance(start, date) != isinstance(end, date):
        raise TypeError("Nelze míchat čísla dní a kalendářní data.")
    if isinstance(start, date):
        return [start + timedelta(d) for d in range((end - start).days + 1)]
    return list(range(int(start), int(end) + 1))


def _kind(value):
    return date if isinstance(value, date) else int


def _check_request_kind(cal, start, end):
    """Požadovaný termín musí používat stejný druh dnů jako kalendář (čísla, nebo data)."""
    kinds = {_kind(start), _kind(end)} | ({cal.day_kind} if cal.day_kind else set())
    if len(kinds) > 1:
        raise TypeError("Nelze míchat čísla dní a kalendářní data.")


def _shift(day, n):
    return day + timedelta(n) if isinstance(day, date) else day + n


def overlap(a_start, a_end, b_start, b_end) -> bool:
    """Překrývají se dva intervaly (oba konce včetně)? Navazující intervaly se nepřekrývají."""
    return a_start <= b_end and b_start <= a_end


# ---------------------------------------------------------------------------
# Sestavení kalendáře
# ---------------------------------------------------------------------------

def build_calendar(units, assignments=(), open_jobs=()) -> Calendar:
    """Zkontroluje vstupy a připraví kalendář. Nic nemění ani nepřesouvá.

    Neplatné záznamy se nezahodí potichu: zapíšou se do `issues` a dotčený typ stroje
    se označí jako nejistý (kapacita nebo obsazenost „nelze určit“).
    """
    cal = Calendar(units=[])
    for u in units:
        problem = None
        if u.status not in UNIT_STATUSES:
            problem = f"kus {u.unit_id}: neznámý stav „{u.status}“"
        elif any(v is not None and not _is_day(v) for v in (u.available_from, u.available_to)):
            problem = f"kus {u.unit_id}: neplatné datum dostupnosti"
        elif u.available_from is not None and u.available_to is not None \
                and not _valid_interval(u.available_from, u.available_to):
            problem = f"kus {u.unit_id}: neplatné období dostupnosti"
        if problem:
            cal.issues.append((u.type_id, problem))
            cal.unknown_capacity.add(u.type_id)
            continue
        cal.units.append(u)

    unit_type = {u.unit_id: u.type_id for u in cal.units}
    known_types = set(unit_type.values()) | cal.unknown_capacity
    for a in assignments:
        if a.status in IGNORED:
            continue
        problem = None
        if a.kind not in KIND_LABEL:
            problem = f"{a.ref}: neznámý druh využití „{a.kind}“"
        elif a.status not in (BLOCKING,) + NON_BLOCKING:
            problem = f"{a.label}: neznámý stav „{a.status}“"
        elif not _valid_interval(a.start, a.end):
            problem = f"{a.label}: chybí nebo je neplatný termín"
        elif not _is_count(a.quantity):
            problem = f"{a.label}: neplatný počet kusů"
        elif a.unit_id is not None and a.unit_id not in unit_type:
            problem = f"{a.label}: kus {a.unit_id} není ve flotile"
        elif a.unit_id is not None and unit_type[a.unit_id] != a.type_id:
            problem = f"{a.label}: kus {a.unit_id} není typu {a.type_id}"
        elif a.unit_id is not None and a.quantity != 1:
            problem = f"{a.label}: u konkrétního kusu musí být počet 1"
        elif a.type_id not in known_types:
            problem = f"{a.label}: typ stroje {a.type_id} není ve flotile"
        if problem:
            cal.issues.append((a.type_id, problem))
            if a.status not in NON_BLOCKING:
                # potvrzený záznam nebo záznam s neznámým stavem (např. překlep) může kapacitu
                # blokovat → obsazenost typu nelze spolehlivě určit
                cal.unknown_occupancy.add(a.type_id)
            continue
        cal.assignments.append(a)

    for j in open_jobs:
        window_ok = _is_day(j.earliest) and (
            j.latest is None or _valid_interval(j.earliest, j.latest))
        if not (window_ok and _is_count(j.duration) and _is_count(j.quantity)):
            cal.issues.append((j.type_id, f"zakázka {j.ref}: chybí platné časové okno, délka nebo počet"))
            cal.unknown_occupancy.add(j.type_id)
            continue
        cal.open_jobs.append(j)

    values = [v for u in cal.units for v in (u.available_from, u.available_to) if v is not None]
    values += [v for a in cal.assignments for v in (a.start, a.end)]
    values += [v for j in cal.open_jobs for v in (j.earliest, j.latest) if v is not None]
    kinds = {_kind(v) for v in values}
    if len(kinds) > 1:
        raise TypeError("Nelze míchat čísla dní a kalendářní data.")
    cal.day_kind = kinds.pop() if kinds else None
    return cal


def _unit_available(u: Unit, day) -> bool:
    if u.status == "ordered" and u.available_from is None:
        return False                      # objednaný kus bez data dodání se nepočítá
    if u.available_from is not None and day < u.available_from:
        return False
    if u.available_to is not None and day > u.available_to:
        return False
    return True


def _unit(cal: Calendar, unit_id: str) -> Unit:
    return next(u for u in cal.units if u.unit_id == unit_id)


def _counts(cal: Calendar, a: Assignment, day) -> bool:
    """Ubírá využití v daný den kapacitu typu?

    Servis kusu, který ten den není ve flotile, ji podruhé neubírá – kus už v počtu není.
    Pronájem nebo zakázka konkrétního kusu mimo jeho dostupnost se ale NEZAHAZUJE: počítá se
    (stroj je slíbený) a find_conflicts ho vypíše jako konflikt.
    """
    if not (a.start <= day <= a.end):
        return False
    if a.unit_id is not None and a.kind == "maintenance":
        return _unit_available(_unit(cal, a.unit_id), day)
    return True


# ---------------------------------------------------------------------------
# A + B. Denní dostupnost a obsazenost
# ---------------------------------------------------------------------------

COLUMNS = ["day", "type_id", "units", "maintenance", "rental", "job", "free", "tentative",
           "open_job_risk", "status"]


def daily_capacity(cal: Calendar, start, end, type_ids=None) -> pd.DataFrame:
    """Kapacita po dnech a typech strojů (počty kusů v daný den, ne strojodny).

    Sloupce: day, type_id, units (kusy k dispozici), maintenance, rental, job (potvrzené),
    free (= units − maintenance − rental − job; záporné = přetížení), tentative (předběžné
    nároky), open_job_risk (kusy, které může zabrat zakázka bez pevného termínu),
    status: volno | plno | překročeno | nelze určit.
    """
    _check_request_kind(cal, start, end)
    types = sorted(type_ids) if type_ids else cal.type_ids()
    rows = []
    for day in days(start, end):
        for t in types:
            units = sum(1 for u in cal.units if u.type_id == t and _unit_available(u, day))
            use = {"maintenance": 0, "rental": 0, "job": 0}
            tentative = 0
            for a in cal.assignments:
                if a.type_id == t and _counts(cal, a, day):
                    if a.status == BLOCKING:
                        use[a.kind] += a.quantity
                    else:
                        tentative += a.quantity
            risk = sum(j.quantity for j in cal.open_jobs if j.type_id == t and j.may_occupy(day))
            free = units - sum(use.values())
            if free < 0 and t not in cal.unknown_capacity:
                status = OVER                      # přetížení je jisté i bez chybějících údajů
            elif t in cal.unknown_capacity or t in cal.unknown_occupancy:
                status = UNKNOWN
            elif free == 0:
                status = FULL
            elif free - risk < 1:
                status = UNKNOWN                   # volný kus může zabrat zakázka bez pevného termínu
            else:
                status = FREE                      # aspoň jeden kus je volný i v nejhorším případě
            rows.append(dict(day=day, type_id=t, units=units, **use, free=free,
                             tentative=tentative, open_job_risk=risk, status=status))
    return pd.DataFrame(rows, columns=COLUMNS)


def unit_occupancy(cal: Calendar, start, end) -> pd.DataFrame:
    """Stav jednotlivých fyzických kusů po dnech: mimo flotilu / servis / pronájem / zakázka /
    bez přiřazení. „bez přiřazení“ NEZNAMENÁ volno: rezervace známé jen po typech se konkrétním
    kusům nepřiřazují (počty kusů viz daily_capacity)."""
    _check_request_kind(cal, start, end)
    rows = []
    for day in days(start, end):
        for u in cal.units:
            hits = [a for a in cal.assignments if a.unit_id == u.unit_id and a.status == BLOCKING
                    and a.start <= day <= a.end]
            used = " + ".join(KIND_LABEL[a.kind] for a in hits)
            if not _unit_available(u, day):
                planned = " + ".join(KIND_LABEL[a.kind] for a in hits if a.kind != "maintenance")
                state = "mimo flotilu" + (f" (přesto naplánováno: {planned})" if planned else "")
            else:
                state = used or "bez přiřazení"
            rows.append(dict(day=day, unit_id=u.unit_id, type_id=u.type_id, state=state))
    return pd.DataFrame(rows, columns=["day", "unit_id", "type_id", "state"])


# ---------------------------------------------------------------------------
# C. Konflikty
# ---------------------------------------------------------------------------

def _conflict_kind(kinds: set) -> str:
    if "maintenance" in kinds:
        return "konflikt se servisem"
    if kinds == {"rental"}:
        return "překryv rezervací"
    if kinds == {"rental", "job"}:
        return "konflikt pronájmu a vlastní zakázky"
    if kinds == {"job"}:
        return "překryv vlastních zakázek"
    return "nedostatek kusů ve flotile"


def find_conflicts(cal: Calendar, start=None, end=None) -> list:
    """Konflikty POTVRZENÝCH využití. Nic se neřeší ani nepřesouvá – jen se vypíší.

    Každý konflikt: {kind, scope ('kus' | 'typ'), id, days, sources, certain[, max_over]}.
      - scope 'kus': stejný konkrétní kus využitý dvakrát ve stejný den, nebo naplánovaný
        na den, kdy není ve flotile,
      - scope 'typ': v daný den je potřeba víc kusů typu, než kolik jich je k dispozici;
        max_over = o kolik kusů je kapacita nejvíce překročena,
      - certain = False: počet kusů typu není kvůli neplatným údajům jistý.
    """
    firm = [a for a in cal.assignments if a.status == BLOCKING]
    if not firm:
        return []
    start = start if start is not None else min(a.start for a in firm)
    end = end if end is not None else max(a.end for a in firm)
    _check_request_kind(cal, start, end)
    out = []
    for a in firm:
        if a.unit_id is None or a.kind == "maintenance":
            continue
        unit = _unit(cal, a.unit_id)
        outside = [d for d in days(max(a.start, start), min(a.end, end)) if not _unit_available(unit, d)]
        if outside:
            out.append(dict(kind="využití kusu mimo dobu, kdy je ve flotile", scope="kus", id=a.unit_id,
                            days=outside, sources=[a.label], certain=True))
    by_unit = {}
    for a in firm:
        if a.unit_id:
            by_unit.setdefault(a.unit_id, []).append(a)
    for unit_id, items in sorted(by_unit.items()):
        for i, a in enumerate(items):
            for b in items[i + 1:]:
                s, e = max(a.start, b.start, start), min(a.end, b.end, end)
                if s <= e:
                    out.append(dict(kind=_conflict_kind({a.kind, b.kind}), scope="kus",
                                    id=unit_id, days=days(s, e), sources=[a.label, b.label],
                                    certain=True))
    cap = daily_capacity(cal, start, end)
    over = cap[cap["free"] < 0]
    for t, grp in over.groupby("type_id", sort=True):
        current = None
        for r in grp.itertuples():
            active = [a for a in firm if a.type_id == t and _counts(cal, a, r.day)]
            src = sorted(a.label for a in active)
            if current and current["sources"] == src and _shift(current["days"][-1], 1) == r.day:
                current["days"].append(r.day)
                current["max_over"] = max(current["max_over"], int(-r.free))
            else:
                current = dict(kind="překročení kapacity typu – " + _conflict_kind({a.kind for a in active}),
                               scope="typ", id=t, days=[r.day], sources=src, max_over=int(-r.free),
                               certain=t not in cal.unknown_capacity)
                out.append(current)
    return out


# ---------------------------------------------------------------------------
# D. Vyhodnocení nového požadavku
# ---------------------------------------------------------------------------

def evaluate_request(cal: Calendar, type_id: str, start, end, quantity: int = 1,
                     include_tentative: bool = False) -> dict:
    """Lze přijmout požadavek na `quantity` kusů typu `type_id` v termínu start–end (včetně)?

    Nic nerezervuje ani neposouvá. Vrací slovník:
      status       … realizovatelné | nerealizovatelné | nelze určit
      min_free     … nejmenší počet volných kusů v termínu (potvrzené nároky)
      free_by_day  … {den: volné kusy}
      problem_days … [(den, volné kusy, požadováno, příčiny – rezervace / zakázky / servis)]
      warnings     … předběžné nároky, které by konflikt způsobily, kdyby se potvrdily
      reasons      … proč „nelze určit“ (neúplná data, zakázky bez pevného termínu)
      data_issues  … všechny neplatné / neúplné záznamy v kalendáři (i u jiných typů)
    include_tentative=True započte předběžné nároky jako blokující (opatrnější pohled).
    """
    def result(status, reasons, cap=None, problems=(), warnings=()):
        free = {} if cap is None else {d: int(v) for d, v in cap["free"].items()}
        return dict(status=status, min_free=min(free.values()) if free else None, free_by_day=free,
                    problem_days=list(problems), warnings=list(warnings), reasons=list(reasons),
                    data_issues=[text for _, text in cal.issues])

    if not _valid_interval(start, end):
        return result(UNKNOWN, ["neplatný požadovaný termín"])
    if not _is_count(quantity):
        return result(UNKNOWN, ["neplatný požadovaný počet kusů"])
    if type_id not in cal.type_ids():
        return result(INFEASIBLE, [f"typ stroje {type_id} není ve flotile"])

    cap = daily_capacity(cal, start, end, [type_id]).set_index("day")
    if include_tentative:
        cap["free"] = cap["free"] - cap["tentative"]
    active = [a for a in cal.assignments if a.type_id == type_id]
    problems, warnings = [], []
    worst_short = False
    for day, r in cap.iterrows():
        today = [a for a in active if _counts(cal, a, day)]
        firm = [a.label for a in today if a.status == BLOCKING]
        tent = [a.label for a in today if a.status != BLOCKING]
        free = int(r.free)
        if free < quantity:
            problems.append((day, free, quantity, firm + (tent if include_tentative else [])))
            continue
        if not include_tentative and r.tentative and free - r.tentative < quantity:
            warnings.append((day, "kdyby se potvrdily předběžné nároky: " + ", ".join(tent)))
        if free - r.open_job_risk < quantity:
            worst_short = True

    reasons = []
    if type_id in cal.unknown_capacity or type_id in cal.unknown_occupancy:
        reasons.append("neplatné nebo neúplné údaje pro tento typ stroje: "
                       + "; ".join(cal.issues_for(type_id)))
    if worst_short:
        reasons += [j.label for j in cal.open_jobs if j.type_id == type_id and j.may_overlap(start, end)]

    if problems and type_id not in cal.unknown_capacity:
        status = INFEASIBLE          # kusy chybí už podle jistých údajů; další nejistota to nezlepší
    elif problems or reasons:
        status = UNKNOWN
    else:
        status = FEASIBLE
    return result(status, reasons, cap, problems, warnings)


# ---------------------------------------------------------------------------
# Napojení na datový model E1
# ---------------------------------------------------------------------------

INVALID = "neplatná hodnota"          # značka neplatné hodnoty z CSV (neprojde kontrolou)


def _int(value):
    s = "" if value is None else str(value).strip()
    if s in ("", "nan", "None", "<NA>"):
        return None
    try:
        f = float(s)
    except ValueError:
        return INVALID
    return int(f) if f.is_integer() else INVALID


def _day(value, plan_start: Optional[date]):
    """Číslo dne z dat E1 → číslo dne, nebo kalendářní datum, je-li zadán plan_start (= den 1)."""
    n = _int(value)
    if not isinstance(n, int):
        return n
    return plan_start + timedelta(n - 1) if plan_start else n


def _end(start, length):
    """Poslední den intervalu dané délky (včetně), nebo None, nejde-li spočítat."""
    return _shift(start, length - 1) if _is_day(start) and _is_count(length) else None


def from_model(model, plan_start: Optional[date] = None, job_starts: Optional[dict] = None,
               assume_deadlines_met: bool = False) -> Calendar:
    """Kalendář z datového modelu E1 (datamodel.DataModel). Model se nijak nemění.

    - kusy z fleet_units, servis z maintenance (konkrétní kus), rezervace z rental_bookings
      (jen typ stroje – konkrétní kus se nevymýšlí),
    - potvrzené zakázky: s termínem z `job_starts` {job_id: první den} jako pevné využití,
      bez něj jako zakázka bez pevného termínu, která může stroj obsadit kdykoli od
      nejdřívějšího začátku. Termín dokončení je jen smluvní (zakázka může skončit později
      s penále, jak to dovoluje i plán v záložce 6), proto se konec neomezuje. Jen při
      `assume_deadlines_met=True` se možné obsazení omezí do termínu (nejméně na celou délku
      zakázky) – je to předpoklad, ne jistota,
    - nepotvrzené zakázky (tentative, opportunity) se vykazují jako předběžné, jen mají-li termín.
    `plan_start` převede čísla dní na kalendářní data.
    """
    units = [Unit(str(r.unit_id), str(r.type_id), str(r.status),
                  _day(r.available_from_day, plan_start), _day(r.available_to_day, plan_start))
             for r in model.fleet_units.itertuples()]
    unit_type = {u.unit_id: u.type_id for u in units}
    assignments = []
    for r in model.maintenance.itertuples():
        s = _day(r.start_day, plan_start)
        assignments.append(Assignment("maintenance", f"{r.unit_id} od dne {r.start_day}",
                                      unit_type.get(str(r.unit_id), "?"), s, _end(s, _int(r.days)),
                                      1, BLOCKING, str(r.unit_id)))
    for r in model.rental_bookings.itertuples():
        assignments.append(Assignment("rental", str(r.booking_id), str(r.type_id),
                                      _day(r.start_day, plan_start), _day(r.end_day, plan_start),
                                      _int(r.units), str(r.status)))
    open_jobs = []
    req = model.job_requirements[model.job_requirements["resource_kind"] == "machine"]
    job_starts = job_starts or {}
    for j in model.jobs.itertuples():
        dur = _int(j.duration_days)
        for r in req[req["job_id"] == j.job_id].itertuples():
            q = _int(r.quantity)
            if j.job_id in job_starts:
                s = _day(job_starts[j.job_id], plan_start)
                assignments.append(Assignment("job", str(j.job_id), str(r.resource_id), s,
                                              _end(s, dur), q, str(j.status)))
            elif j.status == BLOCKING:
                es, dl = _day(j.earliest_start_day, plan_start), _day(j.deadline_day, plan_start)
                latest = None
                if assume_deadlines_met:
                    min_end = _end(es, dur)
                    latest = max(dl, min_end) if _is_day(dl) and _is_day(min_end) else min_end
                open_jobs.append(OpenJob(str(j.job_id), str(r.resource_id), es, latest, dur, q))
    return build_calendar(units, assignments, open_jobs)
