"""Ekonomická evidence závazků za období (E2.2 – K3.1: period_ledger) a varianty řešení
konfliktů (K3.2: conflict_options). Bez Streamlitu.

Propojuje kapacitní fakta z fleet_usage.py (nad capacity.py) s příspěvky z economics.py.
Nic nového nepočítá jinak: strojodny a konflikty jsou z fleet_usage, částky z economics,
vzorce s neznámými a porovnání z econ_formula, kontrola posunu zakázky z planner/capacity.
K3.3 (period_result) skládá evidenci a varianty konfliktů s fixními náklady a režií z economics.
Zatím se nikde nezobrazuje a nic nerozhoduje (porovnání využití – K3.4 a dál – tu záměrně NENÍ).

ZÁSADY
------
- Každý pronájem (rental_bookings) a každá zakázka (jobs) má právě jeden řádek evidence,
  i když je mimo období, zrušený nebo neplatný – nic nezmizí potichu.
- Stav kapacity závazku (právě jeden):
    zrušeno            … stav cancelled, nepočítá se,
    nelze určit        … neplatný / chybějící termín nebo počet, potvrzená zakázka bez termínu
                         v plánu, nejistý počet kusů nebo obsazenost typu, možný střet se
                         zakázkou bez pevného termínu,
    mimo období        … závazek období vůbec nezasahuje,
    předběžné          … předběžná rezervace / nabídka (kapacitu neblokuje),
    konflikt           … některý den závazku (i mimo období) je v konfliktu podle fleet_usage,
    přesah mimo období … závazek začíná nebo končí mimo období – vyžaduje pravidlo časového
                         rozlišení (OTEVŘENÉ – ekonomika se poměrně NEDĚLÍ),
    jisté využití      … potvrzený závazek celý v období bez konfliktu.
  Nejednoznačné přiřazení konkrétního kusu je samostatný údaj (unit_assignment): ekonomika
  je po typech strojů, proto stav „jisté využití“ neruší – jen se výslovně uvede.
- Příspěvek je vždy za CELÝ závazek (economics.rental_contribution / job_contribution) se
  stavem úplnosti. Chybějící údaj zůstává chybějícím (nikdy 0); neplatný = nelze spočítat.
- Mezi jisté příspěvky se sčítají jen závazky ve stavu „jisté využití“. Konfliktní, předběžné,
  přesahující a neurčité závazky se vykazují zvlášť a NEsčítají se s jistými.
- Plán zakázek je NÁVRH (date_source = „návrh plánu“), ne smluvní závazek; rezervace jsou
  evidence. Modelová poptávka (rental_demand) se nepoužívá.
- Priorita pronájmu a zakázky se neurčuje. Období je povinný parametr (žádný 30denní měsíc).
"""

import calendar
import dataclasses
import itertools
from dataclasses import dataclass, field
from datetime import date, timedelta

import pandas as pd

import capacity as cp
import econ_formula as ef
import economics as ec
import fleet_usage as fu
import planner as pl

CANCELLED, UNKNOWN, OUTSIDE = "zrušeno", "nelze určit", "mimo období"
TENTATIVE, CONFLICT, CROSSES, CERTAIN = "předběžné", "konflikt", "přesah mimo období", "jisté využití"
STATES = [CERTAIN, CONFLICT, TENTATIVE, CROSSES, UNKNOWN, OUTSIDE, CANCELLED]

UNIT_EXACT, UNIT_DERIVED, UNIT_AMBIGUOUS, UNIT_NONE = (
    "konkrétní kus", "kus odvozen – jediná možnost", "nejednoznačné přiřazení", "—")
FROM_BOOKING, FROM_PLAN, NO_DATE = "rezervace", "návrh plánu", "bez termínu"

COLUMNS = ["ref", "kind", "name", "booking_status", "date_source", "start", "end", "needs",
           "state", "reason", "crosses_period", "conflict_days", "unit_assignment",
           "machine_days", "md_in_period", "md_outside", "md_conflict_in_period",
           "contribution", "contribution_status", "missing", "invalid", "counted_as_certain"]

OPEN_RULES = [
    "časové rozlišení závazků přesahujících období (pronájmy i zakázky) – nerozhodnuto, "
    "příspěvek je uveden za celý závazek a do jistých se nezapočítává",
    "plán zakázek je návrh – jistota „jisté využití“ u zakázek platí jen pro tento plán",
    "řešení konfliktů (který závazek ustoupí) – nerozhoduje se, rozhodne majitel",
    "penále potvrzené zakázky bez termínu v plánu nelze určit (příspěvek je horní mez)",
]


@dataclass
class Ledger:
    """Výsledek period_ledger.

    status   … úplný | neúplný | nelze spočítat (neplatné období)
    entries  … DataFrame (COLUMNS), jeden řádek na závazek
    results  … {(druh, ref): economics.Result} – příspěvek celého závazku
    totals   … {skupina: economics.Result} – součty po skupinách (nesčítat mezi sebou)
    usage    … fleet_usage.Usage za období (strojodny, konflikty)
    notes, open_rules … nejistoty a nerozhodnutá pravidla
    """
    status: str
    start: object = None
    end: object = None
    entries: pd.DataFrame = field(default_factory=lambda: pd.DataFrame(columns=COLUMNS, dtype=object))
    results: dict = field(default_factory=dict)
    totals: dict = field(default_factory=dict)
    usage: object = None
    notes: list = field(default_factory=list)
    open_rules: list = field(default_factory=list)


def period_ledger(model, start, end, job_starts=None, plan_start=None) -> Ledger:
    """Evidence pronájmů a zakázek za období start–end (oba dny včetně).

    job_starts … návrh plánu zakázek {job_id: první den} (např. capacity_view.job_starts_from_plan);
    plan_start … datum dne 1 (dny se pak převedou na kalendářní data, start/end jsou data).
    Model ani vstupy se nemění.
    """
    job_starts = dict(job_starts or {})
    cal = cp.from_model(model, plan_start=plan_start, job_starts=job_starts)
    usage = fu.machine_days(cal, start, end)
    if usage.status == fu.NOT_COMPUTABLE:
        return Ledger(fu.NOT_COMPUTABLE, start, end, usage=usage, notes=list(usage.notes),
                      open_rules=list(OPEN_RULES))

    items = [a for a in cal.assignments if a.kind in ("rental", "job")]
    # konflikty hledáme v celém rozsahu dotčených závazků – konflikt mimo období také zpochybní závazek
    lo = min([start] + [a.start for a in items])
    hi = max([end] + [a.end for a in items])
    span = fu.machine_days(cal, lo, hi) if (lo, hi) != (start, end) else usage
    conflicts = _conflict_days(span)
    ambiguous = _unit_states(usage)

    rows, results = [], {}
    for r in model.rental_bookings.itertuples():
        row, res = _rental(model, cal, r, start, end, conflicts, ambiguous)
        rows.append(row)
        results[("rental", row["ref"])] = res
    for j in model.jobs.itertuples():
        row, res = _job(model, cal, j, job_starts, start, end, conflicts, ambiguous)
        rows.append(row)
        results[("job", row["ref"])] = res

    entries = pd.DataFrame(rows, columns=COLUMNS, dtype=object)      # chybějící hodnota zůstane None, ne NaN
    totals = _totals(entries, results)
    incomplete = (usage.status != fu.COMPLETE
                  or any(not results[(r.kind, r.ref)].complete for r in entries.itertuples()
                         if r.state not in (CANCELLED, OUTSIDE))
                  or entries["state"].isin([UNKNOWN, CROSSES, CONFLICT]).any())
    notes = list(usage.notes) + [
        "modelová poptávka po pronájmu (rental_demand) se do evidence nezahrnuje",
        "příspěvky jsou bez fixních nákladů flotily a bez režie – nejde o zisk",
    ]
    if any(r.fictitious for r in results.values()):
        notes.append("fiktivní ukázková data")
    return Ledger(fu.INCOMPLETE if incomplete else fu.COMPLETE, start, end, entries, results, totals,
                  usage, notes, list(OPEN_RULES))


# ---------------------------------------------------------------------------
# Kapacitní údaje z fleet_usage
# ---------------------------------------------------------------------------

def _sources(text: str) -> set:
    return {s.strip() for s in str(text).split(", ") if s.strip()}


def _conflict_days(usage) -> dict:
    """{(popisek závazku, typ stroje): {dny v konfliktu}} ze seznamu konfliktů fleet_usage."""
    out = {}
    for c in usage.conflicts:
        for label in _sources(c["sources"]):
            out.setdefault((label, c["type_id"]), set()).update(c["days"])
    return out


def _unit_states(usage) -> dict:
    """{(popisek, den, typ): {stavy kusů}} – kde se nárok projevil v tabulce kusů."""
    out = {}
    for r in usage.by_unit.itertuples():
        for label in _sources(r.sources):
            out.setdefault((label, r.day, r.type_id), set()).add(r.state)
    return out


# ---------------------------------------------------------------------------
# Jeden závazek
# ---------------------------------------------------------------------------

def _assignments(cal, kind, ref):
    return [a for a in cal.assignments if a.kind == kind and a.ref == ref]


def _issues(cal, label) -> list:
    return [text for _, text in cal.issues if text.startswith(label + ":") or f" {label} " in f" {text} "]


def _base(kind, ref, name, status, source):
    return dict(ref=ref, kind=kind, name=name, booking_status=status, date_source=source, start=None, end=None,
                needs="", state=UNKNOWN, reason="", crosses_period=False, conflict_days=[],
                unit_assignment=UNIT_NONE, machine_days=None, md_in_period=None, md_outside=None,
                md_conflict_in_period=None, contribution=None, contribution_status="", missing=(),
                invalid=(), counted_as_certain=False)


def _classify(cal, row, items, start, end, conflicts, units, blocking):
    """Doplní termín, strojodny, konflikty a stav podle kapacitních údajů. items = Assignment(y) závazku."""
    label = items[0].label
    s, e = items[0].start, items[0].end
    row.update(start=s, end=e, needs="; ".join(f"{a.type_id}×{a.quantity}" for a in items))
    total = in_period = conflict_in = 0
    conflict_all = set()
    unit_kinds = set()
    for a in items:
        span = cp.days(a.start, a.end)
        inside = [d for d in span if start <= d <= end]
        total += len(span) * a.quantity
        in_period += len(inside) * a.quantity
        days = conflicts.get((label, a.type_id), set()) & set(span)
        conflict_all |= days
        conflict_in += len([d for d in inside if d in days]) * a.quantity
        for d in inside:
            unit_kinds |= units.get((label, d, a.type_id), set())
    row.update(machine_days=total, md_in_period=in_period, md_outside=total - in_period,
               md_conflict_in_period=conflict_in, conflict_days=sorted(conflict_all),
               crosses_period=bool(in_period) and (s < start or e > end))
    if any(a.unit_id is not None for a in items):
        row["unit_assignment"] = UNIT_EXACT
    elif fu.AMBIGUOUS in unit_kinds:
        row["unit_assignment"] = UNIT_AMBIGUOUS
    elif unit_kinds:
        row["unit_assignment"] = UNIT_DERIVED

    types = {a.type_id for a in items}
    risky = sorted(t for t in types if any(j.type_id == t and j.may_overlap(s, e) for j in cal.open_jobs))
    unsure = sorted(t for t in types if t in cal.unknown_capacity or t in cal.unknown_occupancy)
    if not in_period:
        row.update(state=OUTSIDE, reason="závazek období nezasahuje")
    elif not blocking:
        row.update(state=TENTATIVE, reason="předběžný závazek – kapacitu neblokuje, do jistých se nepočítá")
    elif unsure:
        row.update(state=UNKNOWN, reason="počet kusů nebo obsazenost typu nelze spolehlivě určit: "
                   + "; ".join(sorted({x for t in unsure for x in cal.issues_for(t)})))
    elif risky:
        row.update(state=UNKNOWN, reason="může kolidovat s potvrzenou zakázkou bez termínu v plánu ("
                   + ", ".join(sorted({j.ref for j in cal.open_jobs if j.type_id in risky})) + ")")
    elif conflict_all:
        out = sorted(d for d in conflict_all if not (start <= d <= end))
        row.update(state=CONFLICT, reason="konflikt ve dnech " + ", ".join(map(str, sorted(conflict_all)))
                   + (" (část mimo období)" if out else "") + " – priorita se neurčuje")
    elif row["crosses_period"]:
        row.update(state=CROSSES, reason="závazek přesahuje období – chybí pravidlo časového rozlišení")
    else:
        row.update(state=CERTAIN, reason="potvrzený závazek celý v období bez konfliktu")


def _rental(model, cal, r, start, end, conflicts, units):
    ref, status = str(r.booking_id), str(r.status)
    row = _base("rental", ref, "", status, FROM_BOOKING)
    label = f"{cp.KIND_LABEL['rental']} {ref}"
    res = _rental_result(model, r)
    if status in cp.IGNORED:
        row.update(state=CANCELLED, reason="zrušená rezervace – nepočítá se")
    else:
        items = _assignments(cal, "rental", ref)
        if items:
            _classify(cal, row, items, start, end, conflicts, units, status == cp.BLOCKING)
        else:
            row.update(reason="; ".join(_issues(cal, label)) or "neplatný záznam rezervace")
    return _with_result(row, res), res


def _rental_result(model, r) -> ec.Result:
    """Příspěvek celé rezervace z economics.rental_contribution (cena z rezervace, nikdy z ceníku)."""
    fic = ec.is_fictitious(model)
    days = cp._int(r.end_day), cp._int(r.start_day)
    units = cp._int(r.units)
    price = str(r.price_day).strip()
    if price == "":
        return ec.Result(None, invalid=(f"{r.booking_id}: chybí cena rezervace",), fictitious=fic)
    try:
        price_value = float(price.replace(",", "."))
    except ValueError:
        return ec.Result(None, invalid=(f"{r.booking_id}: cena „{price}“ není číslo",), fictitious=fic)
    if not all(isinstance(x, int) for x in (*days, units)) or days[0] < days[1] or units < 1 or price_value < 0:
        return ec.Result(None, invalid=(f"{r.booking_id}: neplatný termín, počet kusů nebo cena",), fictitious=fic)
    if str(r.type_id) not in set(model.machine_types["type_id"]):
        return ec.Result(None, invalid=(f"{r.booking_id}: neznámý typ stroje {r.type_id}",), fictitious=fic)
    return ec.rental_contribution(model, str(r.type_id), (days[0] - days[1] + 1) * units, price_value)


def _job(model, cal, j, job_starts, start, end, conflicts, units):
    ref, status = str(j.job_id), str(j.status)
    planned = ref in job_starts
    row = _base("job", ref, str(j.name), status, FROM_PLAN if planned else NO_DATE)
    label = f"{cp.KIND_LABEL['job']} {ref}"
    res = _job_result(model, j, job_starts.get(ref))
    if status in cp.IGNORED:
        row.update(state=CANCELLED, reason="zrušená zakázka – nepočítá se")
    elif not planned:
        if status == cp.BLOCKING:
            row.update(state=UNKNOWN, reason="potvrzená zakázka nemá termín v návrhu plánu – obsazení nelze určit")
        elif status in cp.NON_BLOCKING:
            row.update(state=TENTATIVE, reason="nepotvrzená zakázka bez termínu – do jistých se nepočítá")
        else:
            row.update(reason=f"neznámý stav zakázky „{status}“")
    else:
        items = _assignments(cal, "job", ref)
        if items:
            _classify(cal, row, items, start, end, conflicts, units, status == cp.BLOCKING)
        else:
            row.update(reason="; ".join(_issues(cal, label)) or "neplatný termín nebo údaje zakázky")
    return _with_result(row, res), res


def _job_result(model, j, first_day) -> ec.Result:
    """Příspěvek celé zakázky z economics.job_contribution; penále podle termínu z návrhu plánu."""
    fic = ec.is_fictitious(model)
    dur, deadline = cp._int(j.duration_days), cp._int(j.deadline_day)
    if not isinstance(dur, int) or dur < 1:
        return ec.Result(None, invalid=(f"{j.job_id}: neplatná délka zakázky",), fictitious=fic)
    if first_day is None:
        try:
            base = ec.job_contribution(model, str(j.job_id), 0)
        except (ValueError, TypeError, KeyError) as exc:
            return ec.Result(None, invalid=(f"{j.job_id}: {exc}",), fictitious=fic)
        note = f"{j.job_id}: termín v plánu – penále nelze určit (příspěvek je horní mez)"
        return ec.Result(base.value, ec._unique(base.missing, (note,)), base.invalid, base.fictitious, base.parts)
    first = cp._int(first_day)
    if not isinstance(first, int) or not isinstance(deadline, int):
        return ec.Result(None, invalid=(f"{j.job_id}: neplatný začátek v plánu nebo termín dokončení",),
                         fictitious=fic)
    late = max(0, first + dur - 1 - deadline)
    try:
        return ec.job_contribution(model, str(j.job_id), late)
    except (ValueError, TypeError, KeyError) as exc:
        return ec.Result(None, invalid=(f"{j.job_id}: {exc}",), fictitious=fic)


def _with_result(row, res: ec.Result) -> dict:
    row.update(contribution=res.value, contribution_status=res.status, missing=tuple(res.missing),
               invalid=tuple(res.invalid), counted_as_certain=row["state"] == CERTAIN)
    return row


# ---------------------------------------------------------------------------
# Součty po skupinách (nesčítat mezi sebou)
# ---------------------------------------------------------------------------

def _totals(entries, results) -> dict:
    def group(mask):
        return ec.total(results[(r.kind, r.ref)] for r in entries[mask].itertuples())

    certain = entries["state"] == CERTAIN
    return {
        "jisté – pronájmy (rezervace)": group(certain & (entries["kind"] == "rental")),
        "jisté – zakázky (návrh plánu)": group(certain & (entries["kind"] == "job")),
        "v konfliktu – nelze sčítat s jistými ani mezi sebou": group(entries["state"] == CONFLICT),
        "přesah mimo období – celé závazky, chybí pravidlo": group(entries["state"] == CROSSES),
        "předběžné – jen potenciál": group(entries["state"] == TENTATIVE),
        "nelze určit": group(entries["state"] == UNKNOWN),
    }


# ===========================================================================
# K3.2 – varianty řešení konfliktů (conflict_options)
# ===========================================================================
#
# Pro každý konflikt potvrzených závazků (fleet_usage) sestaví podklad pro rozhodnutí majitele:
# dotčené závazky, stroje, dny, příspěvky a varianty řešení. ŽÁDNÁ varianta se nevybírá.
#
# Varianty (jen podložené daty):
#   posun zakázky      … zakázka se přesune na nejbližší začátek, kdy se vejdou VŠECHNY její stroje
#                        vedle ostatních závazků, servisu a zakázek (planner._check nad
#                        capacity.evaluate_request); penále z economics.job_contribution. Výsledek
#                        se ověří přepočtem kalendáře – zbývající konflikt = „částečně“, nikdy „vyřešeno“.
#   uvolnění pronájmu  … zrušení rezervace nebo její zkrácení o konfliktní dny (jen na kraji termínu) –
#                        POUZE podmíněně: storno / smluvní náhrada je neznámá (ve vzorci jako neznámá).
#   posun servisu, externí zapůjčení / náhradní stroj … nevyhodnotitelné – chybí údaje (vypíše se co).
#   ruční rozhodnutí   … vždy; konflikt trvá, ekonomika nevyčíslitelná (oba závazky nelze splnit).
# Ekonomika varianty = součet příspěvků VŠECH dotčených závazků (každý jednou, ve stejném scénáři
# nastavení direct_cost_includes_machines) jako vzorec z econ_formula; servis nemá příspěvek.
# Dny jsou čísla dní plánu (stejně jako planner.py).

FEASIBLE, CONDITIONAL, NOT_EVALUABLE, MANUAL = "proveditelná", "podmíněná", "nevyhodnotitelná", "ruční rozhodnutí"
PARTIAL = "proveditelná jen částečně – konflikt trvá"
RESOLVED, PARTLY, NOT_RESOLVED = "ano", "částečně", "ne"
ECON_COMPLETE, ECON_INCOMPLETE, ECON_NONE = "úplný", "neúplný", "nevyčíslitelný"

CONFLICT_OPEN_RULES = [
    "storno poplatek / smluvní náhrada při zrušení nebo zkrácení pronájmu – neznámé (souhlas zákazníka?)",
    "zda lze pronájem zkrátit nebo rozdělit – neznámé",
    "zda a kdy lze přesunout servis – neznámé",
    "externí zapůjčení a zaměnitelnost strojů – chybí ceny a pravidla",
    "přeplánování více zakázek najednou (kaskáda) – nevyhodnocuje se; posouvá se vždy jen jedna zakázka",
    "priorita pronájmu, zakázky a servisu – nerozhoduje se, rozhodne majitel",
]


@dataclass
class ConflictOptions:
    """Výsledek conflict_options.

    status … úplný | neúplný | nelze spočítat (neplatné období / dny nejsou čísla plánu)
    cases  … [případ konfliktu – viz _case]; selected je vždy None (nic se nevybírá)
    """
    status: str
    start: object = None
    end: object = None
    cases: list = field(default_factory=list)
    ledger: object = None
    notes: list = field(default_factory=list)
    open_rules: list = field(default_factory=list)


def conflict_options(model, start, end, job_starts=None) -> ConflictOptions:
    """Varianty řešení konfliktů potvrzených závazků, které zasahují období start–end (čísla dní plánu).

    job_starts … návrh plánu zakázek {job_id: první den}. Model, plán ani vstupy se nemění.
    """
    job_starts = dict(job_starts or {})
    rules = list(OPEN_RULES) + CONFLICT_OPEN_RULES
    if not all(isinstance(x, int) and not isinstance(x, bool) for x in (start, end)):
        return ConflictOptions(fu.NOT_COMPUTABLE, start, end,
                               notes=["období musí být čísla dní plánu (planner.py pracuje s čísly dní)"],
                               open_rules=rules)
    ledger = period_ledger(model, start, end, job_starts=job_starts)
    if ledger.status == fu.NOT_COMPUTABLE:
        return ConflictOptions(fu.NOT_COMPUTABLE, start, end, ledger=ledger, notes=list(ledger.notes),
                               open_rules=rules)
    cal = cp.from_model(model, job_starts=job_starts)
    span = _span(cal, start, end)
    before = fu.machine_days(cal, *span).conflicts
    relevant = [c for c in before if any(start <= d <= end for d in c["days"])]
    cases = [_case(model, cal, ledger, group, before, span, i + 1)
             for i, group in enumerate(_components(relevant))]
    incomplete = any(v["econ_status"] != ECON_COMPLETE for c in cases for v in c["variants"]
                     if v["feasibility"] in (FEASIBLE, CONDITIONAL, PARTIAL))
    notes = ["žádná varianta se automaticky nevybírá – podklad pro rozhodnutí majitele",
             "příspěvky jsou bez fixních nákladů flotily a bez režie – nejde o zisk"]
    if not cases:
        notes.append("v období nejsou konflikty potvrzených závazků")
    return ConflictOptions(fu.INCOMPLETE if incomplete else fu.COMPLETE, start, end, cases, ledger, notes, rules)


# ---------------------------------------------------------------------------
# Konflikty → případy
# ---------------------------------------------------------------------------

def _span(cal, start, end) -> tuple:
    """Rozsah všech závazků a období – konflikty i posuny se hledají v celém rozsahu."""
    items = cal.assignments
    return min([start] + [a.start for a in items]), max([end] + [a.end for a in items])


def _components(conflicts) -> list:
    """Konflikty se společným závazkem tvoří jeden případ (řeší se společně)."""
    groups = []
    for c in conflicts:
        labels = _sources(c["sources"])
        hit = [g for g in groups if g["labels"] & labels]
        merged = dict(labels=set(labels), items=[c])
        for g in hit:
            merged["labels"] |= g["labels"]
            merged["items"] = g["items"] + merged["items"]
            groups.remove(g)
        groups.append(merged)
    return sorted((g["items"] for g in groups), key=lambda items: min(d for c in items for d in c["days"]))


def _key(c) -> tuple:
    return (c["scope"], c["id"], c["kind"], c["sources"], tuple(c["days"]))


def _commitments(cal, labels) -> tuple:
    """(závazky {popisek: [Assignment]}, servisy [popisek]) – jen pronájmy a zakázky mají ekonomiku."""
    out, services = {}, []
    for label in sorted(labels):
        items = [a for a in cal.assignments if a.label == label]
        if items and items[0].kind == "maintenance":
            services.append(label)
        elif items:
            out[label] = items
    return out, services


def _request(model, items) -> pl.Request:
    """Závazek z dat jako planner.Request (pro econ_formula a kontrolu kapacity)."""
    a = items[0]
    needs = tuple((x.type_id, x.quantity) for x in items)
    if a.kind == "rental":
        row = model.rental_bookings[model.rental_bookings["booking_id"] == a.ref].iloc[0]
        return pl.Request(a.ref, "rental", pl.NEW, needs, a.start, a.end, price_day=pl._float(row["price_day"]),
                          in_data=True)
    row = model.jobs[model.jobs["job_id"] == a.ref].iloc[0]
    return pl.Request(a.ref, "job", pl.NEW, needs, a.start, a.end, name=str(row["name"]),
                      earliest=pl._int(row["earliest_start_day"]), deadline=pl._int(row["deadline_day"]),
                      in_data=True)


def _formula(model, req) -> ef.Contribution:
    try:
        return ef.request_formula(model, req)
    except ArithmeticError as exc:          # pojistka econ_formula: nelineární příspěvek
        return ef.Contribution(None, {}, (str(exc),))


def _result(model, req) -> ec.Result:
    """Příspěvek z economics.py (cena z rezervace; penále podle termínu požadavku)."""
    v = pl.contribution(model, req)
    return v.result if v.result is not None else ec.Result(None, invalid=("neplatný údaj",))


def _penalty_unknown(ref, what) -> ef.Contribution:
    """Neznámá smluvní náhrada (storno / zkrácení) – ve vzorci jako neznámý údaj, nikdy 0.

    Náhrada za zrušení a za zkrácení jsou RŮZNÉ neznámé (různý klíč) – v porovnání se nevyruší.
    """
    key = (f"contract_compensation_{what}", ref)
    unknown = ef.Unknown(key, f"{ref}: {what} (storno / smluvní náhrada)", ef.DATA, "Kč")
    return ef.Contribution({ef.ONLY: ef.Formula.of(0.0, {key: -1.0})}, {key: unknown})


def _econ_status(contrib: ef.Contribution, result: ec.Result) -> str:
    if not contrib.valid or result.value is None or result.invalid:
        return ECON_NONE
    has_terms = any(f.terms for f in contrib.scenarios.values())
    return ECON_INCOMPLETE if (has_terms or result.missing or len(contrib.scenarios) > 1) else ECON_COMPLETE


# ---------------------------------------------------------------------------
# Jeden případ konfliktu
# ---------------------------------------------------------------------------

def _case(model, cal, ledger, items, before, span, number) -> dict:
    labels = set().union(*(_sources(c["sources"]) for c in items))
    commits, services = _commitments(cal, labels)
    reqs = {label: _request(model, a) for label, a in commits.items()}
    base_f = {label: _formula(model, r) for label, r in reqs.items()}
    base_r = {label: _result(model, r) for label, r in reqs.items()}

    commitments = []
    for label, r in reqs.items():
        entry = ledger.entries[(ledger.entries["ref"] == r.ref) & (ledger.entries["kind"] == r.kind)]
        commitments.append(dict(
            label=label, ref=r.ref, kind=r.kind, start=r.start, end=r.end,
            needs="; ".join(f"{t}×{q}" for t, q in r.needs),
            machine_days=sum((r.end - r.start + 1) * q for _, q in r.needs),
            contribution=base_r[label].value, contribution_status=base_r[label].status,
            missing=tuple(base_r[label].missing), formula=ef.describe(base_f[label]),
            ledger_state=entry.iloc[0]["state"] if len(entry) else None))

    days = {}
    for c in items:
        days.setdefault(c["type_id"], set()).update(c["days"])
    ctx = dict(model=model, cal=cal, labels=labels, before={_key(c) for c in before}, span=span,
               reqs=reqs, base_f=base_f, base_r=base_r)

    variants = []
    for label, r in reqs.items():
        if r.kind == "job":
            variants.append(_shift_job(ctx, label))
    for label, r in reqs.items():
        if r.kind == "rental":
            conflict_days = sorted(set().union(*(set(c["days"]) for c in items if label in _sources(c["sources"]))))
            variants.append(_release_rental(ctx, label, conflict_days, whole=True))
            variants.append(_release_rental(ctx, label, conflict_days, whole=False))
    for label in services:
        variants.append(dict(key=f"servis {label}", kind="posun servisu", feasibility=NOT_EVALUABLE,
                             resolves=None, description=f"přesunout {label}", changes={}, formula=None,
                             formula_text=[], result=None, econ_status=ECON_NONE, bounds=(None, None),
                             warnings=[], blockers=[], missing=["zda a kdy lze servis přesunout",
                                                                "náklady na přesun servisu"]))
    variants.append(dict(key="externí zapůjčení", kind="externí zapůjčení / náhradní stroj",
                         feasibility=NOT_EVALUABLE, resolves=None,
                         description="zajistit stroj jinde (pronájem od třetí strany nebo jiný typ stroje)",
                         changes={}, formula=None, formula_text=[], result=None, econ_status=ECON_NONE,
                         bounds=(None, None), warnings=[], blockers=[],
                         missing=["ceny externího zapůjčení", "zaměnitelnost typů strojů"]))
    variants.append(dict(key="ruční rozhodnutí", kind="ponechat k ručnímu rozhodnutí", feasibility=MANUAL,
                         resolves=NOT_RESOLVED, description="beze změny – konflikt trvá, rozhodne majitel",
                         changes={}, formula=None, formula_text=[], result=None, econ_status=ECON_NONE,
                         bounds=(None, None), warnings=[], blockers=[],
                         missing=["oba závazky nelze splnit současně – výsledek nelze vyčíslit"]))

    comparable = [v for v in variants if v["formula"] is not None and v["formula"].valid
                  and v["feasibility"] in (FEASIBLE, CONDITIONAL)]
    comparisons = []
    for i, a in enumerate(comparable):
        for b in comparable[i + 1:]:
            r = ef.compare(a["formula"], b["formula"], (a["key"], b["key"]))
            comparisons.append(dict(a=a["key"], b=b["key"], status=r["status"], text=r["text"]))
    return dict(
        id=f"K{number}", kinds=sorted({c["kind"] for c in items}), types=sorted(days),
        days={t: sorted(d) for t, d in days.items()}, conflicts=items, services=services,
        commitments=commitments, variants=variants, comparisons=comparisons, selected=None,
        note="žádná varianta se nevybírá – rozhodne majitel")


def _after(ctx, assignments) -> tuple:
    """(vyřešeno?, zbývající konflikty případu, nové konflikty) po změně závazků – přepočtem fleet_usage."""
    cal = ctx["cal"]
    test = cp.build_calendar(cal.units, assignments, cal.open_jobs)
    lo = min([ctx["span"][0]] + [a.start for a in assignments])
    hi = max([ctx["span"][1]] + [a.end for a in assignments])
    after = fu.machine_days(test, lo, hi).conflicts
    remaining = [c for c in after if _sources(c["sources"]) & ctx["labels"]]
    new = [c for c in after if _key(c) not in ctx["before"] and c not in remaining]
    if new:
        return NOT_RESOLVED, remaining, new
    return (RESOLVED if not remaining else PARTLY), remaining, new


def _variant_economics(ctx, replaced: dict, extra=()) -> tuple:
    """Vzorec, Result a meze varianty: všechny závazky případu, každý jednou (replaced = změněné)."""
    forms = [replaced.get(label, (f, None))[0] for label, f in ctx["base_f"].items()]
    results = [replaced.get(label, (None, r))[1] for label, r in ctx["base_r"].items()]
    contrib = ef.combine(list(forms) + list(extra))
    result = ec.total(results)
    if extra:
        result = ec.Result(result.value, ec._unique(result.missing, tuple(
            u.label for x in extra for u in x.unknowns.values())), result.invalid, result.fictitious)
    return contrib, result, ef.bounds(contrib)


def _shift_job(ctx, label) -> dict:
    """Zachovat ostatní závazky a posunout zakázku na nejbližší začátek, kdy se vejde celá."""
    model, cal, req = ctx["model"], ctx["cal"], ctx["reqs"][label]
    dur = req.end - req.start + 1
    others = [a for a in cal.assignments if a.label != label]
    rest = cp.build_calendar(cal.units, others, cal.open_jobs)
    v = dict(key=f"posun {req.ref}", kind="zachovat ostatní závazky a posunout zakázku", changes={},
             formula=None, formula_text=[], result=None, econ_status=ECON_NONE, bounds=(None, None),
             warnings=[], blockers=[], missing=[], resolves=None,
             description=f"posunout zakázku {req.ref} ({req.name}) – všechny její stroje současně")
    if req.earliest is None or req.deadline is None:
        return dict(v, feasibility=NOT_EVALUABLE, missing=["nejdřívější začátek nebo termín zakázky"])
    last = max([a.end for a in cal.assignments] + [req.end]) + 1
    found, blockers, undetermined = None, {}, []
    for s in range(max(req.earliest, 1), last + 1):
        if s == req.start:
            continue
        status, problems, reasons = pl._check(rest, req, s, s + dur - 1)
        if status == pl.OK_ALONE:
            found = s
            break
        if status == pl.UNDETERMINED:
            undetermined += reasons
        for t, day, _, _, causes in problems:
            for cause in causes or ["kus není k dispozici"]:
                blockers.setdefault(cause, set()).add(day)
    v["blockers"] = [dict(source=k, days=sorted(d)) for k, d in sorted(blockers.items())]
    if found is None:
        return dict(v, feasibility=NOT_EVALUABLE, missing=list(dict.fromkeys(
            undetermined or ["zakázku nelze posunout – kapacita ani po všech závazcích není jistá"])))
    new = dataclasses.replace(req, start=found, end=found + dur - 1)
    moved = others + [cp.Assignment("job", req.ref, t, new.start, new.end, q) for t, q in req.needs]
    resolves, remaining, created = _after(ctx, moved)
    for t, q in req.needs:
        for day, text in cp.evaluate_request(rest, t, new.start, new.end, q)["warnings"]:
            v["warnings"].append(f"{t} den {day}: {text}")
    free_window = [s for s in range(max(req.earliest, 1), req.deadline - dur + 2)
                   if s != req.start and pl._check(rest, req, s, s + dur - 1)[0] == pl.OK_ALONE]
    f_new, r_new = _formula(model, new), _result(model, new)
    contrib, result, bounds = _variant_economics(ctx, {label: (f_new, r_new)})
    late_old = max(0, req.end - req.deadline)
    late_new = max(0, new.end - req.deadline)
    old_r = ctx["base_r"][label]
    v.update(
        feasibility=FEASIBLE if resolves == RESOLVED else PARTIAL if resolves == PARTLY else NOT_EVALUABLE,
        resolves=resolves,
        changes=dict(ref=req.ref, start=req.start, end=req.end, new_start=new.start, new_end=new.end,
                     days_late=late_old, new_days_late=late_new, penalty_free_starts=free_window,
                     contribution_change=(None if r_new.value is None or old_r.value is None
                                          else r_new.value - old_r.value)),
        formula=contrib, formula_text=ef.describe(contrib), result=result,
        econ_status=_econ_status(contrib, result), bounds=bounds,
        missing=list(result.missing) + [f"zbývá konflikt: {c['kind']} ({c['sources']})" for c in remaining]
        + [f"nový konflikt: {c['kind']} ({c['sources']})" for c in created])
    return v


def _release_rental(ctx, label, conflict_days, whole: bool) -> dict:
    """Zachovat zakázky a uvolnit kapacitu z pronájmu – jen podmíněný scénář (smluvní podmínky neznámé)."""
    model, cal, req = ctx["model"], ctx["cal"], ctx["reqs"][label]
    what = "zrušení celé rezervace" if whole else "zkrácení rezervace o konfliktní dny"
    v = dict(key=f"{'zrušení' if whole else 'zkrácení'} {req.ref}", kind=f"zachovat zakázku – {what}",
             description=f"{what} {req.ref}", changes={}, formula=None, formula_text=[], result=None,
             econ_status=ECON_NONE, bounds=(None, None), warnings=[], blockers=[], resolves=None,
             missing=["souhlas zákazníka / smluvní možnost " + ("zrušení" if whole else "zkrácení")])
    others = [a for a in cal.assignments if a.label != label]
    if whole:
        new_req, moved = None, others
    else:
        span = cp.days(req.start, req.end)
        keep = [d for d in span if d not in set(conflict_days)]
        if not keep or keep != span[len(span) - len(keep):] and keep != span[:len(keep)]:
            reason = ("konflikt zasahuje celý pronájem – zkrácení = zrušení" if not keep
                      else "konfliktní dny jsou uvnitř termínu – pronájem by se musel rozdělit")
            return dict(v, feasibility=NOT_EVALUABLE, missing=v["missing"] + [reason])
        new_req = dataclasses.replace(req, start=keep[0], end=keep[-1])
        moved = others + [cp.Assignment("rental", req.ref, t, new_req.start, new_req.end, q) for t, q in req.needs]
    resolves, remaining, created = _after(ctx, moved)
    if new_req is None:
        f_new, r_new = ef.NOTHING, ec.Result(0.0, fictitious=ec.is_fictitious(model))
    else:
        f_new, r_new = _formula(model, new_req), _result(model, new_req)
    extra = [_penalty_unknown(req.ref, "zrušení" if whole else "zkrácení")]
    contrib, result, bounds = _variant_economics(ctx, {label: (f_new, r_new)}, extra)
    v.update(
        feasibility=CONDITIONAL if resolves == RESOLVED else PARTIAL if resolves == PARTLY else NOT_EVALUABLE,
        resolves=resolves,
        changes=dict(ref=req.ref, start=req.start, end=req.end,
                     new_start=None if new_req is None else new_req.start,
                     new_end=None if new_req is None else new_req.end,
                     released_machine_days=sum(q for _, q in req.needs) * (
                         (req.end - req.start + 1) - (0 if new_req is None else new_req.end - new_req.start + 1))),
        formula=contrib, formula_text=ef.describe(contrib), result=result,
        econ_status=_econ_status(contrib, result), bounds=bounds,
        missing=v["missing"] + [m for m in result.missing if m not in v["missing"]]
        + [f"zbývá konflikt: {c['kind']} ({c['sources']})" for c in remaining]
        + [f"nový konflikt: {c['kind']} ({c['sources']})" for c in created])
    return v



# ===========================================================================
# K3.3 – ekonomický výsledek flotily za období (period_result)
# ===========================================================================
#
# výsledek z jistých závazků = jisté příspěvky pronájmů + jisté příspěvky zakázek
#                              − fixní náklady flotily za období − režie za období
# - jisté příspěvky = jen závazky ve stavu „jisté využití“ z period_ledger (každý jednou),
# - konfliktní, předběžné, přesahující a neurčité závazky se do něj NEZAPOČÍTÁVAJÍ – jsou zvlášť,
# - fixní náklady = economics.fleet_fixed_costs (vlastněné a leasingové kusy; servis je jejich
#   součástí – nepřičítá se znovu), odečtou se JEDNOU; režie = economics.overhead,
# - přepočet měsíčních nákladů na období jen výslovně zadanou metodou (fixed_cost_method):
#     None            … nezadáno → fixní náklady a režie jsou CHYBĚJÍCÍ údaj (výsledek neúplný),
#     "months"        … počet měsíců zadá volající (months > 0), např. 1 pro celý měsíc,
#     "calendar_days" … podle kalendáře: Σ (dny období v měsíci ÷ dny měsíce); vyžaduje plan_start
#                       (datum dne 1 plánu). Žádný 30denní měsíc se nepředpokládá.
# - scénáře konfliktů: kombinace proveditelných a podmíněných variant z conflict_options, které
#   konflikt vyřeší. Kombinace se ověří kapacitně (společně) a započte se jen tehdy, když všechny
#   dotčené závazky zůstanou celé v období. Rozpětí = meze vzorců (econ_formula) přes tyto
#   kombinace; NIC se nevybírá. Výsledek s nevyřešenými konflikty je vždy neúplný.
# - modelová poptávka (rental_demand) se nepoužívá.

FIXED_METHODS = (None, "months", "calendar_days")
MAX_SCENARIOS = 256


@dataclass
class PeriodResult:
    """Výsledek period_result. Částky jsou economics.Result (hodnota + chybějící / neplatné údaje).

    status          … úplný | neúplný | nelze spočítat
    rentals, jobs   … jisté příspěvky pronájmů / zakázek (jen „jisté využití“)
    fixed_costs, overhead … za období podle zadané metody (nezadáno = chybějící údaj)
    result          … výsledek z jistých závazků (bez konfliktních, předběžných, přesahujících)
    conflicts       … přehled případů konfliktů (příspěvky obou stran – NEsčítat) a jejich variant
    tentative, uncertain … přehled předběžných a nejistých závazků (nezapočteno)
    scenarios       … ověřené kombinace variant konfliktů s mezemi výsledku
    range           … (dolní, horní) mez výsledku přes započitatelné scénáře; None = nelze určit
    missing         … co chybí, aby šlo o přesný výsledek
    """
    status: str
    start: object = None
    end: object = None
    rentals: object = None
    jobs: object = None
    fixed_costs: object = None
    overhead: object = None
    result: object = None
    months: object = None
    fixed_cost_basis: str = ""
    conflicts: list = field(default_factory=list)
    tentative: list = field(default_factory=list)
    uncertain: list = field(default_factory=list)
    scenarios: list = field(default_factory=list)
    range: tuple = (None, None)
    missing: list = field(default_factory=list)
    notes: list = field(default_factory=list)
    options: object = None


def period_result(model, start, end, job_starts=None, fixed_cost_method=None, months=None,
                  plan_start=None) -> PeriodResult:
    """Ekonomický výsledek flotily za období start–end (čísla dní plánu, oba včetně).

    job_starts        … návrh plánu zakázek {job_id: první den},
    fixed_cost_method … None | "months" | "calendar_days" (viz výše); months jen pro "months",
                        plan_start (datum dne 1) jen pro "calendar_days".
    Nic se nevybírá ani nemění.
    """
    opts = conflict_options(model, start, end, job_starts)
    if opts.status == fu.NOT_COMPUTABLE:
        return PeriodResult(fu.NOT_COMPUTABLE, start, end, notes=list(opts.notes), options=opts)
    led = opts.ledger
    fic = ec.is_fictitious(model)

    m, basis, problem = _months(start, end, fixed_cost_method, months, plan_start)
    if problem:
        return PeriodResult(fu.NOT_COMPUTABLE, start, end, notes=[problem], options=opts)
    if m is None:
        label = "metoda přepočtu měsíčních fixních nákladů a režie na období (nezadána)"
        fixed = ec.Result(0.0, missing=(f"fixní náklady flotily: {label}",), fictitious=fic,
                          parts={"fixní náklady flotily": None})
        overhead = ec.Result(0.0, missing=(f"režie: {label}",), fictitious=fic, parts={"režie": None})
    else:
        fixed = ec.fleet_fixed_costs(model, m)["celkem"].scaled(1, "fixní náklady flotily")
        overhead = ec.overhead(model, m)

    rentals = led.totals["jisté – pronájmy (rezervace)"]
    jobs = led.totals["jisté – zakázky (návrh plánu)"]
    base = rentals + jobs - fixed - overhead

    entries = led.entries
    excluded = []
    for state, text in ((CONFLICT, "řešení konfliktů není rozhodnuto – konfliktní závazky nejsou započteny"),
                        (CROSSES, "časové rozlišení závazků přesahujících období – nezapočteny"),
                        (UNKNOWN, "závazky, jejichž obsazenost nelze určit – nezapočteny")):
        refs = sorted(entries.loc[entries["state"] == state, "ref"])
        if refs:
            excluded.append(f"{text} ({', '.join(refs)})")
    result = ec.Result(base.value, ec._unique(base.missing, tuple(excluded)), base.invalid, base.fictitious,
                       dict(base.parts))

    cal = cp.from_model(model, job_starts=dict(job_starts or {}))
    certain_f = _certain_formula(model, cal, entries)
    costs_f = ef.combine([_cost_formula(fixed, ("fixed_costs_period",), "fixní náklady flotily za období"),
                          _cost_formula(overhead, ("overhead_period",), "režie za období")])
    scenarios, rng, scen_notes = _scenarios(opts, cal, certain_f, costs_f, start, end)

    def overview(state):
        return [dict(ref=r.ref, kind=r.kind, state=r.state, start=r.start, end=r.end, reason=r.reason,
                     contribution=r.contribution, contribution_status=r.contribution_status)
                for r in entries.itertuples() if r.state == state]

    conflicts = [dict(id=c["id"], kinds=c["kinds"], days=c["days"], services=c["services"],
                      commitments=[dict(ref=x["ref"], kind=x["kind"], contribution=x["contribution"],
                                        contribution_status=x["contribution_status"]) for x in c["commitments"]],
                      variants=[dict(key=v["key"], feasibility=v["feasibility"], resolves=v["resolves"],
                                     econ_status=v["econ_status"], bounds=v["bounds"]) for v in c["variants"]],
                      note="příspěvky obou stran konfliktu nelze sčítat – rozhodne majitel")
                 for c in opts.cases]
    missing = list(result.missing)
    status = fu.NOT_COMPUTABLE if result.invalid else (fu.COMPLETE if result.complete else fu.INCOMPLETE)
    notes = [
        "výsledek z jistých závazků: jisté příspěvky − fixní náklady flotily − režie; "
        "konfliktní, předběžné a přesahující závazky jsou zvlášť a nezapočítávají se",
        "fixní náklady obsahují i průměrný servis – servis se znovu neodečítá",
        "modelová poptávka po pronájmu (rental_demand) se nepoužívá",
        "žádná varianta konfliktu se nevybírá – rozpětí je jen přehled",
    ] + scen_notes
    if status != fu.COMPLETE:
        notes.insert(0, "NEJDE o přesný čistý zisk – výsledek je neúplný (viz chybějící údaje)")
    return PeriodResult(status, start, end, rentals, jobs, fixed, overhead, result, m, basis, conflicts,
                        overview(TENTATIVE), overview(UNKNOWN) + overview(CROSSES), scenarios, rng,
                        missing, notes, opts)


def _months(start, end, method, months, plan_start):
    """(počet měsíců | None, popis, chyba) – jen výslovně zadaná metoda, nic se nepředpokládá."""
    if method not in FIXED_METHODS:
        return None, "", f"neznámá metoda přepočtu fixních nákladů „{method}“ (povoleno: months, calendar_days)"
    if method is None:
        return None, "nezadáno – fixní náklady a režie za období nelze určit", None
    if method == "months":
        if isinstance(months, bool) or not isinstance(months, (int, float)) or not months > 0:
            return None, "", "metoda „months“ vyžaduje kladný počet měsíců (months)"
        return float(months), f"zadáno: {float(months):.4f} měsíce × měsíční náklady", None
    if not isinstance(plan_start, date):
        return None, "", "metoda „calendar_days“ vyžaduje datum prvního dne plánu (plan_start)"
    first, last = plan_start + timedelta(start - 1), plan_start + timedelta(end - 1)
    total, parts, day = 0.0, [], first
    while day <= last:
        in_month = calendar.monthrange(day.year, day.month)[1]
        month_end = date(day.year, day.month, in_month)
        n = (min(last, month_end) - day).days + 1
        total += n / in_month
        parts.append(f"{n}/{in_month} ({day.month}/{day.year})")
        day = month_end + timedelta(1)
    return total, f"kalendářní dny: {' + '.join(parts)} = {total:.4f} měsíce × měsíční náklady", None


def _certain_formula(model, cal, entries) -> ef.Contribution:
    """Vzorec jistých příspěvků (stejné závazky a stejné částky jako ledger.totals)."""
    parts = []
    for r in entries[entries["counted_as_certain"]].itertuples():
        items = [a for a in cal.assignments if a.kind == r.kind and a.ref == r.ref]
        parts.append(_formula(model, _request(model, items)))
    return ef.combine(parts)


def _cost_formula(cost: ec.Result, key, label) -> ef.Contribution:
    """Náklad jako vzorec: známá část se odečte, chybějící část je neznámý údaj (nikdy 0)."""
    if cost.invalid or cost.value is None:
        return ef.Contribution(None, {}, tuple(cost.invalid) or (f"{label}: neplatný údaj",))
    if not cost.missing:
        return ef.Contribution({ef.ONLY: ef.Formula.of(-cost.value, {})})
    unknown = ef.Unknown(key, f"{label} – chybí: " + "; ".join(cost.missing), ef.DATA, "Kč")
    return ef.Contribution({ef.ONLY: ef.Formula.of(-cost.value, {key: -1.0})}, {key: unknown})


def _scenarios(opts, cal, certain_f, costs_f, start, end) -> tuple:
    """Ověřené kombinace variant konfliktů a rozpětí výsledku (dolní, horní mez; None = nelze určit)."""
    base_f = ef.combine([certain_f, costs_f])
    if not opts.cases:
        return ([dict(variants={}, valid=True, in_period=True, formula=ef.describe(base_f),
                      bounds=ef.bounds(base_f))], ef.bounds(base_f), [])
    choices = []
    for c in opts.cases:
        usable = [v for v in c["variants"] if v["feasibility"] in (FEASIBLE, CONDITIONAL)
                  and v["resolves"] == RESOLVED and v["formula"] is not None and v["formula"].valid]
        choices.append([(c, v) for v in usable])
    notes = []
    if any(not ch for ch in choices):
        empty = [c["id"] for c, ch in zip(opts.cases, choices) if not ch]
        return [], (None, None), [f"případ {', '.join(empty)}: žádná vyčíslitelná varianta, která konflikt "
                                  "vyřeší – rozpětí výsledku nelze určit"]
    combos = list(itertools.islice(itertools.product(*choices), MAX_SCENARIOS + 1))
    if len(combos) > MAX_SCENARIOS:
        return [], (None, None), [f"příliš mnoho kombinací variant (víc než {MAX_SCENARIOS}) – rozpětí nelze určit"]
    before = {_key(c) for c in fu.machine_days(cal, *_span(cal, start, end)).conflicts}
    out = []
    for combo in combos:
        valid, in_period, reason = _check_combo(cal, combo, before, start, end)
        f = ef.combine([base_f] + [v["formula"] for _, v in combo])
        out.append(dict(variants={c["id"]: v["key"] for c, v in combo}, valid=valid, in_period=in_period,
                        reason=reason, formula=ef.describe(f), bounds=ef.bounds(f)))
    counted = [s for s in out if s["valid"] and s["in_period"]]
    if not counted:
        notes.append("žádná kombinace variant nezůstane celá v období a bez konfliktu – rozpětí nelze určit "
                     "(chybí pravidlo časového rozlišení přesahujících závazků)")
        return out, (None, None), notes
    los = [s["bounds"][0] for s in counted]
    his = [s["bounds"][1] for s in counted]
    rng = (None if any(x is None for x in los) else min(los), None if any(x is None for x in his) else max(his))
    skipped = len(out) - len(counted)
    if skipped:
        notes.append(f"{skipped} kombinací variant se do rozpětí nezapočítává (přesah mimo období nebo "
                     "společně nevyřeší konflikty)")
    return out, rng, notes


def _check_combo(cal, combo, before, start, end) -> tuple:
    """Platí kombinace variant společně? (bez konfliktu, bez nového konfliktu; závazky celé v období)."""
    assignments = list(cal.assignments)
    labels, placed = set(), []
    for case, v in combo:
        labels |= {x["label"] for x in case["commitments"]} | set(case["services"])
        ch = v["changes"]
        if v["kind"].startswith("zachovat ostatní"):
            label = f"{cp.KIND_LABEL['job']} {ch['ref']}"
        else:
            label = f"{cp.KIND_LABEL['rental']} {ch['ref']}"
        old = [a for a in assignments if a.label == label]
        assignments = [a for a in assignments if a.label != label]
        if ch.get("new_start") is not None:
            assignments += [dataclasses.replace(a, start=ch["new_start"], end=ch["new_end"]) for a in old]
    for case, _ in combo:
        for x in case["commitments"]:
            placed += [a for a in assignments if a.label == x["label"]]
    test = cp.build_calendar(cal.units, assignments, cal.open_jobs)
    lo = min([start] + [a.start for a in assignments])
    hi = max([end] + [a.end for a in assignments])
    after = fu.machine_days(test, lo, hi).conflicts
    clash = [c for c in after if _sources(c["sources"]) & labels or _key(c) not in before]
    if clash:
        return False, False, "varianty společně nevyřeší konflikty: " + "; ".join(
            f"{c['kind']} ({c['sources']})" for c in clash)
    outside = sorted({a.ref for a in placed if a.start < start or a.end > end})
    if outside:
        return True, False, "přesah mimo období: " + ", ".join(outside) + " – chybí pravidlo časového rozlišení"
    return True, True, ""
