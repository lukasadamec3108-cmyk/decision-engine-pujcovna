"""Ekonomická evidence závazků za období (etapa E2.2 – K3.1: period_ledger). Bez Streamlitu.

Propojuje kapacitní fakta z fleet_usage.py (nad capacity.py) s příspěvky z economics.py.
Nic nového nepočítá jinak: strojodny a konflikty jsou z fleet_usage, částky z economics.
Zatím se nikde nezobrazuje a nic nerozhoduje (další kroky K3 – varianty konfliktů, výsledek
období, porovnání využití – tu záměrně NEJSOU).

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

from dataclasses import dataclass, field

import pandas as pd

import capacity as cp
import economics as ec
import fleet_usage as fu

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
