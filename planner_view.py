"""Zobrazení rozhodovacího plánovače (etapa E2.5b) – záložka „Rozhodovací plánovač“.

Modul NIC nerozhoduje ani nepočítá: vše počítá planner.py (nad capacity.py a economics.py).
Tady se jen převádí formulář na požadavek a výsledek plánovače na obyčejné tabulky, štítky
a srozumitelná vysvětlení (obyčejné typy → lze uložit do mezipaměti Streamlitu).

Štítek požadavku (schválená pravidla E2.5b):
  Doporučeno     … požadavek je ve variantě, kterou plánovač ekonomicky doporučil
                   (lepší i v nejhorším případě než každá jiná proveditelná varianta),
  Proveditelné   … vejde se do kapacity; samostatný požadavek bez soupeře nikdy není
                   „Doporučeno“ – ekonomické hodnocení je uvedené zvlášť,
  Neproveditelné … kapacitu drží potvrzené závazky nebo servis,
  Nelze rozhodnout … chybí údaje, nebo se ze soupeřících požadavků vejde jen část a
                   ekonomika chybí.
Nic se neukládá, nepotvrzuje ani nemění – požadavky existují jen v relaci uživatele.
"""

import html
from typing import Optional

import pandas as pd

import datamodel as dm
import planner as pl

RECOMMENDED, FEASIBLE, NOT_FEASIBLE, UNDECIDED = "Doporučeno", "Proveditelné", "Neproveditelné", "Nelze rozhodnout"
BADGE_CLASS = {RECOMMENDED: "rec", FEASIBLE: "ok", NOT_FEASIBLE: "no", UNDECIDED: "und"}
MAX_REQUESTS = 5
REF_PREFIX = "NP-"


# ---------------------------------------------------------------------------
# Vstupy z formuláře
# ---------------------------------------------------------------------------

def md_escape(text) -> str:
    """Text (i od uživatele) bezpečně do st.markdown: bez HTML a bez markdownových odkazů či formátu."""
    out = html.escape(str(text), quote=False).replace("\\", "\\\\")
    for ch in "`*_{}[]()#+-.!|~":
        out = out.replace(ch, "\\" + ch)
    return out


def _days(n: int) -> str:
    return f"{n} " + ("den" if n == 1 else "dny" if 2 <= n <= 4 else "dní")


def _num(v) -> Optional[float]:
    s = "" if v is None else str(v).strip()
    if s in ("", "nan", "None"):
        return None
    try:
        return float(s.replace(",", "."))
    except ValueError:
        return None


def next_ref(model: dm.DataModel, existing) -> str:
    """Jedinečné označení nového požadavku (nekoliduje s daty ani s ostatními požadavky)."""
    taken = set(existing) | set(model.jobs["job_id"].astype(str)) | set(model.rental_bookings["booking_id"].astype(str))
    i = 1
    while f"{REF_PREFIX}{i}" in taken:
        i += 1
    return f"{REF_PREFIX}{i}"


def list_price(model: dm.DataModel, type_id: str) -> Optional[float]:
    """Ceník pronájmu typu stroje (None = v datech chybí)."""
    row = model.machine_types[model.machine_types["type_id"] == type_id]
    return None if row.empty else _num(row.iloc[0]["rental_price_day"])


def job_template(model: dm.DataModel, job_id: str) -> dict:
    """Údaje existující zakázky jako šablona NOVÉHO požadavku (původní zakázka se nemění).

    Chybějící hodnota zůstává None – nikdy se nenahrazuje nulou.
    """
    j = model.jobs[model.jobs["job_id"] == job_id].iloc[0]
    req = model.job_requirements
    req = req[(req["job_id"] == job_id) & (req["resource_kind"] == "machine")]
    dur = _num(j["duration_days"])
    return dict(source=job_id, name=f"{j['name']} (kopie údajů {job_id})",
                needs=[(str(r.resource_id), int(float(r.quantity))) for r in req.itertuples()],
                duration=int(dur) if dur else None,
                earliest=int(_num(j["earliest_start_day"])) if _num(j["earliest_start_day"]) else None,
                deadline=int(_num(j["deadline_day"])) if _num(j["deadline_day"]) else None,
                revenue=_num(j["revenue"]), direct_cost=_num(j["direct_cost"]),
                penalty_per_day=_num(j["penalty_per_day"]))


def rental_request(ref: str, type_id: str, qty: int, start: int, end: int, price_day=None) -> dict:
    """Požadavek na pronájem jako obyčejný slovník (do relace i do mezipaměti)."""
    return dict(ref=ref, kind="rental", needs=((type_id, int(qty)),), start=int(start), end=int(end),
                name="", earliest=None, deadline=None, price_day=_num(price_day),
                revenue=None, direct_cost=None, penalty_per_day=None)


def job_request(ref: str, name: str, needs, start: int, duration: int, earliest=None, deadline=None,
                revenue=None, direct_cost=None, penalty_per_day=None) -> dict:
    """Požadavek na vlastní zakázku. needs = [(type_id, kusů)] – prázdné řádky se vynechají."""
    merged = {}
    for t, q in needs:
        if t:
            merged[t] = merged.get(t, 0) + int(q)
    start, duration = int(start), int(duration)
    return dict(ref=ref, kind="job", needs=tuple(merged.items()), start=start, end=start + duration - 1,
                name=name or ref, earliest=None if earliest is None else int(earliest),
                deadline=None if deadline is None else int(deadline), price_day=None,
                revenue=_num(revenue), direct_cost=_num(direct_cost), penalty_per_day=_num(penalty_per_day))


def form_problems(req: dict) -> list:
    """Zjevné chyby formuláře (plánovač vstupy kontroluje i sám)."""
    out = []
    if not req["needs"]:
        out.append("Vyberte aspoň jeden stroj.")
    if req["end"] < req["start"]:
        out.append("Konec je před začátkem.")
    if req["kind"] == "job" and req["earliest"] is not None and req["deadline"] is not None \
            and req["deadline"] < req["earliest"]:
        out.append("Termín dokončení je před nejdřívějším začátkem.")
    if req["kind"] == "job" and req["earliest"] is not None and req["start"] < req["earliest"]:
        out.append("Začátek je před nejdřívějším začátkem okna zakázky.")
    return out


def to_request(d: dict) -> pl.Request:
    return pl.Request(d["ref"], d["kind"], pl.NEW, tuple(tuple(x) for x in d["needs"]), d["start"], d["end"],
                      name=d.get("name") or "", earliest=d.get("earliest"), deadline=d.get("deadline"),
                      price_day=d.get("price_day"), revenue=d.get("revenue"), direct_cost=d.get("direct_cost"),
                      penalty_per_day=d.get("penalty_per_day"))


def freeze(requests: list) -> tuple:
    """Seznam požadavků → n-tice (klíč mezipaměti)."""
    return tuple(tuple(sorted((k, tuple(map(tuple, v)) if k == "needs" else v) for k, v in r.items()))
                 for r in requests)


def unfreeze(frozen: tuple) -> list:
    return [dict(items) for items in frozen]


# ---------------------------------------------------------------------------
# Výsledek plánovače → obyčejné typy
# ---------------------------------------------------------------------------

def badge(entry: dict, group: Optional[dict]) -> tuple:
    """(štítek, proč) – štítek a srozumitelné zdůvodnění podle schválených pravidel."""
    op = entry["operational"]
    if op == pl.NOT_FEASIBLE:
        return NOT_FEASIBLE, ("Požadavek se v zadaném termínu nevejde – kapacitu drží potvrzené závazky "
                              "nebo servis (viz „Co blokuje“).")
    if op == pl.UNDETERMINED:
        return UNDECIDED, "Provozní proveditelnost nelze určit – chybí nebo jsou neplatné údaje (viz důvody)."
    if group is None or group["verdict"] == pl.ALL_FIT:
        return FEASIBLE, ("Požadavek se vejde do zbývající kapacity. Doporučení se nevydává – není s čím "
                          "porovnat; ekonomika je uvedená zvlášť.")
    if group["verdict"] == pl.RECOMMEND:
        if entry["ref"] in (group["choice"] or ()):
            return RECOMMENDED, ("Ze soupeřících požadavků se vejde jen část; varianta s tímto požadavkem je "
                                 "ekonomicky lepší i v nejhorším případě.")
        return FEASIBLE, ("Sám se vejde, ale soupeří o stejné stroje; ekonomicky lepší je jiná varianta "
                          f"({', '.join(group['choice'])}). Nic se neruší – rozhodne uživatel.")
    if group["verdict"] == pl.NONE_PAYS:
        return FEASIBLE, ("Sám se vejde, ale se soupeřícími požadavky se vejde jen část a každá taková varianta "
                          "je ztrátová – ekonomicky nejlépe vychází nepřijmout žádný z nich. Nic se neruší – "
                          "rozhodne uživatel.")
    if group["verdict"] == pl.CONDITIONAL:
        return FEASIBLE, ("Sám se vejde, ale se soupeřícími požadavky se vejde jen část a která varianta je "
                          "ekonomicky nejlepší, závisí na neznámých nákladech – hranice zlomu jsou uvedené u "
                          "soupeřících požadavků. Nic se neruší – rozhodne uživatel.")
    if group["verdict"] == pl.EQUAL:
        return FEASIBLE, ("Sám se vejde; se soupeřícími požadavky mají varianty shodný ekonomický přínos – "
                          "rozhodne uživatel.")
    return UNDECIDED, ("Sám se vejde, ale se soupeřícími požadavky se vejde jen část a ekonomicky nelze "
                       "rozhodnout – chybí údaje (viz Chybějící údaje).")


def _group_of(p: pl.Plan, ref: str) -> Optional[dict]:
    return next((g for g in p.groups if ref in g["requests"]), None)


def _operational_text(e: dict) -> str:
    if e["operational"] == pl.OK_ALONE:
        return "Provozně: vejde se do zbývající kapacity (po potvrzených závazcích a servisu)."
    if e["operational"] == pl.NOT_FEASIBLE:
        return "Provozně: nevejde se – " + e["decision"] + "."
    return "Provozně: nelze určit – " + "; ".join(e["reasons"] or [e["decision"]]) + "."


def _late_text(model: dm.DataModel, r) -> str:
    """U zakázky po termínu: kolik dní a jaké penále je v příspěvku (nebo že sazba chybí)."""
    if r.kind != "job" or r.deadline is None or r.end <= r.deadline:
        return ""
    late = r.end - r.deadline
    rate = r.penalty_per_day
    if r.in_data:
        row = model.jobs[model.jobs["job_id"] == r.ref]
        rate = _num(row.iloc[0]["penalty_per_day"]) if not row.empty else None
    if rate is None:
        return f" Zakázka končí {_days(late)} po termínu – sazba penále chybí."
    czk = f"{late * rate:,.0f}".replace(",", " ")
    return f" Zakázka končí {_days(late)} po termínu – penále {czk} Kč je v příspěvku započtené."


def _economic_text(e: dict, late: str = "") -> str:
    if e.get("economic_note"):
        return "Ekonomicky (jen informace, odděleně od provozu): " + e["economic_note"] + "." + late
    if e["operational"] == pl.OK_ALONE:
        return "Ekonomicky: nehodnoceno."
    return "Ekonomicky: nehodnotí se – požadavek není provozně proveditelný."


def evaluate(model: dm.DataModel, job_starts: dict, requests: list) -> dict:
    """Spustí planner.plan a vrátí výsledek jen z obyčejných typů (bez objektů plánovače)."""
    p = pl.plan(model, job_starts, [to_request(r) for r in requests])
    mine = [r["ref"] for r in requests]
    competing = {ref for g in p.groups if set(g["requests"]) & set(mine) for ref in g["requests"]}
    cards = []
    for ref, e in p.requests.items():
        if ref not in mine and ref not in competing:
            continue                                   # předběžné bez vztahu k požadavkům jen v přehledu
        g = _group_of(p, ref)
        label, why = badge(dict(e, ref=ref), g)
        r = e["request"]
        cards.append(dict(
            ref=ref, title=e["label"], kind=r.kind, own=ref in mine, name=r.name,
            period=(r.start, r.end), needs=[tuple(x) for x in r.needs], badge=label, why=why,
            operational=_operational_text(e), economic=_economic_text(e, _late_text(model, r)),
            # záporné volno se nezobrazuje: den je už přetížený potvrzenými závazky
            blocking=[dict(Stroj=t, Den=int(d), **{"Volné kusy": max(0, int(free)), "Potřeba": int(need),
                                                   "Co stroj drží": ("; ".join(c) or "kus není k dispozici")
                                                   + (f" – už přetíženo o {-int(free)} ks" if free < 0 else "")})
                      for t, d, free, need, c in e["problems"]],
            alternatives=[a["text"] for a in e["alternatives"]],
            note=e["note"]))
    cards.sort(key=lambda c: (not c["own"], c["ref"]))      # nejdřív požadavky uživatele
    groups = [dict(requests=list(g["requests"]), verdict=g["verdict"], text=g["text"],
                   options=[list(o) for o in g["options"]], choice=list(g["choice"]) if g.get("choice") else None)
              for g in p.groups if set(g["requests"]) & set(mine)]
    types = {t for r in requests for t, _ in r["needs"]}
    lo = min((r["start"] for r in requests), default=None)
    hi = max((r["end"] for r in requests), default=None)
    manual = []
    unit_type = dict(zip(model.fleet_units["unit_id"].astype(str), model.fleet_units["type_id"].astype(str)))
    for m in p.manual:
        c = m["conflict"]
        t = c["id"] if c["scope"] == "typ" else unit_type.get(c["id"], c["id"])
        related = t in types and lo is not None and c["days"][0] <= hi and c["days"][-1] >= lo
        manual.append(dict(text=m["text"], alternatives=[a["text"] for a in m["alternatives"]], related=related))
    tentative = [dict(ref=r.ref, type_id=r.needs[0][0], start=r.start, end=r.end)
                 for r in pl.tentative_requests(model, job_starts)]
    return dict(cards=cards, groups=groups, manual=manual, missing=list(p.missing_data), notes=list(p.notes),
                explanation=pl.explain(p), tentative=tentative)


# ---------------------------------------------------------------------------
# HTML karty
# ---------------------------------------------------------------------------

CSS = """
<style>
.dp-card{border:1px solid #E3E6EA;border-radius:10px;padding:.8rem 1rem;margin:.5rem 0 .3rem 0;background:#fff}
.dp-head{display:flex;flex-wrap:wrap;align-items:center;gap:.5rem;margin-bottom:.35rem}
.dp-title{font-weight:700;color:#1F2933}
.dp-sub{color:#4B5563;font-size:.85rem}
.dp-badge{display:inline-block;font-weight:700;font-size:.8rem;border-radius:999px;padding:.12rem .6rem}
.dp-badge.rec{background:#1E7B4A;color:#fff}
.dp-badge.ok{background:#E3F1E8;color:#14532D;border:1px solid #1E7B4A}
.dp-badge.no{background:#FDECEA;color:#8A1C12;border:1px solid #B42318}
.dp-badge.und{background:#FFF4DB;color:#7A4E00;border:1px solid #B7791F}
.dp-why{color:#1F2933;margin:.2rem 0 .35rem 0}
.dp-line{color:#374151;font-size:.9rem;margin:.1rem 0}
.dp-tag{display:inline-block;font-size:.72rem;color:#4B5563;border:1px solid #D1D5DB;border-radius:4px;padding:0 .3rem}
</style>
"""


def _h(text) -> str:
    """Text do HTML karty: escapování HTML a závorek (ani Markdown z něj neudělá odkaz)."""
    return html.escape(str(text)).replace("[", "&#91;").replace("]", "&#93;").replace("(", "&#40;").replace(")", "&#41;")


def card_html(card: dict, names: dict) -> str:
    machines = ", ".join(f"{names.get(t, t)} × {q}" for t, q in card["needs"])
    origin = "váš požadavek" if card["own"] else "předběžná rezervace z dat – neruší se"
    lines = "".join(f'<div class="dp-line">{_h(x)}</div>' for x in (card["operational"], card["economic"]))
    return (f'<div class="dp-card"><div class="dp-head"><span class="dp-badge {BADGE_CLASS[card["badge"]]}">'
            f'{_h(card["badge"])}</span><span class="dp-title">{_h(card["title"])}</span>'
            f'<span class="dp-tag">{_h(origin)}</span></div>'
            f'<div class="dp-sub">{_h(machines)} · dny {card["period"][0]}–{card["period"][1]}'
            + (f' · {_h(card["name"])}' if card["name"] else "") + "</div>"
            f'<div class="dp-why"><b>Proč:</b> {_h(card["why"])}</div>{lines}</div>')


def blocking_table(card: dict, names: dict) -> pd.DataFrame:
    df = pd.DataFrame(card["blocking"], columns=["Stroj", "Den", "Volné kusy", "Potřeba", "Co stroj drží"])
    df["Stroj"] = df["Stroj"].map(lambda t: names.get(t, t))
    return df
