"""Testy zobrazovací vrstvy „Ekonomika flotily“ (E2.2 – K3.4a, fleet_view.py). Jen fiktivní data.

Spuštění: python -m pytest tests/test_fleet_view.py
"""

import ast
import json
import sys
from datetime import date
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))
import capacity_view as cv  # noqa: E402
import economics as ec  # noqa: E402
import fleet_economics as fe  # noqa: E402
import fleet_view as fv  # noqa: E402
from formatting import NBSP  # noqa: E402
from test_fleet_economics import BAGR, BASE, DUMPER, STARTS, booking, job, model, need  # noqa: E402
from test_period_result import full  # noqa: E402

NAMES = cv.type_names(BASE)
K = NBSP + "Kč"


def view(m, start=1, end=10, starts=None, **kw):
    return fv.build(fe.period_result(m, start, end, starts, **kw), cv.type_names(m))


def months(m, start=1, end=10, starts=None, n=1.0):
    return view(m, start, end, starts, fixed_cost_method="months", months=n)


def texts(obj):
    """Všechny texty ve výstupu (rekurzivně)."""
    if isinstance(obj, str):
        yield obj
    elif isinstance(obj, dict):
        for v in obj.values():
            yield from texts(v)
    elif isinstance(obj, (list, tuple)):
        for v in obj:
            yield from texts(v)


def statuses(obj):
    if isinstance(obj, dict):
        if "status" in obj and "tone" in obj:
            yield obj["status"]
        for v in obj.values():
            yield from statuses(v)
    elif isinstance(obj, (list, tuple)):
        for v in obj:
            yield from statuses(v)


def kpi(v, key):
    return next(k for k in v["kpis"] if k["key"] == key)


def row(v, item):
    return next(r for r in v["breakdown"] if r["item"] == item)


# ---------------------------------------------------------------------------
# Částky a stavy
# ---------------------------------------------------------------------------

def test_money_formatting():
    assert fv.money(0) == "0" + K and fv.money(389500) == f"389{NBSP}500" + K
    assert fv.money(-20000) == f"−20{NBSP}000" + K and fv.money(None) == "–"


def test_contribution_states():
    assert fv.contribution_cell(ec.Result(1000.0))["status"] == fv.COMPLETE
    assert fv.contribution_cell(ec.Result(0.0))["text"] == "0" + K               # skutečná nula
    upper = fv.contribution_cell(ec.Result(1000.0, missing=("Mini bagr: variabilní náklad pronájmu/den",)))
    assert upper["status"] == fv.UPPER and upper["text"] == f"nejvýše 1{NBSP}000" + K
    revenue = fv.contribution_cell(ec.Result(0.0, missing=("Zakázka: tržba",)))
    assert revenue["status"] == fv.MISSING and revenue["text"] == "nelze určit"     # bez tržby ani horní mez
    price = fv.contribution_cell(ec.Result(0.0, missing=("Mini bagr: cena pronájmu/den",)))
    assert price["status"] == fv.MISSING
    bad = fv.contribution_cell(ec.Result(None, invalid=("tržba = „abc“",)))
    assert bad["status"] == fv.MISSING and "abc" in bad["explanation"]
    assert fv.contribution_cell(None)["status"] == fv.MISSING


def test_cost_states_distinguish_zero_from_missing():
    zero = fv.cost_cell(ec.Result(0.0))
    assert zero["status"] == fv.COMPLETE and zero["text"] == "0" + K
    missing = fv.cost_cell(ec.Result(0.0, missing=("společná režie/měs.",)))
    assert missing["status"] == fv.MISSING and missing["text"] == "nezadáno" and "0" not in missing["text"]
    part = fv.cost_cell(ec.Result(5000.0, missing=("Bagr: nerozdělené fixní/měs.",)))
    assert part["status"] == fv.MISSING and part["text"].startswith("nejméně")
    assert fv.cost_cell(ec.Result(None, invalid=("x",)))["status"] == fv.MISSING


def test_bounds_states():
    assert fv.bounds_cell(100.0, 100.0)["status"] == fv.COMPLETE
    assert fv.bounds_cell(None, 100.0)["status"] == fv.UPPER
    assert fv.bounds_cell(50.0, 100.0)["explanation"].startswith("rozpětí")
    cond = fv.bounds_cell(None, 100.0, conditional=True, conditions=["souhlas zákazníka"])
    assert cond["status"] == fv.CONDITIONAL and "souhlas zákazníka" in cond["explanation"]
    assert fv.bounds_cell(10.0, None)["status"] == fv.MISSING
    assert fv.bounds_cell(None, None)["text"] == "nelze určit"


def test_day_texts():
    assert fv.days_text([15, 16, 17, 18]) == "dny 15–18" and fv.days_text([11]) == "den 11"
    assert fv.days_text([3, 5]) == "dny 3, 5" and fv.days_text([]) == "–" and fv.day_text(None) == "–"
    assert fv.day_text(3, date(2026, 1, 30)) == "den 3 (1. 2. 2026)"            # přes hranici měsíce
    assert fv.days_text([1, 2, 3], date(2026, 1, 30)) == "dny 1 (30. 1. 2026)–3 (1. 2. 2026)"


def test_warning_and_basis_texts():
    w = ["SMYKOVY_NAKLADAC den 20: kdyby se potvrdily předběžné nároky: pronájem FIKT-P07",
         "SMYKOVY_NAKLADAC den 21: kdyby se potvrdily předběžné nároky: pronájem FIKT-P07", "jiné"]
    assert fv.warnings_text(w, NAMES) == [
        "Smykový nakladač, dny 20–21: kdyby se potvrdily předběžné nároky: pronájem FIKT-P07", "jiné"]
    assert fv.basis_text("zadáno: 0.8333 měsíce") == "zadáno: 0,8333 měsíce" and fv.basis_text(None) == ""


# ---------------------------------------------------------------------------
# Celkový pohled
# ---------------------------------------------------------------------------

def test_complete_result_without_conflicts():
    m = full(bookings=[booking("P1", BAGR, 1, 4)], jobs=[job("J1", 3, deadline=4)], needs=[need("J1", DUMPER)])
    v = months(m, starts={"J1": 1})
    assert v["status"] == fv.COMPLETE and v["computable"]
    assert not any(fv.NOT_NET_PROFIT in b["text"] for b in v["banners"])
    res = row(v, "Výsledek z jistých závazků")
    expected = 4 * (6500 - 500) + (100000 - 40000) - 127000 - 20000
    assert res["status"] == fv.COMPLETE and res["text"] == fv.money(expected) and res["total"]
    assert row(v, "Fixní náklady flotily za období")["text"] == "−" + fv.money(127000)
    assert kpi(v, "konflikty")["text"] == "0 případů" and kpi(v, "konflikty")["status"] == fv.COMPLETE
    assert v["missing"] == [] and v["method"]["chosen"]
    assert v["scenarios"]["applicable"] is False and v["scenarios"]["cell"]["status"] == fv.COMPLETE


def test_real_zero_overhead_versus_missing_overhead():
    zero = months(full(overhead="0", bookings=[booking("P1", BAGR, 1, 4)]))
    empty = months(full(overhead=None, bookings=[booking("P1", BAGR, 1, 4)]))
    assert row(zero, "Režie za období")["status"] == fv.COMPLETE and row(zero, "Režie za období")["text"] == "0" + K
    assert zero["status"] == fv.COMPLETE
    assert row(empty, "Režie za období")["status"] == fv.MISSING and row(empty, "Režie za období")["text"] == "nezadáno"
    assert empty["status"] == fv.UPPER and row(empty, "Výsledek z jistých závazků")["text"].startswith("nejvýše")
    assert any(fv.NOT_NET_PROFIT in b["text"] for b in empty["banners"])
    assert "společná režie/měs." in empty["missing"][0]["items"]


def test_no_method_means_costs_are_missing_and_result_is_upper_bound():
    v = view(full(bookings=[booking("P1", BAGR, 1, 4)]))
    assert row(v, "Fixní náklady flotily za období")["text"] == "nezadáno"
    assert row(v, "Výsledek z jistých závazků")["status"] == fv.UPPER
    assert any("metoda přepočtu" in b["text"] for b in v["banners"]) and not v["method"]["chosen"]
    assert any(g["group"] == "Nastavení" for g in v["missing"])


def test_missing_revenue_is_not_an_upper_bound():
    m = full(jobs=[job("J1", 3, revenue="")], needs=[need("J1", BAGR)])
    v = months(m, starts={"J1": 1})
    assert row(v, "Jisté příspěvky vlastních zakázek")["status"] == fv.MISSING
    assert row(v, "Výsledek z jistých závazků")["text"] == "nelze určit"
    assert v["status"] == fv.MISSING


def test_complete_costs_but_unresolved_conflict_is_not_complete():
    m = full(bookings=[booking("P1", BAGR, 1, 5), booking("P2", DUMPER, 1, 2)],
             jobs=[job("J1", 3, deadline=10)], needs=[need("J1", BAGR)])
    v = months(m, end=12, starts={"J1": 4})
    res = row(v, "Výsledek z jistých závazků")
    assert res["status"] == fv.MISSING and "bez nezapočítaných závazků" in res["text"]
    assert "nezapočítáno: J1, P1" in res["explanation"]
    assert kpi(v, "konflikty")["status"] == fv.EXCLUDED
    [case] = v["conflicts"]
    assert all(c["status"] == fv.EXCLUDED for c in case["commitments"])
    assert [x["title"] for x in case["variants"]] == ["Posun zakázky", "Zrušení pronájmu", "Zkrácení pronájmu",
                                                      "Externí zapůjčení / náhradní stroj",
                                                      "Ponechat k ručnímu rozhodnutí"]
    shift = case["variants"][0]
    assert shift["amount"]["status"] == fv.COMPLETE and shift["feasibility"] == "proveditelná"
    cancel = case["variants"][1]
    assert cancel["amount"]["status"] == fv.CONDITIONAL and "souhlas zákazníka" in cancel["amount"]["explanation"]
    assert case["variants"][-1]["amount"]["status"] == fv.EXCLUDED
    assert v["scenarios"]["applicable"] and v["scenarios"]["cell"]["status"] == fv.CONDITIONAL
    assert {g["group"] for g in v["missing"]} == {"Rozhodnutí majitele"}


def test_invalid_inputs_are_reported_not_hidden():
    bad_period = fv.build(fe.period_result(BASE, 5, 1, STARTS))
    assert bad_period["computable"] is False and bad_period["banners"][0]["tone"] == "bad"
    assert bad_period["kpis"] == [] and bad_period["status"] == fv.MISSING
    bad_method = fv.build(fe.period_result(BASE, 1, 10, STARTS, fixed_cost_method="months", months=0))
    assert bad_method["computable"] is False and "months" in bad_method["banners"][0]["text"]
    invalid_job = months(full(jobs=[job("J1", 3, revenue="abc")], needs=[need("J1", BAGR)]), starts={"J1": 1})
    assert invalid_job["computable"] is False
    assert fv.build(None)["computable"] is False


# ---------------------------------------------------------------------------
# Ukázková data, zákaz „zisku“, tvar výstupu
# ---------------------------------------------------------------------------

SAMPLE = fv.build(fe.period_result(BASE, 1, 25, STARTS), NAMES)
SAMPLE_M = fv.build(fe.period_result(BASE, 1, 25, STARTS, fixed_cost_method="months", months=25 / 30), NAMES)


def test_sample_data_view():
    v = SAMPLE_M
    assert v["status"] == fv.UPPER
    assert kpi(v, "vysledek")["text"] == f"nejvýše 389{NBSP}500" + K
    assert kpi(v, "prispevky")["text"] == f"nejvýše 829{NBSP}500" + K
    assert row(v, "Fixní náklady flotily za období")["text"] == f"−440{NBSP}000" + K
    assert "0,8333 měsíce" in row(v, "Fixní náklady flotily za období")["explanation"]
    assert row(v, "Režie za období")["text"] == "nezadáno"
    assert [c["title"] for c in v["conflicts"]] == [
        "Kolový bagr · den 11 · J005 × servis", "Silniční válec · dny 15–18 · FIKT-P05 × J006"]
    shift = v["conflicts"][1]["variants"][0]
    assert shift["change"].startswith("J006: dny 11–18 → dny 20–27 · zpoždění 2 → 11 dní")
    assert shift["amount"]["text"] == f"nejvýše 108{NBSP}500" + K and shift["amount"]["status"] == fv.UPPER
    assert shift["warnings"] == ["Smykový nakladač, dny 20–23: kdyby se potvrdily předběžné nároky: pronájem FIKT-P07"]
    j005 = v["conflicts"][0]["variants"][0]
    assert j005["amount"]["text"] == f"nejvýše −20{NBSP}000" + K
    assert v["scenarios"]["cell"]["text"] == "nelze určit"
    assert [p["ref"] for p in v["pending"]] == ["FIKT-P06", "FIKT-P07"]
    assert all(p["status"] == fv.EXCLUDED for p in v["pending"])
    ledger = {r["ref"]: r for r in v["ledger"]}
    assert len(ledger) == 19 and not ledger["FIKT-P05"]["counted"] and ledger["FIKT-P05"]["status"] == fv.EXCLUDED
    assert ledger["J001"]["counted"] and ledger["J001"]["status"] == fv.UPPER
    assert sum(r["capacity"] for r in v["usage"]) == 225


def test_sample_without_method():
    assert kpi(SAMPLE, "vysledek")["text"] == f"nejvýše 829{NBSP}500" + K
    assert kpi(SAMPLE, "naklady")["status"] == fv.MISSING


@pytest.mark.parametrize("v", [SAMPLE, SAMPLE_M], ids=["bez metody", "s metodou"])
def test_incomplete_result_is_never_presented_as_profit(v):
    assert any(fv.NOT_NET_PROFIT in b["text"] for b in v["banners"])
    for text in texts(v):
        low = text.lower()
        assert low.count("zisk") == low.count("nejde o čistý zisk"), text


def test_profit_wording_absent_in_all_views():
    cases = [SAMPLE, SAMPLE_M,
             months(full(bookings=[booking("P1", BAGR, 1, 4)])),
             months(full(jobs=[job("J1", 3, revenue="")], needs=[need("J1", BAGR)]), starts={"J1": 1}),
             fv.build(fe.period_result(BASE, 5, 1, STARTS))]
    for v in cases:
        for text in texts(v):
            low = text.lower()
            assert low.count("zisk") == low.count("nejde o čistý zisk"), text


def test_only_allowed_status_labels():
    for v in (SAMPLE, SAMPLE_M):
        found = set(statuses(v))
        assert found <= set(fv.STATUSES) and {fv.UPPER, fv.MISSING, fv.EXCLUDED} <= found


def test_output_is_plain_and_inputs_unchanged():
    pr = fe.period_result(BASE, 1, 25, STARTS, fixed_cost_method="months", months=1)
    before = (pr.result, pr.rentals, pr.jobs, pr.range, len(pr.scenarios), pr.options.ledger.entries.copy())
    v = fv.build(pr, NAMES)
    json.dumps(v, ensure_ascii=False)                                    # jen obyčejné hodnoty
    assert (pr.result, pr.rentals, pr.jobs, pr.range, len(pr.scenarios)) == before[:5]
    assert pr.options.ledger.entries.equals(before[5])


def test_module_is_standalone_and_not_connected_to_ui():
    tree = ast.parse((ROOT / "fleet_view.py").read_text(encoding="utf-8"))
    imported = {n.module.split(".")[0] for n in ast.walk(tree) if isinstance(n, ast.ImportFrom) and n.module}
    imported |= {a.name.split(".")[0] for n in ast.walk(tree) if isinstance(n, ast.Import) for a in n.names}
    assert "streamlit" not in imported
    # K3.4c: app.py používá zobrazovací vrstvu jen na stránce Ekonomika flotily; jinde nikde
    from app_pages import outside_economics_page
    assert outside_economics_page("fleet_view", "fv") == []
    for path in [ROOT / "navigace.py", *ROOT.glob("*_view.py"), *(ROOT / "stranky").glob("*.py")]:
        if path.name != "fleet_view.py":
            assert "fleet_view" not in path.read_text(encoding="utf-8"), path


# ---------------------------------------------------------------------------
# K3.4c – hlavní ukazatele stránky a rozpad výsledku
# ---------------------------------------------------------------------------

def head(v, key):
    return next(h for h in v["headline"] if h["key"] == key)


def comp(v, item):
    return next(r for g in v["components"] for r in g["rows"] if r["item"] == item)


def test_headline_has_four_indicators_with_meaning_period_and_state():
    v = SAMPLE_M
    assert [h["label"] for h in v["headline"]] == ["Výsledek z pronájmu strojů", "Výsledek z vlastních zakázek",
                                                    "Fixní náklady flotily", "Celkový výsledek flotily"]
    for h in v["headline"]:
        assert h["status"] in fv.STATUSES and h["meaning"] and h["period"] == "dny 1–25 (25 dní)"
    assert head(v, "pronajem")["text"] == f"nejvýše 194{NBSP}500" + K and head(v, "pronajem")["status"] == fv.UPPER
    assert head(v, "zakazky")["text"] == f"nejvýše 635{NBSP}000" + K
    assert head(v, "fixni")["text"] == f"440{NBSP}000" + K and head(v, "fixni")["status"] == fv.COMPLETE
    assert head(v, "celkem")["text"] == f"nejvýše 389{NBSP}500" + K
    assert "režie není známá" in head(v, "celkem")["meaning"]
    assert head(SAMPLE, "fixni")["text"] == "nezadáno" and head(SAMPLE, "fixni")["status"] == fv.MISSING


def test_sample_components_sum_up_without_double_counting():
    v = SAMPLE_M
    assert comp(v, "Tržby z pronájmu")["value"] == 194500 and comp(v, "Tržby z pronájmu")["status"] == fv.COMPLETE
    assert comp(v, "Variabilní náklady pronájmu")["text"] == "nezadáno"            # chybí sazby – ne 0 Kč
    assert comp(v, "Tržby ze zakázek")["value"] == 1425000
    assert comp(v, "Přímé náklady zakázek")["text"] == f"−700{NBSP}000" + K
    assert comp(v, "Penále za zpoždění")["text"] == f"−90{NBSP}000" + K
    assert comp(v, "Variabilní náklady strojů na zakázkách")["status"] == fv.MISSING
    # tržby − přímé náklady − penále = příspěvek zakázek (stejné závazky, každý jednou)
    assert 1425000 - 700000 - 90000 == comp(v, "Příspěvek vlastních zakázek")["value"] == 635000
    assert comp(v, "Fixní náklady flotily")["text"] == f"−440{NBSP}000" + K
    assert comp(v, "Celkový výsledek flotily")["value"] == 194500 + 635000 - 440000
    fixed_rows = [r for g in v["components"] for r in g["rows"] if "Fixní" in r["item"]]
    assert len(fixed_rows) == 1                                                    # fixní náklady jen jednou
    assert not any("servis" in r["item"].lower() for g in v["components"] for r in g["rows"])


def test_components_with_complete_inputs_and_true_zero():
    m = full(bookings=[booking("P1", BAGR, 1, 4)], jobs=[job("J1", 3, deadline=4)], needs=[need("J1", DUMPER)])
    v = months(m, starts={"J1": 1})
    assert comp(v, "Tržby z pronájmu")["value"] == 4 * 6500
    assert comp(v, "Variabilní náklady pronájmu")["text"] == "−" + fv.money(4 * 500)
    assert comp(v, "Penále za zpoždění")["text"] == "0" + K and comp(v, "Penále za zpoždění")["status"] == fv.COMPLETE
    # sazba stroje na zakázce je skutečná 0 (dumper) → úplná nula, ne „nezadáno“
    machine = comp(v, "Variabilní náklady strojů na zakázkách")
    assert machine["status"] == fv.COMPLETE and machine["text"] == "0" + K
    assert head(v, "celkem")["status"] == fv.COMPLETE and "režie započtena" in head(v, "celkem")["meaning"]
    assert comp(v, "Celkový výsledek flotily")["value"] == 4 * 6000 + 60000 - 127000 - 20000


def test_machine_costs_included_in_direct_costs_are_not_subtracted_again():
    m = full(jobs=[job("J1", 3)], needs=[need("J1", BAGR)])
    st = m.settings.copy()
    st.loc[st["key"] == "direct_cost_includes_machines", "value"] = "ano"
    m.settings = st
    v = months(m, starts={"J1": 1})
    machine = comp(v, "Variabilní náklady strojů na zakázkách")
    assert machine["text"] == fv.MACHINES_IN_DIRECT and machine["status"] == fv.COMPLETE
    assert comp(v, "Příspěvek vlastních zakázek")["value"] == 100000 - 40000


def test_missing_revenue_component_is_not_zero():
    v = months(full(jobs=[job("J1", 3, revenue="")], needs=[need("J1", BAGR)]), starts={"J1": 1})
    assert comp(v, "Tržby ze zakázek")["text"] == "nelze určit" and comp(v, "Tržby ze zakázek")["status"] == fv.MISSING
    assert head(v, "zakazky")["status"] == fv.MISSING


def test_part_total_marks_missing_parts():
    known = ec.Result(1000.0, parts={"tržba z pronájmu": 1200.0, "variabilní náklady pronájmu": -200.0})
    unknown = ec.Result(800.0, missing=("Mini bagr: variabilní náklad pronájmu/den",),
                        parts={"tržba z pronájmu": 800.0, "variabilní náklady pronájmu": -0.0})
    rev = fv.part_total([known, unknown], "tržba z pronájmu")
    assert rev.value == 2000 and rev.complete
    var = fv.part_total([known, unknown], "variabilní náklady pronájmu", cost=True)
    assert var.value == 200 and var.missing == ("Mini bagr: variabilní náklad pronájmu/den",)
    assert fv.cost_cell(var)["text"].startswith("nejméně")
    bad = fv.part_total([ec.Result(None, invalid=("x",))], "přímé náklady", cost=True)
    assert bad.value is None and bad.invalid == ("x",)


def test_not_computable_view_has_empty_page_sections():
    v = fv.build(fe.period_result(BASE, 5, 1, STARTS))
    assert v["headline"] == [] and v["components"] == [] and v["period"] == "neplatné období"



# ---------------------------------------------------------------------------
# K3.4d – porovnání variant a podklady pro odkazy
# ---------------------------------------------------------------------------

def test_conflict_view_has_machine_codes_jobs_and_comparison_notes():
    k1, k2 = SAMPLE["conflicts"]
    assert (k1["types"], k1["jobs"], k2["types"], k2["jobs"]) == (["KOLOVY_BAGR"], ["J005"], ["SILNICNI_VALEC"], ["J006"])
    assert k1["comparisons"] == [] and "vyčíslit lze jen jednu variantu" in k1["comparison_note"]
    assert {c["kind"] for c in k2["comparisons"]} == {"nelze porovnat"} and len(k2["comparisons"]) == 3
    assert "nelze spolehlivě porovnat" in k2["comparison_note"]
    assert "doporučení" in fv.COMPARISON_DISCLAIMER and "rozhoduje majitel" in fv.COMPARISON_DISCLAIMER


def test_comparison_with_complete_data_is_labelled_but_not_selected():
    m = full(bookings=[booking("P1", BAGR, 1, 5)], jobs=[job("J1", 3, deadline=10)], needs=[need("J1", BAGR)])
    v = months(m, end=12, starts={"J1": 4})
    [case] = v["conflicts"]
    kinds = {c["kind"] for c in case["comparisons"]}
    assert kinds <= set(fv.COMPARISON_KIND.values())
    assert "nelze porovnat" in kinds                                  # podmíněné varianty: neznámá náhrada
    assert case["comparison_note"] == ""                              # aspoň jedna dvojice je podložená/podmíněná
    assert "selected" not in case and all("doporuč" not in c["text"].lower() for c in case["comparisons"])



# ---------------------------------------------------------------------------
# K3.4e1 – předběžné a nejisté závazky, evidence závazků
# ---------------------------------------------------------------------------

def test_pending_rows_are_not_counted_and_have_reason():
    rows = SAMPLE["pending"]
    assert [r["ref"] for r in rows] == ["FIKT-P06", "FIKT-P07"]
    for r in rows:
        assert r["status"] == fv.EXCLUDED and not r["counted"] and r["reason"] == "předběžný – jen potenciál"
    assert rows[0]["machines_text"] == "Pásový dumper" and rows[0]["text"] == f"nejvýše 40{NBSP}000" + K
    assert "nejde o potvrzený výsledek" in fv.PENDING_NOTE


def test_ledger_groups_cover_every_commitment_exactly_once():
    for v in (SAMPLE, SAMPLE_M):
        refs = [r["ref"] for g in v["ledger_groups"] for r in g["rows"]]
        assert len(refs) == len(set(refs)) == len(v["ledger"]) == 19
        assert [g["state"] for g in v["ledger_groups"]] == [fe.CERTAIN, fe.CONFLICT, fe.TENTATIVE]
        for g in v["ledger_groups"]:
            for r in g["rows"]:
                assert r["state"] == g["state"]
                assert r["counted"] == (g["state"] == fe.CERTAIN)
                assert (r["status"] == fv.EXCLUDED) == (g["state"] != fe.CERTAIN)


def test_crossing_and_unknown_commitments_are_pending():
    v14 = fv.build(fe.period_result(BASE, 1, 14, STARTS), NAMES)
    assert [(r["ref"], r["state"]) for r in v14["pending"]] == [
        ("FIKT-P01", fe.CROSSES), ("FIKT-P03", fe.CROSSES), ("J009", fe.CROSSES)]
    assert all(r["status"] == fv.EXCLUDED for r in v14["pending"])
    m = model(bookings=[booking("P1", BAGR, 1, 2)], jobs=[job("J1", 2), job("J2", 2)],
              needs=[need("J1", BAGR), need("J2", BAGR)])
    v = fv.build(fe.period_result(m, 1, 10, {"J1": 4}))
    unknown = [r for r in v["pending"] if r["state"] == fe.UNKNOWN]
    assert {r["ref"] for r in unknown} >= {"J2"} and all(r["status"] == fv.EXCLUDED for r in unknown)
    j2 = next(r for r in v["ledger"] if r["ref"] == "J2")
    assert j2["period"] == "bez termínu" and j2["date_source"] == "bez termínu"


def test_machines_text_uses_names_and_quantities():
    assert fv._machines_text("MINI_BAGR×1; PASOVY_DUMPER×2", NAMES) == "Mini bagr, Pásový dumper ×2"
    assert fv._machines_text("", NAMES) == "" and fv._machines_text("X×1", None) == "X"
