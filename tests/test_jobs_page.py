"""Stránka Vlastní zakázky (modernizace UI): stejný plán a výpočty, nový vzhled.

KPI, srovnání, harmonogram, prioritní zakázka a tabulky jen zobrazují výsledky scheduling (plan_jobs,
summary, evaluate, conflicts, utilization) – test ověřuje shodu hodnot mezi částmi stránky a s plánem.
Heslo je náhodné, jen pro test.
"""

import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))
from app_pages import open_page  # noqa: E402
from test_dashboard import dash_kpi, plain  # noqa: E402
from test_navigation import app, ok  # noqa: E402

_STATE = {}


def jobs_app():
    if "at" not in _STATE:
        at = app()
        open_page(at, "zakazky")
        ok(at)
        _STATE["at"] = at
    return _STATE["at"]


def texts(at) -> str:
    return " ".join(m.value for m in at.main.markdown if "<style>" not in m.value)


def compare_rows(at) -> dict:
    html_all = [m.value for m in at.main.markdown if 'class="dd-cmp"' in m.value][0]
    rows = {}
    for tr in re.findall(r"<tr>(.*?)</tr>", html_all.split("<tbody>")[1]):
        cells = [plain(c) for c in re.findall(r"<td[^>]*>(.*?)</td>", tr)]
        rows[cells[0]] = cells[1:]
    return rows


def test_kpis_match_plan_and_comparison():
    at = jobs_app()
    assert dash_kpi(at, "Zisk zakázek po penále") == "892 000 Kč"            # demo plán beze změny
    assert dash_kpi(at, "Zakázky v termínu") == "9z 12"
    assert dash_kpi(at, "Délka plánu") == "20 dní"
    rows = compare_rows(at)
    assert rows["Zisk po penále"][1] == "892 000 Kč" and rows["Zakázky po termínu"][1] == "3 z 12"
    assert rows["Délka plánu"][1] == "20 dní"
    old = int(rows["Zisk po penále"][0].replace(" Kč", "").replace(" ", ""))
    gain = 892000 - old
    assert dash_kpi(at, "Zlepšení proti původnímu plánu") == "+" + f"{gain:,}".replace(",", " ") + " Kč"
    assert rows["Zisk po penále"][2] == "✓ +" + f"{gain:,}".replace(",", " ") + " Kč"   # změna se symbolem
    assert f"o {gain:,} Kč výhodnější".replace(",", " ") in plain(texts(at))
    assert "12 naplánovaných zakázek" in texts(at)


def test_warnings_rentals_service_and_not_net_profit():
    shown = plain(texts(jobs_app()))
    assert "Plán zatím nezohledňuje pronájmy a servis" in shown
    assert "jen vlastní zakázky" in shown and "nejde o čistý zisk půjčovny" in shown
    assert "Nejde o čistý zisk firmy" in shown                                   # metodika beze změny


def test_best_job_uses_original_contribution():
    at = jobs_app()
    contrib = [d.value for d in at.main.dataframe if "Čistý přínos" in d.value.columns][0]
    best = contrib.iloc[0]
    box = plain(" ".join(m.value for m in at.main.markdown if 'class="dd-kv"' in m.value
                         or ("dd-rows" in m.value and "Čistý přínos pro plán" in m.value)))
    assert f"★ {best['Zakázka']}" in box
    assert plain(best["Čistý přínos"]) in box and plain(best["Zisk"]) in box


def test_detail_table_and_status_column():
    at = jobs_app()
    detail = [d.value for d in at.main.dataframe if "Konec (den)" in d.value.columns][0]
    assert list(detail.columns[:7]) == ["Zakázka", "Stav", "Začátek (den)", "Konec (den)", "Termín (den)",
                                        "Zpoždění (dny)", "Zisk po penále"]
    assert len(detail) == 12
    assert ((detail["Zpoždění (dny)"] > 0) == (detail["Stav"] == "✕ Po termínu")).all()
    assert (detail["Stav"] == "✕ Po termínu").sum() == 3
    assert (detail["Konec (den)"] - detail["Termín (den)"]).clip(lower=0).tolist() == detail["Zpoždění (dny)"].tolist()


def test_charts_show_on_time_late_deadlines_and_bottleneck():
    at = jobs_app()
    charts = [json.loads(c.proto.spec) for c in at.get("plotly_chart")]
    assert len(charts) == 2
    gantt, usage = charts
    names = [t.get("name") for t in gantt["data"]]
    assert names == ["✓ Včas", "✕ Po termínu", "Termín"]
    late = [t for t in gantt["data"] if t.get("name") == "✕ Po termínu"][0]
    assert all(x.startswith("+") and x.endswith(" d") for x in late["text"]) and len(late["text"]) == 3
    assert any(str(y).startswith("★ ") for t in gantt["data"] for y in t["y"])
    colors = usage["data"][0]["marker"]["color"]
    assert colors[0] != colors[1]                                                # zvýrazněné úzké hrdlo


def test_methodology_expander_kept():
    at = jobs_app()
    exp = [e for e in at.main.expander if e.label == "Předpoklady a metodika"][0]
    text = " ".join(m.value for m in exp.markdown)
    for item in ("Cíl plánu", "Pojmy", "Stroje", "Jak se plán hledá", "Co model nezahrnuje"):
        assert f'<div class="t">{item}</div>' in text
    assert any(e.label == "Přínos všech zakázek pro plán" for e in at.main.expander)
