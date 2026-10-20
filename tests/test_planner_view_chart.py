"""Testy grafu hranic zlomu v panelu „Ekonomické porovnání“ (planner_view, E2.5c krok 4D-A).

Graf smí ukazovat jen to, co plyne z již vypočtených hranic – tady se ověřuje proti skutečným
hodnotám příspěvků. Jen fiktivní testovací data.
Spuštění: python -m pytest tests   nebo   python tests/test_planner_view_chart.py
"""

import random
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))
import planner as pl  # noqa: E402
import planner_view as pv  # noqa: E402
from test_econ_formula import filled, files_hash  # noqa: E402
from test_planner import BAGR, model  # noqa: E402
from test_planner_view import DEMO, STARTS  # noqa: E402

NBSP = " "
SMYK = "SMYKOVY_NAKLADAC"


def rent(ref, s, e, price=None, t=BAGR):
    return pv.rental_request(ref, t, 1, s, e, price)


def job(ref, s, dur, revenue=150000, direct=60000):
    return pv.job_request(ref, ref, [(BAGR, 1)], s, dur, None, None, revenue, direct, None)


def panel_for(res, ref):
    return next(p for p in res["panels"] if ref in p["requests"])


def plain(html):
    return html.replace(NBSP, " ")


# ---------------------------------------------------------------------------
# 1–2) Kontrolní scénář FIKT-P07 × NP-3 a rovnost přesně v hranici
# ---------------------------------------------------------------------------

def test_fikt_p07_vs_np3_chart():
    res = pv.evaluate(DEMO, STARTS, [rent("NP-3", 20, 23, 2000, SMYK)])
    pn = panel_for(res, "NP-3")
    (block,) = pn["conditions"]
    ch = block["chart"]
    assert ch["unknown"] == "Smykový nakladač: variabilní náklad pronájmu (Kč/den)"
    assert [(s["range"], s["winners"], s["nothing"]) for s in ch["segments"]] == [
        ("pod 7 500 Kč/den", ["FIKT-P07"], False), ("nad 7 500 Kč/den", ["nepřijmout nic"], True)]
    assert ch["boundaries"] == [dict(n=1, value="7 500 Kč/den", winners=["FIKT-P07", "nepřijmout nic"], others=[],
                                     text="FIKT-P07 a nepřijmout nic vycházejí stejně (rovnocenné)")]
    assert ch["never"] == ["NP-3 není nejvýhodnější při žádné výši nákladu."]
    out = plain(pv.panel_html(pn))
    assert "Nejvýhodnější varianta podle: Smykový nakladač: variabilní náklad pronájmu &#40;Kč/den&#41;" in out
    assert "<b>Hranice 1 – přesně 7 500 Kč/den:</b> FIKT-P07 a nepřijmout nic vycházejí stejně &#40;rovnocenné&#41;." in out
    assert 'class="dp-seg nothing"' in out and "nepřijmout nic &#40;0 Kč&#41;" in out          # „nic“ je vidět
    assert "NP-3 není nejvýhodnější při žádné výši nákladu." in out
    assert "Úseky nejsou v měřítku" in out
    # slovní vysvětlení zůstává beze změny pod grafem
    assert block["lines"][:3] == ["pod 7 500 Kč/den: FIKT-P07",
                                  "přesně 7 500 Kč/den: FIKT-P07 a nepřijmout nic vycházejí stejně (rovnocenné)",
                                  "nad 7 500 Kč/den: nepřijmout nic"]
    # skutečné hodnoty: 7 499 → FIKT-P07, 7 500 → shoda s „nic“, 7 501 → „nic“
    p = pl.plan(DEMO, STARTS, [pv.to_request(rent("NP-3", 20, 23, 2000, SMYK))])
    p07 = p.requests["FIKT-P07"]["request"]
    assert [pl.contribution(filled(DEMO, rental={SMYK: x}), p07).lo for x in (7499, 7500, 7501)] == [5.0, 0.0, -5.0]


def test_three_way_tie_at_boundary_is_explained():
    pn = panel_for(pv.evaluate(DEMO, STARTS, [rent("NP-1", 22, 23, None, "PASOVY_DUMPER")]), "NP-1")
    ch = pn["conditions"][0]["chart"]
    assert ch["boundaries"][0]["text"] == "FIKT-P06, NP-1 a nepřijmout nic vycházejí stejně (rovnocenné)"
    assert ch["never"] == ["NP-1 není nejvýhodnější při žádné výši nákladu (stejně jako nejlepší jen přesně při "
                           "8 000 Kč/den)."]


# ---------------------------------------------------------------------------
# 3–4) Desetinná hranice, více hranic a variant
# ---------------------------------------------------------------------------

def test_decimal_boundary_is_not_rounded():
    pn = panel_for(pv.evaluate(model(), {}, [rent("NP-1", 1, 7, 6100), rent("NP-2", 1, 4, 9000)]), "NP-1")
    ch = pn["conditions"][0]["chart"]
    assert [b["value"] for b in ch["boundaries"]] == ["2 233,33 Kč/den", "9 000 Kč/den"]
    assert [s["range"] for s in ch["segments"]] == ["pod 2 233,33 Kč/den", "mezi 2 233,33 a 9 000 Kč/den",
                                                    "nad 9 000 Kč/den"]


def test_multiple_boundaries_and_variants():
    pn = panel_for(pv.evaluate(model(), {}, [rent("NP-1", 1, 5, 6500), rent("NP-2", 1, 3, 7000),
                                             rent("NP-3", 4, 5, 6000)]), "NP-1")
    ch = pn["conditions"][0]["chart"]
    assert [(s["range"], s["label"]) for s in ch["segments"]] == [
        ("pod 6 000 Kč/den", "NP-2 + NP-3"), ("mezi 6 000 a 7 000 Kč/den", "NP-2"),
        ("nad 7 000 Kč/den", "nepřijmout nic (0 Kč)")]
    assert [(b["n"], b["value"]) for b in ch["boundaries"]] == [(1, "6 000 Kč/den"), (2, "7 000 Kč/den")]
    assert ch["boundaries"][0]["text"] == "NP-2 + NP-3 a NP-2 vycházejí stejně (rovnocenné)"
    out = plain(pv.chart_html(ch))
    assert out.count('class="dp-bnd"') == 2 and len(re.findall(r'class="dp-seg[ "]', out)) == 3
    # barvy se neopakují pro různé varianty v jednom grafu a „nic“ nemá barvu varianty
    assert [s["color"] for s in ch["segments"]] == [0, 1, None]


def test_many_tied_variants_are_condensed_but_complete():
    # čtyři jednodenní pronájmy á 7 000 Kč: při 7 000 Kč/den vychází každá jejich kombinace stejně (0 Kč)
    reqs = [rent(f"NP-{i}", i, i, 7000) for i in range(1, 5)] + [rent("NP-5", 1, 5, 6000)]
    pn = panel_for(pv.evaluate(model(), {}, reqs), "NP-5")
    ch = pn["conditions"][0]["chart"]
    b2 = ch["boundaries"][1]
    assert b2["value"] == "7 000 Kč/den" and len(b2["winners"]) > 10
    assert b2["text"].startswith("NP-1 + NP-2 + NP-3 + NP-4 a nepřijmout nic vycházejí stejně (rovnocenné)")
    assert f"stejně vychází i dalších {len(b2['others'])} variant" in b2["text"]
    assert pv._variants(1) == "další 1 varianta" and pv._variants(3) == "další 3 varianty"
    assert set(b2["others"]) | {"NP-1 + NP-2 + NP-3 + NP-4", "nepřijmout nic"} == set(b2["winners"])
    assert len(ch["never"]) == pv.MAX_CHART_NEVER and ch["never_more"]          # zbytek sbalený, nic se neztratí
    out = plain(pv.chart_html(ch))
    assert "Zobrazit všechny shodné varianty" in out and "Další varianty, které nejsou nejvýhodnější nikdy" in out
    for w in b2["winners"]:
        assert w in out                                                       # úplný výčet je v rozbalení


# ---------------------------------------------------------------------------
# 5–6) Bez vypočitatelné hranice, více neznámých
# ---------------------------------------------------------------------------

def test_no_chart_without_computable_boundary():
    for m, reqs in ((model(var_rental=1500, var_job=2000, includes="ne"), [job("NP-1", 1, 10), rent("NP-2", 1, 5, 6500)]),
                    (model(var_rental=7000), [rent("NP-1", 1, 3, 6000), rent("NP-2", 2, 4, 6500)]),
                    (model(var_rental=500), [rent("NP-1", 1, 3, 500), rent("NP-2", 2, 4, 0)]),
                    (model(var_rental=0, includes="ano"),
                     [pv.job_request("NP-1", "J", [(BAGR, 1)], 1, 3, None, None, None, 1000, None),
                      rent("NP-2", 2, 4, 6500)])):
        res = pv.evaluate(m, {}, reqs)
        for pn in res["panels"]:
            assert pn["kind"] != "conditional"
            assert "dp-chart" not in pv.panel_html(pn)                       # jistý výsledek → žádný graf
    assert pv.profile_chart(dict(unknown=None, segments=[dict(start=0.0, end=None, winners=["A"])], boundaries=[],
                                 at_zero=["A"]), []) is None
    one = dict(unknown=dict(label="x/den", unit=pv.KC_DEN), segments=[dict(start=0.0, end=None, winners=["A"])],
               boundaries=[], at_zero=["A"])
    assert pv.profile_chart(one, [("A",)]) is None                          # vítěz nezávisí na nákladu


def test_multiple_unknowns_show_no_misleading_chart():
    pn = panel_for(pv.evaluate(model(), {}, [job("NP-1", 1, 10), rent("NP-2", 1, 5, 6500),
                                             rent("NP-3", 6, 10, 7000)]), "NP-1")
    blocks = {b["title"]: b for b in pn["conditions"]}
    ne = blocks["Pokud přímé náklady zakázek neobsahují stroje"]
    assert ne["chart"] is None and ne["chart_note"] == pv.CHART_MULTI_NOTE
    ano = blocks["Pokud přímé náklady zakázek obsahují stroje"]
    assert ano["chart"] is None and ano["chart_note"] == ""               # jeden vítěz při jakékoli výši
    out = plain(pv.panel_html(pn))
    assert "Graf se nezobrazuje: rozhoduje víc nezadaných nákladů najednou" in out and "dp-bar" not in out


# ---------------------------------------------------------------------------
# 7) Graf nesmí říkat nic jiného než skutečný výpočet (náhodně)
# ---------------------------------------------------------------------------

def _num(text):
    return float(text.replace(" Kč/den", "").replace(" ", "").replace(NBSP, "").replace(",", "."))


def test_chart_matches_true_values_random():
    rnd = random.Random(404)
    charts = 0
    for _ in range(40):
        reqs = []
        for i in range(rnd.randint(2, 4)):
            s = rnd.randint(1, 6)
            reqs.append(rent(f"NP-{i + 1}", s, s + rnd.randint(0, 4), rnd.choice([None, 2000, 4500, 6100, 9000])))
        m = model(units=rnd.randint(1, 2))
        res = pv.evaluate(m, {}, reqs)
        p = pl.plan(m, {}, [pv.to_request(r) for r in reqs])
        by = {r["ref"]: pv.to_request(r) for r in reqs}
        for pn in res["panels"]:
            for block in pn["conditions"]:
                ch = block.get("chart")
                if not ch:
                    continue
                charts += 1
                g = next(g for g in p.groups if set(g["requests"]) == set(pn["requests"]))

                def best_at(x):
                    vals = {v: sum(pl.contribution(filled(m, rental={BAGR: x}), by[r]).lo for r in v)
                            for v in g["candidates"]}
                    top = max(vals.values())
                    return {pv._vlabel(v) for v, val in vals.items() if abs(val - top) < 1e-6}
                cuts = [_num(b["value"]) for b in ch["boundaries"]]
                assert cuts == sorted(cuts) and len(cuts) == len(ch["segments"]) - 1
                edges = [0.0] + cuts + [None]
                for i, seg in enumerate(ch["segments"]):
                    lo, hi = edges[i], edges[i + 1]
                    probe = lo + 1.0 if hi is None else (lo + hi) / 2
                    assert set(seg["winners"]) == best_at(probe)                # vítěz úseku
                    assert seg["nothing"] == (seg["winners"] == ["nepřijmout nic"])
                prof = next(x for x in res["groups"] if set(x["requests"]) == set(pn["requests"]))["economy"]
                at = {round(x["at"], 2): x["at"] for x in next(pr for pr in prof["profiles"])["boundaries"]}
                for b in ch["boundaries"]:
                    assert len(b["winners"]) >= 2                               # v hranici vždy rovnost
                    if "," not in b["value"]:                                   # celá hranice → přesně ověřitelná
                        assert set(b["winners"]) == best_at(_num(b["value"]))
                    assert abs(_num(b["value"]) - at[round(_num(b["value"]), 2)]) < 0.006
                # graf a slovní pásma říkají totéž
                for seg in ch["segments"]:
                    assert any(line.startswith(seg["range"] + ":") for line in block["lines"])
    assert charts > 10


def test_tie_at_non_terminating_boundary_is_kept():
    # hranice 12 500 / 3 = 4 166,666… Kč/den: rovnost NP-4 a NP-1 se nesmí ztratit zaokrouhlením
    import econ_formula as ef
    key = ("variable_cost_rental_day", BAGR)
    prof = ef.profile({"NP-4": ef.Formula(30500.0, ((key, -5.0),)), "NP-1": ef.Formula(18000.0, ((key, -2.0),)),
                       "nepřijmout nic": ef.ZERO})
    assert prof["boundaries"] == [(4166.666667, ("NP-4", "NP-1")), (9000.0, ("NP-1", "nepřijmout nic"))]
    assert ef.rate_text(prof["boundaries"][0][0]) == "4 166,67 Kč"


# ---------------------------------------------------------------------------
# 8) Vzhled: mobil, bez zelené „doporučeno“, escapování, data
# ---------------------------------------------------------------------------

def test_chart_css_mobile_and_no_recommend_green():
    css = pv.CSS
    assert "@media (max-width: 640px)" in css and ".dp-bar{flex-direction:column}" in css
    chart_css = "".join(line for line in css.splitlines() if line.startswith((".dp-seg", ".dp-bnd", ".dp-chart")))
    assert "#1E7B4A" not in chart_css                                     # zelená jen pro „Doporučeno“
    seg_rule = css.split(".dp-seg{")[1].split("}")[0]
    assert ".dp-seg{flex:1 1 0;min-width:0" in css
    assert "overflow-wrap:break-word" in seg_rule and "anywhere" not in seg_rule    # slova se nelámou uprostřed


def test_many_segments_stack_vertically_on_medium_width():
    # víc než 3 úseky: na střední šířce (do 900 px) svisle; do 3 úseků zůstává vodorovně až do 640 px
    css = pv.CSS
    medium = css.split("@media (max-width: 900px){")[1].split("\n}")[0]
    assert ".dp-chart.many .dp-bar{flex-direction:column}" in medium
    five = panel_for(pv.evaluate(model(), {}, [rent("NP-1", 1, 6, 6000), rent("NP-2", 1, 4, 7500),
                                               rent("NP-3", 1, 2, 11000), rent("NP-4", 1, 1, 15000)]), "NP-1")
    ch = five["conditions"][0]["chart"]
    assert len(ch["segments"]) == 5 and pv.chart_html(ch).startswith('<div class="dp-chart many"')
    two = panel_for(pv.evaluate(DEMO, STARTS, [rent("NP-3", 20, 23, 2000, SMYK)]), "NP-3")["conditions"][0]["chart"]
    assert len(two["segments"]) == 2 and pv.chart_html(two).startswith('<div class="dp-chart" ')
    # obsah grafu se uspořádáním nemění (stejné úseky, hranice i texty)
    assert [s["range"] for s in ch["segments"]] == ["pod 3 000 Kč/den", "mezi 3 000 a 4 000 Kč/den",
                                                    "mezi 4 000 a 7 000 Kč/den", "mezi 7 000 a 15 000 Kč/den",
                                                    "nad 15 000 Kč/den"]


def test_chart_html_escapes_text():
    ch = dict(unknown="<b>x</b>", summary="<i>", never=["<script>"], zero_note="", boundaries=[
        dict(n=1, value="1 Kč/den", winners=["a", "b"], text="[a](https://example.com)")],
        segments=[dict(range="pod 1 Kč/den", winners=["<u>"], tie=False, nothing=False, color=0, label="<u>"),
                  dict(range="nad 1 Kč/den", winners=["nepřijmout nic"], tie=False, nothing=True, color=None,
                       label="nepřijmout nic (0 Kč)")])
    out = pv.chart_html(ch)
    assert "<b>x" not in out and "<script>" not in out and "<u>" not in out and "](https://example.com)" not in out


def test_data_unchanged():
    before = files_hash()
    pv.evaluate(DEMO, STARTS, [rent("NP-3", 20, 23, 2000, SMYK)])
    assert files_hash() == before


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK", name)
