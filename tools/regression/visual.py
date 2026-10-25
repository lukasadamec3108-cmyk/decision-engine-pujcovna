"""Vizuální regrese: pixelové porovnání záložek dvou běžících verzí aplikace (např. main × větev).

Jen vývojový nástroj – potřebuje Playwright (Chromium) a Pillow, které nejsou v requirements.txt
a aplikace je nepotřebuje.

Postup:
  1. Spusťte obě verze lokálně na různých portech, každou s vlastním lokálním .streamlit/secrets.toml
     se STEJNÝM testovacím heslem (soubor je v .gitignore – nikdy ho necommitujte):
        streamlit run app.py --server.port 8501      (verze „před“)
        streamlit run app.py --server.port 8502      (verze „po“)
  2. Heslo předejte proměnnou prostředí (nikdy ne parametrem ani v kódu):
        DE_TEST_PASSWORD=… python tools/regression/visual.py http://localhost:8501 http://localhost:8502 vystup/
     Volitelně DE_CHROMIUM=/cesta/k/chrome, DE_WIDTHS=1440,900,768,390 a DE_HEIGHT (výška okna, výchozí
     900; vyšší okno – např. 4000 – omezí skládání dlouhých prvků z více částí, při kterém se lepkavé
     hlavičky tabulek vykreslují podle polohy hranice okna).

Porovnává se obsah každé záložky (celá výška) na zvolených šířkách, navíc záložka 7 v pohledu
„Přehled využití“ a záložka 8 se vzorovým požadavkem.

Vícestránková aplikace (etapa M2) se ovládá levým menu (na mobilu se nejdřív otevře panel).
Původní záložky hledá tam, kam se přesunuly (Investice → záložky 1–5, ostatní → stránky) a navíc
nafotí nové stránky (p_…). Při porovnání verze se záložkami s vícestránkovou se pixelově porovnávají
jen pohledy se shodným výřezem (záložky Investic a kalendář kapacity); ostatní se jen uloží
(„jen snímek“) – stránka má jinou hlavičku a okolí než dřívější záložka. Před snímkem se myš přesune do rohu, aby se
nezachytily efekty najetí myší. Odchylky barvy nejvýš DE_NOISE (výchozí 2 z 255) jsou šum vyhlazování
hran v prohlížeči – vzniká i při porovnání verze se sebou samou; vypíší se, ale nejsou chybou.
Výsledek: seznam shod a rozdílů, u rozdílu obrázek „…_rozdil.png“. Návratový kód 1 při rozdílu.
"""

import os
import sys
from pathlib import Path


def _idle(p):
    p.wait_for_timeout(1500)
    p.wait_for_function("!document.querySelector('[data-testid=stStatusWidget]')", timeout=180000)
    p.wait_for_timeout(1500)


def _views(p):
    """Pořadí pohledů: (název, akce před snímkem, výřez, porovnatelné se záložkami). Akce jen kliká."""
    pages = p.locator("[data-testid=stSidebarNav]").count() > 0
    panel = "[role=tabpanel]:visible"

    def go(title):
        def act():
            expand = p.locator("[data-testid=stExpandSidebarButton]")
            opened = bool(expand.count() and expand.first.is_visible())   # sbalený panel (mobil, tablet)
            if opened:
                expand.first.click()
                p.wait_for_timeout(600)
            p.locator("[data-testid=stSidebarNav] a").filter(has_text=title).first.click()
            _idle(p)
            close = p.locator("[data-testid=stSidebarCollapseButton] button")
            if opened and not (expand.count() and expand.first.is_visible()) and close.count():
                close.first.click()                                         # panel zpět sbalit jako předtím
                p.wait_for_timeout(600)
        return act

    def tab(i):
        return lambda: p.locator("[role=tab]").nth(i).click()

    def seq(*acts):
        def act():
            for a in acts:
                a()
                _idle(p)
        return act

    def usage():
        p.locator("[role=tabpanel]:visible" if not pages else "[data-testid=stMain]").get_by_text("Přehled využití").click()

    def calendar_back():
        p.locator("[role=tabpanel]:visible" if not pages else "[data-testid=stMain]").get_by_text(
            "Kalendář dostupnosti").click()

    def request():
        p.locator(".st-key-dp_r_type input").click()
        p.locator(".st-key-dp_r_type input").fill("Pásový dumper")
        p.keyboard.press("Enter")
        _idle(p)
        p.get_by_role("button", name="Vyhodnotit samostatně").click()

    if not pages:
        views = [(f"t{i + 1}", tab(i), panel, True) for i in range(8)]
        views += [("t7_vyuziti", seq(tab(6), usage), panel, True), ("t7_kalendar_zpet", calendar_back, panel, True),
                  ("t8_pozadavek", seq(tab(7), request), panel, False)]
        return views
    main, cap = "[data-testid=stMainBlockContainer]", ".st-key-de_tab_kapacita"
    views = [(f"t{i + 1}", seq(go("Investice"), tab(i)) if i == 0 else tab(i), panel, True) for i in range(5)]
    views += [("t6", go("Vlastní zakázky"), main, False), ("t7", go("Plánování kapacity"), cap, True),
              ("t7_vyuziti", usage, cap, True), ("t7_kalendar_zpet", calendar_back, cap, True),
              ("t8", go("Rozhodovací centrum"), main, False), ("t8_pozadavek", request, main, False)]
    views += [(f"p_{k}", go(t), main, False) for k, t in (
        ("dashboard", "Dashboard"), ("pronajem", "Pronájem strojů"), ("strojovy_park", "Strojový park"),
        ("nastaveni", "Nastavení"))]
    views += [("p_investice_panel", go("Investice"), "[data-testid=stSidebar]", False)]
    return views


def shoot(url: str, width: int, out: Path, tag: str, password: str, chromium: str) -> list:
    from playwright.sync_api import sync_playwright

    files = []
    with sync_playwright() as pw:
        b = pw.chromium.launch(executable_path=chromium) if chromium else pw.chromium.launch()
        mobile = width < 500
        p = b.new_page(viewport={"width": width, "height": int(os.environ.get("DE_HEIGHT", "900"))}, is_mobile=mobile, has_touch=mobile)
        p.goto(url)
        p.wait_for_selector("input[type=password]", timeout=90000)
        p.fill("input[type=password]", password)
        p.keyboard.press("Enter")
        p.wait_for_selector(".de-header", timeout=180000)
        _idle(p)
        for name, action, area, comparable in _views(p):
            action()
            p.mouse.move(0, 0)            # bez najetí myší: jinak se např. u posuvníku ukazuje stupnice
            _idle(p)
            f = out / f"{tag}_{width}_{name}.png"
            if area == "[data-testid=stSidebar]" and width < 500:
                p.locator("[data-testid=stExpandSidebarButton]").first.click()
                _idle(p)
            p.locator(area).first.screenshot(path=str(f), animations="disabled")
            files.append((name, (f, comparable, area)))
        b.close()
    return files


def _offset(ia, ib, noise, limit=40):
    """Svislý posun (px), při kterém je menší snímek pixelově shodný s výřezem většího, jinak None.

    Stejná šířka, rozdíl výšky nejvýš `limit` px – např. stránka nemá nad obsahem prázdný blok stylů,
    který dřív byl uvnitř záložky. Jakýkoli jiný rozdíl obsahu se dál hlásí jako ROZDÍL.
    """
    from PIL import ImageChops

    big, small = (ia, ib) if ia.size[1] >= ib.size[1] else (ib, ia)
    extra = big.size[1] - small.size[1]
    if big.size[0] != small.size[0] or extra > limit:
        return None
    for off in range(extra + 1):
        part = big.crop((0, off, big.size[0], off + small.size[1]))
        hist = ImageChops.difference(part, small).convert("L").histogram()
        if max(i for i, v in enumerate(hist) if v) <= noise:
            return off
    return None


def main(argv):
    if len(argv) != 4:
        print(__doc__)
        return 2
    password = os.environ.get("DE_TEST_PASSWORD")
    if not password:
        print("Chybí proměnná prostředí DE_TEST_PASSWORD (testovací heslo obou lokálních instancí).")
        return 2
    from PIL import Image, ImageChops

    before, after, out = argv[1], argv[2], Path(argv[3])
    out.mkdir(parents=True, exist_ok=True)
    widths = [int(w) for w in os.environ.get("DE_WIDTHS", "1440,900,768,390").split(",")]
    chromium = os.environ.get("DE_CHROMIUM", "")
    noise = int(os.environ.get("DE_NOISE", "2"))      # odchylka barvy (0–255), která se bere jako šum vykreslení
    differences = 0
    for w in widths:
        a = dict(shoot(before, w, out, "pred", password, chromium))
        b = dict(shoot(after, w, out, "po", password, chromium))
        for name in sorted(set(a) | set(b), key=lambda n: (n not in a, n)):
            if name not in a or name not in b:
                print(f"{w} px {name}: jen snímek ({(a.get(name) or b.get(name))[0].name})")
                continue
            (fa, ca, aa), (fb, cb, ab) = a[name], b[name]
            if not (ca and cb) and aa != ab:          # jiný výřez (záložka × stránka) – nelze porovnat pixelově
                print(f"{w} px {name}: jen snímek – jiný výřez ({fa.name}, {fb.name})")
                continue
            ia, ib = Image.open(fa).convert("RGB"), Image.open(fb).convert("RGB")
            if ia.size != ib.size:
                off = _offset(ia, ib, noise)
                if off is None:
                    differences += 1
                    print(f"{w} px {name}: ROZDÍL – jiná velikost {ia.size} → {ib.size}")
                else:
                    print(f"{w} px {name}: shodný obsah, výřez posunutý o {off} px svisle {ia.size} → {ib.size}")
                continue
            diff = ImageChops.difference(ia, ib)
            box = diff.getbbox()
            if box is None:
                print(f"{w} px {name}: shodné pixelově")
                continue
            hist = diff.convert("L").histogram()
            changed = sum(hist[1:])
            worst = max(i for i, v in enumerate(hist) if v)
            if worst <= noise:
                # vyhlazování hran v prohlížeči: jednotky pixelů o 1–2 odstíny, okem neviditelné
                # (stejný šum vzniká i při porovnání verze se sebou samou)
                print(f"{w} px {name}: shodné (šum vykreslení: {changed} px, odchylka nejvýš {worst}/255)")
                continue
            differences += 1
            diff.save(out / f"{w}_{name}_rozdil.png")
            print(f"{w} px {name}: ROZDÍL v oblasti {box} – {changed} px, odchylka až {worst}/255")
    print("VÝSLEDEK:", "bez viditelných rozdílů" if not differences else f"{differences} rozdílů")
    return 1 if differences else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
