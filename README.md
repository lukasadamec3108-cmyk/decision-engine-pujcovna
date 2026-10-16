# Decision Engine – rozhodovací nástroj pro půjčovnu stavebních strojů

**Interaktivní webová aplikace pro hodnocení investic do stavebních strojů, analýzu rizika a plánování zakázek.**
Vytvořeno v Pythonu a Streamlitu.

> ⚠️ **Stav projektu: funkční pilotní prototyp připravovaný k odbornému a uživatelskému ověření.**
> Nejde o hotový produkční systém.
>
> Aplikace pracuje **výhradně s fiktivními demonstračními hodnotami**. Stroje, ceny, náklady,
> zakázky i všechny zobrazené výsledky slouží jen k ukázce funkčnosti a **nepředstavují
> skutečné ekonomické výsledky žádné společnosti**.

---

## Představení projektu

Decision Engine je samostatný analytický a vývojový projekt. Ukazuje, jak lze typická
manažerská rozhodnutí provozovatele stavebních strojů podpořit kvantitativním modelem
a přehlednou webovou aplikací.

Projekt vzniká jako pilotní řešení pro společnost působící v oblasti stavební mechanizace
a je připravován za odborných konzultací. Cílem současné fáze je:

- vytvořit funkční prototyp, na kterém lze přístup srozumitelně předvést,
- získat odbornou a uživatelskou zpětnou vazbu,
- připravit základ pro případné ověření na reálných datech.

## Podnikatelský problém

Provozovatel půjčovny stavebních strojů pravidelně řeší otázky, které přímo ovlivňují ziskovost:

- **Do kterého stroje investovat?** Nákup stroje je vysoká jednorázová investice, jejíž návratnost
  závisí na nejisté poptávce, ceně pronájmu, servisních nákladech a poruchovosti.
- **Jak velké je riziko ztráty?** Průměrný odhad nestačí. Je potřeba vědět, jak pravděpodobný
  je ztrátový měsíc a jak nepříznivé mohou být horší scénáře.
- **Při jaké ceně a vytížení se stroj vyplatí?** Kde leží bod zvratu a jaká je minimální denní sazba.
- **Jak rozdělit investiční rozpočet** mezi více typů strojů.
- **Jak seřadit a naplánovat zakázky**, když více zakázek potřebuje stejné stroje.

Decision Engine tyto otázky spojuje do jednoho interaktivního nástroje. Uživatel mění vstupní
předpoklady a okamžitě vidí dopad na očekávaný zisk, riziko a výsledné doporučení.

## Funkce aplikace

Vstupy se nastavují v postranním panelu:
- výběr dvou porovnávaných strojů,
- fixní měsíční náklad,
- scénář trhu (pesimistický / realistický / optimistický),
- investiční rozpočet,
- v části „Pokročilé nastavení“: odhad poptávky pro rychlý výpočet (dny v měsíci) a počet simulací,
- kapacitní scénář (nákup druhého pásového dumperu) – zatím vypnutý, zapojení je plánováno.

Výchozí hodnoty tvoří **ukázkový scénář** s fiktivními čísly (realistický trh, fixní náklad 100 000 Kč).

Výsledky jsou rozděleny do šesti záložek:

| Záložka | Obsah |
|---|---|
| 🏆 **Doporučení** | Verdikt jednou větou, karty obou strojů (očekávaný zisk, šance ztrátového měsíce, bod zvratu, minimální cena), klíčová srovnání, rychlý odhad bez simulace, bodové hodnocení a **export investičního reportu do formátu DOCX**. |
| 📈 **Možné výsledky** | Simulace měsíčního zisku (Monte Carlo): slabý / typický / dobrý měsíc (P10 / medián / P90), histogram; v detailu pro analytiky kumulativní pravděpodobnost (CDF) a interval nejistoty. |
| ⚠️ **Riziko** | Semafor rizika, šance ztrátového měsíce, nejhorší měsíc z 20 (VaR 5 %), pravděpodobnost poruchy a graf riziko vs. výnos. |
| 🎯 **Co ovlivňuje zisk** | Tornádo graf citlivosti na změnu vstupů o ±10 %, minimální denní cena pronájmu, zisk podle počtu dní (bod zvratu) a mapa ziskovosti (cena × počet dní). |
| 🛒 **Nákup strojů** | Optimální nákup strojů v rámci rozpočtu, křivka rozpočet vs. maximální zisk; v detailu kontrola lineární (LP) a diskrétní (DP) optimalizací. |
| 🗓️ **Plán zakázek** | Harmonogram (Ganttův diagram) se smluvními termíny, zakázky po termínu a penále za zpoždění, detail zakázek, vytížení strojů, úzké hrdlo a prioritní zakázka. |

Čísla jsou zobrazena v českém formátu, grafy jsou interaktivní (Plotly) a vzhled je přizpůsoben i menším obrazovkám.

Přístup do aplikace je chráněn heslem.

## Analytické metody a výpočty

**Monte Carlo simulace měsíčního výsledku** (100–10 000 simulovaných měsíců) pro každý stroj zvlášť:
- poptávka ve dnech: trojúhelníkové rozdělení v rozmezí daném scénářem trhu,
- servisní náklady: normální rozdělení (oříznuté na nezáporné hodnoty),
- poruchy: náhodný výskyt podle pravděpodobnosti poruchy, délka odstávky z Poissonova rozdělení,
- denní cena: normální rozdělení kolem ceníkové ceny (směrodatná odchylka 8 %, omezeno na ±20 %),
- měsíční zisk = skutečně pronajaté dny × denní cena − fixní náklad − servisní náklad.

**Rizikové metriky:** pravděpodobnost ztráty, percentily P10 / P50 / P90, Value at Risk (5% kvantil)
a pravděpodobnost, že jeden stroj vydělá více než druhý.

**Ekonomické ukazatele:** bod zvratu (počet dní pronájmu potřebný k pokrytí fixních nákladů),
minimální denní cena pronájmu a doba návratnosti investice (využitá ve scorecard).

**Citlivostní analýza:** jednorozměrná citlivost na cenu, vytížení, fixní a servisní náklady (±10 %)
a dvourozměrná mapa ziskovosti cena × dny.

**Optimalizace:**
- lineární programování (`scipy.optimize.linprog`, metoda HiGHS) pro rozdělení kapacity mezi dva stroje,
- diskrétní optimalizace úplným prohledáním celočíselných variant,
- výběr portfolia strojů v rámci investičního rozpočtu: prohledání kombinací 0–2 kusů
  u pěti strojů s nejvyšším očekávaným ziskem.

**Rozhodovací pravidla:** bodová scorecard (očekávaný zisk, riziko ztráty, VaR, návratnost)
a textové zdůvodnění doporučení generované **pevně danými pravidly** podle vypočtených metrik.
Aplikace nevyužívá generativní umělou inteligenci.

**Plánování zakázek:** zakázky se řadí podle rozhodovacího skóre (zisk upravený o riziko, priorita,
dostupnost strojů, penále za zpoždění) a plánují se postupně tak, aby nedošlo ke kolizi
v obsazenosti potřebných strojů.

## Ukázkový scénář rozhodování

*Všechny hodnoty v tomto scénáři jsou fiktivní a pocházejí z demonstračních dat aplikace.*

Provozovatel zvažuje, zda dokoupit **mini bagr**, nebo **pásový dumper**.

1. **Nastavení vstupů:** v postranním panelu vybere oba stroje, zadá fixní měsíční náklad,
   zvolí scénář trhu (např. realistický) a odhad poptávky v dnech za měsíc.
2. **Shrnutí:** aplikace porovná očekávaný měsíční zisk, pravděpodobnost ztráty a bod zvratu
   a označí relativně výhodnější variantu. Pokud jsou při daném nastavení obě varianty ztrátové,
   výslovně na to upozorní.
3. **Monte Carlo a Riziko:** uživatel vidí rozptyl možných výsledků. Nejen průměr, ale i nepříznivý
   scénář (P10), hodnotu v riziku a pravděpodobnost ztrátového měsíce.
4. **Citlivost:** zjistí, které vstupy mají na výsledek největší vliv, při jakém vytížení se stroj
   dostává do zisku a jaká je minimální denní cena pronájmu.
5. **Optimalizace:** ověří, zda se doporučení shoduje napříč modely, a podívá se, jakou kombinaci strojů
   by bylo výhodné pořídit v rámci zadaného rozpočtu.
6. **Zakázky:** posoudí, zda by druhý kus stroje uvolnil úzké hrdlo v plánování zakázek.
7. **Report:** stáhne investiční report ve formátu DOCX jako podklad pro další diskusi.

Výstupem je **strukturovaný podklad pro rozhodnutí**, ne automatické investiční doporučení.

## Technologie

| Oblast | Nástroje |
|---|---|
| Jazyk | Python |
| Webová aplikace | Streamlit, nasazení na Streamlit Community Cloud |
| Práce s daty | pandas, NumPy |
| Simulace a optimalizace | NumPy (náhodná rozdělení), SciPy (`linprog`) |
| Vizualizace | Plotly |
| Reporting | python-docx |
| Správa kódu a zabezpečení | Git, GitHub (větve, pull requesty), Streamlit secrets |

## Technické informace

### Struktura projektu

```
├── app.py                           # Streamlit aplikace (výpočty, záložky, export reportu)
├── ui_theme.py                      # vzhled: barvy, české formátování čísel, karty, šablona grafů
├── assets/logo_placeholder.svg      # zástupné logo (firemní logo se do repozitáře nevkládá)
├── data/
│   ├── machines_demo.csv            # fiktivní parametry 8 typů strojů
│   ├── jobs_demo.csv                # fiktivní zásobník 12 zakázek
│   └── *.xlsx                       # rozšířený fiktivní datový model (aplikace ho zatím nenačítá)
├── .streamlit/config.toml           # barevné téma aplikace
├── .streamlit/secrets.toml.example  # šablona pro nastavení přístupového hesla
├── requirements.txt
└── README.md
```

### Demonstrační data

Složka `data/` obsahuje **výhradně fiktivní hodnoty** vytvořené pro testování a ukázku:
typy strojů, denní sazby, pořizovací ceny, náklady, parametry poptávky a poruchovosti a vzorové zakázky.
Neobsahuje žádná skutečná data žádné společnosti ani osobní údaje.

### Spuštění lokálně

```bash
git clone <URL tohoto repozitáře>
cd decision-engine-pujcovna
pip install -r requirements.txt

cp .streamlit/secrets.toml.example .streamlit/secrets.toml
# v souboru .streamlit/secrets.toml nastavte vlastní heslo

streamlit run app.py
```

### Nastavení přístupového hesla

Heslo se načítá výhradně ze Streamlit secrets (klíč `password`) a není součástí zdrojového kódu.

- **Lokálně:** soubor `.streamlit/secrets.toml`. Je uvedený v `.gitignore`, takže se do repozitáře nedostane.
- **Streamlit Community Cloud:** nastavení aplikace → *Settings* → *Secrets* → řádek `password = "..."`.

### Logo společnosti

V hlavičce aplikace je připravené místo pro logo. Dokud logo není dodáno, zobrazuje se neutrální
zástupný obrázek. Firemní logo se do veřejného repozitáře **nevkládá**:

- **Streamlit Community Cloud:** do *Secrets* přidejte řádek `logo_url = "https://…"` (odkaz na obrázek).
- **Lokálně:** uložte soubor jako `assets/logo.png` (případně `.jpg` / `.svg`) – je uvedený v `.gitignore`.

### Online verze

Pilotní verze běží na adrese <https://decision-engine-pujcovna.streamlit.app>.
Aplikace je **chráněna heslem**. Přístup je poskytován pouze pro účely ověření prototypu.

## Omezení současného prototypu

- Model pracuje s **fiktivními demonstračními daty** a nebyl ověřen na skutečných provozních datech.
- Model používá zjednodušené předpoklady. Například zadaný fixní náklad je stejný pro oba porovnávané stroje
  a model nepočítá se sezónností poptávky.
- Simulace nepoužívá pevný seed, takže se výsledky při každém přepočtu mírně liší.
- Celá aplikace je zatím v jednom souboru `app.py`. Výpočty nejsou pokryty automatickými testy.
- Výstupy slouží jako **podklad pro odbornou diskusi**, ne jako investiční doporučení.

## Plánované další kroky

1. **Odborná validace** předpokladů, vstupních parametrů a metodiky.
2. **Uživatelské ověření**: předvedení prototypu budoucím uživatelům a sběr zpětné vazby.
3. **Případné ověření na reálných datech**, pouze se souhlasem dotčené společnosti
   a s odpovídajícím zabezpečením. Reálná data nebudou ukládána do veřejného repozitáře.
4. **Technický rozvoj:**
   - rozdělení kódu na moduly (výpočty / uživatelské rozhraní / exporty),
   - automatické testy výpočtů a opakovatelné simulace,
   - možnost nahrát vlastní data a ukládat projekty.

## Přínos z hlediska datové a business analytiky

Projekt propojuje programování, datovou analýzu a business analytiku na konkrétním podnikatelském problému:

- **převod obchodní otázky na kvantitativní model:** investice, riziko, kapacity,
- **stochastické modelování nejistoty** pomocí Monte Carlo simulace a práce s rizikovými metrikami (VaR, percentily),
- **optimalizační úlohy:** lineární a diskrétní optimalizace, výběr portfolia v rámci rozpočtu,
- **citlivostní analýza** jako nástroj pro identifikaci klíčových faktorů ziskovosti,
- **interaktivní aplikace pro netechnické uživatele**, kde se výsledky okamžitě přepočítávají podle předpokladů,
- **automatizovaný manažerský reporting** ve formátu DOCX,
- **profesionální vývojový postup:** Git, pull requesty a bezpečná správa přístupových údajů.
