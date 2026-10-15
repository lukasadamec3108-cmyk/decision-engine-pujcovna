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
- odhad reálné poptávky (dny v měsíci),
- počet simulací,
- investiční rozpočet,
- kapacitní scénář (nákup druhého pásového dumperu pro plánování zakázek).

Výsledky jsou rozděleny do šesti záložek:

| Záložka | Obsah |
|---|---|
| 📊 **Shrnutí** | Očekávaný měsíční zisk, pravděpodobnost ztráty a bod zvratu. Preferovaná varianta, rozdíly mezi stroji, bodová scorecard, realistický scénář při zadané poptávce, doporučený další postup a **export investičního reportu do formátu DOCX**. |
| 📈 **Monte Carlo** | Rozdělení simulovaného měsíčního zisku (histogram), kumulativní pravděpodobnost (CDF), percentily P10 / medián / P90 a interval nejistoty. |
| ⚠ **Riziko** | Semafor rizika ztráty, pravděpodobnost poruchy, Value at Risk (5 %) a mapa riziko vs. výnos. |
| 📉 **Citlivost** | Zisk v závislosti na vytížení (bod zvratu), tornádo graf citlivosti na změnu vstupů o ±10 %, heatmapa ziskovosti (cena × počet dní) a minimální denní cena pronájmu. |
| 🗺 **Optimalizace** | Lineární programování a diskrétní optimalizace rozdělení dní mezi dva stroje, porovnání výsledků modelů, optimální nákup strojů v rámci rozpočtu a křivka rozpočet vs. maximální zisk. |
| 📋 **Zakázky** | Rozhodovací skóre zakázek, harmonogram s ohledem na kapacitu strojů (Ganttův diagram), detail zakázek, vytížení strojů a identifikace úzkého hrdla. |

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
| Vizualizace | Matplotlib, Plotly |
| Reporting | python-docx |
| Správa kódu a zabezpečení | Git, GitHub (větve, pull requesty), Streamlit secrets |

## Technické informace

### Struktura projektu

```
├── app.py                           # Streamlit aplikace (UI, výpočty, export reportu)
├── data/
│   ├── machines_demo.csv            # fiktivní parametry 8 typů strojů
│   ├── jobs_demo.csv                # fiktivní zásobník 12 zakázek
│   └── *.xlsx                       # rozšířený fiktivní datový model (aplikace ho zatím nenačítá)
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
