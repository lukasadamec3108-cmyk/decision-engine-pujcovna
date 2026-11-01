# Ekonomika v Decision Engine – mapa výpočtů (E2.2)

Přehled všech ekonomických výpočtů aplikace, jejich zdrojů a vzájemných vztahů. Slouží jako podklad
pro sjednocení ekonomiky (etapa E2.2). Pojmy jsou vysvětlené ve [slovníku](SLOVNIK.md).

> **Všechna data jsou fiktivní ukázka.** Žádná částka v aplikaci není čistý zisk půjčovny.
> Neúplný výsledek (chybí některý náklad) je jen **horní mez**, ne skutečný výsledek.

## Tři úrovně ekonomiky

1. **Model jednoho stroje** – očekávaný průměrný měsíc jednoho kusu pro rozhodnutí o nákupu
   (Investice, Nákup strojů, Strojový park). Obsahuje fixní náklad kusu. **Nesčítá se za flotilu.**
2. **Příspěvek** pronájmu nebo zakázky – tržba − variabilní náklady (u zakázky i přímé náklady
   a penále). **Bez fixních nákladů** a bez režie (`economics.py`).
3. **Výsledek celé flotily** – příspěvky z evidovaných pronájmů a zakázek − fixní náklady flotily
   (jednou) − režie, za stejné období (`economics.total_result`). **V aplikaci se zatím nezobrazuje.**

## Kde se co počítá

| Výpočet | Stránka | Modul | Vstupy | Model / evidence | Vztah k `economics.py` (test) |
|---|---|---|---|---|---|
| Simulace měsíčního zisku dvou strojů, riziko, bod zvratu, minimální cena, citlivost | Investice (Doporučení – Co ovlivňuje zisk), Dashboard, export DOCX | `investice.py` | ceník, scénář trhu, servis (průměr, rozptyl), poruchy, **fixní náklad z panelu** | model | se stejným fixním nákladem jako v datech = tržby − fixní náklad kusu (`test_investment_quick_estimate_equals_economics_when_fixed_cost_from_data`) |
| Ekonomika jednoho kusu, nákupní sestava | Investice – Nákup strojů, Strojový park, Dashboard | `portfolio.py` | ceník, modelová poptávka, poruchy, fixní náklad a servis **z dat** | model | náklady = `unit_fixed_total` × kusy; tržby = `rental_contribution` (`test_purchase_model_*`) |
| Plán zakázek, zisk po penále, přínos zakázky | Vlastní zakázky, Dashboard | `scheduling.py` | tržba, přímé náklady, délka, termín, penále zakázek | evidence + optimalizace | = `job_contribution(..., includes_machines=True)` (`test_job_profit_in_plan_equals_economics_contribution`) |
| Pořadí starého plánu (jen pro srovnání) | Vlastní zakázky | `app.py` | (tržba − přímé náklady) × koeficient rizika | pravidlo | – (jen řazení, žádná zobrazená částka) |
| Příspěvek pronájmu a zakázky, verdikt | Pronájem strojů, Rozhodovací centrum | `planner.py`, `econ_formula.py` → **`economics.py`** | rezervace, požadavky uživatele, data | evidence + vstup | přímo `economics.py` (`test_planner_*`) |
| Kapacita, využití, konflikty | Plánování kapacity, Dashboard | `capacity.py` | flotila, rezervace, servis, plán | evidence | žádné částky (strojodny) |
| Fixní náklady flotily, režie, celkový výsledek | – (nezobrazuje se) | `economics.py` | data | – | `test_total_result_is_incomplete_on_current_data` |
| Využití fyzických strojů za období (strojodny po kusech a typech, konflikty) | – (nezobrazuje se, K2) | `fleet_usage.py` nad `capacity.py` | flotila, rezervace, plán zakázek, servis | evidence | žádné částky; shoda s kalendářem a přehledem využití (`tests/test_fleet_usage.py`) |

## Pravidla proti dvojímu započtení

- **Fixní náklady stroje** (`fixed_cost_month_unsplit` + servis) se ve výsledku flotily odečítají
  **jen jednou** za flotilu. Model jednoho stroje je už obsahuje – „zisky strojů“ se proto nesčítají
  a k součtu se znovu neodečítají fixní náklady flotily
  (`test_purchase_model_profit_already_contains_fixed_costs_of_fleet`).
- **Příspěvek** pronájmu ani zakázky fixní náklady neobsahuje (`test_contributions_never_contain_fixed_costs`).
- **Každý den stroje má jedno využití** (pronájem, zakázka, servis, volno) – modelová poptávka po
  pronájmu se nesmí sčítat se zakázkami ve stejných dnech.
- **Výpadky:** modely (Investice, Nákup strojů) odečítají dny poruch; výsledek flotily z evidence
  smí použít jen skutečné servisní odstávky z kalendáře – ne obojí.
- **Náklady strojů v zakázkách** se odečítají jen tehdy, když nejsou v přímých nákladech zakázky
  (nastavení `direct_cost_includes_machines`).
- **Stejné období:** příspěvky, fixní náklady a režie se sčítají jen za stejné období.

## Známé rozdíly (zdokumentované, neopravené)

Testy `tests/test_econ_consistency.py` s označením „ZNÁMÝ ROZDÍL“ popisují současné chování.
Změna kteréhokoli z nich mění zobrazené částky a vyžaduje **schválení majitele**.

| # | Rozdíl | Dopad | Test |
|---|---|---|---|
| 1 | Investice berou fixní náklad z panelu (výchozí 100 000 Kč pro oba stroje), Nákup strojů, Strojový park a `economics.py` z dat (např. mini bagr 45 000 Kč). | Na Dashboardu stojí vedle sebe „očekávaný zisk“ (43 169 Kč, fixní 100 000 Kč) a „doporučená sestava“ (199 215 Kč, fixní z dat). | `test_known_difference_investment_fixed_cost_comes_from_panel` |
| 2 | Investice a Nákup strojů nepočítají s variabilními náklady pronájmu; `economics.py` je bez nich označí jako chybějící. | Modelové tržby/zisky jsou horní mez, pokud variabilní náklady existují. | `test_purchase_model_revenue_equals_economics_rental_revenue` |
| 3 | Plán zakázek nezná náklady strojů na zakázce; v datech chybí `direct_cost_includes_machines`. | „Zisk zakázek po penále“ (892 000 Kč) je dílčí údaj – `economics.py` ho bez nastavení označí jako neúplný. | `test_known_difference_plan_does_not_know_machine_costs` |
| 4 | Bod zvratu = fixní náklad ÷ cena (bez servisu), minimální cena = (fixní náklad + servis) ÷ dny. Aplikace to uvádí v textu. | Při bodu zvratu je zisk se servisem záporný. | `test_known_difference_break_even_ignores_service_but_minimum_price_includes_it` |
| 5 | Období: plán zakázek trvá ~20 dní, rezervace dny 1–25, fixní náklady a režie jsou měsíční. | `total_result` nad současnými daty míchá období – proto se nezobrazuje. | – (rozhodnutí majitele) |
| 6 | Servis: v simulaci Investic náhodně kolísá (průměr + rozptyl), jinde se bere průměr; `economics.py` ho řadí k provozním fixním nákladům. | Průměrně stejná částka, jiný rozptyl. | `test_legacy_fixed_and_service_costs_equal_economics_unit_costs` |

Poznámka: v `investice.py` se prostá návratnost (pořizovací cena ÷ měsíční zisk, v měsících) jmenuje
`roi_*` – jde jen o název proměnné, v aplikaci je správně „prostá návratnost“.

## Otevřené otázky pro majitele (podklad pro E2.2 K2–K6)

| Otázka | Proč je důležitá | Navržená výchozí varianta (do potvrzení) |
|---|---|---|
| Jsou měsíční fixní náklady strojů správné a jak se dělí na vlastnické a provozní? | Největší nákladová položka flotily. | Celek ponechat, rozdělení doplnit později. |
| Je servis pevný měsíční náklad, nebo závisí na využití? | Fixní × variabilní náklad. | Pevný měsíční průměr (jako dnes). |
| Variabilní náklad za den pronájmu a za den práce na zakázce? | Bez nich je příspěvek jen horní mez. | Nezadáno – **nedosazovat 0**. |
| Obsahují přímé náklady zakázek vlastní stroje? | Zabrání dvojímu započtení nebo vynechání nákladů. | Neznámé; doporučeno evidovat přímé náklady **bez** strojů. |
| Společná režie měsíčně? | Bez ní nejde mluvit o zisku půjčovny. | Nezadáno → výsledek neúplný; nerozpočítávat. |
| Za jaké období ukazovat výsledek flotily? | Fixní náklady jsou měsíční, plán 20–25 dní. | Za zobrazené období, fixní náklady poměrně (dny ÷ 30) – **předpoklad ke schválení**. |
| Počítat předběžné rezervace a nepotvrzené zakázky? | Mohou zkreslit výsledek. | Jen potvrzené, předběžné zvlášť. |
| Fixní náklad v Investicích z panelu, nebo z dat? | Dva základy vedle sebe na Dashboardu. | Beze změny, později volitelný přepínač. |

## Využití strojů za období (K2)

Funkce `fleet_usage.machine_days(kalendář, začátek, konec)` (případně `fleet_usage.from_model(...)`
přímo z dat) vyhodnotí každý fyzický kus v každém dni zvoleného období. Zatím se **nikde nezobrazuje**
a nic v aplikaci nemění – je připravená pro výsledek flotily za období (K3).

**Strojodny.** 1 strojoden = 1 kus na 1 den; období zahrnuje oba krajní dny, dny mohou být čísla
plánu i kalendářní data (přes hranici měsíce a roku). Dostupná kapacita = kusy ve flotile, které jsou
ten den k dispozici (objednaný kus až od dodání). Každý dostupný kus se v každém dni započítá
**právě jednou** do jedné složky:

> pronájem + vlastní zakázka + servis + konflikt + nelze určit + volno = dostupná kapacita

**Stavy záznamů** (stejně jako Plánování kapacity): potvrzené blokují kapacitu; předběžné rezervace
a nabídky ji neblokují a vykazují se zvlášť (volbou `include_tentative=True` je lze započítat);
zrušené se ignorují; neznámý stav nebo neplatný termín je chyba dat – volné strojodny typu jsou pak
„nelze určit“, nikdy volno.

**Konflikty** se nikdy tiše neodstraní ani nezapočítají dvakrát. Přednost pronájmu, zakázky ani
servisu se neurčuje. Konflikt je:
- dva různé nároky na stejný konkrétní kus ve stejný den (překrývající se servisní záznamy téhož kusu
  konflikt nejsou – kus je v servisu jednou, vypíše se upozornění na možnou duplicitu),
- víc nároků na typ stroje, než je dostupných kusů – v konfliktu jsou pak všechny dostupné kusy typu
  v tom dni (bez priority nelze říct, co ustoupí),
- konkrétní kus naplánovaný na den, kdy není ve flotile.

Konfliktní strojoden se započte jednou (složka „konflikt“); nároky, které se nevejdou, se vykazují
zvlášť jako **nárok nad kapacitu** a do kapacity se nepřičítají.

**Konkrétní kus × typ stroje.** Rezervace i zakázky v datech uvádějí jen typ stroje. Počty strojodnů
po typech jsou proto jisté, ale u kusů je stav **„nejednoznačné přiřazení“** – kromě případu, kdy jsou
všechny zbývající kusy typu obsazené stejným druhem využití (druh je jistý, konkrétní rezervace ne).
Výsledek to uvádí ve sloupci `unit_level` a v poznámkách.

**Výsledek:** stav (úplný / neúplný / nelze spočítat), tabulka kusů po dnech, tabulka typů po dnech,
součty po typech a za flotilu (s kontrolou součtu), seznam konfliktů a poznámky o nejistotách.
Neplatné období (chybí, konec před začátkem, míchání čísel dní a dat) = „nelze spočítat“. Pokud nelze
spolehlivě určit počet kusů typu, je jeho součet výslovně neurčitelný a do součtu flotily se nezapočte.

**Ukázková data** (plán zakázek ze stránky Vlastní zakázky, dny 1–25): 225 strojodnů, z toho pronájem
25, vlastní zakázky 110, servis 8, konflikt 5, volno 77; předběžné rezervace 10 strojodnů zvlášť.
Konflikty: kolový bagr den 11 (servis × zakázka J005), silniční válec dny 15–18 (pronájem FIKT-P05 ×
zakázka J006) – stejné jako na stránce Plánování kapacity.

**Chybějící vstupy:**
- přiřazení rezervací a zakázek ke konkrétním kusům (sloupec kusu v rezervacích / požadavcích zakázek),
- stav „zrušeno“ v datech (model ho zná, data E1 ho zatím nemají),
- pracovní dny a horizont plánování (`settings`: nezadáno – počítají se kalendářní dny),
- skutečné termíny zakázek: berou se z optimalizovaného plánu, ne z evidence,
- datum prvního dne plánu (bez něj jsou dny jen čísla 1, 2, …),
- pravidla priority mezi pronájmem, zakázkou a servisem (záměrně nedefinována).

## Evidence závazků za období (K3.1)

`fleet_economics.period_ledger(model, začátek, konec, job_starts=…, plan_start=…)` propojí kapacitu
z `fleet_usage.py` s příspěvky z `economics.py`. **Nezobrazuje se** a nic nerozhoduje.

- Každý pronájem a každá zakázka z dat má právě jeden řádek: termín, zdroj termínu (rezervace /
  **návrh plánu** / bez termínu), strojodny celkem, v období, mimo období a v konfliktu, přiřazení kusu,
  příspěvek **celého** závazku a jeho úplnost.
- Stav (právě jeden): jisté využití · konflikt · předběžné · přesah mimo období · nelze určit · mimo
  období · zrušeno. Nejednoznačné přiřazení kusu je samostatný údaj – ekonomika je po typech strojů.
- Mezi **jisté** příspěvky se sčítají jen závazky ve stavu „jisté využití“. Konfliktní (obě strany,
  bez priority), předběžné, přesahující a neurčité závazky jsou ve vlastních skupinách a nesčítají se.
- Chybějící náklad zůstává chybějícím (stav „neúplný“, částka = horní mez), nikdy 0.
- Období je povinný parametr; přesah období se **neděli poměrně** – pravidlo časového rozlišení je otevřené.
- Plán zakázek je návrh; bez něj zůstávají zakázky „nelze určit“. Modelová poptávka se nepoužívá.

Ukázková data (dny 1–25, návrh plánu): jisté pronájmy 194 500 Kč a jisté zakázky 635 000 Kč (horní meze),
v konfliktu FIKT-P05, J005 a J006 (304 500 Kč – nesčítat), předběžné FIKT-P06 a FIKT-P07 (77 500 Kč).

Otevřená pravidla: časové rozlišení přesahujících závazků, závaznost plánu zakázek, řešení konfliktů,
penále potvrzené zakázky bez termínu.

## Varianty řešení konfliktů (K3.2)

`fleet_economics.conflict_options(model, začátek, konec, job_starts)` (čísla dní plánu) připraví pro
každý konflikt potvrzených závazků podklad pro rozhodnutí majitele. **Nezobrazuje se a nic nevybírá.**

- Případ konfliktu: dotčené pronájmy a zakázky (příspěvek, úplnost, vzorec), servis, stroje a dny.
  Konflikty se společným závazkem tvoří jeden případ.
- **Posun zakázky** (proveditelná): nejbližší začátek, kdy se vejdou všechny stroje zakázky vedle
  ostatních závazků, servisu a zakázek (kontrola plánovače nad `capacity.evaluate_request`). Penále
  z `economics.job_contribution`. Vypíše se, co bližší termíny blokuje (kaskáda), termíny bez penále
  a varování na předběžné rezervace. Posouvá se vždy jen jedna zakázka.
- **Uvolnění pronájmu** – zrušení, nebo zkrácení o konfliktní dny na kraji termínu – je jen
  **podmíněné**: smluvní náhrada je ve vzorci jako neznámá a chybí souhlas zákazníka.
- **Posun servisu** a **externí zapůjčení** jsou nevyhodnotitelné – vypíše se, co chybí.
- **Ruční rozhodnutí** je vždy k dispozici (konflikt trvá, výsledek nelze vyčíslit).
- Každá varianta se ověří přepočtem kalendáře: zbývající konflikt = „proveditelná jen částečně“,
  nikdy „vyřešeno“; částečné varianty se ekonomicky nesrovnávají.
- Ekonomika varianty = vzorec (`econ_formula`) součtu **všech** dotčených závazků, každého jednou,
  ve stejném scénáři nastavení přímých nákladů. Stav: úplný / neúplný / nevyčíslitelný; meze místo
  vymyšlené částky. Porovnání dvojic je jen informace.

Ukázková data (dny 1–25, návrh plánu):

| Konflikt | Varianta | Stav | Příspěvek dotčených závazků |
|---|---|---|---|
| FIKT-P05 × J006 (válec, dny 15–18) | posun J006 na dny 20–27 (zpoždění 11 místo 2 dní, −81 000 Kč) | proveditelná; varování FIKT-P07 | nejvýše 108 500 Kč |
| | zrušení FIKT-P05 | podmíněná | 142 000 Kč − náhrada |
| | zkrácení FIKT-P05 na den 19 | podmíněná | 151 500 Kč − náhrada |
| servis × J005 (kolový bagr, den 11) | posun J005 na dny 21–26 (zpoždění 15 dní, −135 000 Kč) | proveditelná; varování FIKT-P06 | nejvýše −20 000 Kč |
| | posun servisu | nevyhodnotitelná | – |

Všechny částky jsou horní meze (chybí variabilní náklady a nastavení přímých nákladů). Porovnání
s podmíněnými variantami nelze rozhodnout, dokud není známa smluvní náhrada.

## Výsledek flotily za období (K3.3)

`fleet_economics.period_result(model, začátek, konec, job_starts, fixed_cost_method=…, months=…, plan_start=…)`
skládá evidenci (K3.1) a varianty konfliktů (K3.2) s fixními náklady a režií z `economics.py`.
**Nezobrazuje se a nic nevybírá.**

**Výsledek z jistých závazků** = jisté příspěvky pronájmů + jisté příspěvky zakázek − fixní náklady
flotily za období − režie za období.
- Jisté příspěvky = jen závazky ve stavu „jisté využití“ (každý jednou). Konfliktní (obě strany),
  předběžné, přesahující a neurčité závazky jsou ve vlastních přehledech a **nezapočítávají se**;
  dokud konflikty nejsou rozhodnuté, je výsledek vždy **neúplný**.
- Fixní náklady = `economics.fleet_fixed_costs` (vlastněné a leasingové kusy), odečtou se **jednou**;
  průměrný servis je jejich součástí a znovu se neodečítá. Režie = `economics.overhead`.
- **Přepočet měsíčních nákladů na období jen výslovně zvolenou metodou:**
  - nezadáno (výchozí) → fixní náklady i režie jsou chybějící údaj, výsledek je jen horní mez,
  - `"months"` → počet měsíců zadá volající (např. 1 = celý měsíc),
  - `"calendar_days"` → Σ (dny období v měsíci ÷ počet dní měsíce), vyžaduje datum dne 1 plánu.
  Žádný 30denní měsíc se nepředpokládá.
- Chybějící údaj není nula (režie „0“ = úplný výsledek, prázdná režie = neúplný). U neúplného
  výsledku se výslovně uvádí, že **nejde o přesný čistý zisk**, a seznam chybějících údajů.

**Scénáře konfliktů:** kombinace proveditelných a podmíněných variant z K3.2, které konflikt vyřeší.
Kombinace se ověří kapacitně **společně** (varianty různých konfliktů se nesmí srazit) a do rozpětí
se započte jen tehdy, když všechny dotčené závazky zůstanou celé v období. Rozpětí = dolní a horní mez
vzorců (`econ_formula`) přes tyto kombinace; neznámá mez = nelze určit. Modelová poptávka se nepoužívá.

Ukázková data (dny 1–25, návrh plánu, `months` = 25/30 pouze jako příklad zadané metody):

| Položka | Částka | Stav |
|---|---|---|
| Jisté příspěvky pronájmů | 194 500 Kč | neúplné (chybí variabilní náklady) – horní mez |
| Jisté příspěvky zakázek | 635 000 Kč | neúplné (chybí `direct_cost_includes_machines`) – horní mez |
| Fixní náklady flotily (25/30 měsíce) | 440 000 Kč | úplné při zadané metodě |
| Režie | – | chybí |
| **Výsledek z jistých závazků** | **nejvýše 389 500 Kč** | neúplný – nejde o čistý zisk |
| V konfliktu (nezapočteno) | FIKT-P05, J005, J006 | rozhodne majitel |
| Předběžné (nezapočteno) | FIKT-P06, FIKT-P07 – 77 500 Kč | jen potenciál |
| Rozpětí scénářů | nelze určit | posuny J005 a J006 přesahují den 25 |

Pro období 1–31 jsou posuny celé v období a horní mez rozpětí je 521 000 Kč (dolní mez nelze určit –
chybí náklady a smluvní náhrady).

## Zobrazovací vrstva „Ekonomika flotily“ (K3.4a)

`fleet_view.build(period_result, názvy_strojů, plan_start)` převádí výsledek K3.3 na řádky, texty,
stavy a upozornění pro budoucí stránku „Ekonomika flotily“ (K3.4b+). **Nic nepočítá jinak a zatím
se nikde nezobrazuje.** Výstup jsou jen obyčejné hodnoty (lze uložit do mezipaměti).

Každá částka má právě jeden z pěti stavů:

| Stav | Kdy | Text |
|---|---|---|
| **úplný** | všechny údaje známé (i skutečná nula) | „389 500 Kč“, „0 Kč“ |
| **horní mez** | chybí jen náklad (variabilní, přímé náklady, penále, režie, fixní náklady) | „nejvýše 389 500 Kč“ |
| **podmíněný** | platí jen za neověřitelné podmínky (souhlas zákazníka, storno, smluvní náhrada) | „nejvýše 142 000 Kč“ + podmínky |
| **chybí údaje** | chybí tržba / cena, neplatný údaj, nezadaný náklad, nerozhodnuté konflikty | „nelze určit“, „nezadáno“ |
| **nezapočítáno** | závazek v konfliktu, předběžný, přesahující období, s nejistou obsazeností | částka jen pro informaci |

- Chybějící náklad se nikdy nezobrazí jako „0 Kč“ (režie „0“ = úplný, prázdná = „nezadáno“).
- Chybějící tržba nebo cena znamená „nelze určit“ – horní mez by neplatila.
- Výsledek je **úplný** jen při úplném výpočtu K3.3. Při nerozhodnutých konfliktech je i s úplnými
  náklady „chybí údaje“ (částka platí jen pro jisté závazky).
- Slovo „zisk“ se ve výstupu objevuje jen v upozornění „Neúplný výsledek – nejde o čistý zisk.“,
  které je u každého neúplného výsledku (hlídá test).
- Varianty konfliktů mají pevné pořadí podle druhu (ne podle částky) – nic se nedoporučuje.

## Stránka „Ekonomika flotily“ (K3.4b–K3.4c)

Stránka ve skupině Přehled pod Dashboardem zobrazuje výsledek `fleet_economics.period_result`
převedený `fleet_view.build`. Výpočet běží v mezipaměti (`fleet_economics_view` v `app.py`) – nic jiného
se nepočítá a žádná jiná stránka se nemění.

- **Vstupy:** termíny zakázek z plánu stránky Vlastní zakázky (stejná funkce `plan_jobs`), období
  v dnech plánu (výchozí jako Plánování kapacity), volitelné datum dne 1 a **metoda přepočtu
  fixních nákladů a režie** – výchozí „nezadávat“ (fixní náklady a režie se neodečtou, výsledek je
  horní mez), „zadat počet měsíců“, nebo „podle kalendáře“ (vyžaduje datum dne 1).
- **Čtyři ukazatele** (každý s významem, obdobím a stavem): výsledek z pronájmu strojů, výsledek
  z vlastních zakázek, fixní náklady flotily, celkový výsledek flotily (režie jen pokud je zadaná).
- **Rozpad výsledku:** tržby, variabilní náklady pronájmu, přímé náklady, penále, variabilní náklady
  strojů na zakázkách, fixní náklady, režie. Součty jen z příspěvků jistých závazků z `economics.py`;
  položka, kterou `economics.py` označí jako chybějící, je „nezadáno“ – nikdy 0 Kč.
- **Co chybí k přesnému výsledku:** konkrétní chybějící údaje a rozhodnutí z modelu.
- **Kapacitní konflikty a varianty** (K3.2) včetně rozpětí scénářů (K3.3) – jen přehled, nic se
  nevybírá ani nemění, rozhodne majitel.
- Neúplný výsledek má vždy upozornění „Neúplný výsledek – nejde o čistý zisk.“
- **K3.4d – u každého konfliktu:** závazky, které blokují bližší termín posunu zakázky (kaskáda);
  rozbalovací blok **„Porovnání variant – jen informace“** (porovnání dvojic z `econ_formula`,
  výslovně „nelze porovnat“, když chybí údaje; není to doporučení ani rozhodnutí); odkaz
  **„Zobrazit v kalendáři“** (Plánování kapacity s filtrem na stroj konfliktu – mění se jen filtr
  zobrazení) a odkaz **„Vlastní zakázky“**. Na ukázkových datech nelze varianty porovnat: u J005 lze
  vyčíslit jen posun, u FIKT-P05 × J006 chybí storno / smluvní náhrada.
- **K3.4e1 – předběžné a nejisté závazky:** samostatná sekce s předběžnými, přesahujícími a nejistými
  závazky (období, stroje, částka, důvod nezapočítání) označenými „nezapočítáno – nejde o potvrzený
  výsledek“; částky se nesčítají s výsledkem ani mezi sebou. Rozbalovací blok **„Evidence závazků“**
  obsahuje každý pronájem a zakázku z dat právě jednou, ve skupinách podle stavu (jisté využití –
  započteno; konflikt; předběžné; přesah; nejistá obsazenost; mimo období; zrušené), s termínem, jeho
  zdrojem a přiřazením kusu. Na mobilu řádky pod sebou. Nic se nepočítá znovu – data jsou z K3.1.
- **K3.4e2 – využití strojů a metodika:** kompaktní souhrn strojodnů za období (pronájem, vlastní
  zakázky, servis, konflikt, případně „nelze určit“, volná kapacita) převzatý z `fleet_usage`
  (`Usage.fleet`) – stejná čísla jako Plánování kapacity → Přehled využití. Kategorie se nepřekrývají
  a dohromady dávají kapacitu; nárok nad kapacitu a předběžné rezervace jsou uvedené zvlášť. Nové
  procento využití se nepočítá. Tlačítko **„Podrobné využití v Plánování kapacity“** otevře Přehled
  využití se stejným obdobím (a datem dne 1) pro všechny stroje. Rozbalovací blok **„Jak se počítá“**
  popisuje stavy částek, příspěvek × výsledek, fixní náklady jednou, metody přepočtu, nezapočtené
  závazky, strojodny a proč nejde o čistý zisk (texty ve `fleet_view.METHODOLOGY`).
- **K3.4e3a – vzhled a mobil (jen prezentace):** navigační odkazy („Zobrazit v kalendáři – stroj“,
  „Podrobné využití v Plánování kapacity“, „Vlastní zakázky“) mají jednotný styl odkazu; na mobilu je
  v řádcích vždy popisek nahoře a částka pod ním; identifikátory (např. FIKT-P05) se nelámou na
  spojovníku; informační upozornění o metodě přepočtu a o konfliktech jsou v jednom bloku (texty beze
  změny). Výpočty, popisky stavů a počet chybějících údajů se nemění.
- **K3.4e3b – srozumitelné popisky (jen zobrazení, `fleet_view`):** technické názvy chybějících údajů
  v přirozené češtině (např. „údaj, zda přímé náklady zakázek už obsahují náklady vlastních strojů“);
  nezvolená metoda přepočtu hlášená fixními náklady i režií je jedna volba, proto jedna položka
  (počet chybějících údajů na ukázkových datech bez metody 8 → 7, s metodou beze změny 7); servis
  jako „servis – Kolový bagr, kus 01“; v evidenci „termín z rezervace / z návrhu plánu zakázek“ a
  srozumitelné přiřazení konkrétního stroje („určen jednoznačně (jediný kus tohoto typu)“ jen u
  jednokusových typů, jinak „nelze určit…“); porovnání variant s názvy karet; metodika jako krátké
  body. Štítky z `economics.py` / `fleet_economics.py`, výpočty i stavy se nemění.

Na ukázkových datech (dny 1–25) jsou všechny ekonomické částky **horní meze** nebo „chybí údaje“
(chybí variabilní náklady pronájmu, nastavení `direct_cost_includes_machines`, režie a rozhodnutí
konfliktů); přesný výsledek ani čistý zisk z nich určit nelze.

## Regrese

- `tests/test_econ_consistency.py` – shoda modulů s `economics.py`, ochrana před dvojím započtením,
  známé rozdíly a reference současných výsledků (`tests/fixtures/econ_regression_reference.json`):
  plán zakázek, přínos zakázek, ekonomika jednoho kusu, nákupní sestava, fixní náklady flotily
  a export DOCX výchozího scénáře.
- `tests/test_investice.py`, `tests/test_report_docx.py` – bitová shoda výpočtů Investic a DOCX.
  Determinismus DOCX se ověřuje porovnáním pořadí, názvů, komprese a obsahu všech částí souboru,
  nikoli celých bajtů: knihovna zipfile dává každé položce aktuální čas (přesnost 2 s), takže se
  dvě jinak shodná sestavení mohla lišit jen časovou značkou (dříve občasné selhání testu).
- `tools/regression/` – snímek celé aplikace pro porovnání větví.
