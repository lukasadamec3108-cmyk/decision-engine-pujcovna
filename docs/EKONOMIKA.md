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
