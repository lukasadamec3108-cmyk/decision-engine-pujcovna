# Slovník pojmů Decision Engine

Interní slovník pro texty aplikace. Při každé úpravě rozhraní používejte tyto pojmy ve stejném významu.
Definice popisují **současné výpočty** – pokud se výpočet změní, upravte i slovník.

Všechna data v aplikaci jsou ukázková (fiktivní).

## Ekonomické výsledky

Aplikace počítá několik různých „zisků“. **Nejsou zaměnitelné** a nesmějí se vydávat jeden za druhý.
Přehled všech výpočtů, jejich zdrojů a známých rozdílů je v [`EKONOMIKA.md`](EKONOMIKA.md).

### Tři úrovně ekonomiky

| Úroveň | Co vyjadřuje | Kde | Smí se sčítat za flotilu? |
|---|---|---|---|
| **Model jednoho stroje** | Očekávaný průměrný měsíc **jednoho kusu** pro rozhodnutí o nákupu. Fixní náklad kusu už obsahuje. | Investice, Nákup strojů, Strojový park | **Ne** – součet by dvakrát započítal fixní náklady a „prodal“ stejné dny stroje pronájmu i zakázce. |
| **Příspěvek** pronájmu / zakázky | Tržba − variabilní náklady (u zakázky i přímé náklady a penále). **Bez fixních nákladů** strojů a bez režie. | Pronájem strojů, Rozhodovací centrum (`economics.py`); zisk zakázek ve Vlastních zakázkách je stejný vzorec | Ano, ale výsledek není zisk – fixní náklady a režie ještě chybí. |
| **Výsledek celé flotily** | Příspěvky z evidovaných pronájmů a zakázek − fixní náklady flotily (jednou) − režie, za stejné období. | Stránka **Ekonomika flotily** (přes `fleet_economics.period_result` – jen jisté závazky, konflikty a předběžné zvlášť) | Je to sám součet – nic dalšího se k němu nepřičítá. |

> **Neúplný výsledek není čistý zisk.** Pokud chybí některý náklad (např. variabilní náklad pronájmu,
> režie nebo informace, zda přímé náklady zakázek obsahují stroje), známá částka je jen **horní mez**.
> Aplikace ji označí jako neúplnou („nejvýše …“, „dílčí údaj“) – nikdy ji nevydávejte za zisk půjčovny.

### Pojmy v aplikaci

| Pojem v aplikaci | Kde | Jak se počítá | Co nezahrnuje |
|---|---|---|---|
| **Očekávaný měsíční zisk stroje** (simulace) | Investice – Doporučení, Možné výsledky, Riziko; Dashboard; DOCX | Průměr simulovaných měsíců: tržby − fixní náklad (z panelu „Nastavení investičního scénáře“) − servis. Simulace mění počet dní pronájmu podle scénáře trhu, poruchy, cenu (±) a servisní náklady. | Společné náklady půjčovny (prostory, mzdy, administrativa), daně, variabilní náklady. |
| **Zisk při N dnech** (rychlý odhad) | Investice – Doporučení (rozbalovací blok), Co ovlivňuje zisk | Ceníková cena × N dní − fixní náklad − průměrný servis. Bez simulace. | Poruchy, kolísání ceny a servisu, rozpětí dní podle scénáře – proto se liší od simulace. |
| **Tržby − fixní náklad** | Investice – Co ovlivňuje zisk (graf „Kdy stroj pokryje fixní náklad“, mapa ziskovosti) | Denní cena × dny − fixní náklad. | Servis a poruchy. Není to zisk. |
| **Očekávaný měsíční provozní zisk** (jednoho kusu / sestavy) | Investice – Nákup strojů, Strojový park, Dashboard (doporučená sestava) | Tržby z průměrné poptávky po odečtení poruch − fixní náklady jednotlivých strojů **z dat** − průměrný servis. | Fixní náklad z panelu (nepoužívá se), variabilní náklady, společné náklady, úroky, daně, sezónnost. Na Strojovém parku jen **jeden kus** – nesčítá se za flotilu. |
| **Zisk zakázky**, **zisk zakázek po penále** | Vlastní zakázky, Dashboard („Ekonomický výsledek“ – dílčí údaj) | Tržba zakázky − přímé náklady zakázky (z dat); po penále navíc − penále za zpoždění. Stejný vzorec jako příspěvek zakázky v `economics.py`. | Fixní náklady strojů, režie, pronájmy. Data neuvádějí, zda přímé náklady obsahují náklady vlastních strojů – proto **není to čistý zisk**. |
| **Příspěvek** | Pronájem strojů, Rozhodovací centrum | Pronájem: tržba − variabilní náklad pronájmu × strojodny. Zakázka: tržba − přímé náklady − variabilní náklady strojů (jen pokud nejsou v přímých nákladech) − penále. | Fixní náklady flotily a režie. **Není to čistý zisk.** |
| **Výsledek flotily** | Ekonomika flotily | Viz tři úrovně výše. Neúplný výsledek je jen horní mez. | – |
| **Čistý zisk firmy** | aplikace ho nepočítá | – | Chybí režie, daně a úplné náklady. Tento pojem v aplikaci nepoužívejte. |

## Náklady

- **Fixní náklad** – platí se každý měsíc bez ohledu na počet dní pronájmu (např. splátka, pojištění).
  Investice (záložky Doporučení až Co ovlivňuje zisk) berou jednu ukázkovou hodnotu z panelu
  „Nastavení investičního scénáře“ pro oba stroje. Nákup strojů, Strojový park a `economics.py`
  berou individuální hodnoty strojů z ukázkových dat. Ve výsledku flotily se fixní náklady
  odečítají **jen jednou** za celou flotilu, nikdy znovu u pronájmu či zakázky.
- **Servisní náklad** – průměrný měsíční servis stroje z ukázkových dat (v simulaci kolísá).
- **Variabilní náklad** – náklad za den, kdy stroj pracuje (pronájem nebo zakázka). V ukázkových
  datech chybí, proto je ekonomika v Rozhodovacím centru často „Záleží na nákladech“ a příspěvek
  pronájmu jen „nejvýše …“.
- **Režie půjčovny** – společné náklady firmy (kancelář, mzdy, pojištění firmy). V datech zatím
  nezadaná; nerozpočítává se na stroje ani zakázky, odečítá se jen ve výsledku flotily.
- **Přímé náklady zakázky** – náklady konkrétní zakázky z dat.
- **Ukázkové náklady** – všechny náklady v aplikaci jsou zatím ukázkové, nejde o ověřené náklady firmy.

## Riziko a návratnost

- **Šance ztrátového měsíce** – podíl simulovaných měsíců, kdy stroj skončí ve ztrátě.
- **Velmi nepříznivý měsíc (VaR 5 %)** – hranice 5 % nejhorších simulovaných měsíců.
- **Bod zvratu** – počet dní pronájmu měsíčně, při kterém tržby pokryjí fixní náklad (bez servisu a poruch).
- **Minimální denní cena** – cena, při které se při zadaném počtu dní pokryje fixní náklad i průměrný servis.
- **Prostá návratnost** – pořizovací cena ÷ měsíční zisk; bez úroků a společných nákladů.

## Kapacita a vytížení

- **Strojoden** – 1 kus stroje na 1 den. Jednotka kapacity a využití (dříve také „kusoden“).
- **Kapacita dne** – kusy ve flotile k dispozici ten den minus kusy v servisu.
- **Vytížení zakázkami** (Vlastní zakázky) – dny strojů na vlastních zakázkách ÷ (kusy × délka plánu).
  Pronájmy a servis nezahrnuje.
- **Obsazenost** (Plánování kapacity, kalendář) – pronájmy, zakázky i servis.
- **Využití** (Plánování kapacity, přehled využití) – pronájmy a vlastní zakázky; servis se nepočítá.
- **Konflikt** – potvrzené nároky (pronájmy, zakázky z plánu, servis) převyšují počet kusů.
- **Kolize mezi zakázkami plánu** (Vlastní zakázky) – překrytí vlastních zakázek navzájem. Plán zakázek
  pronájmy a servis nezná, proto může mít 0 kolizí, i když kalendář ukazuje konflikty.
- **Nelze určit** – kapacitu nelze spočítat (chybí nebo je neplatný údaj); nikdy to neznamená volno.
- **Nejednoznačné přiřazení** – kolik kusů typu je obsazeno, je jisté, ale který konkrétní kus, z dat
  určit nelze (rezervace a zakázky uvádějí jen typ stroje).
- **Nárok nad kapacitu** – nároky, které se do dostupných kusů nevejdou; vykazují se zvlášť, ne jako strojodny kapacity.

## Rezervace a závazky

- **Potvrzená rezervace / závazek** – blokuje kapacitu.
- **Předběžná rezervace** – kapacitu zatím neblokuje.
  - Kalendář (Plánování kapacity) ji vyznačí zvlášť jako riziko; při ověření požadavku ji lze volbou započítat.
  - Rozhodovací plánovač (Pronájem strojů, Rozhodovací centrum) ji bere jako požadavek, který soutěží o volnou kapacitu stejně
    jako nové požadavky.
  - Nikdy se automaticky neruší.

## Kvalita údajů a typy výsledků

- **Skutečná / zadaná data** – údaj, který zadal uživatel nebo je v datech vyplněný.
- **Ukázková data** – fiktivní hodnoty pro předvedení aplikace.
- **Chybějící údaj (nezadáno)** – prázdná hodnota. **Nikdy se nepočítá jako 0 Kč.**
- **Skutečně zadaná nula** – výslovně zadaná 0; počítá se jako 0 a aplikace ji tak označí.
- **Neuplatňuje se** – nezadaný údaj, který na daný výsledek nemá vliv.
- **Hypotetický výpočet „Co kdyby“** – dosazení zkusmé hodnoty nezadaného nákladu. Nemění
  doporučení, data ani rezervace a nic se neukládá.
- **Podmíněné doporučení** (štítek „⇄ Záleží na nákladech“) – nejlepší varianta závisí na výši
  nezadaného nákladu; aplikace ukáže hranice zlomu.
- **Hranice zlomu** – výše nákladu, při které vycházejí dvě varianty stejně.
- **Nelze rozhodnout / nelze určit** – chybí údaje, výsledek nelze spolehlivě určit.

## Štítky v Rozhodovacím plánovači

| Štítek | Význam |
|---|---|
| Provoz: ✓ Vejde se / ✕ Nevejde se / ? Provoz nelze určit | Kapacita po potvrzených závazcích a servisu |
| ★ Doporučeno | Varianta je prokazatelně nejlepší i v nejhorším případě |
| Méně výhodné | Lepší je jiná varianta |
| Nevyplatí se | Každá varianta, která se vejde, je ztrátová |
| = Rovnocenné | Varianty vycházejí stejně |
| ⇄ Záleží na nákladech | Podmíněné doporučení – rozhoduje nezadaný náklad |
| ? Nelze rozhodnout | Chybí údaje |
| Bez porovnání | Požadavek nemá soupeře – doporučení se nevydává |
