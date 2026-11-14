# Insyn

Lokalt rapporteringssystem för bokföringsdata från Visma Control 14: resultat- och balansräkning, noter, egna dashboards, budget och prognos med import från Excel och redigering i systemet, behörighet per avdelning och konto.

Allt körs på den egna datorn. Servern lyssnar bara på `127.0.0.1` och använder bara Pythons standardbibliotek. Webbsidan gör inga anrop ut på nätet.

> **Repot innehåller bara påhittad testdata.** Riktiga exporter (`databas_visma.xlsx`) och `installningar.json` (användare, lösenordshashar, budgetar) ska aldrig in här. `.gitignore` blockerar alla Excelfiler utom de namngivna testfilerna.

## Kom igång (Windows)

1. Ladda ner repot (Code → Download ZIP) och packa upp.
2. Öppna mappen `app` och dubbelklicka på `start.bat`.
3. Webbläsaren öppnas på http://127.0.0.1:8765. Logga in som `admin` / `admin123`.

Andra testkonton är `ekonomi` / `ekonomi123`, `skador` / `skador123` (avdelning SKS), `salj` / `salj123` (avdelning FSG) och `drift` / `drift123` (bara konton 6xxx).

Det går också att öppna `app/insyn.html` direkt i webbläsaren utan server. Då sparas allt i webbläsaren i stället för i `installningar.json`.

## Demo på GitHub Pages

Workflowen `.github/workflows/pages.yml` publicerar `app/insyn.html` som demo varje gång filen ändras. Aktivera det en gång under **Settings → Pages → Source: GitHub Actions**. Adressen blir `https://tommykalmer.github.io/insyn/`.

Demon är öppen för alla som har länken, även om repot är privat. Den visar bara den inbyggda påhittade testdatan, och inläsning av bokföringsfiler är avstängd där. Budget- och prognosimport fungerar och sparas bara i den egna webbläsaren.

## Testdata

`app/databas.xlsx` innehåller påhittade siffror för ett skadeförsäkringsbolag med två bolag (SAK, ALLIV). Flikarna har samma format som Visma-exporten (`vourowx`, `account`, `object`, `fiscalyear`). Perioden är 2025 helår och januari–oktober 2026.

| Kontoserie | Innehåll |
|---|---|
| 3000–3499 | Premieinkomst per försäkringsklass (3010 motor … 3060 ansvar) |
| 3500–3599 | Premier för avgiven återförsäkring |
| 3600–3999 | Förändring i avsättning för ej intjänade premier |
| 5000–5399 | Utbetalda försäkringsersättningar per klass |
| 5400–5499 | Återförsäkrares andel |
| 5500–5699 | Förändring i avsättning för oreglerade skador |
| 5700–5999 | Skaderegleringskostnader |
| 6000–6099 | Personalkostnader |
| 6100–6199 | Lokalkostnader |
| 6200–6799, 6900–6999 | Övriga externa kostnader |
| 6800–6899 | Avskrivningar |
| 7000–7499 | Placeringsintäkter (utdelningar, räntor, hyror, realisations- och värdeförändringar) |
| 7500–7999 | Placeringskostnader |
| 1xxx–2xxx | Balanskonton |

Sex rader har `altered = 2` och beloppet 999 999 på konto 3010 i mars. De ska inte synas i rapporterna. Syns de fungerar inte filtret.

## Testa budget och prognos

Gå till **Budget och prognos → Importera** och ladda upp filerna i `importfiler/`:

| Fil | Typ / år | Förväntat |
|---|---|---|
| `budget_2027.xlsx` | Budget, 2027 | Godkänd, resultat helår ca 2,8 MEUR |
| `prognos_2026_oktober.xlsx` | Prognos, 2026 | Godkänd: utfall jan–okt plus prognos nov–dec |
| `budget_med_fel.xlsx` | Budget, 2027 | Stoppas med fem fel: okänt konto, okänd Kst, text i stället för tal, dubblett och rad utan konto |

Efter importen kan versionen redigeras direkt i systemet via **Redigera**.

## Struktur

```
app/          det som körs: server.py, rapportering.html, start.bat, testdata och mall
src/app.html  källkoden till insyn.html (med platshållare för data och Excelläsare)
vendor/       SheetJS xlsx.mini 0.18.5 (Apache-2.0), bäddas in i HTML-filen
testdata/     skript som skapar den påhittade testdatan
importfiler/  testfiler för budget- och prognosimport
mall/         skript som skapar databas_visma_mall.xlsx
bygg.py       bygger app/insyn.html (och med --testdata även all testdata)
```

Efter ändringar i `src/app.html` kör du `python bygg.py`. Kör `python bygg.py --testdata` för att även skapa om testdatan; det kräver `pip install pandas openpyxl`.

## Koppla på riktig data

Lägg exporten från Visma som `databas_visma.xlsx` i samma mapp som `server.py` och starta om servern. Det är bara fliken `vourowx` som krävs; `databas_visma_mall.xlsx` visar formatet. Se även `app/LASMIG.txt`.
