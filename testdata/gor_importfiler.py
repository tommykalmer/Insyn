"""
Skapar testfiler för import av budget och prognos (påhittade siffror).
  budget_2027.xlsx          - budget för 2027, ca +4 % mot utfall 2025/2026
  prognos_2026_oktober.xlsx - utfall jan–okt 2026 + prognos nov–dec
  budget_med_fel.xlsx       - avsiktliga fel för att testa valideringen
Belopp anges som resultatpåverkan: intäkter positiva, kostnader negativa.
Kör: python gor_importfiler.py [utmapp]
"""
import random, sqlite3, sys
from collections import defaultdict
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

random.seed(11)
HERE = Path(__file__).parent
OUT = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE
con = sqlite3.connect(HERE / "testdata.sqlite")
fy = dict(con.execute("SELECT DISTINCT year, CAST(substr(startdate,1,4) AS INT) FROM fiscalyear"))
names = dict(con.execute("SELECT DISTINCT account, name FROM account"))
HEAD = ["Bolag", "Konto", "Kontonamn", "Klass", "Kst", "Avdelning"] + [str(m) for m in range(1, 13)]

# utfall per (år, bolag, konto, klass, kst, avdelning, period), resultatpåverkan (= -belopp)
act = defaultdict(float)
for y, p, cid, acc, o1, o3, o7, amt in con.execute(
        "SELECT year, period, cid, account, o1, o3, o7, amount FROM vourow "
        "WHERE altered<>2 AND account>='3' AND account<'8'"):
    act[(fy[y], cid, acc, o1, o3, o7, p)] -= amt
keys = sorted({k[1:6] for k in act})


def write(path, rows, title):
    wb = Workbook(); ws = wb.active; ws.title = "Import"
    ws.append(HEAD)
    for r in rows:
        ws.append(r)
    for c in ws[1]:
        c.font = Font(name="Arial", bold=True, size=10); c.fill = PatternFill("solid", fgColor="E7F1FB")
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = Font(name="Arial", size=10)
            if c.column > 6:
                c.number_format = "#,##0;-#,##0;-"
    for col, w in zip("ABCDEF", [9, 8, 44, 7, 6, 11]):
        ws.column_dimensions[col].width = w
    for i in range(7, 19):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = 10
    ws.freeze_panes = "G2"; ws.auto_filter.ref = ws.dimensions
    info = wb.create_sheet("Läs mig")
    for line in [title,
                 "Påhittade siffror – bara för test av importfunktionen.",
                 "Systemet läser FÖRSTA fliken (Import). En rad per bolag + konto + dimensionskombination.",
                 "Kolumner 1–12 = månader. Belopp = resultatpåverkan: intäkter positiva, kostnader negativa.",
                 "Dimensionskolumner (Klass, Kst, Avdelning, Bolag) är valfria – tomma celler betyder ingen dimension.",
                 "Importera via Budget och prognos → Importera, välj typ och år och ladda upp filen."]:
        info.append([line]); info.cell(info.max_row, 1).font = Font(name="Arial", size=10, bold=info.max_row == 1)
    info.column_dimensions["A"].width = 110
    wb.save(path); print(path.name, len(rows), "rader")


# --- Budget 2027: snitt av 2025 och 2026 (2026 uppräknat till helår), +4 %, jämn eller säsong
SEASONAL = ("3040", "5040", "3600", "7010")
bud = []
for k in keys:
    cid, acc, o1, o3, o7 = k
    m25 = [act.get((2025, *k, p), 0) for p in range(1, 13)]
    t26 = sum(act.get((2026, *k, p), 0) for p in range(1, 11)) * 12 / 10
    base = (sum(m25) + t26) / 2 * 1.04
    if abs(base) < 500:
        continue
    tot25 = sum(m25)
    shares = [x / tot25 for x in m25] if acc in SEASONAL and tot25 else [1 / 12] * 12
    vals = [round(base * s * (1 + random.uniform(-0.02, 0.02)), -1) for s in shares]
    bud.append([cid, acc, names.get(acc, ""), o1, o3, o7, *vals])
write(OUT / "budget_2027.xlsx", bud, "Testbudget 2027 – importera som typ Budget, år 2027")

# --- Prognos 2026 (oktober): utfall jan–okt, nov–dec = snitt jan–okt * 1,02 (säsongskonton från 2025)
prg = []
for k in keys:
    cid, acc, o1, o3, o7 = k
    a = [act.get((2026, *k, p), 0) for p in range(1, 11)]
    avg = sum(a) / 10 * 1.02
    if acc in SEASONAL:
        rest = [act.get((2025, *k, p), 0) * 1.04 for p in (11, 12)]
    else:
        rest = [avg, avg]
    vals = [round(x, -1) for x in a + rest]
    if any(vals):
        prg.append([cid, acc, names.get(acc, ""), o1, o3, o7, *vals])
write(OUT / "prognos_2026_oktober.xlsx", prg, "Testprognos 2026 (oktober) – importera som typ Prognos, år 2026")

# --- Fil med avsiktliga fel
bad = [r[:] for r in bud[:6]]
bad[1][1] = "3999"                        # konto finns inte
bad[2][4] = "999"                         # Kst finns inte
bad[3][8] = "tusen"                       # inte ett tal
bad.append(bad[0][:])                     # dubblett
bad.append(["ALANDS", "", "Rad utan konto", "", "400", "FSG", *[1000] * 12])
write(OUT / "budget_med_fel.xlsx", bad, "Fil med avsiktliga fel – importen ska stoppas med 5 felmeddelanden")
