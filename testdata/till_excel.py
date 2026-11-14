import sqlite3, pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from pathlib import Path
HERE = Path(__file__).parent
con = sqlite3.connect(HERE / "testdata.sqlite")
import sys
out = sys.argv[1] if len(sys.argv) > 1 else str(HERE / "testdata.xlsx")
sheets = {t: pd.read_sql(f"SELECT * FROM {t}", con) for t in ["vourow", "account", "object", "fiscalyear"]}
sheets["vourowx"] = sheets.pop("vourow")
sheets = {k: sheets[k] for k in ["vourowx", "account", "object", "fiscalyear"]}
sheets["vourowx"]["date"] = pd.to_datetime(sheets["vourowx"]["date"])
info = pd.DataFrame({"Läs mig": [
    "Påhittad testdata för rapporteringssystemet. Ingen riktig bokföring.",
    "vourowx: verifikationsrader med samma kolumner som Visma Control 14. 2025 (year=16) och jan–okt 2026 (year=17).",
    "amount = belopp i EUR (bolagets valuta), foramount = belopp i transaktionsvaluta (SEK-kurs ca 11,05).",
    "Positivt belopp = debet, negativt = kredit. Varje verifikation (year+serie+vouno) summerar till 0.",
    "Två påhittade bolag (cid ALANDS och ALLIV). Dimensioner: o1 Klass, o2 Slag, o3 Kst, o4 Placering, o5 Moment, o6 Undergrp, o7 Avdelning, o8 oanvänd. Koderna är påhittade – namn finns i fliken object.",
    "Kontoplan: 3xxx premier, 5xxx skador, 6xxx driftskostnader, 7xxx placeringsverksamheten, 1xxx–2xxx balans.",
    "Sex rader har altered = 2 (makulerade, belopp 999 999 på konto 3010 i mars). De ska INTE synas i rapporterna – gör de det fungerar inte filtret.",
    "account: påhittad kontoplan. type: T=tillgång, S=skuld/eget kapital, I=intäkt, K=kostnad, F=finansiell.",
    "fiscalyear: koppling mellan year-index och kalenderår.",
    f"Skapad {datetime.now():%Y-%m-%d}."]})
with pd.ExcelWriter(out, engine="openpyxl") as w:
    info.to_excel(w, sheet_name="Läs mig", index=False)
    for n, df in sheets.items():
        df.to_excel(w, sheet_name=n, index=False)

wb = load_workbook(out)
for ws in wb:
    for row in ws.iter_rows():
        for c in row:
            c.font = Font(name="Arial", size=10, bold=c.row == 1)
            if c.row == 1:
                c.fill = PatternFill("solid", fgColor="E7F1FB")
    if ws.title != "Läs mig":
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
    for i, col in enumerate(ws.columns, 1):
        width = max(len(str(c.value or "")) for c in col[:200])
        ws.column_dimensions[get_column_letter(i)].width = min(max(width + 2, 8), 110 if ws.title == "Läs mig" else 40)
ws = wb["vourowx"]
hdr = [c.value for c in ws[1]]
for name, fmt in [("date", "yyyy-mm-dd"), ("amount", "#,##0.00"), ("foramount", "#,##0.00")]:
    col = hdr.index(name) + 1
    for r in range(2, ws.max_row + 1):
        ws.cell(r, col).number_format = fmt
wb.save(out)
print(ws.max_row - 1, "rader")
