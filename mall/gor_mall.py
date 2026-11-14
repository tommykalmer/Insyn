import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

F = "Arial"
H = Font(name=F, bold=True, color="FFFFFF", size=10)
HF = PatternFill("solid", fgColor="3563C9")
EX = PatternFill("solid", fgColor="F9EEDA")
B = Font(name=F, size=10)
BB = Font(name=F, size=10, bold=True)
T = Font(name=F, size=14, bold=True)
thin = Side(style="thin", color="E3E7EE")

wb = Workbook()

def sheet(ws, header, rows, widths, example=True, textcols=()):
    ws.append(header)
    for c in ws[1]:
        c.font, c.fill, c.alignment = H, HF, Alignment(vertical="center")
    for r in rows:
        ws.append(r)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = B
            if example: c.fill = EX
            if c.column in textcols: c.number_format = "@"
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

# ---------- Läs mig
ws = wb.active; ws.title = "Läs mig"
ws["A1"] = "Mall för databas_visma.xlsx"; ws["A1"].font = T
lines = [
    ("Syfte", "Filen läses av rapporteringssystemet (server.py). Spara den som databas_visma.xlsx i samma mapp som server.py."),
    ("Viktigast", "Fliken vourowx – verifikationsraderna. De andra flikarna är valfria men ger namn på konton, dimensioner och bolag."),
    ("Exempelrader", "Gula rader i varje flik är exempel. Ta bort dem och klistra in era egna data från rad 2."),
    ("Rubriker", "Rubrikerna i rad 1 måste finnas kvar och stavas exakt som i mallen (versaler/gemener spelar ingen roll). Kolumnernas ordning spelar ingen roll."),
    ("Extra kolumner", "Kolumner som systemet inte använder kan finnas kvar – de hoppas över."),
    ("Uppdatera", "Ersätt innehållet, spara filen och tryck Läs om (Inställningar → Datakälla). Systemet läser också in filen automatiskt inom 20 sekunder."),
    ("Stäng filen", "Om Excel håller filen öppen medan den sparas kan inläsningen misslyckas en gång – försök igen när filen är sparad."),
    ("", ""),
    ("Flik", "Innehåll"),
    ("vourowx", "OBLIGATORISK. En rad per verifikationsrad, direkt från Visma Control. Se fliken Kolumner för vad varje kolumn betyder."),
    ("account", "Valfri. Kontoplan: account, name, (type). Ger kontonamn i rapporterna."),
    ("object", "Valfri. Namn på dimensionskoder: dim (o1–o8 eller cid), code, name. Ger t.ex. ”410 Försäljning företag” i stället för bara ”410”, och namn på bolagen (dim = cid)."),
    ("fiscalyear", "Valfri. Behövs bara om year är ett index (t.ex. 17) OCH räkenskapsåret inte följer kalenderåret. Annars räknas året fram ur datumet."),
    ("Kolumner", "Beskrivning av kolumnerna i vourowx (endast information – läses inte av systemet)."),
    ("", ""),
    ("Belopp", "amount = belopp i bolagets valuta (EUR) med bokföringstecken: debet positivt, kredit negativt. Decimaltecken punkt eller komma fungerar."),
    ("Period", "1–12 = månad. 0 = ingående balans (används i balansräkningen). 13 = bokslutsperiod (räknas i december)."),
    ("Konton och koder", "Lagra gärna konto och dimensionskoder som text så att inledande nollor behålls (t.ex. 0100)."),
    ("Säkerhet", "Filen innehåller riktig bokföring. Lägg mappen där bara du har åtkomst och skicka inte filen vidare."),
]
for i, (a, b) in enumerate(lines, 3):
    ws.cell(i, 1, a).font = BB
    ws.cell(i, 2, b).font = B
    ws.cell(i, 2).alignment = Alignment(wrap_text=True, vertical="top")
    ws.cell(i, 1).alignment = Alignment(vertical="top")
    if a in ("Flik",):
        for c in (ws.cell(i, 1), ws.cell(i, 2)): c.font, c.fill = H, HF
ws.column_dimensions["A"].width = 18; ws.column_dimensions["B"].width = 110

# ---------- vourowx
cols = ["cid","year","period","serie","vouno","rowno","date","altered","account","resacc","currency","amount","foramount","number",
        "o1","o2","o3","o4","o5","o6","o7","o8","xfreetext","extraamount","group1","group2","group3","group4","vatcode","crtype","creator","basecurrency","baseamount"]
def vr(cid, year, per, serie, vouno, rowno, d, acc, cur, amt, famt, dims, text, creator):
    o = (dims + [""]*8)[:8]
    return [cid, year, per, serie, vouno, rowno, d, 0, acc, "", cur, amt, famt, 0, *o, text, 0, 0, "", "", "", 0, 4, creator, "", 0]
d1 = dt.datetime(2026, 1, 1); d2 = dt.datetime(2026, 1, 15)
rows = [
    vr("SAK", 17, 0, "IB", 1, 0, d1, "1910", "EUR", 1250000.00, 1250000.00, [], "Ingående balans", ""),
    vr("SAK", 17, 0, "IB", 1, 1, d1, "2091", "EUR", -1250000.00, -1250000.00, [], "Ingående balans", ""),
    vr("SAK", 17, 1, "PREM", 12, 0, d2, "3010", "EUR", -430000.00, -430000.00, ["10","1","400","","","","FSG"], "Premieinkomst januari", ""),
    vr("SAK", 17, 1, "PREM", 12, 1, d2, "1510", "EUR", 430000.00, 430000.00, [], "Premieinkomst januari", ""),
    vr("SAK", 17, 1, "IMANAL", 1, 0, d1, "2780", "SEK", -2261.42, -25000.00, [], "Exempel fritext 1", "5569387169"),
    vr("SAK", 17, 1, "IMANAL", 1, 1, d1, "1699", "SEK", 2261.42, 25000.00, [], "Exempel fritext 2", "5569387169"),
    vr("ALLIV", 17, 1, "LEV", 3, 0, d2, "6540", "EUR", 4800.00, 4800.00, ["","","300","","","","STAB"], "IT-licenser", "5590065024"),
    vr("ALLIV", 17, 1, "LEV", 3, 1, d2, "2440", "EUR", -4800.00, -4800.00, [], "IT-licenser", "5590065024"),
]
ws = wb.create_sheet("vourowx")
sheet(ws, cols, rows, [9,6,7,8,7,7,12,7,9,7,8,12,12,7]+[8]*8+[26,10,8,8,8,8,8,7,12,10,10], textcols=(1,4,9,15,16,17,18,19,20,21,22,31))
for row in ws.iter_rows(min_row=2):
    row[6].number_format = "yyyy-mm-dd"
    row[11].number_format = row[12].number_format = "#,##0.00"

# ---------- account
ws = wb.create_sheet("account")
sheet(ws, ["account","name","type"], [["1510","Premiefordringar","T"],["1910","Bankkonto","T"],["2091","Balanserad vinst","S"],["3010","Premieinkomst motorförsäkring","I"],["6540","IT-tjänster och licenser","K"]], [10,40,8], textcols=(1,))
ws["E1"] = "type är valfri: T = tillgång, S = skuld/eget kapital, I = intäkt, K = kostnad, F = finansiell"; ws["E1"].font = Font(name=F, size=9, italic=True)

# ---------- object
ws = wb.create_sheet("object")
sheet(ws, ["dim","code","name"], [
    ["o1","10","Motorfordon"],["o2","1","Direkt försäkring"],["o3","400","Försäljning privat"],["o3","300","IT"],["o7","FSG","Försäljning"],["o7","STAB","Stab"],
    ["cid","SAK","Sak Ab"],["cid","ALLIV","Testbolaget Liv Ab"]], [8,12,36], textcols=(1,2))
ws["E1"] = "dim = o1 Klass, o2 Slag, o3 Kst, o4 Placering, o5 Moment, o6 Undergrp, o7 Avdelning, o8, cid = bolag"; ws["E1"].font = Font(name=F, size=9, italic=True)

# ---------- fiscalyear
ws = wb.create_sheet("fiscalyear")
sheet(ws, ["year","startdate"], [[16, dt.datetime(2025,1,1)],[17, dt.datetime(2026,1,1)]], [8,14])
for row in ws.iter_rows(min_row=2): row[1].number_format = "yyyy-mm-dd"
ws["D1"] = "Valfri. Kopplar year-index till räkenskapsårets startdatum. Utan fliken räknas året fram ur kolumnen date."; ws["D1"].font = Font(name=F, size=9, italic=True)

# ---------- Kolumner (information)
info = {
 "cid":("Ja","Bolag – blir dimensionen Bolag","Text"),
 "year":("Ja","Räkenskapsår, index (t.ex. 17) eller årtal","Tal"),
 "period":("Ja","1–12 månad, 0 ingående balans, 13 bokslut","Tal"),
 "serie":("Ja","Verifikationsserie","Text"),
 "vouno":("Ja","Verifikationsnummer","Tal"),
 "rowno":("Ja","Radnummer inom verifikationen","Tal"),
 "date":("Ja","Verifikationsdatum – används för år om fiscalyear saknas","Datum"),
 "account":("Ja","Bokföringskonto","Text"),
 "currency":("Ja","Transaktionsvaluta","Text"),
 "amount":("Ja","Belopp i bolagets valuta, debet +, kredit −","Tal"),
 "foramount":("Ja","Belopp i transaktionsvaluta","Tal"),
 **{f"o{i}":("Ja",f"Dimension {n}","Text") for i,n in zip(range(1,9),["Klass","Slag","Kst","Placering","Moment","Undergrp","Avdelning","(oanvänd)"])},
 "xfreetext":("Ja","Verifikationstext – visas vid drill-down","Text"),
 "creator":("Ja","Skapare/motpart – visas vid drill-down","Text"),
}
ws = wb.create_sheet("Kolumner")
sheet(ws, ["Kolumn","Används","Betydelse","Format"], [[c, *info.get(c, ("Nej","Läses inte av systemet (får finnas kvar)",""))] for c in cols], [14,10,60,10], example=False)
for row in ws.iter_rows(min_row=2):
    if row[1].value == "Nej":
        for c in row: c.font = Font(name=F, size=10, color="8A93A5")

wb.active = 1
import sys
wb.save(sys.argv[1] if len(sys.argv) > 1 else "databas_visma_mall.xlsx")
print("klar")
