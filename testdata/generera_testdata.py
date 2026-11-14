"""
Genererar en testdatabas (SQLite) som efterliknar verifikationsradstabellen i Visma Control 14.
All data är påhittad. Kör: python generera_testdata.py  ->  skapar testdata.sqlite

Tabeller:
  vourow   - verifikationsrader, samma kolumner som i Visma Control
  account  - kontoplan (påhittad, försäkringsbolagslik)
  object   - objekt/dimensioner o1-o8 och bolag (cid)
fiscalyear - year-index -> kalenderår

Kontoplan: 3xxx premier, 5xxx skador, 6xxx driftskostnader, 7xxx placeringsverksamheten.
"""
import random
import sqlite3
from datetime import date, timedelta
from pathlib import Path

random.seed(42)
DB = Path(__file__).with_name("testdata.sqlite")
CID = "ALANDS"
SCALE = 1.0
COMPANIES = [("ALANDS", "Testbolaget Skade Ab", 1.0), ("ALLIV", "Testbolaget Liv Ab", 0.35)]  # påhittade bolag
YEARS = {2025: 16, 2026: 17}          # kalenderår -> Visma 'year'-index (2026 = 17 enligt exemplet)
LAST_PERIOD = {2025: 12, 2026: 10}    # 2026 bokfört t.o.m. oktober
SEK_RATE = 11.05                      # SEK per EUR, ungefär som i exemplet

# --- Kontoplan (påhittad, försäkringsbolag) ----------------------------------
# 3xxx premier, 5xxx skador, 6xxx driftskostnader, 7xxx placeringsverksamheten
ACCOUNTS = [
    # konto, namn, typ (T=tillgång, S=skuld/EK, I=intäkt, K=kostnad, F=finansiell)
    ("1200", "Inventarier", "T"),
    ("1209", "Ackumulerade avskrivningar inventarier", "T"),
    ("1510", "Premiefordringar", "T"),
    ("1600", "Fordringar på återförsäkrare", "T"),
    ("1699", "Övriga fordringar", "T"),
    ("1710", "Aktier och andelar", "T"),
    ("1720", "Obligationer och räntebärande värdepapper", "T"),
    ("1910", "Bankkonto", "T"),
    ("2091", "Balanserat resultat", "S"),
    ("2400", "Skulder avseende återförsäkring", "S"),
    ("2440", "Leverantörsskulder", "S"),
    ("2710", "Personalskatt", "S"),
    ("2730", "Lagstadgade sociala avgifter och pension", "S"),
    ("2780", "Skulder till försäkringstagare", "S"),
    ("2810", "Avsättning för ej intjänade premier", "S"),
    ("2820", "Avsättning för oreglerade skador", "S"),
    ("3010", "Premieinkomst motorfordon", "I"),
    ("3020", "Premieinkomst hem och villa", "I"),
    ("3030", "Premieinkomst företag", "I"),
    ("3040", "Premieinkomst båt", "I"),
    ("3050", "Premieinkomst olycksfall", "I"),
    ("3060", "Premieinkomst ansvar", "I"),
    ("3500", "Premier för avgiven återförsäkring", "I"),
    ("3600", "Förändring i avsättning för ej intjänade premier", "I"),
    ("5010", "Utbetalda försäkringsersättningar motorfordon", "K"),
    ("5020", "Utbetalda försäkringsersättningar hem och villa", "K"),
    ("5030", "Utbetalda försäkringsersättningar företag", "K"),
    ("5040", "Utbetalda försäkringsersättningar båt", "K"),
    ("5050", "Utbetalda försäkringsersättningar olycksfall", "K"),
    ("5060", "Utbetalda försäkringsersättningar ansvar", "K"),
    ("5400", "Återförsäkrares andel av ersättningar", "K"),
    ("5500", "Förändring i avsättning för oreglerade skador", "K"),
    ("5700", "Skaderegleringskostnader", "K"),
    ("6010", "Löner tjänstemän", "K"),
    ("6020", "Sociala avgifter och pensionskostnader", "K"),
    ("6100", "Lokalhyra", "K"),
    ("6200", "IT-tjänster och licenser", "K"),
    ("6300", "Konsultarvoden", "K"),
    ("6400", "Marknadsföring", "K"),
    ("6500", "Resekostnader", "K"),
    ("6600", "Kontorsmaterial", "K"),
    ("6700", "Provisioner till förmedlare", "K"),
    ("6800", "Avskrivningar inventarier", "K"),
    ("6900", "Övriga driftskostnader", "K"),
    ("7010", "Utdelningar på aktier och andelar", "F"),
    ("7020", "Ränteintäkter", "F"),
    ("7030", "Hyresintäkter fastigheter", "F"),
    ("7040", "Realisationsvinster placeringar", "F"),
    ("7050", "Orealiserade värdeförändringar placeringar", "F"),
    ("7500", "Realisationsförluster placeringar", "F"),
    ("7600", "Kapitalförvaltningskostnader", "F"),
    ("7700", "Fastighetskostnader", "F"),
]

COSTCENTERS = [
    ("100", "Ledning"), ("200", "Ekonomi och placeringar"), ("300", "IT"),
    ("400", "Försäljning privat"), ("410", "Försäljning företag"),
    ("500", "Skadereglering"), ("600", "Kundservice"),
]
PROJECTS = [("", ""), ("P01", "Nytt försäkringssystem"), ("P02", "Digital skadeanmälan")]
# Dimensioner (påhittade koder):  o1 Klass, o2 Slag, o3 Kst, o4 Placering, o5 Moment, o6 Undergrp, o7 Avdelning, o8 oanvänd
KLASS = [("10", "Motorfordon"), ("20", "Hem och villa"), ("30", "Företag"), ("40", "Båt"),
         ("50", "Olycksfall"), ("60", "Ansvar")]
SLAG = [("1", "Direkt försäkring"), ("2", "Mottagen återförsäkring"), ("3", "Avgiven återförsäkring")]
PLACERING = [("AKT", "Aktier"), ("RNT", "Räntebärande"), ("FAST", "Fastigheter"), ("BNK", "Bankmedel")]
MOMENT = [("UTB", "Utbetalning"), ("RES", "Reservförändring")]
UNDERGRP = [("P01", "Nytt försäkringssystem"), ("P02", "Digital skadeanmälan")]
AVDELNING = [("STAB", "Stab"), ("FSG", "Försäljning"), ("SKS", "Skador och service")]
AVD_AV_KST = {"100": "STAB", "200": "STAB", "300": "STAB", "400": "FSG", "410": "FSG", "500": "SKS", "600": "SKS"}
# försäkringsklass per premie-/skadekonto (3010 och 5010 = motor osv.)
KLASS_AV_KONTO = {f"{p}0{n}0": k for p in "35" for n, (k, _) in enumerate(KLASS, 1)}
PLACERING_AV_KONTO = {"7010": ["AKT"], "7020": ["RNT", "BNK"], "7030": ["FAST"], "7040": ["AKT", "RNT"],
                      "7050": ["AKT", "RNT"], "7500": ["AKT", "RNT"], "7600": ["AKT", "RNT"], "7700": ["FAST"]}


def dims_for(account, kst, proj):
    d = dict(o1="", o2="", o3=kst, o4="", o5="", o6=proj, o7=AVD_AV_KST.get(kst, ""), o8="")
    if account[0] in "35" and account != "5700":
        d["o1"] = KLASS_AV_KONTO.get(account) or random.choice(KLASS)[0]
        d["o2"] = "3" if account in ("3500", "5400") else ("2" if random.random() < 0.10 else "1")
    if account == "5700":
        d["o1"] = random.choice(KLASS)[0]
    if account[:2] == "50":
        d["o5"] = "RES" if random.random() < 0.2 else "UTB"
    if account == "5500":
        d["o5"] = "RES"
    if account[0] == "7":
        d["o4"] = random.choice(PLACERING_AV_KONTO[account])
    return d

# Månadsnivåer (EUR). Positivt = debet (kostnad), negativt = kredit (intäkt).
# (medel, spridning, kostnadsställen). Spridning > 1 ger konton som kan byta tecken.
MONTHLY = {
    # premier
    "3010": (-420_000, 0.08, ["400"]), "3020": (-260_000, 0.06, ["400"]),
    "3030": (-310_000, 0.10, ["410"]), "3040": (-90_000, 0.40, ["400"]),
    "3050": (-70_000, 0.10, ["400", "410"]), "3060": (-45_000, 0.15, ["410"]),
    "3500": (85_000, 0.05, ["100"]),
    "3600": (0, 0, ["100"]),            # säsongsberoende, se nedan
    # skador
    "5010": (240_000, 0.25, ["500"]), "5020": (130_000, 0.35, ["500"]),
    "5030": (150_000, 0.40, ["500"]), "5040": (40_000, 0.60, ["500"]),
    "5050": (35_000, 0.35, ["500"]), "5060": (20_000, 0.70, ["500"]),
    "5400": (-55_000, 0.30, ["500"]),
    "5500": (10_000, 4.0, ["500"]),
    "5700": (35_000, 0.15, ["500", "600"]),
    # driftskostnader
    "6010": (190_000, 0.03, [c for c, _ in COSTCENTERS]),
    "6100": (28_000, 0.0, ["100"]), "6200": (26_000, 0.2, ["300"]),
    "6300": (12_000, 0.6, ["200", "300"]), "6400": (18_000, 0.4, ["400", "410"]),
    "6500": (6_000, 0.5, ["100", "400", "410", "500"]), "6600": (1_500, 0.4, ["200", "600"]),
    "6700": (22_000, 0.2, ["400", "410"]), "6800": (9_000, 0.0, ["300"]),
    "6900": (4_000, 0.5, ["100", "200", "600"]),
    # placeringsverksamheten
    "7010": (-12_000, 0.5, ["200"]), "7020": (-14_000, 0.2, ["200"]),
    "7030": (-8_000, 0.05, ["200"]), "7040": (-20_000, 1.0, ["200"]),
    "7050": (-10_000, 6.0, ["200"]), "7500": (8_000, 1.0, ["200"]),
    "7600": (4_000, 0.2, ["200"]), "7700": (3_500, 0.4, ["200"]),
}
SEASON_BOAT = {5: 2.5, 6: 2.0, 7: 1.4, 8: 1.2}   # båtpremier/skador högre på sommaren
DIVIDEND = {3: 1.5, 4: 6.0, 5: 3.0}             # utdelningar främst under våren
PREMIUM_RESERVE = {1: 160_000, 2: 90_000, 3: 40_000, 4: 10_000, 5: -10_000, 6: -30_000,
                   7: -40_000, 8: -50_000, 9: -40_000, 10: -40_000, 11: -40_000, 12: -50_000}

ORGNR = [f"{random.randint(5560000000, 5599999999)}" for _ in range(25)]
COLUMNS = ["cid", "year", "period", "serie", "vouno", "rowno", "date", "altered", "account",
           "resacc", "currency", "amount", "foramount", "number", "o1", "o2", "o3", "o4", "o5",
           "o6", "o7", "o8", "xfreetext", "extraamount", "group1", "group2", "group3", "group4",
           "vatcode", "crtype", "creator", "basecurrency", "baseamount"]


def row(year, period, serie, vouno, rowno, d, account, amount, text, kst="", proj="",
        currency="EUR", creator=""):
    dims = dims_for(account, kst, proj) if kst or account[0] in "3567" else dict(o1="", o2="", o3="", o4="", o5="", o6="", o7="", o8="")
    foramount = round(amount * SEK_RATE) if currency == "SEK" else amount
    return dict(cid=CID, year=YEARS[year], period=period, serie=serie, vouno=vouno, rowno=rowno,
                date=d.strftime("%Y-%m-%d 00:00:00"), altered=0, account=account, resacc="",
                currency=currency, amount=round(amount, 2), foramount=round(foramount, 2), number=0,
                **dims, xfreetext=text,
                extraamount=0, group1=0, group2="", group3="", group4="", vatcode=0, crtype=4,
                creator=creator, basecurrency="", baseamount=0)


COUNTER = {"3500": "2400", "3600": "2810", "5400": "1600", "5500": "2820", "6010": "2710",
           "6020": "2730", "6800": "1209", "7050": "1710", "7600": "2440", "7700": "2440"}


def counter_account(acc):
    if acc in COUNTER:
        return COUNTER[acc]
    if acc[0] == "3":
        return "1510"   # premie -> premiefordran
    if acc[0] in "57":
        return "1910"   # skadeutbetalningar och placeringsflöden via bank
    return "2440"       # övriga driftskostnader -> leverantörsskuld


def generate():
    rows, vouno = [], {}

    def next_no(year, serie):
        vouno[(year, serie)] = vouno.get((year, serie), 0) + 1
        return vouno[(year, serie)]

    for year in YEARS:
        growth = 1.0 if year == 2025 else 1.04
        for period in range(1, LAST_PERIOD[year] + 1):
            first = date(year, period, 1)
            last = (date(year + (period == 12), period % 12 + 1, 1) - timedelta(days=1))
            for acc, (mean, spread, ccs) in MONTHLY.items():
                season = SEASON_BOAT.get(period, 0.6) if acc in ("3040", "5040") else 1.0
                if acc == "7010":
                    season = DIVIDEND.get(period, 0.2)
                if acc == "3600":
                    mean = PREMIUM_RESERVE[period]
                    spread = 0.2
                total = mean * growth * season * SCALE * (1 + random.uniform(-spread, spread))
                if not total:
                    continue
                # dela upp på 1-4 verifikationer och kostnadsställen
                split_cc = acc == "6010"   # löner fördelas på alla kostnadsställen
                parts = len(ccs) if split_cc else random.randint(1, 4) if acc[0] == "5" else 1
                weights = [random.uniform(0.5, 1.5) for _ in range(parts)]
                for p in range(parts):
                    a = round(total * weights[p] / sum(weights), 2)
                    cc = ccs[p] if split_cc else random.choice(ccs)
                    proj = random.choice(PROJECTS)[0] if acc in ("6200", "6300") else ""
                    serie = {"3": "PREM", "5": "SKAD", "7": "PLAC"}.get(acc[0], "LEV")
                    if acc == "6010":
                        serie = "LON"
                    if acc in ("3600", "5500", "6800", "7050"):
                        serie = "MAN"
                    d = random.choice([first + timedelta(days=random.randint(0, 27)), last])
                    no = next_no(year, serie)
                    creator = random.choice(ORGNR) if serie == "LEV" else ""
                    cur = "SEK" if serie == "LEV" and random.random() < 0.15 else "EUR"
                    text = f"{dict((x[0], x[1]) for x in ACCOUNTS)[acc]} {period:02d}/{year}"
                    rows.append(row(year, period, serie, no, 0, d, acc, a, text, cc, proj, cur, creator))
                    rows.append(row(year, period, serie, no, 1, d, counter_account(acc), -a, text,
                                    "", "", cur, creator))
            # sociala avgifter på löner, en verifikation per period
            no = next_no(year, "LON")
            soc = round(MONTHLY["6010"][0] * growth * SCALE * 0.31, 2)
            rows.append(row(year, period, "LON", no, 0, last, "6020", soc, f"Sociala avgifter {period:02d}/{year}", "100"))
            rows.append(row(year, period, "LON", no, 1, last, "2730", -soc, f"Sociala avgifter {period:02d}/{year}"))
            # import-serie som i exemplet: 2780 mot 1699
            for _ in range(random.randint(3, 6)):
                no = next_no(year, "IMANAL")
                sek = random.randint(10_000, 150_000)
                eur = round(sek / SEK_RATE, 2)
                org = random.choice(ORGNR)
                rows.append(row(year, period, "IMANAL", no, 0, first, "2780", -eur, f"Import {org}", currency="SEK", creator=org))
                rows.append(row(year, period, "IMANAL", no, 1, first, "1699", eur, f"Import {org}", currency="SEK", creator=org))
                rows[-2]["foramount"], rows[-1]["foramount"] = -sek, sek
    # några rader med altered = 2 (ändrade/makulerade i Visma) – ska INTE räknas.
    # Stora belopp gör att det syns direkt i rapporterna om filtret inte fungerar.
    for year in YEARS:
        for _ in range(3):
            no = next_no(year, "PREM")
            d = date(year, 3, 15)
            r1 = row(year, 3, "PREM", no, 0, d, "3010", -999_999, "Makulerad rad (altered=2)", "400")
            r2 = row(year, 3, "PREM", no, 1, d, "1510", 999_999, "Makulerad rad (altered=2)")
            r1["altered"] = r2["altered"] = 2
            rows += [r1, r2]
    return rows


def main():
    if DB.exists():
        DB.unlink()
    con = sqlite3.connect(DB)
    con.executescript("""
    CREATE TABLE vourow (
        cid TEXT, year INTEGER, period INTEGER, serie TEXT, vouno INTEGER, rowno INTEGER,
        date TEXT, altered INTEGER, account TEXT, resacc TEXT, currency TEXT,
        amount REAL, foramount REAL, number INTEGER,
        o1 TEXT, o2 TEXT, o3 TEXT, o4 TEXT, o5 TEXT, o6 TEXT, o7 TEXT, o8 TEXT,
        xfreetext TEXT, extraamount REAL, group1 INTEGER, group2 TEXT, group3 TEXT, group4 TEXT,
        vatcode INTEGER, crtype INTEGER, creator TEXT, basecurrency TEXT, baseamount REAL,
        PRIMARY KEY (cid, year, serie, vouno, rowno)
    );
    CREATE TABLE account (cid TEXT, account TEXT, name TEXT, type TEXT, PRIMARY KEY (cid, account));
    CREATE TABLE object (cid TEXT, dim TEXT, code TEXT, name TEXT, PRIMARY KEY (cid, dim, code));
    CREATE TABLE fiscalyear (cid TEXT, year INTEGER, startdate TEXT, enddate TEXT, PRIMARY KEY (cid, year));
    """)
    global CID, SCALE
    rows = []
    for cid, _, sc in COMPANIES:
        CID, SCALE = cid, sc
        rows += generate()
    con.executemany(f"INSERT INTO vourow VALUES ({','.join('?' * len(COLUMNS))})",
                    [[r[c] for c in COLUMNS] for r in rows])
    con.executemany("INSERT INTO account VALUES (?,?,?,?)", [(c, *a) for c, _, _ in COMPANIES for a in ACCOUNTS])
    con.executemany("INSERT INTO object VALUES (?,?,?,?)",
                    [(CID, "o1", c, n) for c, n in KLASS] + [(CID, "o2", c, n) for c, n in SLAG] +
                    [(CID, "o3", c, n) for c, n in COSTCENTERS] + [(CID, "o4", c, n) for c, n in PLACERING] +
                    [(CID, "o5", c, n) for c, n in MOMENT] + [(CID, "o6", c, n) for c, n in UNDERGRP] +
                    [(CID, "o7", c, n) for c, n in AVDELNING] + [(CID, "cid", c, n) for c, n, _ in COMPANIES])
    con.executemany("INSERT INTO fiscalyear VALUES (?,?,?,?)",
                    [(c, idx, f"{y}-01-01", f"{y}-12-31") for c, _, _ in COMPANIES for y, idx in YEARS.items()])
    con.commit()

    # kontroll: varje verifikation ska balansera
    bad = con.execute("""SELECT year, serie, vouno, ROUND(SUM(amount),2) s FROM vourow
                         GROUP BY year, serie, vouno HAVING ABS(s) > 0.005""").fetchall()
    n = con.execute("SELECT COUNT(*) FROM vourow").fetchone()[0]
    print(f"{n} rader skrivna till {DB.name}. Obalanserade verifikationer: {len(bad)}")
    con.close()


if __name__ == "__main__":
    main()
