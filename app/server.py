"""
Insyn – lokal server
===========================
Kör rapporteringssystemet Insyn helt lokalt på den här datorn.

  * Läser bokföringsdata från en Excelfil: databas_visma.xlsx (flik vourowx) i samma mapp,
    eller databas.xlsx (testdata) om Visma-exporten inte finns.
  * Sparar inställningar, användare, rapporter och budgetar i installningar.json i samma mapp.
  * Lyssnar bara på 127.0.0.1 – ingen annan dator kan ansluta, och ingenting skickas ut på nätet.
  * Kräver bara Python 3.8+ (inga extra paket).
  * Verifikationsraderna läses med alla dimensioner o1–o8 (Klass, Slag, Kst, Placering, Moment, Undergrp, Avdelning)
    samt bolag (cid). Objektfliken kan ange bolagsnamn med dim = cid.
  * Rader med altered = 2 läses inte in.

Start:  python server.py                 (eller dubbelklicka på start.bat)
Annan fil: python server.py "C:\\mapp\\export.xlsx"
Annan port: python server.py --port 8800
"""
import argparse
import datetime as dt
import json
import os
import re
import sys
import threading
import webbrowser
import zipfile
import xml.etree.ElementTree as ET
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

HERE = os.path.dirname(os.path.abspath(__file__))
HTML = os.path.join(HERE, "insyn.html")
if not os.path.exists(HTML) and os.path.exists(os.path.join(HERE, "rapportering.html")):
    HTML = os.path.join(HERE, "rapportering.html")   # äldre filnamn
STATE = os.path.join(HERE, "installningar.json")
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

# ---------------------------------------------------------------- Excel-läsare (utan extra paket)
def _col(ref):
    n = 0
    for ch in ref:
        if ch.isalpha():
            n = n * 26 + (ord(ch.upper()) - 64)
        else:
            break
    return n - 1

def read_xlsx(path):
    """Returnerar {fliknamn: [[cell, ...], ...]} med råvärden (tal som float, text som str)."""
    out = {}
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        shared = []
        if "xl/sharedStrings.xml" in names:
            for si in ET.fromstring(z.read("xl/sharedStrings.xml")).iter(NS + "si"):
                shared.append("".join(t.text or "" for t in si.iter(NS + "t")))
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        target = {r.get("Id"): r.get("Target") for r in rels}
        for sh in wb.iter(NS + "sheet"):
            t = target.get(sh.get(RNS + "id"), "")
            t = t.lstrip("/")
            p = t if t.startswith("xl/") else "xl/" + t
            if p not in names:
                continue
            rows = []
            with z.open(p) as f:
                for ev, el in ET.iterparse(f):
                    if el.tag != NS + "row":
                        continue
                    row = []
                    for c in el.iter(NS + "c"):
                        i = _col(c.get("r", "A"))
                        typ = c.get("t")
                        v = c.find(NS + "v")
                        if typ == "s":
                            val = shared[int(v.text)] if v is not None else ""
                        elif typ == "inlineStr":
                            val = "".join(x.text or "" for x in c.iter(NS + "t"))
                        elif typ in ("str", "e"):
                            val = v.text if v is not None else ""
                        elif typ == "b":
                            val = (v.text == "1") if v is not None else ""
                        else:
                            if v is None or v.text is None:
                                val = ""
                            else:
                                try:
                                    val = float(v.text)
                                except ValueError:
                                    val = v.text
                        while len(row) < i:
                            row.append("")
                        row.append(val)
                    rows.append(row)
                    el.clear()
            out[sh.get("name")] = rows
    return out

# ---------------------------------------------------------------- tolkning till systemets format
def _num(v):
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace("\u00a0", "").replace(" ", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None

def _str(v):
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    s = str(v if v is not None else "").strip()
    return s[:-2] if s.endswith(".0") else s

def _date(v):
    if isinstance(v, float) and 20000 < v < 80000:
        return dt.date(1899, 12, 30) + dt.timedelta(days=int(v))
    s = str(v or "").strip()
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})", s)
    if m:
        return dt.date(int(m[1]), int(m[2]), int(m[3]))
    m = re.match(r"^(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})", s)
    if m:
        return dt.date(int(m[3]), int(m[2]), int(m[1]))
    return None

def _table(rows):
    """Första raden är rubriker -> lista av dict med gemena kolumnnamn."""
    if not rows:
        return [], []
    head = [str(h).strip().lower() for h in rows[0]]
    data = []
    for r in rows[1:]:
        if not any(str(x).strip() for x in r):
            continue
        data.append({head[i]: (r[i] if i < len(r) else "") for i in range(len(head))})
    return head, data

def parse_bookkeeping(path):
    sheets = read_xlsx(path)
    heads = {n: [str(h).strip().lower() for h in (rows[0] if rows else [])] for n, rows in sheets.items()}
    has = lambda n, cols: all(c in heads[n] for c in cols)
    vs = next((n for n in sheets if n.lower() == "vourowx" and has(n, ["account", "amount"])), None) \
        or next((n for n in sheets if n.lower() == "vourow" and has(n, ["account", "amount"])), None) \
        or next((n for n in sheets if has(n, ["account", "amount", "period"])), None)
    if not vs:
        raise ValueError("Hittar ingen flik med kolumnerna account, amount och period (verifikationsrader).")
    fy = {}
    fs = next((n for n in sheets if has(n, ["year", "startdate"])), None)
    if fs:
        for r in _table(sheets[fs])[1]:
            d = _date(r.get("startdate"))
            if d:
                fy[_str(r.get("year"))] = d.year
    rows, skipped, warn = [], 0, []
    cids, p13, altered2 = set(), 0, 0
    for r in _table(sheets[vs])[1]:
        acc, amt = _str(r.get("account")), _num(r.get("amount"))
        try:
            p = int(float(r.get("period")))
        except (TypeError, ValueError):
            p = 0
        if _num(r.get("altered")) == 2:   # altered = 2 ska inte med i rapporteringen
            altered2 += 1
            continue
        if not acc or amt is None or not (0 <= p <= 13):   # 0 = ingående balans, 13 = bokslutsperiod
            skipped += 1
            continue
        d = _date(r.get("date"))
        y = fy.get(_str(r.get("year")))
        if not y:
            try:
                yi = int(float(r.get("year")))
            except (TypeError, ValueError):
                yi = 0
            y = yi if yi > 1900 else (d.year if d else None)
        if not y:
            skipped += 1
            continue
        fa = _num(r.get("foramount"))
        try:
            vouno = int(float(r.get("vouno") or 0))
            rowno = int(float(r.get("rowno") or 0))
        except ValueError:
            vouno, rowno = 0, 0
        cids.add(_str(r.get("cid")))
        p13 += p == 13
        rows.append([y, min(p, 12), _str(r.get("serie")), vouno, rowno, d.isoformat() if d else "", acc,
                     _str(r.get("currency")) or "EUR", amt, fa if fa is not None else amt,
                     *[_str(r.get(f"o{i}")) for i in range(1, 9)], _str(r.get("cid")), str(r.get("xfreetext") or ""), _str(r.get("creator"))])
    if not rows:
        raise ValueError("Filen innehåller inga läsbara verifikationsrader.")
    accounts, objects = [], []
    acs = next((n for n in sheets if n != vs and has(n, ["account", "name"])), None)
    if acs:
        accounts = [[_str(r.get("account")), str(r.get("name") or ""), _str(r.get("type"))] for r in _table(sheets[acs])[1] if _str(r.get("account"))]
    else:
        warn.append("Ingen flik med kontoplan (kolumnerna account och name) – kontona visas utan namn.")
    obs = next((n for n in sheets if has(n, ["dim", "code", "name"])), None)
    if obs:
        objects = [[_str(r.get("dim")), _str(r.get("code")), str(r.get("name") or "")] for r in _table(sheets[obs])[1] if _str(r.get("code"))]
    else:
        warn.append("Ingen flik med objekt (kolumnerna dim, code och name) – dimensionernas koder visas utan namn.")
    cids.discard("")
    if len(cids) > 1:
        pass   # flera bolag hanteras som dimensionen Bolag
    if altered2:
        warn.append(f"{altered2} rader med altered = 2 exkluderades.")
    if p13:
        warn.append(f"{p13} rader i period 13 (bokslut) räknas in i december.")
    ob = sum(1 for r in rows if r[1] == 0)
    if ob:
        warn.append(f"{ob} rader med ingående balans (period 0) – används i balansräkningen, inte i resultatet.")
    if skipped:
        warn.append(f"{skipped} rader saknade konto, belopp, period eller år och hoppades över.")
    vb = {}
    for r in rows:
        k = (r[18], r[0], r[2], r[3])   # bolag, år, serie, verifikationsnummer
        vb[k] = vb.get(k, 0) + r[8]
    unbal = sum(1 for v in vb.values() if abs(v) > 0.01)
    if unbal:
        warn.append(f"{unbal} verifikationer summerar inte till noll.")
    info = {"name": os.path.basename(path), "rows": len(rows), "years": sorted({r[0] for r in rows}),
            "accounts": len({r[6] for r in rows}), "cc": len({r[12] for r in rows if r[12]}),
            "dims": [len({r[10 + i] for r in rows if r[10 + i]}) for i in range(9)],
            "vouchers": len(vb), "warn": warn, "sheet": vs, "cids": sorted(cids)}
    return {"rows": rows, "accounts": accounts, "objects": objects, "name": os.path.basename(path)}, info

# ---------------------------------------------------------------- cache: läs om bara när filen ändrats
class DataCache:
    def __init__(self, path):
        self.path = path
        self.lock = threading.Lock()
        self.mtime = None
        self.payload = None

    def status(self):
        try:
            st = os.stat(self.path)
            return {"file": self.path, "name": os.path.basename(self.path), "exists": True,
                    "mtime": dt.datetime.fromtimestamp(st.st_mtime).isoformat(timespec="seconds"), "size": st.st_size}
        except FileNotFoundError:
            return {"file": self.path, "name": os.path.basename(self.path), "exists": False}

    def get(self, force=False):
        with self.lock:
            st = self.status()
            if not st["exists"]:
                return {"error": f"Hittar inte datafilen {self.path}", "status": st}
            if force or self.payload is None or st["mtime"] != self.mtime:
                t0 = dt.datetime.now()
                try:
                    ds, info = parse_bookkeeping(self.path)
                except PermissionError:
                    return {"error": "Filen är låst – stäng den i Excel eller vänta tills den sparats klart och försök igen.", "status": st}
                except (zipfile.BadZipFile, KeyError):
                    return {"error": "Filen kunde inte läsas som .xlsx. Spara den som Excel-arbetsbok (.xlsx).", "status": st}
                except ValueError as e:
                    return {"error": str(e), "status": st}
                ds["fileDate"] = st["mtime"]
                ds["loadedAt"] = dt.datetime.now().isoformat(timespec="seconds")
                info["seconds"] = round((dt.datetime.now() - t0).total_seconds(), 1)
                self.payload = json.dumps({"dataset": ds, "info": info, "status": st}, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
                self.mtime = st["mtime"]
                print(f"[{dt.datetime.now():%H:%M:%S}] Läste {info['rows']} rader från {st['name']} på {info['seconds']} s")
            return self.payload

# ---------------------------------------------------------------- webbserver (bara localhost)
class Handler(BaseHTTPRequestHandler):
    server_version = "Insyn"

    def log_message(self, fmt, *args):
        pass

    def _local_only(self):
        host = (self.headers.get("Host") or "").split(":")[0]
        if self.client_address[0] not in ("127.0.0.1", "::1") or host not in ("127.0.0.1", "localhost"):
            self.send_error(403, "Endast lokal åtkomst")
            return False
        origin = self.headers.get("Origin")
        if origin and urlparse(origin).hostname not in ("127.0.0.1", "localhost"):
            self.send_error(403, "Fel ursprung")
            return False
        return True

    def _send(self, code, body, ctype="application/json; charset=utf-8"):
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Content-Security-Policy", "default-src 'self' 'unsafe-inline' data: blob:; connect-src 'self'")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if not self._local_only():
            return
        u = urlparse(self.path)
        if u.path in ("/", "/index.html", "/insyn.html", "/rapportering.html"):
            try:
                with open(HTML, "rb") as f:
                    return self._send(200, f.read(), "text/html; charset=utf-8")
            except FileNotFoundError:
                return self._send(404, "insyn.html saknas i mappen", "text/plain; charset=utf-8")
        if u.path == "/api/ping":
            return self._send(200, json.dumps({"ok": True, "version": 1}))
        if u.path == "/api/status":
            return self._send(200, json.dumps(self.server.cache.status(), ensure_ascii=False))
        if u.path == "/api/data":
            force = "reload" in parse_qs(u.query)
            res = self.server.cache.get(force)
            if isinstance(res, dict):
                return self._send(200, json.dumps(res, ensure_ascii=False))
            return self._send(200, res)
        if u.path == "/api/state":
            try:
                with open(STATE, "rb") as f:
                    return self._send(200, f.read())
            except FileNotFoundError:
                return self._send(200, "null")
        self._send(404, '{"error":"saknas"}')

    def do_PUT(self):
        if not self._local_only():
            return
        if urlparse(self.path).path != "/api/state":
            return self._send(404, '{"error":"saknas"}')
        n = int(self.headers.get("Content-Length") or 0)
        if n > 50_000_000:
            return self._send(413, '{"error":"för stor"}')
        body = self.rfile.read(n)
        try:
            json.loads(body)
        except ValueError:
            return self._send(400, '{"error":"ogiltig json"}')
        tmp = STATE + ".tmp"
        with open(tmp, "wb") as f:
            f.write(body)
        os.replace(tmp, STATE)
        self._send(200, '{"ok":true}')

def main():
    ap = argparse.ArgumentParser(description="Insyn – lokal server")
    ap.add_argument("datafil", nargs="?", default=None, help="Excelfil med verifikationsrader (standard: databas_visma.xlsx, annars databas.xlsx)")
    ap.add_argument("--port", type=int, default=8765)
    ap.add_argument("--ingen-webblasare", action="store_true", help="Öppna inte webbläsaren automatiskt")
    a = ap.parse_args()
    if a.datafil:
        path = a.datafil
    else:  # riktig export i första hand, annars testdatan
        path = next((os.path.join(HERE, f) for f in ("databas_visma.xlsx", "databas.xlsx") if os.path.exists(os.path.join(HERE, f))), os.path.join(HERE, "databas_visma.xlsx"))
    if not os.path.isabs(path):
        path = os.path.join(HERE, path)
    srv = ThreadingHTTPServer(("127.0.0.1", a.port), Handler)
    srv.cache = DataCache(path)
    url = f"http://127.0.0.1:{a.port}/"
    print("Insyn körs lokalt på", url)
    print("Datafil:", path, "" if os.path.exists(path) else "(saknas ännu – testdata visas tills filen finns)")
    print("Inställningar sparas i:", STATE)
    print("Stäng fönstret eller tryck Ctrl+C för att avsluta.\n")
    if os.path.exists(path):
        threading.Thread(target=srv.cache.get, daemon=True).start()
    if not a.ingen_webblasare:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("Avslutar.")

if __name__ == "__main__":
    if sys.version_info < (3, 8):
        sys.exit("Python 3.8 eller senare krävs.")
    main()
