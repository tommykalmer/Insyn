"""Bygger data.json (inbyggd testdata i HTML-filen) från testdata.sqlite. Körs av bygg.py."""
import sqlite3, json, random, sys
from pathlib import Path
HERE = Path(__file__).parent
from collections import defaultdict
random.seed(7)
con = sqlite3.connect(HERE / "testdata.sqlite")
fy = {r[0]: int(r[1][:4]) for r in con.execute("SELECT DISTINCT year, startdate FROM fiscalyear")}
DIMS = ["o1","o2","o3","o4","o5","o6","o7","o8"]
rows = []
for r in con.execute(f"SELECT year,period,serie,vouno,rowno,substr(date,1,10),account,currency,amount,foramount,{','.join(DIMS)},cid,xfreetext,creator FROM vourow WHERE altered<>2 ORDER BY cid,year,period,serie,vouno,rowno"):
    r = list(r); r[0] = fy[r[0]]; rows.append(r)
accounts = [list(r) for r in con.execute("SELECT account,name,type FROM account ORDER BY account")]
objects = [list(r) for r in con.execute("SELECT DISTINCT dim,code,name FROM object ORDER BY dim,code")]
accounts = [list(r) for r in con.execute("SELECT DISTINCT account,name,type FROM account ORDER BY account")]
# utfall per (år, konto, dimensioner, period) för resultatkonton
agg = defaultdict(float)
for r in rows:
    if r[6][0] in "345678": agg[(r[0], r[6], tuple(r[10:19]), r[1])] += r[8]
def annual(y):
    a = defaultdict(float); m = defaultdict(float)
    for (yy,acc,d,p),v in agg.items():
        if yy==y: a[(acc,d)] += v; m[(acc,d,p)] += v
    return a, m
a25, m25 = annual(2025); a26, m26 = annual(2026)
SEASONAL = ("3040","5040","3600","7010")
def budget(base_a, base_m, growth, noise):
    out = []
    for (acc,d),tot in base_a.items():
        for p in range(1,13):
            share = base_m[(acc,d,p)]/tot if acc in SEASONAL and tot else 1/12
            v = round(tot*growth*share*(1+random.uniform(-noise,noise)),-1)
            if v: out.append([acc,p,v,*d])
    return out
bud26 = budget(a25, m25, 1.03, 0.03)
bud25 = budget(a25, m25, 0.98, 0.06)
bud26m = defaultdict(float)
for r in bud26: bud26m[(r[0],tuple(r[3:12]),r[1])] += r[2]
prg = []
for (acc,d) in sorted(set(list(a25.keys())+list(a26.keys()))):
    act6 = sum(m26.get((acc,d,p),0) for p in range(1,7)); bud6 = sum(bud26m.get((acc,d,p),0) for p in range(1,7))
    k = max(0.8, min(1.2, act6/bud6)) if bud6 else 1
    for p in range(1,13):
        v = m26.get((acc,d,p),0) if p<=6 else bud26m.get((acc,d,p),0)*k
        v = round(v,-1)
        if v: prg.append([acc,p,v,*d])
versions = [
  {"id":"b2025","name":"Budget 2025","type":"budget","year":2025,"locked":True,"source":"Inbyggd testdata","rows":bud25},
  {"id":"b2026","name":"Budget 2026","type":"budget","year":2026,"locked":True,"source":"Inbyggd testdata","rows":bud26},
  {"id":"p2026-1","name":"Prognos 1 2026 (juni)","type":"prognos","year":2026,"locked":False,"source":"Inbyggd testdata","rows":prg},
]
data = {"company":"ALANDS","currency":"EUR","fiscalyears":fy,"rows":rows,"accounts":accounts,"objects":objects,"versions":versions}
s = json.dumps(data, ensure_ascii=False, separators=(",",":"))
open(sys.argv[1] if len(sys.argv) > 1 else HERE / "data.json","w").write(s); print(len(s)//1024, "KB", len(rows), "rader", len(bud26), "budgetrader")
