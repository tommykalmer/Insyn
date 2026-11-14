"""
Bygger om allt som genereras i repot – bara påhittad testdata.

    python bygg.py            # bygger app/insyn.html från src/app.html
    python bygg.py --testdata # skapar dessutom om testdatabasen, app/databas.xlsx och importfilerna

Kräver Python 3.10+. För --testdata behövs pandas och openpyxl (pip install pandas openpyxl).
"""
import subprocess
import sys
from pathlib import Path

ROT = Path(__file__).parent
TD = ROT / "testdata"


def kor(*args):
    print(">", " ".join(str(a) for a in args))
    subprocess.run([sys.executable, *map(str, args)], check=True, cwd=ROT)


def bygg_html():
    html = (ROT / "src" / "app.html").read_text(encoding="utf-8")
    data = (TD / "data.json").read_text(encoding="utf-8")
    xlsx = (ROT / "vendor" / "xlsx.mini.min.js").read_text(encoding="utf-8").replace("</script", "<\\/script")
    assert "/*__DATA__*/" in html and "/*__XLSX__*/" in html, "platshållare saknas i src/app.html"
    ut = html.replace("/*__DATA__*/", data).replace("/*__XLSX__*/", xlsx)
    (ROT / "app" / "insyn.html").write_text(ut, encoding="utf-8")
    print(f"app/insyn.html: {len(ut) // 1024} KB")


if __name__ == "__main__":
    if "--testdata" in sys.argv or not (TD / "data.json").exists():
        kor(TD / "generera_testdata.py")
        kor(TD / "till_excel.py", ROT / "app" / "databas.xlsx")
        kor(TD / "gor_importfiler.py", ROT / "importfiler")
        kor(TD / "bygg_data.py", TD / "data.json")
        kor(ROT / "mall" / "gor_mall.py", ROT / "app" / "databas_visma_mall.xlsx")
    bygg_html()
