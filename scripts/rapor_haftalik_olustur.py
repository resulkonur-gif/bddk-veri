"""veri/haftalik/csv/*.csv dosyalarından haftalık HTML raporu üretir (docs/haftalik.html).

Veri yapısı ve HTML birleştirme yerel/bddk_indir.py içindeki ortak fonksiyonlardan gelir;
böylece GitHub sayfası ile yerel betiğin ürettiği rapor aynıdır.
"""
import argparse
import csv
import json
import sys
from pathlib import Path

KOK = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(KOK / "yerel"))
import bddk_indir as b  # noqa: E402


def kayitlar():
    meta = json.loads((KOK / "veri" / "haftalik" / "kalemler.json").read_text("utf-8"))
    ad = {(int(n), k["id"]): (k["sira"], k["ad"]) for n, t in meta.items() for k in t["kalemler"]}
    tablo = {int(n): t["ad"] for n, t in meta.items()}
    out = []
    for f in sorted((KOK / "veri" / "haftalik" / "csv").glob("*.csv")):
        with open(f, encoding="utf-8-sig", newline="") as fh:
            for r in csv.DictReader(fh):
                t, k = int(r["tablo_no"]), int(r["kalem_id"])
                sira, kalem = ad[(t, k)]
                out.append({"tarih": r["tarih"], "tablo_no": t, "tablo": tablo[t], "grup": r["grup"], "sira": sira,
                            "kalem": kalem, **{c: b.sayi(r[c]) for c in b.SUTUNLAR}})
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--cikti", default=str(KOK / "docs" / "haftalik.html"))
    a = ap.parse_args()
    D = b.haftalik_rapor_verisi(kayitlar())
    html = b.haftalik_rapor_html(D, (KOK / "rapor_haftalik_sablonu.html").read_text("utf-8"), None, "index.html")
    out = Path(a.cikti)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(html, "utf-8")
    print(f"{out}: {out.stat().st_size / 1e6:.1f} MB, {len(D['per'])} hafta, {len(D['tab'])} tablo")


if __name__ == "__main__":
    main()
