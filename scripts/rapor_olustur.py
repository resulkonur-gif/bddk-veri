"""veri/aylik/uzun/*.csv dosyalarından tek dosyalık etkileşimli HTML rapor üretir.

Kullanım:
  python scripts/rapor_olustur.py                  # -> docs/index.html
  python scripts/rapor_olustur.py --cikti rapor.html
"""
import argparse
import base64
import gzip
import json
import re
from datetime import datetime, timezone, timedelta
from pathlib import Path

import numpy as np
import pandas as pd

KOK = Path(__file__).resolve().parent.parent
AYLAR = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]
GRUP_SIRA = ["Sektör", "Mevduat", "Mevduat-Yerli Özel", "Mevduat-Kamu", "Mevduat-Yabancı",
             "Katılım", "Kalkınma ve Yatırım", "Yerli Özel", "Kamu", "Yabancı"]
# Adı zaman içinde değişen kalemler -> güncel ad
AD_ESLE = {
    "Finansal Kiralama Borçları (Net)": "Kiralama İşlemlerinden Yükümlülükler (Net)",
    "Sermaye Hesaplamasına Dahil Edilecek Borçlanma Araçları": "Sermaye Hesaplamasına Dahil Edilecek Borçlanma Araçları (Net)",
}
ORAN_TABLOLARI = {15, 16, 17}


def kalem_adi(s):
    s = re.sub(r"\*+", "", str(s))
    # formül parantezlerini kaldır: (37+38), (11+12+13+29), (2 den 24'e)
    s = re.sub(r"\s*\((?:\d+\s*\+[\d+\s]*|\d+ den [^)]*)\)", "", s)
    s = re.sub(r"\s+", " ", s).strip()
    return AD_ESLE.get(s, s)


def yil_den(y):
    son, onlar = y % 10, y % 100
    if son == 0:
        ek = {0: "dan", 10: "dan", 20: "den", 30: "dan", 40: "tan", 50: "den", 60: "tan", 70: "ten", 80: "den", 90: "dan"}[onlar]
    else:
        ek = {1: "den", 2: "den", 3: "ten", 4: "ten", 5: "ten", 6: "dan", 7: "den", 8: "den", 9: "dan"}[son]
    return f"{y}’{ek}"


def yil_e(y):
    son, onlar = y % 10, y % 100
    if son == 0:
        ek = {0: "e", 10: "a", 20: "ye", 30: "a", 40: "a", 50: "ye", 60: "a", 70: "e", 80: "e", 90: "a"}[onlar]
    else:
        ek = {1: "e", 2: "ye", 3: "e", 4: "e", 5: "e", 6: "ya", 7: "ye", 8: "e", 9: "a"}[son]
    return f"{y}’{ek}"


def veri_hazirla():
    df = pd.concat(pd.read_csv(f, encoding="utf-8-sig", dtype={"donem": str})
                   for f in sorted((KOK / "veri" / "aylik" / "uzun").glob("*.csv")))
    df["ad"] = df.kalem.map(kalem_adi)
    bilgi = json.loads((KOK / "veri" / "aylik" / "tablolar.json").read_text("utf-8"))
    per = sorted(df.donem.unique())
    grp = [g for g in GRUP_SIRA if g in set(df.grup)] + sorted(set(df.grup) - set(GRUP_SIRA))
    pi = {p: i for i, p in enumerate(per)}
    gi = {g: i for i, g in enumerate(grp)}
    tablolar = []
    for no, d in df.groupby("tablo_no"):
        b = bilgi.get(str(no), {})
        cols = list(dict.fromkeys(b.get("sutunlar", []) + list(d.sutun.unique())))
        # kalem sırası: son dönemdeki sıra; yalnız eski dönemlerde olanlar en son göründükleri sıraya göre araya girer
        son = d.sort_values(["donem", "sira"]).drop_duplicates("ad", keep="last")
        son = son.assign(_p=son.donem.map(pi)).sort_values(["_p", "sira"], ascending=[False, True])
        sira_son = {a: s for a, s in zip(son[son._p == son._p.max()].ad, son[son._p == son._p.max()].sira)}
        items = list(sira_son)
        for _, r in son[son._p < son._p.max()].iterrows():
            # yeni sıra: aynı dönemde kendinden önce gelen kalemin hemen arkası
            onceki = d[(d.donem == r.donem) & (d.sira < r.sira)].sort_values("sira").ad.unique()
            yer = next((items.index(a) + 1 for a in reversed(list(onceki)) if a in items), len(items))
            items.insert(yer, r.ad)
        ii = {a: i for i, a in enumerate(items)}
        ci = {c: i for i, c in enumerate(cols)}
        arr = np.full((len(items), len(cols), len(grp), len(per)), np.nan)
        arr[d.ad.map(ii).values, d.sutun.map(ci).values, d.grup.map(gi).values, d.donem.map(pi).values] = d.deger.values
        dec = 3 if no in ORAN_TABLOLARI else 1
        v = [None if np.isnan(x) else (round(float(x), dec) if dec else int(round(x))) for x in arr.ravel()]
        v = [int(x) if isinstance(x, float) and x.is_integer() else x for x in v]
        tablolar.append({"no": int(no), "ad": b.get("tablo", d.tablo.iloc[0]), "birim": b.get("birim") or {15: "%", 17: "%", 16: "adet"}.get(int(no)),
                         "cols": cols, "items": [{"ad": a} for a in items], "v": v})
    tr = datetime.now(timezone(timedelta(hours=3)))
    meta = {"guncelleme": f"{tr.day} {AYLAR[tr.month - 1]} {tr.year}"}
    return {"per": per, "grp": grp, "tab": tablolar, "meta": meta}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--cikti", default=str(KOK / "docs" / "index.html"))
    a = ap.parse_args()
    D = veri_hazirla()
    js = json.dumps(D, ensure_ascii=False, separators=(",", ":"))
    b64 = base64.b64encode(gzip.compress(js.encode("utf-8"), 9, mtime=0)).decode()
    p0, p1 = D["per"][0], D["per"][-1]
    y0, m0 = map(int, p0.split("-"))
    y1, m1 = map(int, p1.split("-"))
    html = (KOK / "rapor_sablonu.html").read_text("utf-8")
    html = (html.replace("{{ARALIK_DEN_YA}}", f"{AYLAR[m0 - 1]} {yil_den(y0)} {AYLAR[m1 - 1]} {yil_e(y1)}")
                .replace("{{ARALIK}}", f"{AYLAR[m0 - 1]} {y0} – {AYLAR[m1 - 1]} {y1}")
                .replace("{{N}}", str(len(D["per"])))
                .replace("{{VERI}}", b64))
    out = Path(a.cikti)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(html, "utf-8")
    print(f"{out}: {out.stat().st_size / 1e6:.1f} MB, {len(D['per'])} dönem, {len(D['tab'])} tablo, "
          f"{sum(len(t['items']) for t in D['tab'])} kalem")


if __name__ == "__main__":
    main()
