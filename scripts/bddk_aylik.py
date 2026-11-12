"""BDDK Aylık Bülten verilerini çeker ve veri/aylik altında arşivler.

Kullanım:
  python scripts/bddk_aylik.py --baslangic 2018-01 --bitis 2026-09   # belirli aralık
  python scripts/bddk_aylik.py --guncelle                             # yeni dönem var mı bak, varsa son 12 ayı yenile

Çıktılar:
  veri/aylik/uzun/bddk_aylik_<yıl>.csv   uzun biçim: donem,tablo_no,tablo,grup,sira,kalem,sutun,deger
  veri/aylik/excel/BDDK_Aylik_Arsiv_<yıl>.xlsx   Bilgi sayfası + her tablo için bir sayfa
  veri/aylik/tablolar.json               tablo adları, birimleri ve sütunları
  veri/aylik/son_donem.txt               arşivdeki en son dönem
"""
import argparse
import io
import json
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timezone
from pathlib import Path

import pandas as pd
import requests
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Font, PatternFill

TEMEL = "https://www.bddk.org.tr/BultenAylik/tr/Home/BasitExceleAktar"
TABLOLAR = {
    1: "Bilanço", 2: "Kar Zarar", 3: "Krediler", 4: "Tüketici Kredileri",
    5: "Sektörel Kredi Dağılımı", 6: "KOBİ Kredileri", 7: "Sendikasyon Seküritizasyon Kredileri",
    8: "Menkul Kıymetler", 9: "Mevduat Türler İtibarıyla", 10: "Mevduat Vade İtibarıyla",
    11: "Likidite Durumu", 12: "Sermaye Yeterliliği", 13: "Yabancı Para Pozisyonu",
    14: "Bilanço Dışı İşlemler", 15: "Rasyolar", 16: "Diğer Bilgiler", 17: "Yurt Dışı Şube Rasyoları",
}
GRUPLAR = [10001, 10002, 10008, 10009, 10010, 10003, 10004, 10005, 10006, 10007]
GRUP_SIRA = ["Sektör", "Mevduat", "Mevduat-Yerli Özel", "Mevduat-Kamu", "Mevduat-Yabancı",
             "Katılım", "Kalkınma ve Yatırım", "Yerli Özel", "Kamu", "Yabancı"]
KOK = Path(__file__).resolve().parent.parent / "veri" / "aylik"
UZUN_SUTUNLAR = ["donem", "tablo_no", "tablo", "grup", "sira", "kalem", "sutun", "deger"]

oturum = requests.Session()
oturum.headers["User-Agent"] = "Mozilla/5.0 (bddk-veri arsiv; github.com/resulkonur-gif/bddk-veri)"


def indir(tablo_no, yil, ay, deneme=4):
    params = [("tabloNo", tablo_no), ("yil", yil), ("ay", ay), ("paraBirimi", "TL")]
    params += [(f"taraf[{i}]", g) for i, g in enumerate(GRUPLAR)]
    for k in range(deneme):
        try:
            r = oturum.get(TEMEL, params=params, timeout=90)
            if r.status_code == 200 and r.content[:2] == b"PK":
                return r.content
            hata = f"HTTP {r.status_code}, {len(r.content)} bayt"
        except requests.RequestException as e:
            hata = str(e)
        time.sleep(2 * (k + 1))
    raise RuntimeError(f"tablo {tablo_no} {yil}/{ay}: {hata}")


def ayristir(icerik):
    """xlsx -> (başlık satırı, veri satırları)."""
    ws = load_workbook(io.BytesIO(icerik), read_only=True).active
    satirlar = [list(r) for r in ws.iter_rows(values_only=True)]
    return satirlar[0], [r for r in satirlar[1:] if r and r[0]]


def sayi(x):
    if x is None or x == "":
        return None
    try:
        return float(str(x).replace(",", "."))
    except ValueError:
        return None


def donem_cek(yil, ay):
    """Bir dönemin 17 tablosunu çeker; uzun biçimli DataFrame ve tablo bilgisi döner."""
    kayitlar, bilgi = [], {}
    for t in TABLOLAR:
        baslik, satirlar = ayristir(indir(t, yil, ay))
        sutunlar = [str(c) for c in baslik[3:] if c not in (None, "")]
        m = re.search(r"\(([^)]*)\)", str(baslik[2]))
        bilgi[t] = {"tablo": TABLOLAR[t], "baslik": str(baslik[2]), "birim": m.group(1) if m else None,
                    "sutunlar": sutunlar}
        for r in satirlar:
            for j, s in enumerate(sutunlar):
                kayitlar.append((f"{yil}-{ay:02d}", t, TABLOLAR[t], str(r[0]).strip(),
                                 int(float(r[1])), str(r[2]).strip(), s, sayi(r[3 + j])))
    return pd.DataFrame(kayitlar, columns=UZUN_SUTUNLAR), bilgi


def yayinlandi_mi(yil, ay):
    """Dönemin yayımlanıp yayımlanmadığını Bilanço başlığındaki dönemden anlar."""
    try:
        baslik, satirlar = ayristir(indir(1, yil, ay, deneme=2))
    except RuntimeError:
        return False
    return f"Dönem:{yil}/{ay}" in str(baslik[2]) and len(satirlar) > 0


def kontrol(df, donem):
    """Bilançoda her grup için toplam aktif = toplam pasif."""
    b = df[(df.tablo_no == 1) & (df.sutun == "Toplam")]
    sorun = []
    for g, d in b.groupby("grup"):
        a = d.loc[d.kalem == "TOPLAM AKTİFLER", "deger"]
        p = d.loc[d.kalem == "TOPLAM PASİFLER", "deger"]
        if a.empty or p.empty or abs(a.iloc[0] - p.iloc[0]) > 0.01:
            sorun.append(g)
    if sorun:
        print(f"  UYARI {donem}: aktif/pasif tutmuyor: {sorun}", flush=True)


def ay_araligi(bas, bit):
    y, a = bas
    while (y, a) <= bit:
        yield y, a
        y, a = (y + 1, 1) if a == 12 else (y, a + 1)


def yil_csv(yil):
    return KOK / "uzun" / f"bddk_aylik_{yil}.csv"


def kaydet(yeni):
    """Yeni dönemleri yıllık CSV'lere yazar (aynı dönem varsa üzerine yazar)."""
    yeni = yeni.copy()
    yeni["yil"] = yeni.donem.str[:4]
    for yil, d in yeni.groupby("yil"):
        yol = yil_csv(yil)
        yol.parent.mkdir(parents=True, exist_ok=True)
        d = d[UZUN_SUTUNLAR]
        if yol.exists():
            eski = pd.read_csv(yol, encoding="utf-8-sig", dtype={"donem": str})
            d = pd.concat([eski[~eski.donem.isin(d.donem.unique())], d])
        d = d.assign(_g=d.grup.map({g: i for i, g in enumerate(GRUP_SIRA)}).fillna(99))
        d = d.sort_values(["donem", "tablo_no", "_g", "sira"], kind="stable").drop(columns="_g")
        d.to_csv(yol, index=False, encoding="utf-8-sig", float_format="%.6f")
    return sorted(yeni.yil.unique())


def excel_yaz(yil, bilgi):
    df = pd.read_csv(yil_csv(yil), encoding="utf-8-sig", dtype={"donem": str})
    wb = Workbook()
    ws = wb.active
    ws.title = "Bilgi"
    donemler = sorted(df.donem.unique())
    satirlar = [
        ("BDDK Aylık Bülten Arşivi", str(yil)),
        ("Kaynak", "BDDK İnteraktif Aylık Bülten – bddk.org.tr/BultenAylik (Excel dışa aktarımı)"),
        ("Para birimi", "TL. YP sütunları yabancı para bakiyelerin dönem sonu kurla TL karşılığıdır."),
        ("Kapsam", "Yurt dışı şubeler dâhil, toplulaştırılmış veri."),
        ("Dönemler", f"{donemler[0]} – {donemler[-1]} ({len(donemler)} dönem)"),
        ("Gruplar", ", ".join(GRUP_SIRA)),
        ("Güncelleme", datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")),
        ("Not", "BDDK geçmiş dönemleri revize edebilir. Değerler nominaldir; enflasyon ve kur etkisi arındırılmamıştır."),
        ("", ""),
        ("Tablo", "Birim"),
    ]
    satirlar += [(f"{t} {b['tablo']}", b.get("birim") or "") for t, b in sorted(bilgi.items(), key=lambda x: int(x[0]))]
    for r in satirlar:
        ws.append(r)
    ws["A1"].font = Font(bold=True, size=13)
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 100
    for t in TABLOLAR:
        d = df[df.tablo_no == t]
        if d.empty:
            continue
        sira = list(dict.fromkeys(bilgi.get(str(t), bilgi.get(t, {})).get("sutunlar", []) + list(d.sutun.unique())))
        p = d.pivot_table(index=["donem", "grup", "sira", "kalem"], columns="sutun", values="deger",
                          aggfunc="first", sort=False).reset_index()
        p = p[["donem", "grup", "sira", "kalem"] + [s for s in sira if s in p.columns]]
        p["_g"] = p.grup.map({g: i for i, g in enumerate(GRUP_SIRA)}).fillna(99)
        p = p.sort_values(["donem", "_g", "sira"], kind="stable").drop(columns="_g")
        s = wb.create_sheet(re.sub(r"[\\/*?:\[\]]", "", TABLOLAR[t])[:31])
        s.append(["Dönem", "Grup", "Sıra", "Kalem"] + list(p.columns[4:]))
        for c in s[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F3864")
            c.alignment = Alignment(horizontal="center", wrap_text=True)
        for row in p.itertuples(index=False):
            s.append([None if (isinstance(v, float) and pd.isna(v)) else v for v in row])
        for col, w in zip("ABCD", [9, 20, 6, 60]):
            s.column_dimensions[col].width = w
        for j in range(5, s.max_column + 1):
            s.column_dimensions[s.cell(1, j).column_letter].width = 16
        for r in s.iter_rows(min_row=2, min_col=5):
            for c in r:
                c.number_format = "#,##0.000"
        s.freeze_panes = "E2"
        s.auto_filter.ref = s.dimensions
    yol = KOK / "excel" / f"BDDK_Aylik_Arsiv_{yil}.xlsx"
    yol.parent.mkdir(parents=True, exist_ok=True)
    wb.save(yol)


def calistir(donemler, paralel):
    bilgi_yol = KOK / "tablolar.json"
    bilgi = json.loads(bilgi_yol.read_text("utf-8")) if bilgi_yol.exists() else {}
    parcalar = []
    with ThreadPoolExecutor(paralel) as ex:
        isler = {ex.submit(donem_cek, y, a): (y, a) for y, a in donemler}
        for f in as_completed(isler):
            y, a = isler[f]
            df, b = f.result()
            kontrol(df, f"{y}-{a:02d}")
            parcalar.append(df)
            if (y, a) == max(donemler):
                bilgi = {str(k): v for k, v in b.items()}
            print(f"  {y}-{a:02d}: {len(df):,} satır", flush=True)
    yillar = kaydet(pd.concat(parcalar))
    KOK.mkdir(parents=True, exist_ok=True)
    bilgi_yol.write_text(json.dumps(bilgi, ensure_ascii=False, indent=2), "utf-8")
    for yil in yillar:
        excel_yaz(yil, bilgi)
        print(f"  Excel yazıldı: {yil}", flush=True)
    son = max(d.name.split("_")[-1][:4] for d in (KOK / "uzun").glob("*.csv"))
    son_donem = pd.read_csv(yil_csv(son), usecols=["donem"], dtype=str).donem.max()
    (KOK / "son_donem.txt").write_text(son_donem + "\n", "utf-8")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--baslangic", help="YYYY-MM")
    ap.add_argument("--bitis", help="YYYY-MM")
    ap.add_argument("--guncelle", action="store_true", help="Yeni dönem varsa ekle ve son 12 ayı yenile")
    ap.add_argument("--paralel", type=int, default=3)
    a = ap.parse_args()
    ym = lambda s: tuple(int(x) for x in s.split("-"))

    if a.guncelle:
        yol = KOK / "son_donem.txt"
        son = ym(yol.read_text().strip()) if yol.exists() else (2017, 12)
        aday = (son[0] + 1, 1) if son[1] == 12 else (son[0], son[1] + 1)
        bugun = date.today()
        yeni = []
        while aday < (bugun.year, bugun.month) and yayinlandi_mi(*aday):
            yeni.append(aday)
            aday = (aday[0] + 1, 1) if aday[1] == 12 else (aday[0], aday[1] + 1)
        if not yeni:
            print(f"Yeni dönem yok (arşivde son: {son[0]}-{son[1]:02d}).")
            return
        bit = yeni[-1]
        bas_i = bit[0] * 12 + bit[1] - 1 - 11
        donemler = list(ay_araligi((bas_i // 12, bas_i % 12 + 1), bit))
        print(f"Yeni dönem(ler): {[f'{y}-{m:02d}' for y, m in yeni]}; {len(donemler)} ay yenileniyor.")
    else:
        if not (a.baslangic and a.bitis):
            sys.exit("--baslangic ve --bitis ya da --guncelle gerekli")
        donemler = list(ay_araligi(ym(a.baslangic), ym(a.bitis)))
        print(f"{len(donemler)} dönem çekiliyor: {a.baslangic} – {a.bitis}")
    calistir(donemler, a.paralel)


if __name__ == "__main__":
    main()
