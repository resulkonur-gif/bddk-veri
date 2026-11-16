"""BDDK Haftalık Bülten verilerini Gelişmiş Gösterim üzerinden çeker ve veri/haftalik altında arşivler.

Kaynak: https://www.bddk.org.tr/BultenHaftalik/tr/Gelismis
Gelişmiş Gösterim formu (GelismisRaporGetir) seçilen tarih aralığı, kalemler ve taraflar için
tam hassasiyetli (5 ondalık) bir tablo döndürür. Her yıl × tablo için bir istek atılır.

Kullanım:
  python scripts/bddk_haftalik.py --baslangic 2014 --bitis 2026   # yıl aralığı
  python scripts/bddk_haftalik.py --guncelle                       # yeni hafta varsa son 13 haftayı yenile

Çıktılar:
  veri/haftalik/uzun/bddk_haftalik_<yıl>.csv  tarih,tablo_no,tablo,grup,kalem_id,sira,kalem,sutun,deger
  veri/haftalik/excel/BDDK_Haftalik_Arsiv_<yıl>.xlsx
  veri/haftalik/kalemler.json                 tablolar ve kalemleri (id, sıra, ad)
  veri/haftalik/son_tarih.txt                 arşivdeki en son hafta
"""
import argparse
import html
import json
import re
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

import pandas as pd
import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

TEMEL = "https://www.bddk.org.tr/BultenHaftalik"
TARAFLAR = ["10001", "10002", "10008", "10010", "10009", "10004", "10003", "10005", "10007", "10006"]
GRUP_SIRA = ["Sektör", "Mevduat", "Mevduat-Kamu", "Mevduat-Yerli Özel", "Mevduat-Yabancı",
             "Katılım", "Kalkınma ve Yatırım", "Kamu", "Yerli Özel", "Yabancı"]
SUTUNLAR = ["TP", "YP", "Toplam"]
KOK = Path(__file__).resolve().parent.parent / "veri" / "haftalik"
UZUN = ["tarih", "tablo_no", "tablo", "grup", "kalem_id", "sira", "kalem", "sutun", "deger"]
UA = "Mozilla/5.0 (bddk-veri arsiv; github.com/resulkonur-gif/bddk-veri)"


def temiz(s):
    return re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", s))).strip()


def grup_adi(s):
    s = temiz(s)
    s = re.sub(r"\s*-\s*", "-", s) if s.lower().startswith("mevduat") else s
    return s


class Oturum:
    """Kendi çerezi ve doğrulama belirteci olan bir BDDK oturumu."""

    def __init__(self):
        self.s = requests.Session()
        self.s.headers["User-Agent"] = UA
        self.yenile()

    def yenile(self):
        r = self.s.get(f"{TEMEL}/tr/Gelismis", timeout=120)
        r.raise_for_status()
        self.sayfa = r.text
        f = self.sayfa[self.sayfa.find("GelismisRaporGetir"):]
        self.belirtec = re.search(r'name="__RequestVerificationToken" type="hidden" value="([^"]+)"', f).group(1)

    def tablolar(self):
        """{tablo_no: {'bddk_id', 'ad', 'kalemler': [(id, ad), ...]}}"""
        s = self.sayfa
        ids = re.findall(r'id="Kalemler-(\d+)"', s)
        out = {}
        for n, tid in enumerate(ids, 1):
            i = s.find(f'id="Kalemler-{tid}"')
            j = s.find('id="Kalemler-', i + 10)
            blok = s[i:j if j > 0 else len(s)]
            ad = temiz(re.search(r"<h4>(.*?) Tablosuna Ait Kalemler</h4>", blok).group(1))
            kalemler = [(int(k), temiz(a)) for k, a in re.findall(r'id="Kalem-(\d+)"\s+onclick="KalemToggle\(\d+,\s*\'(.*?)\'\)"', blok)]
            out[n] = {"bddk_id": int(tid), "ad": ad, "kalemler": kalemler}
        return out

    def tarihler(self):
        """Bülten haftaları (eskiden yeniye) -> 'g.aa.yyyy 00:00:00' biçiminde form değerleri."""
        i = self.sayfa.find('name="BaslangicTarihi"')
        j = self.sayfa.find("</select>", i)
        vals = re.findall(r'<option value="(\d{1,2}\.\d{2}\.\d{4} 00:00:00)"', self.sayfa[i:j])
        return sorted(set(vals), key=tarih_cevir)

    def rapor(self, bas, bit, kalem_ids, deneme=4):
        veri = [("__RequestVerificationToken", self.belirtec), ("BaslangicTarihi", bas), ("BitisTarihi", bit),
                ("dil", "tr"), ("SeciliParalar", "TL")]
        veri += [("Kalemler", str(k)) for k in kalem_ids]
        veri += [("Taraflar", t) for t in TARAFLAR]
        veri += [("kalemSutun", c) for c in SUTUNLAR]
        for k in range(deneme):
            try:
                r = self.s.post(f"{TEMEL}/tr/Gelismis/GelismisRaporGetir", data=veri, timeout=300)
                if r.status_code == 200 and 'id="TabloExcelGelismis2"' in r.text:
                    return r.text
                hata = f"HTTP {r.status_code}, {len(r.content)} bayt"
            except requests.RequestException as e:
                hata = str(e)
            time.sleep(3 * (k + 1))
            self.yenile()
        raise RuntimeError(f"{bas}–{bit}: {hata}")


def tarih_cevir(s):
    return datetime.strptime(s.split()[0], "%d.%m.%Y").date()


def sayi(s):
    s = s.strip()
    if not s or s in ("-", "&nbsp;"):
        return None
    try:
        return float(s.replace(".", "").replace(",", "."))
    except ValueError:
        return None


def hucreler(satir):
    out = []
    for m in re.finditer(r"<t([hd])([^>]*)>(.*?)</t[hd]>", satir, re.S):
        cs = re.search(r'colspan="?(\d+)', m.group(2))
        out.append((temiz(m.group(3)), int(cs.group(1)) if cs else 1))
    return out


def ayristir(sayfa, tablo_no, tablo_ad, kalemler):
    """TabloExcelGelismis2 (tam hassasiyet) -> kayıt listesi."""
    m = re.search(r'<table[^>]*id="TabloExcelGelismis2".*?</table>', sayfa, re.S)
    satirlar = re.findall(r"<tr.*?</tr>", m.group(0), re.S)
    ust = hucreler(satirlar[0])
    alt = [h for h, _ in hucreler(satirlar[1])]
    # başlık: ilk hücre "Birim: ...", sonra her kalem için colspan kadar sütun (TP / YP / TOPLAM)
    gruplar = ust[1:]
    adlar = [a.split(" / ", 1)[-1] for a, _ in gruplar]
    beklenen = [a for _, a in kalemler]
    if len(adlar) != len(beklenen) or any(re.sub(r"\s+", "", x) != re.sub(r"\s+", "", y) for x, y in zip(adlar, beklenen)):
        raise RuntimeError(f"tablo {tablo_no}: başlıktaki kalemler beklenenle uyuşmuyor\n{adlar[:6]}\n{beklenen[:6]}")
    kolon_kalem = []
    for sira, ((_, cs), (kid, kad)) in enumerate(zip(gruplar, kalemler), 1):
        kolon_kalem += [(kid, sira, kad)] * cs
    alt = alt[len(alt) - len(kolon_kalem):]
    kolon_sutun = [("Toplam" if s.upper().startswith("TOPLAM") else s.upper()[:2]) for s in alt]
    kayit, grup = [], None
    for st in satirlar[2:]:
        h = hucreler(st)
        if not h:
            continue
        if len(h) == 1 or (len(h) <= 2 and not re.match(r"\d{1,2}\.\d{1,2}\.\d{4}$", h[0][0])):
            if h[0][0]:
                grup = grup_adi(h[0][0])
            continue
        t = h[0][0]
        if not re.match(r"\d{1,2}\.\d{1,2}\.\d{4}$", t):
            continue
        tarih = datetime.strptime(t, "%d.%m.%Y").date().isoformat()
        for (deg, _), (kid, sira, kad), sut in zip(h[1:], kolon_kalem, kolon_sutun):
            kayit.append((tarih, tablo_no, tablo_ad, grup, kid, sira, kad, sut, sayi(deg)))
    return kayit


def yil_parcalari(tarihler, yillar):
    out = []
    for y in yillar:
        t = [x for x in tarihler if tarih_cevir(x).year == y]
        if t:
            out.append((y, t[0], t[-1]))
    return out


def cek(parcalar, tablolar, paralel):
    yerel = threading.local()
    isler = [(y, b, e, n) for y, b, e in parcalar for n in tablolar]
    sonuc = []

    def is_yap(y, b, e, n):
        if not hasattr(yerel, "o"):
            yerel.o = Oturum()
        t = tablolar[n]
        for k in range(3):
            s = yerel.o.rapor(b, e, [i for i, _ in t["kalemler"]])
            try:
                return y, n, ayristir(s, n, t["ad"], t["kalemler"])
            except RuntimeError as hata:
                print(f"  {y} tablo {n}: deneme {k + 1} başarısız ({str(hata).splitlines()[0]})", flush=True)
                if k == 2:
                    m = re.search(r'<table[^>]*id="TabloExcelGelismis2".*?</table>', s, re.S)
                    print((m.group(0) if m else s)[:3000], flush=True)
                    raise
                time.sleep(5)
                yerel.o = Oturum()

    with ThreadPoolExecutor(paralel) as ex:
        fs = [ex.submit(is_yap, *i) for i in isler]
        for f in as_completed(fs):
            y, n, kayit = f.result()
            sonuc += kayit
            print(f"  {y} tablo {n}: {len(kayit):,} değer", flush=True)
    return pd.DataFrame(sonuc, columns=UZUN)


def kontrol(df):
    sorun = sorted(set(df.grup.dropna()) - set(GRUP_SIRA))
    if sorun:
        print(f"  UYARI: beklenmeyen grup adları: {sorun}", flush=True)
    d = df[(df.tablo_no == 1) & (df.kalem_id == df[df.tablo_no == 1].kalem_id.min())]
    p = d.pivot_table(index=["tarih", "grup"], columns="sutun", values="deger", aggfunc="first")
    if {"TP", "YP", "Toplam"} <= set(p.columns):
        fark = (p.TP + p.YP - p.Toplam).abs().max()
        print(f"  Kontrol: Toplam Krediler TP+YP-Toplam en büyük fark {fark:.4f}", flush=True)


def kaydet(yeni):
    yeni = yeni.copy()
    yeni["yil"] = yeni.tarih.str[:4]
    gs = {g: i for i, g in enumerate(GRUP_SIRA)}
    for yil, d in yeni.groupby("yil"):
        yol = KOK / "uzun" / f"bddk_haftalik_{yil}.csv"
        yol.parent.mkdir(parents=True, exist_ok=True)
        d = d[UZUN]
        if yol.exists():
            eski = pd.read_csv(yol, encoding="utf-8-sig", dtype={"tarih": str})
            anahtar = set(zip(d.tarih, d.tablo_no))
            eski = eski[[(t, n) not in anahtar for t, n in zip(eski.tarih, eski.tablo_no)]]
            d = pd.concat([eski, d])
        d = d.assign(_g=d.grup.map(gs).fillna(99), _s=d.sutun.map({s: i for i, s in enumerate(SUTUNLAR)}))
        d = d.sort_values(["tarih", "tablo_no", "_g", "sira", "_s"], kind="stable").drop(columns=["_g", "_s"])
        d.to_csv(yol, index=False, encoding="utf-8-sig", float_format="%.5f")
    return sorted(yeni.yil.unique())


def excel_yaz(yil, tablolar):
    df = pd.read_csv(KOK / "uzun" / f"bddk_haftalik_{yil}.csv", encoding="utf-8-sig", dtype={"tarih": str})
    wb = Workbook()
    ws = wb.active
    ws.title = "Bilgi"
    tarihler = sorted(df.tarih.unique())
    for r in [("BDDK Haftalık Bülten Arşivi", str(yil)),
              ("Kaynak", "BDDK Haftalık Bülten – bddk.org.tr/BultenHaftalik (Gelişmiş Gösterim, tam hassasiyet)"),
              ("Birim", "Milyon TL. YP sütunu yabancı para bakiyelerin TL karşılığıdır."),
              ("Haftalar", f"{tarihler[0]} – {tarihler[-1]} ({len(tarihler)} hafta)"),
              ("Gruplar", ", ".join(GRUP_SIRA)),
              ("Güncelleme", datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")),
              ("Not", "Haftalık veriler geçicidir; BDDK sonraki haftalarda revize edebilir. Değerler nominaldir.")]:
        ws.append(r)
    ws["A1"].font = Font(bold=True, size=13)
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 100
    gs = {g: i for i, g in enumerate(GRUP_SIRA)}
    for n, t in sorted(tablolar.items(), key=lambda x: int(x[0])):
        d = df[df.tablo_no == int(n)]
        if d.empty:
            continue
        p = d.pivot_table(index=["tarih", "grup", "sira", "kalem"], columns="sutun", values="deger",
                          aggfunc="first", sort=False).reset_index()
        p = p[["tarih", "grup", "sira", "kalem"] + [s for s in SUTUNLAR if s in p.columns]]
        p = p.assign(_g=p.grup.map(gs).fillna(99)).sort_values(["tarih", "_g", "sira"]).drop(columns="_g")
        s = wb.create_sheet(re.sub(r"[\\/*?:\[\]]", "", t["ad"])[:31])
        s.append(["Tarih", "Grup", "Sıra", "Kalem"] + list(p.columns[4:]))
        for c in s[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F3864")
            c.alignment = Alignment(horizontal="center")
        for row in p.itertuples(index=False):
            s.append([None if (isinstance(v, float) and pd.isna(v)) else v for v in row])
        for col, w in zip("ABCDEFG", [11, 20, 6, 60, 16, 16, 16]):
            s.column_dimensions[col].width = w
        for r in s.iter_rows(min_row=2, min_col=5):
            for c in r:
                c.number_format = "#,##0.000"
        s.freeze_panes = "E2"
        s.auto_filter.ref = s.dimensions
    yol = KOK / "excel" / f"BDDK_Haftalik_Arsiv_{yil}.xlsx"
    yol.parent.mkdir(parents=True, exist_ok=True)
    wb.save(yol)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--baslangic", type=int, help="başlangıç yılı")
    ap.add_argument("--bitis", type=int, help="bitiş yılı")
    ap.add_argument("--guncelle", action="store_true")
    ap.add_argument("--paralel", type=int, default=3)
    a = ap.parse_args()

    o = Oturum()
    tablolar = o.tablolar()
    tarihler = o.tarihler()
    if len(tablolar) != 9 or sum(len(t["kalemler"]) for t in tablolar.values()) < 150:
        sys.exit(f"Beklenmeyen sayfa yapısı: {len(tablolar)} tablo")
    print(f"{len(tablolar)} tablo, {sum(len(t['kalemler']) for t in tablolar.values())} kalem, "
          f"{len(tarihler)} hafta ({tarih_cevir(tarihler[0])} – {tarih_cevir(tarihler[-1])})", flush=True)

    if a.guncelle:
        yol = KOK / "son_tarih.txt"
        son = date.fromisoformat(yol.read_text().strip()) if yol.exists() else date(2013, 12, 31)
        en_yeni = tarih_cevir(tarihler[-1])
        if en_yeni <= son:
            print(f"Yeni hafta yok (arşivde son: {son}).")
            return
        bas = en_yeni - timedelta(weeks=13)
        sec = [t for t in tarihler if tarih_cevir(t) > bas]
        parcalar = [(y, *[[t for t in sec if tarih_cevir(t).year == y][i] for i in (0, -1)])
                    for y in sorted({tarih_cevir(t).year for t in sec})]
        print(f"Yeni hafta: {en_yeni}; {len(sec)} hafta yenileniyor.", flush=True)
    else:
        if not (a.baslangic and a.bitis):
            sys.exit("--baslangic ve --bitis ya da --guncelle gerekli")
        parcalar = yil_parcalari(tarihler, range(a.baslangic, a.bitis + 1))

    df = cek(parcalar, tablolar, a.paralel)
    kontrol(df)
    yillar = kaydet(df)
    KOK.mkdir(parents=True, exist_ok=True)
    meta = {str(n): {"bddk_id": t["bddk_id"], "ad": t["ad"],
                     "kalemler": [{"id": i, "sira": s + 1, "ad": k} for s, (i, k) in enumerate(t["kalemler"])]}
            for n, t in tablolar.items()}
    (KOK / "kalemler.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2), "utf-8")
    for yil in yillar:
        excel_yaz(yil, meta)
        print(f"  Excel yazıldı: {yil}", flush=True)
    son = max(pd.read_csv(f, usecols=["tarih"], dtype=str).tarih.max() for f in (KOK / "uzun").glob("*.csv"))
    (KOK / "son_tarih.txt").write_text(son + "\n", "utf-8")
    print(f"Arşivde son hafta: {son}")


if __name__ == "__main__":
    main()
