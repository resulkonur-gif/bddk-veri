#!/usr/bin/env python3
"""BDDK Aylık ve Haftalık Bülten verilerini indirir; CSV ve Excel olarak kaydeder.

Kurulum (bir kez):
    pip install requests pandas openpyxl

Kullanım:
    python bddk_indir.py                                   # soru sorarak ilerler
    python bddk_indir.py aylik                             # yayımlanmış en son ay
    python bddk_indir.py aylik --baslangic 2018-01 --bitis 2026-09
    python bddk_indir.py aylik --baslangic 2025-01 --bitis 2025-12 --tablolar 1,2,15 --para USD
    python bddk_indir.py haftalik                          # içinde bulunulan yıl
    python bddk_indir.py haftalik --baslangic 2014 --bitis 2026
    python bddk_indir.py haftalik --baslangic 2026 --bitis 2026 --tablolar 1,4

Çıktılar (betiğin yanında, --klasor ile değiştirilebilir):
    BDDK_Veri/aylik/bddk_aylik_<yıl>.csv       uzun biçim: donem,tablo_no,tablo,grup,sira,kalem,sutun,deger
    BDDK_Veri/aylik/BDDK_Aylik_<yıl>.xlsx      her tablo ayrı sayfada
    BDDK_Veri/haftalik/bddk_haftalik_<yıl>.csv tarih,tablo_no,tablo,grup,sira,kalem,TP,YP,Toplam
    BDDK_Veri/haftalik/BDDK_Haftalik_<yıl>.xlsx
Aynı dönem yeniden indirilirse eski kayıtların üzerine yazılır; diğer dönemler korunur.

Kaynak: bddk.org.tr/BultenAylik (Excel dışa aktarımı) ve bddk.org.tr/BultenHaftalik (Gelişmiş Gösterim).
Değerler BDDK'nın yayımladığı haliyle, nominaldir. BDDK geçmiş dönemleri revize edebilir.
"""
import argparse
import html
import io
import re
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime
from pathlib import Path

try:
    import pandas as pd
    import requests
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
except ImportError:
    print("Eksik paket var. Önce şunu çalıştırın:\n    pip install requests pandas openpyxl")
    input("Çıkmak için Enter...") if len(sys.argv) == 1 else None
    sys.exit(1)

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

UA = "Mozilla/5.0 (bddk_indir.py)"
GRUP_SIRA = ["Sektör", "Mevduat", "Mevduat-Kamu", "Mevduat-Yerli Özel", "Mevduat-Yabancı",
             "Katılım", "Kalkınma ve Yatırım", "Kamu", "Yerli Özel", "Yabancı"]
GS = {g: i for i, g in enumerate(GRUP_SIRA)}


def baslik_bicim(ws):
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F3864")
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.freeze_panes = "E2"
    ws.auto_filter.ref = ws.dimensions


def sayfa_adi(ad):
    ad = re.sub(r"[\\/*?:\[\]]", "", ad).replace("Bankalarda Saklanan ", "Saklanan ")
    return ad[:31]


# =====================================================================================
# AYLIK BÜLTEN
# =====================================================================================
AYLIK_URL = "https://www.bddk.org.tr/BultenAylik/tr/Home/BasitExceleAktar"
AYLIK_TABLOLAR = {
    1: "Bilanço", 2: "Kar Zarar", 3: "Krediler", 4: "Tüketici Kredileri",
    5: "Sektörel Kredi Dağılımı", 6: "KOBİ Kredileri", 7: "Sendikasyon Seküritizasyon Kredileri",
    8: "Menkul Kıymetler", 9: "Mevduat Türler İtibarıyla", 10: "Mevduat Vade İtibarıyla",
    11: "Likidite Durumu", 12: "Sermaye Yeterliliği", 13: "Yabancı Para Pozisyonu",
    14: "Bilanço Dışı İşlemler", 15: "Rasyolar", 16: "Diğer Bilgiler", 17: "Yurt Dışı Şube Rasyoları",
}
AYLIK_TARAF = [10001, 10002, 10008, 10009, 10010, 10003, 10004, 10005, 10006, 10007]
AYLIK_SUTUN = ["donem", "tablo_no", "tablo", "grup", "sira", "kalem", "sutun", "deger"]


def _sayi(x):
    if x is None or x == "":
        return None
    try:
        return float(str(x).replace(",", "."))
    except ValueError:
        return None


def aylik_indir(oturum, tablo_no, yil, ay, para="TL", deneme=4):
    params = [("tabloNo", tablo_no), ("yil", yil), ("ay", ay), ("paraBirimi", para)]
    params += [(f"taraf[{i}]", g) for i, g in enumerate(AYLIK_TARAF)]
    for k in range(deneme):
        try:
            r = oturum.get(AYLIK_URL, params=params, timeout=90)
            if r.status_code == 200 and r.content[:2] == b"PK":
                ws = load_workbook(io.BytesIO(r.content), read_only=True).active
                satirlar = [list(x) for x in ws.iter_rows(values_only=True)]
                return satirlar[0], [x for x in satirlar[1:] if x and x[0]]
            hata = f"HTTP {r.status_code}"
        except requests.RequestException as e:
            hata = str(e)
        time.sleep(2 * (k + 1))
    raise RuntimeError(f"Aylık tablo {tablo_no} {yil}/{ay}: {hata}")


def aylik_yayinlandi_mi(oturum, yil, ay):
    try:
        baslik, satirlar = aylik_indir(oturum, 1, yil, ay, deneme=2)
    except RuntimeError:
        return False
    return f"Dönem:{yil}/{ay}" in str(baslik[2]) and len(satirlar) > 0


def aylik_son_donem(oturum):
    bugun = date.today()
    y, a = bugun.year, bugun.month
    for _ in range(4):
        y, a = (y - 1, 12) if a == 1 else (y, a - 1)
        if aylik_yayinlandi_mi(oturum, y, a):
            return y, a
    raise RuntimeError("Son aylık dönem bulunamadı.")


def aylik_donem(oturum, yil, ay, tablolar, para):
    kayit = []
    for t in tablolar:
        baslik, satirlar = aylik_indir(oturum, t, yil, ay, para)
        sutunlar = [str(c) for c in baslik[3:] if c not in (None, "")]
        for r in satirlar:
            for j, s in enumerate(sutunlar):
                kayit.append((f"{yil}-{ay:02d}", t, AYLIK_TABLOLAR[t], str(r[0]).strip(),
                              int(float(r[1])), str(r[2]).strip(), s, _sayi(r[3 + j])))
    return pd.DataFrame(kayit, columns=AYLIK_SUTUN)


def aylik_kaydet(df, klasor, para):
    klasor.mkdir(parents=True, exist_ok=True)
    ek = "" if para == "TL" else f"_{para}"
    df = df.assign(yil=df.donem.str[:4])
    for yil, d in df.groupby("yil"):
        yol = klasor / f"bddk_aylik_{yil}{ek}.csv"
        d = d[AYLIK_SUTUN]
        if yol.exists():
            eski = pd.read_csv(yol, encoding="utf-8-sig", dtype={"donem": str})
            anahtar = set(zip(d.donem, d.tablo_no))
            eski = eski[[(p, t) not in anahtar for p, t in zip(eski.donem, eski.tablo_no)]]
            d = pd.concat([eski, d])
        d = (d.assign(_g=d.grup.map(GS).fillna(99))
              .sort_values(["donem", "tablo_no", "_g", "sira"], kind="stable").drop(columns="_g"))
        d.to_csv(yol, index=False, encoding="utf-8-sig")
        # Excel
        wb = Workbook()
        wb.remove(wb.active)
        for t, dt in d.groupby("tablo_no"):
            sira = list(dict.fromkeys(dt.sutun))
            p = dt.pivot_table(index=["donem", "grup", "sira", "kalem"], columns="sutun", values="deger",
                               aggfunc="first", sort=False).reset_index()
            p = p[["donem", "grup", "sira", "kalem"] + [s for s in sira if s in p.columns]]
            p = p.assign(_g=p.grup.map(GS).fillna(99)).sort_values(["donem", "_g", "sira"]).drop(columns="_g")
            ws = wb.create_sheet(sayfa_adi(AYLIK_TABLOLAR[int(t)]))
            ws.append(["Dönem", "Grup", "Sıra", "Kalem"] + list(p.columns[4:]))
            for row in p.itertuples(index=False):
                ws.append([None if (isinstance(v, float) and pd.isna(v)) else v for v in row])
            for col, w in zip("ABCD", [9, 20, 6, 60]):
                ws.column_dimensions[col].width = w
            baslik_bicim(ws)
        wb.save(klasor / f"BDDK_Aylik_{yil}{ek}.xlsx")
        print(f"  Kaydedildi: {yol.name} ve BDDK_Aylik_{yil}{ek}.xlsx")


def aylik_calistir(baslangic, bitis, tablolar, para, klasor, paralel):
    oturum = requests.Session()
    oturum.headers["User-Agent"] = UA
    if not bitis:
        bitis = aylik_son_donem(oturum)
        print(f"Yayımlanmış en son dönem: {bitis[0]}-{bitis[1]:02d}")
    baslangic = baslangic or bitis
    donemler = []
    y, a = baslangic
    while (y, a) <= bitis:
        donemler.append((y, a))
        y, a = (y + 1, 1) if a == 12 else (y, a + 1)
    print(f"Aylık bülten: {len(donemler)} dönem × {len(tablolar)} tablo × 10 grup indiriliyor ({para})...")
    parcalar = []
    with ThreadPoolExecutor(paralel) as ex:
        fs = {ex.submit(aylik_donem, oturum, y, a, tablolar, para): (y, a) for y, a in donemler}
        for f in as_completed(fs):
            y, a = fs[f]
            try:
                parcalar.append(f.result())
                print(f"  {y}-{a:02d} tamam")
            except RuntimeError as e:
                print(f"  {y}-{a:02d} ATLANDI: {e}")
    if parcalar:
        aylik_kaydet(pd.concat(parcalar), klasor / "aylik", para)


# =====================================================================================
# HAFTALIK BÜLTEN
# =====================================================================================
HAFTALIK_URL = "https://www.bddk.org.tr/BultenHaftalik"
HAFTALIK_TARAF = ["10001", "10002", "10008", "10010", "10009", "10004", "10003", "10005", "10007", "10006"]
HAFTALIK_SUTUN = ["TP", "YP", "Toplam"]
YALNIZ_TOPLAM = {9}  # Yabancı Para Pozisyonu: BDDK yalnız Toplam sütununu dolduruyor


def _temiz(s):
    return re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", s))).strip()


def _tarih(s):
    return datetime.strptime(s.split()[0], "%d.%m.%Y").date()


class HaftalikOturum:
    def __init__(self):
        self.s = requests.Session()
        self.s.headers["User-Agent"] = UA
        self.yenile()

    def yenile(self):
        r = self.s.get(f"{HAFTALIK_URL}/tr/Gelismis", timeout=120)
        r.raise_for_status()
        self.sayfa = r.text
        f = self.sayfa[self.sayfa.find("GelismisRaporGetir"):]
        self.belirtec = re.search(r'name="__RequestVerificationToken" type="hidden" value="([^"]+)"', f).group(1)

    def tablolar(self):
        s, out = self.sayfa, {}
        for n, tid in enumerate(re.findall(r'id="Kalemler-(\d+)"', s), 1):
            i = s.find(f'id="Kalemler-{tid}"')
            j = s.find('id="Kalemler-', i + 10)
            blok = s[i:j if j > 0 else len(s)]
            ad = _temiz(re.search(r"<h4>(.*?) Tablosuna Ait Kalemler</h4>", blok).group(1))
            kalemler = [(int(k), _temiz(a)) for k, a in
                        re.findall(r'id="Kalem-(\d+)"\s+onclick="KalemToggle\(\d+,\s*\'(.*?)\'\)"', blok)]
            out[n] = {"ad": ad, "kalemler": kalemler}
        return out

    def tarihler(self):
        i = self.sayfa.find('name="BaslangicTarihi"')
        j = self.sayfa.find("</select>", i)
        vals = re.findall(r'<option value="(\d{1,2}\.\d{2}\.\d{4} 00:00:00)"', self.sayfa[i:j])
        return sorted(set(vals), key=_tarih)

    def rapor(self, bas, bit, kalem_ids, deneme=4):
        veri = [("__RequestVerificationToken", self.belirtec), ("BaslangicTarihi", bas), ("BitisTarihi", bit),
                ("dil", "tr"), ("SeciliParalar", "TL")]
        veri += [("Kalemler", str(k)) for k in kalem_ids]
        veri += [("Taraflar", t) for t in HAFTALIK_TARAF]
        veri += [("kalemSutun", c) for c in HAFTALIK_SUTUN]
        for k in range(deneme):
            try:
                r = self.s.post(f"{HAFTALIK_URL}/tr/Gelismis/GelismisRaporGetir", data=veri, timeout=300)
                if r.status_code == 200 and 'id="TabloExcelGelismis2"' in r.text:
                    return r.text
                hata = f"HTTP {r.status_code}"
            except requests.RequestException as e:
                hata = str(e)
            time.sleep(3 * (k + 1))
            self.yenile()
        raise RuntimeError(f"Haftalık {bas}–{bit}: {hata}")


def _hucreler(satir):
    out = []
    for m in re.finditer(r"<t([hd])([^>]*)>(.*?)</t[hd]>", satir, re.S):
        cs = re.search(r'colspan="?(\d+)', m.group(2))
        out.append((_temiz(m.group(3)), int(cs.group(1)) if cs else 1))
    return out


def haftalik_ayristir(sayfa, tablo_no, tablo_ad, kalemler):
    m = re.search(r'<table[^>]*id="TabloExcelGelismis2".*?</table>', sayfa, re.S)
    satirlar = [r for r in re.findall(r"<tr.*?</tr>", m.group(0), re.S) if _hucreler(r)]
    ust = _hucreler(satirlar[0])
    alt = [h for h, _ in _hucreler(satirlar[1])]
    gruplar = []
    for ad, cs in ust[1:]:
        if ad:
            gruplar.append([ad, cs])
        elif gruplar:
            gruplar[-1][1] += cs
    adlar = [a.split(" / ", 1)[-1] for a, _ in gruplar]
    beklenen = [a for _, a in kalemler]
    if len(adlar) != len(beklenen) or any(re.sub(r"\s+", "", x) != re.sub(r"\s+", "", y)
                                          for x, y in zip(adlar, beklenen)):
        raise RuntimeError(f"tablo {tablo_no}: başlık beklenen kalemlerle uyuşmuyor")
    kolon_kalem = []
    for sira, ((_, cs), (kid, kad)) in enumerate(zip(gruplar, kalemler), 1):
        kolon_kalem += [(sira, kad)] * cs
    alt = alt[len(alt) - len(kolon_kalem):]
    kolon_sutun = [("Toplam" if s.upper().startswith("TOPLAM") else s.upper()[:2]) for s in alt]
    kayit, grup = [], None
    for st in satirlar[2:]:
        h = _hucreler(st)
        if not re.match(r"\d{1,2}\.\d{1,2}\.\d{4}$", h[0][0]):
            if len(h) <= 2 and h[0][0]:
                ad = h[0][0]
                grup = re.sub(r"\s*-\s*", "-", ad) if ad.lower().startswith("mevduat") else ad
            continue
        t = datetime.strptime(h[0][0], "%d.%m.%Y").date().isoformat()
        for (deg, _), (sira, kad), sut in zip(h[1:], kolon_kalem, kolon_sutun):
            deg = deg.strip()
            try:
                v = float(deg.replace(".", "").replace(",", ".")) if deg and deg != "-" else None
            except ValueError:
                v = None
            kayit.append((t, tablo_no, tablo_ad, grup, sira, kad, sut, v))
    uzun = pd.DataFrame(kayit, columns=["tarih", "tablo_no", "tablo", "grup", "sira", "kalem", "sutun", "deger"])
    w = uzun.pivot_table(index=["tarih", "tablo_no", "tablo", "grup", "sira", "kalem"], columns="sutun",
                         values="deger", aggfunc="first").reindex(columns=HAFTALIK_SUTUN).reset_index()
    if tablo_no in YALNIZ_TOPLAM:
        w[["TP", "YP"]] = float("nan")
    return w.dropna(subset=HAFTALIK_SUTUN, how="all")


def haftalik_kaydet(df, klasor):
    klasor.mkdir(parents=True, exist_ok=True)
    df = df.assign(yil=df.tarih.str[:4])
    kolonlar = ["tarih", "tablo_no", "tablo", "grup", "sira", "kalem"] + HAFTALIK_SUTUN
    for yil, d in df.groupby("yil"):
        yol = klasor / f"bddk_haftalik_{yil}.csv"
        d = d[kolonlar]
        if yol.exists():
            eski = pd.read_csv(yol, encoding="utf-8-sig", dtype={"tarih": str})
            anahtar = set(zip(d.tarih, d.tablo_no))
            eski = eski[[(p, t) not in anahtar for p, t in zip(eski.tarih, eski.tablo_no)]]
            d = pd.concat([eski, d])
        d = (d.assign(_g=d.grup.map(GS).fillna(99))
              .sort_values(["tarih", "tablo_no", "_g", "sira"], kind="stable").drop(columns="_g"))
        d.to_csv(yol, index=False, encoding="utf-8-sig")
        wb = Workbook()
        wb.remove(wb.active)
        for t, dt in d.groupby("tablo_no"):
            cols = ["Toplam"] if int(t) in YALNIZ_TOPLAM else HAFTALIK_SUTUN
            ws = wb.create_sheet(sayfa_adi(dt.tablo.iloc[0]))
            ws.append(["Tarih", "Grup", "Sıra", "Kalem"] + cols)
            for row in dt[["tarih", "grup", "sira", "kalem"] + cols].itertuples(index=False):
                ws.append([None if (isinstance(v, float) and pd.isna(v)) else v for v in row])
            for col, w in zip("ABCD", [11, 20, 6, 60]):
                ws.column_dimensions[col].width = w
            baslik_bicim(ws)
        wb.save(klasor / f"BDDK_Haftalik_{yil}.xlsx")
        print(f"  Kaydedildi: {yol.name} ve BDDK_Haftalik_{yil}.xlsx")


def haftalik_calistir(bas_yil, bit_yil, tablo_sec, klasor, paralel):
    print("Haftalık bülten sayfası açılıyor...")
    o = HaftalikOturum()
    tablolar = o.tablolar()
    tarihler = o.tarihler()
    print(f"  {len(tablolar)} tablo, {sum(len(t['kalemler']) for t in tablolar.values())} kalem, "
          f"{len(tarihler)} hafta ({_tarih(tarihler[0])} – {_tarih(tarihler[-1])})")
    if tablo_sec:
        tablolar = {n: t for n, t in tablolar.items() if n in tablo_sec}
    bas_yil = bas_yil or _tarih(tarihler[-1]).year
    bit_yil = bit_yil or _tarih(tarihler[-1]).year
    isler = []
    for y in range(bas_yil, bit_yil + 1):
        t = [x for x in tarihler if _tarih(x).year == y]
        if t:
            isler += [(y, t[0], t[-1], n) for n in tablolar]
    print(f"Haftalık bülten: {bas_yil}–{bit_yil}, {len(tablolar)} tablo × 10 grup ({len(isler)} istek)...")
    yerel = threading.local()

    def is_yap(y, b, e, n):
        if not hasattr(yerel, "o"):
            yerel.o = HaftalikOturum()
        t = tablolar[n]
        for k in range(3):
            s = yerel.o.rapor(b, e, [i for i, _ in t["kalemler"]])
            try:
                return y, n, haftalik_ayristir(s, n, t["ad"], t["kalemler"])
            except RuntimeError:
                if k == 2:
                    raise
                time.sleep(5)
                yerel.o = HaftalikOturum()

    parcalar = []
    with ThreadPoolExecutor(paralel) as ex:
        fs = [ex.submit(is_yap, *i) for i in isler]
        for f in as_completed(fs):
            try:
                y, n, w = f.result()
                parcalar.append(w)
                print(f"  {y} · {tablolar[n]['ad']}: {len(w):,} satır")
            except RuntimeError as e:
                print(f"  ATLANDI: {e}")
    if parcalar:
        haftalik_kaydet(pd.concat(parcalar), klasor / "haftalik")


# =====================================================================================
def ay_oku(s):
    y, a = s.split("-")
    return int(y), int(a)


def tablo_oku(s):
    return [int(x) for x in s.split(",") if x.strip()] if s else None


def menu(klasor):
    print("BDDK veri indirici\n")
    print("  1) Aylık bülten  – en son yayımlanan ay")
    print("  2) Aylık bülten  – dönem aralığı")
    print("  3) Haftalık bülten – bu yıl")
    print("  4) Haftalık bülten – yıl aralığı")
    sec = input("\nSeçiminiz (1-4): ").strip()
    if sec == "1":
        aylik_calistir(None, None, list(AYLIK_TABLOLAR), "TL", klasor, 3)
    elif sec == "2":
        b = ay_oku(input("Başlangıç (YYYY-AA, ör. 2018-01): ").strip())
        e = ay_oku(input("Bitiş     (YYYY-AA, ör. 2026-09): ").strip())
        aylik_calistir(b, e, list(AYLIK_TABLOLAR), "TL", klasor, 3)
    elif sec == "3":
        haftalik_calistir(None, None, None, klasor, 3)
    elif sec == "4":
        b = int(input("Başlangıç yılı (en erken 2014): ").strip())
        e = int(input("Bitiş yılı: ").strip())
        haftalik_calistir(b, e, None, klasor, 3)
    else:
        print("Geçersiz seçim.")


def main():
    varsayilan = Path(__file__).resolve().parent / "BDDK_Veri"
    if len(sys.argv) == 1:
        try:
            menu(varsayilan)
            print(f"\nDosyalar: {varsayilan}")
        except Exception as e:  # çift tıklayınca pencere hemen kapanmasın
            print(f"\nHATA: {e}")
        input("\nÇıkmak için Enter...")
        return

    ap = argparse.ArgumentParser(description="BDDK Aylık ve Haftalık Bülten indirici")
    ap.add_argument("--klasor", default=str(varsayilan), help="çıktı klasörü (varsayılan: ./BDDK_Veri)")
    ap.add_argument("--paralel", type=int, default=3, help="aynı anda kaç istek (varsayılan 3)")
    alt = ap.add_subparsers(dest="tur", required=True)
    a = alt.add_parser("aylik", help="Aylık Bülten (17 tablo)")
    a.add_argument("--baslangic", help="YYYY-AA (boşsa en son ay)")
    a.add_argument("--bitis", help="YYYY-AA (boşsa en son ay)")
    a.add_argument("--tablolar", help="ör. 1,2,15 (boşsa 17 tablonun hepsi)")
    a.add_argument("--para", default="TL", choices=["TL", "USD"])
    h = alt.add_parser("haftalik", help="Haftalık Bülten (9 tablo, 2014'ten itibaren)")
    h.add_argument("--baslangic", type=int, help="başlangıç yılı (boşsa bu yıl)")
    h.add_argument("--bitis", type=int, help="bitiş yılı (boşsa bu yıl)")
    h.add_argument("--tablolar", help="ör. 1,4 (boşsa 9 tablonun hepsi)")
    g = ap.parse_args()
    klasor = Path(g.klasor)
    if g.tur == "aylik":
        aylik_calistir(ay_oku(g.baslangic) if g.baslangic else None, ay_oku(g.bitis) if g.bitis else None,
                       tablo_oku(g.tablolar) or list(AYLIK_TABLOLAR), g.para, klasor, g.paralel)
    else:
        haftalik_calistir(g.baslangic, g.bitis, tablo_oku(g.tablolar), klasor, g.paralel)
    print(f"\nDosyalar: {klasor}")


if __name__ == "__main__":
    main()
