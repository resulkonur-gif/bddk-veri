#!/usr/bin/env python3
"""BDDK Aylık ve Haftalık Bülten verilerini indirir; CSV ve Excel olarak kaydeder.

Kurulum GEREKMEZ: yalnız Python'un standart kütüphanesini kullanır (Python 3.8+).

Kullanım:
    python bddk_indir.py                                   # menüyle, soru sorarak
    python bddk_indir.py aylik                             # yayımlanmış en son ay
    python bddk_indir.py aylik --baslangic 2018-01 --bitis 2026-09
    python bddk_indir.py aylik --baslangic 2025-01 --bitis 2025-12 --tablolar 1,2,15 --para USD
    python bddk_indir.py haftalik                          # içinde bulunulan yıl
    python bddk_indir.py haftalik --baslangic 2014 --bitis 2026
    python bddk_indir.py haftalik --baslangic 2026 --bitis 2026 --tablolar 1,4
    python bddk_indir.py rapor                             # indirilmiş verilerden aylık + haftalık HTML rapor
    python bddk_indir.py guncelle                          # yalnız yeni/eksik dönemleri indir, çıktıları güncelle
    python bddk_indir.py guncelle --sadece aylik

Hata veren istekler en sonda 10, 30 ve 60 sn arayla yeniden denenir. Yine de inmeyenler
BDDK_Veri/eksikler.txt dosyasına yazılır; 'guncelle' bir sonraki çalıştırmada onları da tamamlar.

Çıktılar (betiğin yanında; --klasor ile değiştirilebilir):
    BDDK_Veri/aylik/bddk_aylik_<yıl>.csv        donem,tablo_no,tablo,grup,sira,kalem,sutun,deger
    BDDK_Veri/aylik/BDDK_Aylik_<yıl>.xlsx       her tablo ayrı sayfada
    BDDK_Veri/haftalik/bddk_haftalik_<yıl>.csv  tarih,tablo_no,tablo,grup,sira,kalem,TP,YP,Toplam
    BDDK_Veri/haftalik/BDDK_Haftalik_<yıl>.xlsx
    BDDK_Veri/BDDK_Aylik_Rapor.html             etkileşimli raporlar (indirmeden sonra kendiliğinden yenilenir;
    BDDK_Veri/BDDK_Haftalik_Rapor.html          internet gerektirmez, grafik kütüphanesi içinde gömülü)
Aynı dönem yeniden indirilirse eski kayıtların üzerine yazılır; diğer dönemler korunur.

Kaynak: bddk.org.tr/BultenAylik (Excel dışa aktarımı) ve bddk.org.tr/BultenHaftalik (Gelişmiş Gösterim).
Değerler BDDK'nın yayımladığı haliyle, nominaldir; birim milyon TL (Sektörel Kredi Dağılımı bin TL,
Rasyolar %, Diğer Bilgiler adet). BDDK geçmiş dönemleri revize edebilir.
"""
import argparse
import csv
import html
import http.cookiejar
import io
import re
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime
from pathlib import Path
from xml.sax.saxutils import escape

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) bddk_indir.py"
GRUP_SIRA = ["Sektör", "Mevduat", "Mevduat-Kamu", "Mevduat-Yerli Özel", "Mevduat-Yabancı",
             "Katılım", "Kalkınma ve Yatırım", "Kamu", "Yerli Özel", "Yabancı"]
GS = {g: i for i, g in enumerate(GRUP_SIRA)}


# =====================================================================================
# HTTP (urllib, çerezli oturum)
# =====================================================================================
class Http:
    def __init__(self):
        self.jar = http.cookiejar.CookieJar()
        self.op = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(self.jar))
        self.op.addheaders = [("User-Agent", UA)]

    def get(self, url, params=None, timeout=120):
        if params:
            url += "?" + urllib.parse.urlencode(params)
        with self.op.open(url, timeout=timeout) as r:
            return r.status, r.read()

    def post(self, url, data, timeout=300):
        body = urllib.parse.urlencode(data).encode()
        req = urllib.request.Request(url, data=body, headers={"Content-Type": "application/x-www-form-urlencoded"})
        with self.op.open(req, timeout=timeout) as r:
            return r.status, r.read()


# =====================================================================================
# İş yürütücü: hata verenleri en sonda yeniden dener
# =====================================================================================
TEKRAR_BEKLEME = [10, 30, 60]  # saniye; her tekrar turundan önce


def _kisa(e):
    m = str(e).strip().splitlines()
    return (type(e).__name__ + (": " + m[0] if m else ""))[:120]


def isleri_calistir(isler, fonk, paralel, etiket, klasor=None, tur_adi=""):
    """isler: argüman demetleri. fonk(*is) bir sonuç döndürür.
    Önce paralel çalıştırır; hata verenleri sona bırakıp TEKRAR_BEKLEME sırasıyla tek tek yeniden dener.
    (sonuçlar, hâlâ başarısız olan işler) döndürür; kalanları klasör/eksikler.txt dosyasına yazar."""
    sonuc, kalan = [], list(isler)
    for tur in range(len(TEKRAR_BEKLEME) + 1):
        if tur:
            bekle = TEKRAR_BEKLEME[tur - 1]
            print(f"\n  {len(kalan)} iş başarısız oldu; {bekle} sn sonra yeniden deneniyor "
                  f"(tekrar {tur}/{len(TEKRAR_BEKLEME)}, tek tek)...", flush=True)
            time.sleep(bekle)
        hatali = []
        with ThreadPoolExecutor(paralel if tur == 0 else 1) as ex:
            fs = {ex.submit(fonk, *i): i for i in kalan}
            for f in as_completed(fs):
                i = fs[f]
                try:
                    r = f.result()
                    sonuc.append(r)
                    print(f"  {etiket(i, r)}", flush=True)
                except Exception as e:  # ağ, zaman aşımı, bozuk yanıt...
                    hatali.append(i)
                    print(f"  {etiket(i, None)} HATA ({_kisa(e)})", flush=True)
        kalan = hatali
        if not kalan:
            break
    if klasor is not None:
        yol = klasor / "eksikler.txt"
        eski = [x for x in (yol.read_text("utf-8").splitlines() if yol.exists() else []) if not x.startswith(tur_adi + " ")]
        yeni = [f"{tur_adi} {etiket(i, None)}" for i in kalan]
        if eski or yeni:
            klasor.mkdir(parents=True, exist_ok=True)
            yol.write_text("\n".join(eski + yeni) + "\n", "utf-8")
        elif yol.exists():
            yol.unlink()
    if kalan:
        print(f"\n  UYARI: {len(kalan)} iş tüm denemelere rağmen inmedi; listesi eksikler.txt dosyasında. "
              f"Daha sonra 'güncelle' çalıştırınca yeniden denenir.", flush=True)
    return sonuc, kalan


# =====================================================================================
# Basit XLSX okuyucu / yazıcı (standart kütüphane)
# =====================================================================================
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def _kolon_no(ref):
    harf = re.match(r"[A-Z]+", ref).group(0)
    n = 0
    for ch in harf:
        n = n * 26 + ord(ch) - 64
    return n - 1


def _kolon_harf(n):
    s = ""
    n += 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def xlsx_oku(icerik):
    """İlk çalışma sayfasını satır listesi olarak döndürür (değerler metin)."""
    z = zipfile.ZipFile(io.BytesIO(icerik))
    ortak = []
    if "xl/sharedStrings.xml" in z.namelist():
        for si in ET.fromstring(z.read("xl/sharedStrings.xml")).iter(NS + "si"):
            ortak.append("".join(t.text or "" for t in si.iter(NS + "t")))
    sayfa = sorted(n for n in z.namelist() if n.startswith("xl/worksheets/") and n.endswith(".xml"))[0]
    satirlar = []
    for row in ET.fromstring(z.read(sayfa)).iter(NS + "row"):
        hucre = {}
        for c in row.iter(NS + "c"):
            t = c.get("t")
            v = c.find(NS + "v")
            if t == "s" and v is not None:
                deger = ortak[int(v.text)]
            elif t == "inlineStr":
                deger = "".join(x.text or "" for x in c.iter(NS + "t"))
            else:
                deger = v.text if v is not None else ""
            hucre[_kolon_no(c.get("r"))] = deger
        satirlar.append([hucre.get(i, "") for i in range(max(hucre) + 1)] if hucre else [])
    return satirlar


_KONTROL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def xlsx_yaz(yol, sayfalar, genislikler=None):
    """sayfalar: [(sayfa_adı, başlık_listesi, satırlar)]. Sayılar sayı, metinler metin olarak yazılır."""
    stiller = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
               '<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.000"/></numFmts>'
               '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>'
               '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts>'
               '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
               '<fill><patternFill patternType="solid"><fgColor rgb="FF1F3864"/><bgColor indexed="64"/></patternFill></fill></fills>'
               '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
               '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
               '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
               '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
               '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>'
               '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
               '</styleSheet>')
    with zipfile.ZipFile(yol, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                   '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
                   + "".join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                             for i in range(1, len(sayfalar) + 1)) + "</Types>")
        z.writestr("_rels/.rels",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                   "</Relationships>")
        z.writestr("xl/workbook.xml",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                   'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
                   + "".join(f'<sheet name="{escape(ad, {chr(34): "&quot;"})}" sheetId="{i}" r:id="rId{i}"/>'
                             for i, (ad, _, _) in enumerate(sayfalar, 1)) + "</sheets></workbook>")
        z.writestr("xl/_rels/workbook.xml.rels",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   + "".join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>'
                             for i in range(1, len(sayfalar) + 1))
                   + f'<Relationship Id="rId{len(sayfalar) + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
                   "</Relationships>")
        z.writestr("xl/styles.xml", stiller)
        for i, (ad, baslik, satirlar) in enumerate(sayfalar, 1):
            parca = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                     '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                     '<sheetViews><sheetView workbookViewId="0"><pane xSplit="4" ySplit="1" topLeftCell="E2" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>']
            gen = genislikler or []
            parca.append("<cols>" + "".join(
                f'<col min="{j + 1}" max="{j + 1}" width="{gen[j] if j < len(gen) else 16}" customWidth="1"/>'
                for j in range(len(baslik))) + "</cols><sheetData>")
            for r, satir in enumerate([baslik] + list(satirlar), 1):
                hucreler = []
                for j, v in enumerate(satir):
                    ref = f"{_kolon_harf(j)}{r}"
                    if v is None or v == "":
                        continue
                    if r == 1:
                        hucreler.append(f'<c r="{ref}" s="1" t="inlineStr"><is><t>{escape(_KONTROL.sub("", str(v)))}</t></is></c>')
                    elif isinstance(v, (int, float)) and not isinstance(v, bool):
                        hucreler.append(f'<c r="{ref}"{" s=" + chr(34) + "2" + chr(34) if isinstance(v, float) else ""}><v>{v!r}</v></c>')
                    else:
                        hucreler.append(f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">{escape(_KONTROL.sub("", str(v)))}</t></is></c>')
                parca.append(f'<row r="{r}">' + "".join(hucreler) + "</row>")
            son = f"{_kolon_harf(len(baslik) - 1)}{len(satirlar) + 1}"
            parca.append(f'</sheetData><autoFilter ref="A1:{son}"/></worksheet>')
            z.writestr(f"xl/worksheets/sheet{i}.xml", "".join(parca))


def sayfa_adi(ad):
    ad = re.sub(r"[\\/*?:\[\]]", "", ad).replace("Bankalarda Saklanan ", "Saklanan ")
    return ad[:31]


def sayi(x):
    if x is None:
        return None
    s = str(x).strip()
    if s == "":
        return None
    try:
        return float(s)
    except ValueError:
        try:
            return float(s.replace(".", "").replace(",", "."))
        except ValueError:
            return None


def tr_sayi(s):
    """BDDK HTML tablosundaki Türkçe biçimli sayı: '18.523.337,65624' ya da '993.454' (=993454)."""
    s = s.strip()
    if s in ("", "-"):
        return None
    try:
        return float(s.replace(".", "").replace(",", "."))
    except ValueError:
        return None


def csv_oku(yol):
    if not yol.exists():
        return []
    with open(yol, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def csv_yaz(yol, kolonlar, satirlar):
    with open(yol, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(kolonlar)
        for s in satirlar:
            w.writerow(["" if s.get(k) is None else (repr(s[k]) if isinstance(s[k], float) else s[k]) for k in kolonlar])


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
AYLIK_KOLON = ["donem", "tablo_no", "tablo", "grup", "sira", "kalem", "sutun", "deger"]


def aylik_indir(tablo_no, yil, ay, para="TL", deneme=4):
    params = [("tabloNo", tablo_no), ("yil", yil), ("ay", ay), ("paraBirimi", para)]
    params += [(f"taraf[{i}]", g) for i, g in enumerate(AYLIK_TARAF)]
    hata = ""
    for k in range(deneme):
        try:
            st, icerik = Http().get(AYLIK_URL, params, timeout=90)
            if st == 200 and icerik[:2] == b"PK":
                satirlar = xlsx_oku(icerik)
                return satirlar[0], [x for x in satirlar[1:] if x and x[0]]
            hata = f"HTTP {st}"
        except Exception as e:  # bağlantı kopması, zaman aşımı, bozuk dosya...
            hata = _kisa(e)
        time.sleep(2 * (k + 1))
    raise RuntimeError(f"Aylık tablo {tablo_no} {yil}/{ay}: {hata}")


def aylik_yayinlandi_mi(yil, ay):
    try:
        baslik, satirlar = aylik_indir(1, yil, ay, deneme=2)
    except RuntimeError:
        return False
    return f"Dönem:{yil}/{ay}" in str(baslik[2]) and len(satirlar) > 0


def aylik_son_donem():
    bugun = date.today()
    y, a = bugun.year, bugun.month
    for _ in range(4):
        y, a = (y - 1, 12) if a == 1 else (y, a - 1)
        if aylik_yayinlandi_mi(y, a):
            return y, a
    raise RuntimeError("Son aylık dönem bulunamadı (bağlantıyı kontrol edin).")


def aylik_tablo(yil, ay, t, para):
    baslik, satirlar = aylik_indir(t, yil, ay, para)
    if f"Dönem:{yil}/{ay}" not in str(baslik[2]) and t not in (15, 16, 17):
        raise RuntimeError(f"{yil}/{ay} dönemi yayımlanmamış görünüyor")
    sutunlar = [str(c) for c in baslik[3:] if c not in (None, "")]
    kayit = []
    for r in satirlar:
        for j, s in enumerate(sutunlar):
            kayit.append({"donem": f"{yil}-{ay:02d}", "tablo_no": t, "tablo": AYLIK_TABLOLAR[t],
                          "grup": str(r[0]).strip(), "sira": int(float(r[1])), "kalem": str(r[2]).strip(),
                          "sutun": s, "deger": sayi(r[3 + j] if 3 + j < len(r) else None)})
    if not kayit:
        raise RuntimeError("boş tablo")
    return kayit


def aylik_kaydet(kayitlar, klasor, para):
    klasor.mkdir(parents=True, exist_ok=True)
    ek = "" if para == "TL" else f"_{para}"
    yillar = sorted({k["donem"][:4] for k in kayitlar})
    for yil in yillar:
        yeni = [k for k in kayitlar if k["donem"][:4] == yil]
        anahtar = {(k["donem"], str(k["tablo_no"])) for k in yeni}
        yol = klasor / f"bddk_aylik_{yil}{ek}.csv"
        eski = [k for k in csv_oku(yol) if (k["donem"], k["tablo_no"]) not in anahtar]
        for k in eski:
            k["tablo_no"], k["sira"], k["deger"] = int(k["tablo_no"]), int(k["sira"]), sayi(k["deger"])
        tum = eski + yeni
        tum.sort(key=lambda k: (k["donem"], k["tablo_no"], GS.get(k["grup"], 99), k["sira"]))
        csv_yaz(yol, AYLIK_KOLON, tum)
        # Excel: her tablo bir sayfa, sütunlar yan yana
        sayfalar = []
        for t in sorted({k["tablo_no"] for k in tum}):
            kt = [k for k in tum if k["tablo_no"] == t]
            sutunlar = list(dict.fromkeys(k["sutun"] for k in kt))
            satir = {}
            for k in kt:
                satir.setdefault((k["donem"], k["grup"], k["sira"], k["kalem"]), {})[k["sutun"]] = k["deger"]
            rows = [[d, g, s, ka] + [v.get(c) for c in sutunlar]
                    for (d, g, s, ka), v in sorted(satir.items(), key=lambda x: (x[0][0], GS.get(x[0][1], 99), x[0][2]))]
            sayfalar.append((sayfa_adi(AYLIK_TABLOLAR[t]), ["Dönem", "Grup", "Sıra", "Kalem"] + sutunlar, rows))
        xlsx_yaz(klasor / f"BDDK_Aylik_{yil}{ek}.xlsx", sayfalar, [9, 20, 6, 60])
        print(f"  Kaydedildi: {yol.name} ve BDDK_Aylik_{yil}{ek}.xlsx")


def aylik_mevcut(klasor, para="TL"):
    """Klasörde zaten bulunan (dönem, tablo) çiftleri."""
    ek = "" if para == "TL" else f"_{para}"
    mevcut = set()
    for f in sorted((klasor / "aylik").glob(f"bddk_aylik_[0-9][0-9][0-9][0-9]{ek}.csv")):
        with open(f, encoding="utf-8-sig", newline="") as fh:
            for r in csv.DictReader(fh):
                mevcut.add((r["donem"], int(r["tablo_no"])))
    return mevcut


def aylik_isleri_yap(isler, para, klasor, paralel):
    """isler: [(yil, ay, tablo)] -> indir, tekrar dene, kaydet, raporu güncelle."""
    if not isler:
        return
    print(f"Aylık bülten: {len({(y, a) for y, a, _ in isler})} dönem, {len(isler)} tablo indirilecek ({para})...")
    sonuclar, _ = isleri_calistir(
        isler, lambda y, a, t: aylik_tablo(y, a, t, para), paralel,
        lambda i, r: f"{i[0]}-{i[1]:02d} {AYLIK_TABLOLAR[i[2]]}" + ("" if r is None else " tamam"),
        klasor, "aylik")
    kayitlar = [k for r in sonuclar for k in r]
    if kayitlar:
        aylik_kaydet(kayitlar, klasor / "aylik", para)
        if para == "TL":
            rapor_olustur(klasor)


def donem_listesi(bas, bit):
    out, (y, a) = [], bas
    while (y, a) <= bit:
        out.append((y, a))
        y, a = (y + 1, 1) if a == 12 else (y, a + 1)
    return out


def aylik_calistir(baslangic, bitis, tablolar, para, klasor, paralel):
    if not bitis:
        print("En son yayımlanan ay aranıyor...")
        bitis = aylik_son_donem()
        print(f"  Yayımlanmış en son dönem: {bitis[0]}-{bitis[1]:02d}")
    baslangic = baslangic or bitis
    isler = [(y, a, t) for y, a in donem_listesi(baslangic, bitis) for t in tablolar]
    aylik_isleri_yap(isler, para, klasor, paralel)


def aylik_guncelle(klasor, paralel, para="TL"):
    """Yalnız eksik ve yeni yayımlanmış dönem/tabloları indirir."""
    print("Aylık bülten: mevcut veriler kontrol ediliyor...")
    mevcut = aylik_mevcut(klasor, para)
    son = aylik_son_donem()
    if mevcut:
        bas = tuple(map(int, min(d for d, _ in mevcut).split("-")))
        print(f"  Klasörde: {len({d for d, _ in mevcut})} dönem ({min(d for d, _ in mevcut)} – "
              f"{max(d for d, _ in mevcut)}); yayımlanmış en son: {son[0]}-{son[1]:02d}")
    else:
        bas = son
        print(f"  Klasörde aylık veri yok; yalnız en son dönem ({son[0]}-{son[1]:02d}) indirilecek. "
              f"Geçmiş için 'dönem aralığı' seçeneğini kullanın.")
    isler = [(y, a, t) for y, a in donem_listesi(bas, son) for t in AYLIK_TABLOLAR
             if (f"{y}-{a:02d}", t) not in mevcut]
    if not isler:
        print("  Aylık bülten güncel; eksik ya da yeni dönem yok.")
        return
    aylik_isleri_yap(isler, para, klasor, paralel)


# =====================================================================================
# HAFTALIK BÜLTEN
# =====================================================================================
HAFTALIK_URL = "https://www.bddk.org.tr/BultenHaftalik"
HAFTALIK_TARAF = ["10001", "10002", "10008", "10010", "10009", "10004", "10003", "10005", "10007", "10006"]
SUTUNLAR = ["TP", "YP", "Toplam"]
YALNIZ_TOPLAM = {9}  # Yabancı Para Pozisyonu: BDDK yalnız Toplam sütununu dolduruyor
HAFTALIK_KOLON = ["tarih", "tablo_no", "tablo", "grup", "sira", "kalem"] + SUTUNLAR


def _temiz(s):
    return re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", s))).strip()


def _tarih(s):
    return datetime.strptime(s.split()[0], "%d.%m.%Y").date()


class HaftalikOturum:
    def __init__(self):
        self.h = Http()
        self.yenile()

    def yenile(self):
        st, b = self.h.get(f"{HAFTALIK_URL}/tr/Gelismis")
        self.sayfa = b.decode("utf-8", "replace")
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
        hata = ""
        for k in range(deneme):
            veri = [("__RequestVerificationToken", self.belirtec), ("BaslangicTarihi", bas), ("BitisTarihi", bit),
                    ("dil", "tr"), ("SeciliParalar", "TL")]
            veri += [("Kalemler", str(i)) for i in kalem_ids]
            veri += [("Taraflar", t) for t in HAFTALIK_TARAF]
            veri += [("kalemSutun", c) for c in SUTUNLAR]
            try:
                st, b = self.h.post(f"{HAFTALIK_URL}/tr/Gelismis/GelismisRaporGetir", veri)
                s = b.decode("utf-8", "replace")
                if st == 200 and 'id="TabloExcelGelismis2"' in s:
                    return s
                hata = f"HTTP {st}"
            except Exception as e:
                hata = _kisa(e)
            time.sleep(3 * (k + 1))
            try:
                self.yenile()
            except Exception:
                pass
        raise RuntimeError(f"Haftalık {bas[:10]}–{bit[:10]}: {hata}")


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
    for sira, ((_, cs), (_, kad)) in enumerate(zip(gruplar, kalemler), 1):
        kolon_kalem += [(sira, kad)] * cs
    alt = alt[len(alt) - len(kolon_kalem):]
    kolon_sutun = [("Toplam" if s.upper().startswith("TOPLAM") else s.upper()[:2]) for s in alt]
    sonuc, grup = {}, None
    for st in satirlar[2:]:
        h = _hucreler(st)
        if not re.match(r"\d{1,2}\.\d{1,2}\.\d{4}$", h[0][0]):
            if len(h) <= 2 and h[0][0]:
                ad = h[0][0]
                grup = re.sub(r"\s*-\s*", "-", ad) if ad.lower().startswith("mevduat") else ad
            continue
        t = datetime.strptime(h[0][0], "%d.%m.%Y").date().isoformat()
        for (deg, _), (sira, kad), sut in zip(h[1:], kolon_kalem, kolon_sutun):
            if tablo_no in YALNIZ_TOPLAM and sut != "Toplam":
                continue
            k = sonuc.setdefault((t, grup, sira), {"tarih": t, "tablo_no": tablo_no, "tablo": tablo_ad,
                                                    "grup": grup, "sira": sira, "kalem": kad})
            k[sut] = tr_sayi(deg)
    return [k for k in sonuc.values() if any(k.get(s) is not None for s in SUTUNLAR)]


def haftalik_kaydet(kayitlar, klasor):
    klasor.mkdir(parents=True, exist_ok=True)
    for yil in sorted({k["tarih"][:4] for k in kayitlar}):
        yeni = [k for k in kayitlar if k["tarih"][:4] == yil]
        anahtar = {(k["tarih"], str(k["tablo_no"])) for k in yeni}
        yol = klasor / f"bddk_haftalik_{yil}.csv"
        eski = [k for k in csv_oku(yol) if (k["tarih"], k["tablo_no"]) not in anahtar]
        for k in eski:
            k["tablo_no"], k["sira"] = int(k["tablo_no"]), int(k["sira"])
            for s in SUTUNLAR:
                k[s] = sayi(k[s])
        tum = eski + yeni
        tum.sort(key=lambda k: (k["tarih"], k["tablo_no"], GS.get(k["grup"], 99), k["sira"]))
        csv_yaz(yol, HAFTALIK_KOLON, tum)
        sayfalar = []
        for t in sorted({k["tablo_no"] for k in tum}):
            kt = [k for k in tum if k["tablo_no"] == t]
            cols = ["Toplam"] if t in YALNIZ_TOPLAM else SUTUNLAR
            rows = [[k["tarih"], k["grup"], k["sira"], k["kalem"]] + [k.get(c) for c in cols] for k in kt]
            sayfalar.append((sayfa_adi(kt[0]["tablo"]), ["Tarih", "Grup", "Sıra", "Kalem"] + cols, rows))
        xlsx_yaz(klasor / f"BDDK_Haftalik_{yil}.xlsx", sayfalar, [11, 20, 6, 60])
        print(f"  Kaydedildi: {yol.name} ve BDDK_Haftalik_{yil}.xlsx")


def haftalik_mevcut(klasor):
    mevcut = set()
    for f in sorted((klasor / "haftalik").glob("bddk_haftalik_[0-9][0-9][0-9][0-9].csv")):
        with open(f, encoding="utf-8-sig", newline="") as fh:
            for r in csv.DictReader(fh):
                mevcut.add((r["tarih"], int(r["tablo_no"])))
    return mevcut


def haftalik_isleri_yap(o, tablolar, isler, klasor, paralel):
    """isler: [(yil, baslangic_form_degeri, bitis_form_degeri, tablo_no)]"""
    if not isler:
        return
    print(f"Haftalık bülten: {len(isler)} istek (yıl × tablo) indirilecek...")
    yerel = threading.local()

    def is_yap(y, b, e, n):
        if not hasattr(yerel, "o"):
            yerel.o = HaftalikOturum()
        t = tablolar[n]
        s = yerel.o.rapor(b, e, [i for i, _ in t["kalemler"]])
        try:
            return haftalik_ayristir(s, n, t["ad"], t["kalemler"])
        except Exception:
            yerel.o = HaftalikOturum()  # bir sonraki deneme yeni oturumla
            raise

    def etiket(i, r):
        y, b, e, n = i
        ad = f"{_tarih(b)}–{_tarih(e)} {tablolar[n]['ad']}"
        return ad if r is None else f"{ad}: {len(r):,} satır"

    sonuclar, _ = isleri_calistir(isler, is_yap, paralel, etiket, klasor, "haftalik")
    kayitlar = [k for r in sonuclar for k in r]
    if kayitlar:
        haftalik_kaydet(kayitlar, klasor / "haftalik")
        haftalik_rapor_olustur(klasor)


def _haftalik_hazirla(tablo_sec=None):
    print("Haftalık bülten sayfası açılıyor...")
    o = HaftalikOturum()
    tablolar = o.tablolar()
    tarihler = o.tarihler()
    print(f"  {len(tablolar)} tablo, {sum(len(t['kalemler']) for t in tablolar.values())} kalem, "
          f"{len(tarihler)} hafta ({_tarih(tarihler[0])} – {_tarih(tarihler[-1])})")
    if tablo_sec:
        tablolar = {n: t for n, t in tablolar.items() if n in tablo_sec}
    return o, tablolar, tarihler


def haftalik_calistir(bas_yil, bit_yil, tablo_sec, klasor, paralel):
    o, tablolar, tarihler = _haftalik_hazirla(tablo_sec)
    bas_yil = bas_yil or _tarih(tarihler[-1]).year
    bit_yil = bit_yil or _tarih(tarihler[-1]).year
    isler = []
    for y in range(bas_yil, bit_yil + 1):
        t = [x for x in tarihler if _tarih(x).year == y]
        if t:
            isler += [(y, t[0], t[-1], n) for n in tablolar]
    haftalik_isleri_yap(o, tablolar, isler, klasor, paralel)


def haftalik_guncelle(klasor, paralel):
    """Yalnız eksik ve yeni yayımlanmış hafta/tabloları indirir."""
    mevcut = haftalik_mevcut(klasor)
    o, tablolar, tarihler = _haftalik_hazirla()
    son = _tarih(tarihler[-1])
    if mevcut:
        ilk = min(t for t, _ in mevcut)
        print(f"  Klasörde: {len({t for t, _ in mevcut})} hafta ({ilk} – {max(t for t, _ in mevcut)}); "
              f"yayımlanmış en son: {son}")
    else:
        ilk = date(son.year, 1, 1).isoformat()
        print(f"  Klasörde haftalık veri yok; {son.year} yılı indirilecek. Geçmiş için 'yıl aralığı' seçeneğini kullanın.")
    isler = []
    for n in tablolar:
        eksik = [x for x in tarihler if _tarih(x).isoformat() >= ilk and (_tarih(x).isoformat(), n) not in mevcut]
        for y in sorted({_tarih(x).year for x in eksik}):
            ey = [x for x in eksik if _tarih(x).year == y]
            isler.append((y, ey[0], ey[-1], n))
    if not isler:
        print("  Haftalık bülten güncel; eksik ya da yeni hafta yok.")
        return
    haftalik_isleri_yap(o, tablolar, isler, klasor, paralel)


# =====================================================================================
# HTML RAPOR (aylık verilerden; şablon ve Chart.js betiğin içinde gömülü)
# =====================================================================================
import base64
import gzip
import json
from datetime import timedelta, timezone

RAPOR_GRUP_SIRA = ["Sektör", "Mevduat", "Mevduat-Yerli Özel", "Mevduat-Kamu", "Mevduat-Yabancı",
                   "Katılım", "Kalkınma ve Yatırım", "Yerli Özel", "Kamu", "Yabancı"]
RAPOR_BIRIM = {5: "bin TL", 15: "%", 16: "adet", 17: "%"}
ORAN_TABLOLARI = {15, 16, 17}
TEKRAR_EK = {11: ["Varlık", "Yükümlülük"]}
AD_ESLE = {
    "Finansal Kiralama Borçları (Net)": "Kiralama İşlemlerinden Yükümlülükler (Net)",
    "Sermaye Hesaplamasına Dahil Edilecek Borçlanma Araçları": "Sermaye Hesaplamasına Dahil Edilecek Borçlanma Araçları (Net)",
}
AYLAR = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]


def _kalem_adi(s):
    s = re.sub(r"\*+", "", str(s))
    s = re.sub(r"\s*[(\[](?=[^)\]]*\d)[\d\s+\-.…]+[)\]]", "", s)
    s = re.sub(r"\s*\(\d+ den [^)]*\)", "", s)
    s = re.sub(r"\s*\[[\d\s+\-.…()]*\]", "", s)
    s = re.sub(r"\s+", " ", s).strip().rstrip("-").strip()
    return AD_ESLE.get(s, s)


def _yil_ek(y, tablo):
    son, onlar = y % 10, y % 100
    if son == 0:
        return tablo[0][onlar]
    return tablo[1][son]


def _yil_den(y):
    return f"{y}’" + _yil_ek(y, ({0: "dan", 10: "dan", 20: "den", 30: "dan", 40: "tan", 50: "den", 60: "tan", 70: "ten", 80: "den", 90: "dan"},
                                {1: "den", 2: "den", 3: "ten", 4: "ten", 5: "ten", 6: "dan", 7: "den", 8: "den", 9: "dan"}))


def _yil_e(y):
    return f"{y}’" + _yil_ek(y, ({0: "e", 10: "a", 20: "ye", 30: "a", 40: "a", 50: "ye", 60: "a", 70: "e", 80: "e", 90: "a"},
                                {1: "e", 2: "ye", 3: "e", 4: "e", 5: "e", 6: "ya", 7: "ye", 8: "e", 9: "a"}))


def rapor_verisi(kayitlar):
    """Uzun biçimli aylık kayıtlardan (dict listesi) raporun beklediği veri yapısını kurar."""
    for k in kayitlar:
        k["ad"] = _kalem_adi(k["kalem"])
    # aynı tabloda aynı dönemde tekrar eden adları ayır (ör. 'a) Gerçek Kişiler' farklı başlıkların altında)
    ilk_sutun, ilk_grup = {}, kayitlar[0]["grup"]
    for k in kayitlar:
        ilk_sutun.setdefault(k["tablo_no"], k["sutun"])
    ilk = {}
    for k in kayitlar:
        if k["grup"] == ilk_grup and k["sutun"] == ilk_sutun[k["tablo_no"]]:
            ilk.setdefault((k["donem"], k["tablo_no"]), []).append((k["sira"], k["ad"]))
    yeni = {}
    for (d, t), x in ilk.items():
        x.sort()
        say = {}
        for _, a in x:
            say[a] = say.get(a, 0) + 1
        hedef = {a for a, n in say.items() if n > 1}
        if not hedef:
            continue
        ust, sayac = None, {}
        for sira, ad in x:
            alt = bool(re.match(r"^[a-zçğıöşü]\)\s", ad))
            if alt and ad in hedef and ust:
                yeni[(d, t, sira)] = f"{ad} · {ust}"
            elif not alt and ad in hedef:
                sayac[ad] = sayac.get(ad, 0) + 1
                ek = TEKRAR_EK.get(t, [])
                yeni[(d, t, sira)] = f"{ad} ({ek[sayac[ad] - 1] if sayac[ad] <= len(ek) else sayac[ad]})"
            if not alt:
                ust = yeni.get((d, t, sira), ad)
    for k in kayitlar:
        k["ad"] = yeni.get((k["donem"], k["tablo_no"], k["sira"]), k["ad"])

    per = sorted({k["donem"] for k in kayitlar})
    gruplar = {k["grup"] for k in kayitlar}
    grp = [g for g in RAPOR_GRUP_SIRA if g in gruplar] + sorted(gruplar - set(RAPOR_GRUP_SIRA))
    pi = {p: i for i, p in enumerate(per)}
    gi = {g: i for i, g in enumerate(grp)}
    tablolar = []
    for no in sorted({k["tablo_no"] for k in kayitlar}):
        d = [k for k in kayitlar if k["tablo_no"] == no]
        son_p = max(k["donem"] for k in d)
        cols = list(dict.fromkeys(k["sutun"] for k in d if k["donem"] == son_p))
        cols += [c for c in dict.fromkeys(k["sutun"] for k in d) if c not in cols]
        # her ad için en son göründüğü dönem ve sıra
        son = {}
        for k in d:
            if k["ad"] not in son or (k["donem"], k["sira"]) > son[k["ad"]]:
                son[k["ad"]] = (k["donem"], k["sira"])
        items = [a for a, (p, s_) in sorted(son.items(), key=lambda x: x[1][1]) if p == son_p]
        eskiler = sorted(((p, s_, a) for a, (p, s_) in son.items() if p != son_p), key=lambda x: (x[0], x[1]), reverse=True)
        donem_sira = {}
        for k in d:
            donem_sira.setdefault(k["donem"], {})[k["sira"]] = k["ad"]
        for p, s_, a in eskiler:
            onceki = [donem_sira[p][x] for x in sorted(donem_sira[p]) if x < s_]
            yer = next((items.index(o) + 1 for o in reversed(onceki) if o in items), len(items))
            items.insert(yer, a)
        ii = {a: i for i, a in enumerate(items)}
        ci = {c: i for i, c in enumerate(cols)}
        ng, np_ = len(grp), len(per)
        v = [None] * (len(items) * len(cols) * ng * np_)
        dec = 3 if no in ORAN_TABLOLARI else 1
        for k in d:
            x = k["deger"]
            if x is None:
                continue
            x = round(x, dec)
            if float(x).is_integer():
                x = int(x)
            v[((ii[k["ad"]] * len(cols) + ci[k["sutun"]]) * ng + gi[k["grup"]]) * np_ + pi[k["donem"]]] = x
        tablolar.append({"no": no, "ad": AYLIK_TABLOLAR.get(no, d[0]["tablo"]), "birim": RAPOR_BIRIM.get(no, "milyon TL"),
                         "cols": cols, "items": [{"ad": a} for a in items], "v": v})
    tr = datetime.now(timezone(timedelta(hours=3)))
    return {"per": per, "grp": grp, "tab": tablolar, "meta": {"guncelleme": f"{tr.day} {AYLAR[tr.month - 1]} {tr.year}"}}


def rapor_olustur(klasor):
    aylik = klasor / "aylik"
    dosyalar = sorted(aylik.glob("bddk_aylik_[0-9][0-9][0-9][0-9].csv"))
    if not dosyalar:
        print("Rapor için aylık veri yok. Önce aylık bülteni indirin (menüde 1 ya da 2).")
        return None
    print(f"Rapor hazırlanıyor ({len(dosyalar)} yıllık dosya)...")
    kayitlar = []
    for f in dosyalar:
        for k in csv_oku(f):
            kayitlar.append({"donem": k["donem"], "tablo_no": int(k["tablo_no"]), "tablo": k["tablo"], "grup": k["grup"],
                             "sira": int(k["sira"]), "kalem": k["kalem"], "sutun": k["sutun"], "deger": sayi(k["deger"])})
    D = rapor_verisi(kayitlar)
    veri = base64.b64encode(gzip.compress(json.dumps(D, ensure_ascii=False, separators=(",", ":")).encode("utf-8"), 9)).decode()
    sablon = gzip.decompress(base64.b64decode(SABLON_GZ_B64)).decode("utf-8")
    chartjs = gzip.decompress(base64.b64decode(CHARTJS_GZ_B64)).decode("utf-8")
    p0, p1 = D["per"][0], D["per"][-1]
    y0, m0 = map(int, p0.split("-"))
    y1, m1 = map(int, p1.split("-"))
    cdn = '<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>'
    sayfa = (sablon.replace("{{ARALIK_DEN_YA}}", f"{AYLAR[m0 - 1]} {_yil_den(y0)} {AYLAR[m1 - 1]} {_yil_e(y1)}")
             .replace("{{ARALIK}}", f"{AYLAR[m0 - 1]} {y0} – {AYLAR[m1 - 1]} {y1}")
             .replace("{{N}}", str(len(D["per"])))
             .replace("{{HAFTALIK_LINK}}", "BDDK_Haftalik_Rapor.html")
             .replace("{{VERI}}", veri)
             .replace(cdn, "<script>" + chartjs.replace("</script", "<\\/script") + "</script>"))
    yol = klasor / "BDDK_Aylik_Rapor.html"
    yol.write_text(sayfa, encoding="utf-8")
    print(f"  Rapor: {yol} ({yol.stat().st_size / 1e6:.1f} MB, {len(D['per'])} dönem, {len(D['tab'])} tablo)")
    return yol


# =====================================================================================
# HAFTALIK HTML RAPOR
# =====================================================================================
def haftalik_rapor_verisi(kayitlar):
    """kayitlar: dict listesi (tarih, tablo_no, tablo, grup, sira, kalem, TP, YP, Toplam).
    Her seri (tablo, sıra, grup, sütun) milyon TL tamsayı ve haftadan haftaya fark olarak saklanır."""
    per = sorted({k["tarih"] for k in kayitlar})
    pi = {p: i for i, p in enumerate(per)}
    gruplar = {k["grup"] for k in kayitlar}
    grp = [g for g in RAPOR_GRUP_SIRA if g in gruplar] + sorted(gruplar - set(RAPOR_GRUP_SIRA))
    gi = {g: i for i, g in enumerate(grp)}
    seriler, tablolar = {}, {}
    for k in kayitlar:
        t, s_ = int(k["tablo_no"]), int(k["sira"])
        tb = tablolar.setdefault(t, {"ad": k["tablo"], "items": {}, "cols": set()})
        it = tb["items"].setdefault(s_, {"ad": k["kalem"], "son": -1, "ilk": 10 ** 9})
        p = pi[k["tarih"]]
        if p >= it["son"]:
            it["ad"] = k["kalem"]
        for c in SUTUNLAR:
            v = k.get(c)
            if v is None:
                continue
            tb["cols"].add(c)
            it["son"], it["ilk"] = max(it["son"], p), min(it["ilk"], p)
            seriler.setdefault((t, s_, gi[k["grup"]], c), {})[p] = int(round(v))
    s_liste = []
    for (t, s_, g, c), d in sorted(seriler.items()):
        p0, p1 = min(d), max(d)
        out, onceki = [], None
        for p in range(p0, p1 + 1):
            v = d.get(p)
            if v is None:
                out.append(None)
                continue
            out.append(v if onceki is None else v - onceki)
            onceki = v
        s_liste.append([t, s_, g, SUTUNLAR.index(c), p0] + out)
    tab = []
    for t in sorted(tablolar):
        tb = tablolar[t]
        tab.append({"no": t, "ad": tb["ad"], "cols": [c for c in SUTUNLAR if c in tb["cols"]],
                    "items": [{"s": s_, "ad": it["ad"], "ilk": it["ilk"], "son": it["son"]}
                              for s_, it in sorted(tb["items"].items()) if it["son"] >= 0]})
    tr = datetime.now(timezone(timedelta(hours=3)))
    return {"per": per, "grp": grp, "cols": SUTUNLAR, "tab": tab, "s": s_liste,
            "meta": {"guncelleme": f"{tr.day} {AYLAR[tr.month - 1]} {tr.year}"}}


def _tarih_uzun(iso):
    y, m, d = map(int, iso.split("-"))
    return f"{d} {AYLAR[m - 1]} {y}"


def haftalik_rapor_html(D, sablon, chartjs=None, aylik_link="BDDK_Aylik_Rapor.html"):
    veri = base64.b64encode(gzip.compress(json.dumps(D, ensure_ascii=False, separators=(",", ":")).encode("utf-8"), 9)).decode()
    y0, y1 = int(D["per"][0][:4]), int(D["per"][-1][:4])
    sayfa = (sablon.replace("{{ARALIK}}", f"{_tarih_uzun(D['per'][0])} – {_tarih_uzun(D['per'][-1])}")
             .replace("{{ARALIK_DEN_YA}}", f"{_tarih_uzun(D['per'][0])[:-4]}{_yil_den(y0)} {_tarih_uzun(D['per'][-1])[:-4]}{_yil_e(y1)}")
             .replace("{{N}}", str(len(D["per"])))
             .replace("{{AYLIK_LINK}}", aylik_link)
             .replace("{{VERI}}", veri))
    if chartjs:
        cdn = '<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>'
        sayfa = sayfa.replace(cdn, "<script>" + chartjs.replace("</script", "<\\/script") + "</script>")
    return sayfa


def haftalik_rapor_olustur(klasor):
    dosyalar = sorted((klasor / "haftalik").glob("bddk_haftalik_[0-9][0-9][0-9][0-9].csv"))
    if not dosyalar:
        print("Haftalık rapor için veri yok. Önce haftalık bülteni indirin.")
        return None
    print(f"Haftalık rapor hazırlanıyor ({len(dosyalar)} yıllık dosya)...")
    kayitlar = []
    for f in dosyalar:
        for k in csv_oku(f):
            kayitlar.append({"tarih": k["tarih"], "tablo_no": int(k["tablo_no"]), "tablo": k["tablo"], "grup": k["grup"],
                             "sira": int(k["sira"]), "kalem": k["kalem"], **{c: sayi(k[c]) for c in SUTUNLAR}})
    D = haftalik_rapor_verisi(kayitlar)
    sablon = gzip.decompress(base64.b64decode(HSABLON_GZ_B64)).decode("utf-8")
    chartjs = gzip.decompress(base64.b64decode(CHARTJS_GZ_B64)).decode("utf-8")
    yol = klasor / "BDDK_Haftalik_Rapor.html"
    yol.write_text(haftalik_rapor_html(D, sablon, chartjs, "BDDK_Aylik_Rapor.html"), encoding="utf-8")
    print(f"  Rapor: {yol} ({yol.stat().st_size / 1e6:.1f} MB, {len(D['per'])} hafta, {len(D['tab'])} tablo)")
    return yol


HSABLON_GZ_B64 = "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"  # <<HSABLON>>


# Aşağıdaki iki değer scripts/yerel_paketle.py ile otomatik doldurulur (rapor_sablonu.html ve Chart.js 4.4.1, MIT).
SABLON_GZ_B64 = "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"  # <<SABLON>>
CHARTJS_GZ_B64 = "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"  # <<CHARTJS>>


# =====================================================================================
def ay_oku(s):
    y, a = s.strip().split("-")
    return int(y), int(a)


def tablo_oku(s):
    return [int(x) for x in s.split(",") if x.strip()] if s else None


def menu(klasor):
    print("BDDK veri indirici\n")
    print("  1) GÜNCELLE – yalnız yeni ve eksik dönemleri indir (aylık + haftalık), raporu yenile")
    print("  2) Aylık bülten  – en son yayımlanan ay")
    print("  3) Aylık bülten  – dönem aralığı")
    print("  4) Haftalık bülten – bu yıl")
    print("  5) Haftalık bülten – yıl aralığı")
    print("  6) HTML raporlarını yeniden oluştur (indirilmiş verilerden)")
    sec = input("\nSeçiminiz (1-6): ").strip()
    if sec == "1":
        aylik_guncelle(klasor, 3)
        print()
        haftalik_guncelle(klasor, 3)
    elif sec == "2":
        aylik_calistir(None, None, list(AYLIK_TABLOLAR), "TL", klasor, 3)
    elif sec == "3":
        b = ay_oku(input("Başlangıç (YYYY-AA, ör. 2018-01): "))
        e = ay_oku(input("Bitiş     (YYYY-AA, ör. 2026-09): "))
        aylik_calistir(b, e, list(AYLIK_TABLOLAR), "TL", klasor, 3)
    elif sec == "4":
        haftalik_calistir(None, None, None, klasor, 3)
    elif sec == "5":
        b = int(input("Başlangıç yılı (en erken 2014): ").strip())
        e = int(input("Bitiş yılı: ").strip())
        haftalik_calistir(b, e, None, klasor, 3)
    elif sec == "6":
        rapor_olustur(klasor)
        haftalik_rapor_olustur(klasor)
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
        try:
            input("\nÇıkmak için Enter...")
        except EOFError:
            pass
        return

    ap = argparse.ArgumentParser(description="BDDK Aylık ve Haftalık Bülten indirici (kurulum gerektirmez)")
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
    alt.add_parser("rapor", help="indirilmiş verilerden aylık ve haftalık HTML raporlarını oluşturur")
    gu = alt.add_parser("guncelle", help="yalnız yeni ve eksik dönemleri indirir, çıktıları ve raporu günceller")
    gu.add_argument("--sadece", choices=["aylik", "haftalik"], help="yalnız birini güncelle")
    g = ap.parse_args()
    klasor = Path(g.klasor)
    if g.tur == "rapor":
        rapor_olustur(klasor)
        haftalik_rapor_olustur(klasor)
    elif g.tur == "guncelle":
        if g.sadece != "haftalik":
            aylik_guncelle(klasor, g.paralel)
        if g.sadece != "aylik":
            print()
            haftalik_guncelle(klasor, g.paralel)
    elif g.tur == "aylik":
        aylik_calistir(ay_oku(g.baslangic) if g.baslangic else None, ay_oku(g.bitis) if g.bitis else None,
                       tablo_oku(g.tablolar) or list(AYLIK_TABLOLAR), g.para, klasor, g.paralel)
    else:
        haftalik_calistir(g.baslangic, g.bitis, tablo_oku(g.tablolar), klasor, g.paralel)
    print(f"\nDosyalar: {klasor}")


if __name__ == "__main__":
    main()
