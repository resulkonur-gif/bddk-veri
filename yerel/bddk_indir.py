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

Çıktılar (betiğin yanında; --klasor ile değiştirilebilir):
    BDDK_Veri/aylik/bddk_aylik_<yıl>.csv        donem,tablo_no,tablo,grup,sira,kalem,sutun,deger
    BDDK_Veri/aylik/BDDK_Aylik_<yıl>.xlsx       her tablo ayrı sayfada
    BDDK_Veri/haftalik/bddk_haftalik_<yıl>.csv  tarih,tablo_no,tablo,grup,sira,kalem,TP,YP,Toplam
    BDDK_Veri/haftalik/BDDK_Haftalik_<yıl>.xlsx
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
        except (urllib.error.URLError, OSError, zipfile.BadZipFile) as e:
            hata = str(e)
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


def aylik_donem(yil, ay, tablolar, para):
    kayit = []
    for t in tablolar:
        baslik, satirlar = aylik_indir(t, yil, ay, para)
        sutunlar = [str(c) for c in baslik[3:] if c not in (None, "")]
        for r in satirlar:
            for j, s in enumerate(sutunlar):
                kayit.append({"donem": f"{yil}-{ay:02d}", "tablo_no": t, "tablo": AYLIK_TABLOLAR[t],
                              "grup": str(r[0]).strip(), "sira": int(float(r[1])), "kalem": str(r[2]).strip(),
                              "sutun": s, "deger": sayi(r[3 + j] if 3 + j < len(r) else None)})
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


def aylik_calistir(baslangic, bitis, tablolar, para, klasor, paralel):
    if not bitis:
        print("En son yayımlanan ay aranıyor...")
        bitis = aylik_son_donem()
        print(f"  Yayımlanmış en son dönem: {bitis[0]}-{bitis[1]:02d}")
    baslangic = baslangic or bitis
    donemler, (y, a) = [], baslangic
    while (y, a) <= bitis:
        donemler.append((y, a))
        y, a = (y + 1, 1) if a == 12 else (y, a + 1)
    print(f"Aylık bülten: {len(donemler)} dönem × {len(tablolar)} tablo × 10 grup indiriliyor ({para})...")
    kayitlar = []
    with ThreadPoolExecutor(paralel) as ex:
        fs = {ex.submit(aylik_donem, y, a, tablolar, para): (y, a) for y, a in donemler}
        for f in as_completed(fs):
            y, a = fs[f]
            try:
                kayitlar += f.result()
                print(f"  {y}-{a:02d} tamam")
            except RuntimeError as e:
                print(f"  {y}-{a:02d} ATLANDI: {e}")
    if kayitlar:
        aylik_kaydet(kayitlar, klasor / "aylik", para)


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
            except (urllib.error.URLError, OSError) as e:
                hata = str(e)
            time.sleep(3 * (k + 1))
            try:
                self.yenile()
            except (urllib.error.URLError, OSError, AttributeError):
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
            except (RuntimeError, AttributeError, IndexError):
                if k == 2:
                    raise RuntimeError(f"{y} {t['ad']}: tablo okunamadı")
                time.sleep(5)
                yerel.o = HaftalikOturum()

    kayitlar = []
    with ThreadPoolExecutor(paralel) as ex:
        fs = [ex.submit(is_yap, *i) for i in isler]
        for f in as_completed(fs):
            try:
                y, n, k = f.result()
                kayitlar += k
                print(f"  {y} · {tablolar[n]['ad']}: {len(k):,} satır")
            except RuntimeError as e:
                print(f"  ATLANDI: {e}")
    if kayitlar:
        haftalik_kaydet(kayitlar, klasor / "haftalik")


# =====================================================================================
def ay_oku(s):
    y, a = s.strip().split("-")
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
        b = ay_oku(input("Başlangıç (YYYY-AA, ör. 2018-01): "))
        e = ay_oku(input("Bitiş     (YYYY-AA, ör. 2026-09): "))
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
