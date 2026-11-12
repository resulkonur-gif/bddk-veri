# BDDK Verileri

BDDK'nın yayımladığı bankacılık sektörü verilerinin otomatik arşivi. Sitede "BDDK Verileri" sekmesinin kaynağıdır. TBB verilerinden ayrı tutulur.

## Aylık Bülten (`veri/aylik/`)

Kaynak: [BDDK İnteraktif Aylık Bülten](https://www.bddk.org.tr/BultenAylik), Excel dışa aktarımı (3 ondalık).

- **Dönem:** 2018/01'den bugüne (TFRS 9 geçişinden itibaren)
- **Tablolar:** 17 tablonun tamamı (Bilanço, Kar Zarar, Krediler, … , Yurt Dışı Şube Rasyoları)
- **Gruplar:** Sektör, Mevduat, Mevduat-Yerli Özel, Mevduat-Kamu, Mevduat-Yabancı, Katılım, Kalkınma ve Yatırım, Yerli Özel, Kamu, Yabancı
- **Para birimi:** TL (YP sütunları dönem sonu kurla TL karşılığı)

| Dosya | İçerik |
|---|---|
| `uzun/bddk_aylik_<yıl>.csv` | Uzun biçim: `donem, tablo_no, tablo, grup, sira, kalem, sutun, deger` (UTF-8 BOM) |
| `excel/BDDK_Aylik_Arsiv_<yıl>.xlsx` | Bilgi sayfası + her tablo için bir sayfa |
| `tablolar.json` | Tablo adları, birimleri, sütunları |
| `son_donem.txt` | Arşivdeki en son dönem |

### Güncelleme

`.github/workflows/bddk-aylik.yml` her gün çalışır. Yeni bir dönem yayımlandıysa onu ekler ve BDDK'nın geriye dönük revizyonlarını yakalamak için son 12 ayı yeniden çeker. Elle çalıştırırken başlangıç ve bitiş dönemi verilerek istenen aralık yeniden çekilebilir.

### Notlar

- Değerler nominaldir; enflasyon ve kur etkisi arındırılmamıştır.
- Birimler tabloya göre değişir: çoğu milyon TL, Sektörel Kredi Dağılımı bin TL, Rasyolar %, Diğer Bilgiler adet.
- Krediler kalemi bankalara kullandırılan kredileri, Mevduat kalemi banka mevduatını içermez.
