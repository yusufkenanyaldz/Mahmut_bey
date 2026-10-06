# CLAUDE.md — Teminat Çözüm Raporu Projesi

Bu klasör, bir YMM ofisinin aylık KDV iadesi teminat çözüm raporlarının **taslağını** üreten Python programıdır.
Ayrıntılı gereksinimler `PROJE_TALIMATI.md` dosyasındadır; mükellef bilgisi içerdiği için **depoda değildir**
(ofis bilgisayarında bu klasöre konur, `.gitignore`'da). Dosya yoksa kullanıcıdan isteyin.

## Çalışma kuralları
- Kullanıcıyla **Türkçe** konuş. Kullanıcı rapor hesaplamalarını bilmiyor; muhasebe kuralı gereken sorularda
  "ofise sorulacak soru" olarak listele, tahminle kural uydurma.
- **Orijinal mükellef dosyalarını asla değiştirme, silme, üzerine yazma.** Çıktılar `CLAUDE TASLAK\` klasörüne, yeni adla.
- **Mükellef verisi depoya girmez** (depo herkese açık): firma ayarları (`firmalar/<firma>/`), girdi dosyaları,
  raporlar, `PROJE_TALIMATI.md`. Testler yalnızca `tests/sentetik.py`'deki uydurma verilerle yazılır;
  commit'ten önce firma / tedarikçi adı, VKN, rapor sayısı sızmadığını kontrol et.
- Hesaplar kesin kurallı Python kodunda yapılır. Bilinmeyen bilgi `[DOLDURULACAK]` + sarı; takdir gerektiren her
  seçim ve her tutarsızlık kontrol listesine (`KontrolListesi`) yazılır — sessizce geçilmez.
- Her değişiklikten sonra: `python -m pytest`; gerçek veri varsa `python -m teminat geriye-donuk ...` ile ofisin
  gerçek raporlarıyla karşılaştır. Tutar farkı kalmamalı; metin farkları açıklanmalı.
- Türkçe dosya adları: karşılaştırmalar `ortak.katla()` ile (NFC/NFD, Türkçe karakter, büyük/küçük harf duyarsız);
  yollar için `pathlib`. Windows'ta çalışmalı.
- `.doc` → `.docx`: `donusum.docx_hazirla()` (LibreOffice, ayrı profil ile).
- Sayı biçimi Türkçe: `1.234.567,89` (`ortak.tr()`); yüzde `% 23,36` / `%96,49` (tablodaki mevcut biçime uy).

## Kod
- `teminat/okuyucular/` — KDV 1 PDF, indirilecek KDV listesi, takip listesi, yüklenilen tutanak, teminat dilekçesi
- `teminat/rapor/olustur.py` — önceki ay raporundan yeni ay taslağı (adım adım; her adımın hatası listeye yazılır)
- `teminat/rapor/tablolar.py` — tabloları içerik/başlıktan bulma (index kullanılmaz)
- `teminat/rapor/docx_araclari.py` — `grubu_esitle`: bir bölümdeki satırları beyandaki kalemlere eşitler
  (eşleşen satır korunur, yeni satır kopyalanır ve sarı yapılır, fazlası silinir)
- `teminat/rapor/safha.py` — 3-4-3 EKLİ seçimi (Word'den bağımsız)
- `teminat/kontroller/` — kontrol listesi, sayısal ve belge kontrolleri, xlsx/html çıktısı
- `teminat/karsilastir.py`, `teminat/geriye_donuk.py` — öğrenme döngüsü
- Komutlar: `python -m teminat {taslak,karsilastir,denetle,geriye-donuk} --help`

## Yol haritası
`PROJE_TALIMATI.md` → "8. Yapılacaklar". Aşama 1'in kod kısmı tamam; kalan: gerçek verilerle geriye dönük test
(2025 Ocak – 2026 Şubat) ve Mart-2026 canlı deneme.
