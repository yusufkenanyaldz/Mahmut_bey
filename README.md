# Finansal Denetim ve Analiz Sistemi (SMMM Modülü)

Faturaları (UBL-TR XML / Excel) ve yevmiye kayıtlarını firma başına ayrı bir yerel SQLite veritabanına
aktarıp firmadaki muhasebe hatalarını raporlayan çevrimdışı masaüstü uygulaması.

## Kurulum ve çalıştırma

```bash
pip install -r requirements.txt
python app.py
```

Testler: `pip install pytest && python -m pytest`

## Önerilen akış

1. **Firma Seç / Yönet** → Firmayı seçin ya da yeni firma oluşturun (VKN'yi girin; alıcı VKN kontrolü için).
2. **UBL-TR (XML) Yükle** → Tek tek, çoklu, ZIP ya da klasör olarak e-Fatura/e-Arşiv XML'leri.
3. **Fatura (Excel) Yükle** → XML'i olmayan faturalar için (şablon butondan indirilebilir).
4. **Yevmiye (Excel) Yükle** → Muhasebe kayıtları.
5. **Genel Denetim Raporu** → Tüm kontroller + çok sayfalı Excel raporu.

## Firma yönetimi

Her firmanın verileri (faturalar, yevmiye, yüklenen dosya kayıtları ve analiz ayarları) ayrı bir veritabanında
tutulur; firma değiştirmek için veri silmek ya da dosya kopyalamak gerekmez.

- Yan menünün üstünde **Aktif Firma** kartı vardır; pencere başlığında da aktif firmanın unvanı yazar.
  **Firma Seç / Yönet** ile firma listesi açılır: **Seç**, **Düzenle**, **Sil** ve **➕ Yeni Firma**.
- Firma kaydı: **kod / kısa ad** (benzersiz, büyük/küçük harf duyarsız), **unvan**, **VKN/TCKN** (isteğe bağlı,
  10 ya da 11 hane), **sektör** (isteğe bağlı) ve oluşturulma tarihi. Alıcı VKN kontrolü aktif firmanın
  VKN'sini kullanır. Kod sonradan değiştirilebilir; firmanın veritabanı dosyası değişmez.
- Analiz ayarları (sapma eşiği, hesap kodları, tolerans, dönem tipi) firma başına saklanır.
- Firma seçilmeden veri yükleme / analiz / ayar ekranları açılmaz; firma seçme ekranına yönlendirilirsiniz.
- Son seçilen firma bir sonraki açılışta otomatik açılır.
- **Tüm Verileri Sil** (Firma ve Veri Ayarları) yalnızca aktif firmanın fatura ve yevmiye kayıtlarını siler.
- **Sil** (firma listesi) firmayı ve veritabanı dosyasını onay sonrası **kalıcı olarak** siler; geri alınamaz.

Dosyalar uygulama klasöründeki `veri/` altında tutulur:

```
veri/firmalar.db     Firma listesi ve son seçilen firma
veri/<KOD>.db        Firma başına denetim veritabanı (yedeklemek için bu dosyayı kopyalamanız yeterli)
```

**Eski sürümden geçiş:** Uygulama klasöründe (ya da programın çalıştırıldığı klasörde) önceki sürümün
`audit_data.db` dosyası varsa ilk açılışta bir kez `AKTARILAN` kodlu firma olarak `veri/` altına kopyalanır.
Unvan eski ayarlardan alınır; yoksa sorulur, boş bırakılırsa "Aktarılan Firma" kullanılır (sonradan
düzenlenebilir). Eski dosya silinmez ve değiştirilmez.

## Şablonlar

**Fatura Excel** — zorunlu: `Fatura_No, Tarih, Tedarikci_VKN, Tedarikci_Ad, Urun_Adi, Miktar, Fiyat`
isteğe bağlı: `Birim, Iskonto, KDV_Orani, Para_Birimi, Kur`.
`Fiyat` KDV hariç birim fiyattır. Aynı `Fatura_No + Tedarikci_VKN` satırları tek faturanın kalemleridir.
Bir faturanın herhangi bir satırı hatalıysa fatura eksik kaydedilmesin diye tamamen atlanır.

**Yevmiye Excel** — zorunlu: `Tarih, Belge_No, Hesap_Kodu` ve `Borc + Alacak` *ya da* `Tutar`; isteğe bağlı `Aciklama`.

Sütun başlıklarında Türkçe karakter, boşluk ve büyük/küçük harf farkı tolere edilir
(ör. `Tedarikçi VKN`, `Borç`, `Hesap Kodu`).

## Kontroller

| Kontrol | Açıklama |
|---|---|
| Fiyat anomalileri | Dönem (aylık/çeyreklik/yıllık) + ürün + birim bazında ağırlıklı ortalama birim fiyattan, kullanıcının belirlediği % eşiği aşan sapmalar (TL, KDV hariç, iadeler hariç) |
| Muhasebeleşmemiş faturalar | Belge numarası yevmiyede hiç geçmeyen faturalar |
| Seçili hesap dışına kaydedilmiş | Yevmiyede var ama girilen hesap kodlarında değil (yanlış hesap şüphesi) |
| Tutar farkları | Fatura KDV hariç tutarı ile seçili hesaplardaki yevmiye toplamı arasındaki tolerans üstü farklar |
| Dönem farkları | Fatura dönemi ile yevmiye kayıt dönemi farklı |
| Faturası bulunmayan yevmiye kayıtları | Seçili hesaplarda olup sistemde faturası olmayan belgeler |
| Belirsiz eşleşme | Aynı fatura numarası birden fazla tedarikçide (yevmiyede VKN olmadığından eşleşme belirsiz) |
| Olası mükerrer faturalar | Aynı tedarikçi, aynı tarih, aynı tutar, farklı numara |
| Fatura hesaplama tutarsızlıkları | XML'de satır toplamları / iskonto / KDV ile belge toplamlarının uyuşmaması |
| Alıcısı firma olmayan faturalar | XML'deki alıcı VKN'si firma VKN'sinden farklı |

**Mutabakat notu:** Hesap kodları önek olarak eşleşir (`153` → `153.01`, `153.02.001` …).
Aynı belgenin karşı hesaplarını (ör. `153` ile `320`) birlikte girmeyin; Borç − Alacak toplamı sıfırlanır.

## Veri güvenliği

- Aynı dosya ikinci kez seçilirse uyarı verilir (yevmiyede tutarların çift sayılmasını önler).
- Aynı `Tedarikçi VKN + Fatura No` ikinci kez kaydedilmez.
- Aynı dosya farklı firmalara ayrı ayrı yüklenebilir (mükerrer dosya kontrolü firma bazındadır).
- V1.1 veritabanı (`audit_data.db`) aktarılırken kopyası yeni şemaya taşınır; eski tablolar kopyada
  `*_v1_yedek` adıyla saklanır.

## Proje yapısı

```
app.py                 Arayüz (customtkinter)
denetim/firms.py       Firma kayıt defteri (firma başına veritabanı, son firma, eski audit_data.db aktarımı)
denetim/database.py    SQLite şeması, kayıt/okuma, eski sürüm taşıma
denetim/export.py      Çok sayfalı Excel rapor çıktısı
denetim/ubl.py         UBL-TR XML ayrıştırıcı
denetim/importers.py   Excel/XML doğrulama ve satır bazlı hata raporu, şablonlar
denetim/checks.py      Fiyat analizi, mutabakat ve tüm denetim kontrolleri
denetim/utils.py       Sayı/tarih/metin/birim normalizasyonu
tests/                 Birim testleri ve örnek XML
```
