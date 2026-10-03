# Finansal Denetim ve Analiz Sistemi (SMMM Modülü)

Faturaları (UBL-TR XML / Excel) ve yevmiye kayıtlarını yerel bir SQLite veritabanına aktarıp
firmadaki muhasebe hatalarını raporlayan çevrimdışı masaüstü uygulaması.

## Kurulum ve çalıştırma

```bash
pip install -r requirements.txt
python app.py
```

Testler: `pip install pytest && python -m pytest`

## Önerilen akış

1. **Firma ve Veri Ayarları** → Firma VKN'sini girin (alıcı VKN kontrolü için).
2. **UBL-TR (XML) Yükle** → Tek tek, çoklu, ZIP ya da klasör olarak e-Fatura/e-Arşiv XML'leri.
3. **Fatura (Excel) Yükle** → XML'i olmayan faturalar için (şablon butondan indirilebilir).
4. **Yevmiye (Excel) Yükle** → Muhasebe kayıtları.
5. **Genel Denetim Raporu** → Tüm kontroller + çok sayfalı Excel raporu.

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
- V1.1 veritabanı (`audit_data.db`) ilk açılışta otomatik olarak yeni şemaya taşınır; eski tablolar
  `*_v1_yedek` adıyla saklanır.

## Proje yapısı

```
app.py                 Arayüz (customtkinter)
denetim/database.py    SQLite şeması, kayıt/okuma, eski sürüm taşıma
denetim/ubl.py         UBL-TR XML ayrıştırıcı
denetim/importers.py   Excel/XML doğrulama ve satır bazlı hata raporu, şablonlar
denetim/checks.py      Fiyat analizi, mutabakat ve tüm denetim kontrolleri
denetim/utils.py       Sayı/tarih/metin/birim normalizasyonu
tests/                 Birim testleri ve örnek XML
```
