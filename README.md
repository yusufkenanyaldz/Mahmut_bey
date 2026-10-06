# Finansal Denetim ve Analiz Sistemi (SMMM Modülü)

Faturaları (UBL-TR XML / Excel) ve yevmiye kayıtlarını firma başına ayrı bir yerel SQLite veritabanına
aktarıp firmadaki muhasebe hatalarını raporlayan çevrimdışı masaüstü uygulaması.

> Bu depoda ayrıca **KDV iadesi teminat çözüm raporu taslak programı** bulunur: [`teminat_cozum/`](teminat_cozum/README.md)
> (kendi bağımlılıkları ve testleri vardır; `cd teminat_cozum && python -m pytest`).

## Kurulum ve çalıştırma

```bash
pip install -r requirements.txt
python app.py
```

Testler: `pip install pytest && python -m pytest`

## Önerilen akış

1. **Firma Seç / Yönet** → Firmayı seçin ya da yeni firma oluşturun (VKN'yi girin; alıcı VKN kontrolü ve satış
   faturalarının ayırt edilmesi için).
2. **UBL-TR (XML) Yükle** → Tek tek, çoklu, ZIP ya da klasör olarak e-Fatura/e-Arşiv XML'leri (gelen ve giden).
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
- Analiz ayarları (sapma eşiği, hesap kodları, TL toleransı, dövizli faturalar için kur toleransı, dönem tipi,
  faturasız kontrolde hariç tutulan belge no önekleri, fiyat analizinin hariç tutma kuralları, kelime listesi ve en az
  alım sayısı, KDV / tevkifat / gelir / hesaplanan KDV hesapları ve maliyete eklenen vergi kodları) ve bulgu inceleme
  kayıtları firma başına saklanır.
- Firma seçilmeden veri yükleme / analiz / ayar ekranları açılmaz; firma seçme ekranına yönlendirilirsiniz.
- Son seçilen firma bir sonraki açılışta otomatik açılır.
- **Tüm Verileri Sil** (Firma ve Veri Ayarları) yalnızca aktif firmanın fatura ve yevmiye kayıtlarını siler; ayarlar,
  kayıtlı sütun eşlemeleri ve bulgu inceleme kayıtları kalır (veriler yeniden yüklendiğinde aynı bulgulara uygulanır).
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
isteğe bağlı: `Birim, Iskonto, KDV_Orani, Para_Birimi, Kur, OTV_Tutari, Fatura_Tipi, Tevkifat_Orani, Yon`.
`Fiyat` KDV hariç birim fiyattır. Aynı `Fatura_No + Tedarikci_VKN` satırları tek faturanın kalemleridir.

- `OTV_Tutari`: satırın ÖTV tutarı (belge para biriminde; KDV matrahına eklenir, 9077 vergi koduyla saklanır).
- `Fatura_Tipi`: `SATIS, IADE, TEVKIFAT, TEVKIFATIADE, ISTISNA, IHRACKAYITLI, OZELMATRAH` ya da `IHRACAT`
  (ihracat senaryosu, istisna faturası). Bilinmeyen değer faturayı düşürmez: uyarıyla SATIS kabul edilir.
- `Tevkifat_Orani`: `4/10`, `%40`, `40` ya da `0,4`; doluysa tevkifat = satır KDV'si × oran ve tip boşsa TEVKIFAT.
- `Yon`: `Alış` / `Satış` (`Gelen` / `Giden`). **Satış satırlarında `Tedarikci_VKN` / `Tedarikci_Ad` alanlarına karşı
  taraf (müşteri)** yazılır; satıcı olarak aktif firmanın VKN'si saklanır. Boşsa yön satıcı VKN'sinden bulunur.
Bir faturanın herhangi bir satırı hatalıysa fatura eksik kaydedilmesin diye tamamen atlanır.

**Yevmiye Excel** — zorunlu: `Tarih, Belge_No, Hesap_Kodu` ve `Borc + Alacak` *ya da* `Tutar`; isteğe bağlı `Aciklama`.

**Belge numarası mutlaka ayrı bir sütunda olmalıdır.** Program belge numarasını açıklama metninin içinden
ayıklamaz. Belge no sütunu olmayan bir döküm yüklendiğinde açık bir hata verilir; muhasebe programından belge
numarasının ayrı sütunda olduğu bir döküm alın (ör. dökümün sütunlarına **Evrak No / Belge No** ekleyerek).
`Fiş No` / `Yevmiye No` belge numarası sayılmaz (fiş sıra numarasıdır).

Sütun başlıklarında Türkçe karakter, boşluk, noktalama ve büyük/küçük harf farkı tolere edilir. Muhasebe
programlarının yaygın başlıkları tanınır, örneğin:

| Alan | Tanınan başlıklar (örnek) |
|---|---|
| Tarih | Tarih, Fiş Tarihi, Yevmiye Tarihi, Kayıt Tarihi, İşlem Tarihi, Evrak Tarihi |
| Belge_No | Belge No, Belge Numarası, Evrak No, Evrak Numarası, Fatura No |
| Hesap_Kodu | Hesap Kodu, Hesap No, Hesap, Muhasebe Hesap Kodu |
| Borc / Alacak | Borç, Borç Tutarı, Borç TL, Borç (TL) / Alacak, Alacak Tutarı, Alacak TL |
| Aciklama | Açıklama, Fiş Açıklaması, Satır Açıklaması |

Fatura Excel'inde de `Fatura No`, `Fatura Tarihi`, `Satıcı VKN/TCKN`, `Satıcı Unvanı`, `Mal/Hizmet`, `Birim Fiyatı`,
`KDV %` gibi başlıklar tanınır.

### İçe aktarma sihirbazı (yevmiye ve fatura Excel'i)

- **Başlık satırı otomatik bulunur:** ilk 20 satır taranır, bilinen sütun adlarıyla en çok eşleşen satır başlık
  kabul edilir. Üstteki firma adı / "YEVMİYE DEFTERİ DÖKÜMÜ" gibi satırlar ve boş satırlar atlanır.
- **Özet satırları veri sayılmaz:** "Toplam", "Genel Toplam", "Ara Toplam", "Nakli Yekün", "Devreden" gibi
  tarihsiz satırlar ve sayfa sonlarında tekrarlanan başlık satırları atlanır; günlükte **[BİLGİ]** olarak
  hangi satırların atlandığı yazılır.
- **Sütun eşleme penceresi:** zorunlu sütunlar bulunamazsa (ya da **🧭 Sütunları Eşle** butonuyla) açılır.
  Algılanan başlık satırı (değiştirilebilir), ilk 10 satırın önizlemesi ve her alan için dosyadaki sütunu
  seçtiren listeler gösterilir. Bir sütun yalnızca bir alana eşlenebilir.
- **Eşleme hatırlanır:** onaylanan eşleme aktif firmanın veritabanında dosyanın sütun başlıkları imzasıyla
  saklanır. Aynı firmadan aynı biçimde (aynı sütun başlıklarıyla) gelen dosyada pencere açılmadan uygulanır ve
  günlüğe "Kayıtlı eşleme kullanıldı" yazılır. **Tüm Verileri Sil** kayıtlı eşlemeleri silmez.

## Kontroller

| Kontrol | Açıklama |
|---|---|
| Fiyat anomalileri | Dönem (aylık/çeyreklik/yıllık) + ürün + birim + para birimi bazında ağırlıklı ortalama birim fiyattan, kullanıcının belirlediği % eşiği aşan sapmalar (belge para biriminde, KDV hariç; iadeler, tevkifatlı faturalar ve hizmet / hakediş kalemleri hariç; aşağıya bakın) |
| Muhasebeleşmemiş faturalar | Yevmiyede hiçbir yöntemle (aşağıya bakın) karşılığı bulunamayan faturalar |
| Seçili hesap dışına kaydedilmiş | Belge no (tam ya da seri+sıra) yevmiyede var ama girilen hesap kodlarında değil (yanlış hesap şüphesi) |
| Tutar farkları | Fatura KDV hariç tutarı ile seçili hesaplardaki yevmiye toplamı arasındaki tolerans üstü farklar (TL faturada sabit TL toleransı, dövizli faturada ayrıca yüzde kur toleransı; aşağıya bakın) |
| Dönem farkları | Fatura dönemi ile yevmiye kayıt dönemi farklı |
| Belge no uyuşmayan eşleşmeler | Belge no tutmadığı için tutar + tarih ile eşleştirilen faturalar (belge no yazım hatası; kontrol edin) |
| Faturası bulunmayan yevmiye kayıtları | Seçili hesaplarda olup hiçbir yöntemle bir faturayla eşleşmeyen, net borç yönlü ve hariç önekle başlamayan belgeler (faturasız gider/alış adayları; öncelik sütunlu, aşağıya bakın) |
| Belirsiz eşleşme (aynı no) | Aynı fatura numarası birden fazla tedarikçide (yevmiyede VKN olmadığından eşleşme belirsiz) |
| Belirsiz eşleşme (seri+sıra) | Kısaltılmış belge no birden fazla faturaya ya da fatura birden fazla belgeye uyuyor |
| Olası mükerrer faturalar | Aynı tedarikçi, aynı tarih, aynı tutar, farklı numara (`Ardisik_Numara`: aynı serinin ardışık numaralarıysa "Evet"; bilgi amaçlı, ardışık olmayanlar önce sıralanır) |
| Fatura hesaplama tutarsızlıkları | XML'de satır toplamları / iskonto / KDV ile belge toplamlarının uyuşmaması |
| Alıcısı firma olmayan faturalar | XML'deki alıcı VKN'si firma VKN'sinden farklı (yalnızca alış faturaları) |
| KDV farkları | Alış faturasının KDV'si ile aynı belgenin KDV hesaplarındaki (191) kaydı farklı (olası neden sütunlu) |
| KDV'si kaydedilmemiş faturalar | Faturada KDV var, belge yevmiyede var ama KDV hesabında kayıt yok (ör. KDV maliyete eklenmiş) |
| Tevkifat kaydı eksik/farklı | Tevkifatlı alış faturasında tevkif edilen KDV tevkifat hesabına (360) alacak yazılmamış / yanlış tutar / ters yön |
| Satış: muhasebeleşmemiş, gelir hesabı dışı, tutar, dönem, KDV farkları | Satış faturaları gelir (600–602) ve hesaplanan KDV (391) hesaplarıyla; aşağıya bakın |
| Faturası bulunmayan gelir kayıtları | Gelir hesaplarında net alacak yönlü, hariç önekle başlamayan, hiçbir satış faturasıyla eşleşmeyen belgeler |

Fiyat analizi, alış mutabakatı, mükerrer fatura ve alıcı VKN kontrolleri yalnızca **alış** faturalarına uygulanır;
hesaplama tutarsızlığı kontrolü tüm XML faturalara uygulanır.

### Fiyat anomalileri

Her satırın birim fiyatı (KDV hariç, iskonto sonrası) aynı dönemdeki aynı **ürün + birim + para birimi** grubunun
ağırlıklı ortalama birim fiyatıyla (AOBF) karşılaştırılır; sapma eşiği (%) aşılırsa satır **YÜKSEK RİSK** olur.

- **Para birimine göre ayrı grup:** EUR ile alınan motorin TL motorinle aynı grupta karşılaştırılmaz. Birim fiyat
  ve AOBF belge para birimindedir, sapma belge para birimindeki fiyat üzerinden hesaplanır (kur değişimi sapma
  yaratmaz). `Birim_Fiyat_TL` (fatura kuruyla) ve `Kur` bilgi için gösterilir.
- **Analize alınmayan satırlar** (sırayla):
  1. **İade** faturaları (`IADE`, `TEVKIFATIADE`) ve miktarı 0 / boş olan satırlar — her zaman.
  2. **Tevkifatlı faturalar** (`InvoiceTypeCode` `TEVKIFAT`): taşeron hakedişi, fason hizmet gibi tevkifata tabi
     işler her ay farklı tutarda faturalanır, birim fiyat karşılaştırması anlamsızdır. Excel'den yüklenen
     faturalarda fatura tipi bilinmediğinden bu kural yalnızca XML faturalarda işler; Excel faturaları için
     anahtar kelime listesi kullanılır.
  3. **Anahtar kelime:** ürün adında listedeki kelimelerden biri **geçen** satırlar (büyük/küçük harf ve Türkçe
     karakter duyarsız: `işçilik` = `İŞÇİLİK` = `ISCILIK`). Varsayılan:
     `HAKEDİŞ, İŞÇİLİK, HİZMET, FASON, KİRALAMA, BAKIM, ONARIM, DANIŞMANLIK`. Kelime ürün adının herhangi bir
     yerinde geçerse eşleşir (`HİZMET` → "Nakliye Hizmeti"); kısa / genel kelimeler eklerken bunu dikkate alın.
  Her iki kural ayrı ayrı açılıp kapatılabilir; kelime listesi boş bırakılırsa kelime süzgeci uygulanmaz,
  **Varsayılan** butonu listeyi varsayılana döndürür.
- **Yetersiz veri:** grupta dönem içinde **en az alım** sayısından (varsayılan 3) az alım varsa sapma eşiği aşsa da
  satır riskli sayılmaz, `Risk_Durumu` = "Yetersiz veri (bilgi)" olarak raporda kalır. İki alımlı bir grupta
  iki fiyat ortalamadan simetrik saptığı için hangisinin hatalı olduğu söylenemez (tek alımda sapma zaten 0'dır).
  `1` girilirse kural kapanır.
- Kurallar, kelime listesi ve en az alım sayısı **Fiyat Risk Analizi** ve **Genel Denetim Raporu** ekranlarından
  düzenlenir ve firma başına saklanır.
- Ekranın ve raporun başında kaç satırın incelendiğini ve kaçının neden analiz dışı kaldığını (iade, tevkifat,
  anahtar kelime, yetersiz veri) gösteren özet satırı yer alır; Excel'de **Fiyat Analizi Özeti** sayfası olarak
  da çıkar. **Analiz dışı satırları Excel'e ekle** seçiliyse (varsayılan) hariç tutulan satırlar ve yetersiz veri
  nedeniyle riskli sayılmayanlar nedenleriyle birlikte **Fiyat Analizi Dışı Satırlar** bilgi sayfasına yazılır.

**Mutabakat notu:** Hesap kodları önek olarak eşleşir (`153` → `153.01`, `153.02.001` …).
Aynı belgenin karşı hesaplarını (ör. `153` ile `320`) birlikte girmeyin; Borç − Alacak toplamı sıfırlanır.

**Fatura no ↔ belge no eşleştirmesi** sırayla üç kademede yapılır; sonuçlarda `Eslesme_Yontemi` sütunu
hangi kademenin kullanıldığını gösterir:

1. **Tam** — boşluklar atılmış, büyük harfe çevrilmiş fatura no ile belge no aynı.
2. **Seri+Sıra** — GİB fatura numarası 3 karakter seri + 4 hane yıl + 9 hane sıra numarasıdır
   (`ABC2024000000123`). Ayraçlar (boşluk, `-`, `/`, `.`, `_`) temizlenip seri, yıl (yazılmışsa) ve sıra
   numarası (tamsayı) çıkarılır; seri ve sıra aynıysa, yıl iki tarafta da yazılmışsa o da aynıysa eşleşir.
   Böylece `ABC123`, `ABC-123`, `ABC 2024 123`, `ABC2024123` yazımları `ABC2024000000123` ile eşleşir.
   Yalnızca tek aday varsa eşleştirilir; birden fazla aday (ör. yılsız `ABC123` hem 2023 hem 2024 faturasına
   uyuyorsa) **Belirsiz eşleşme (seri+sıra)** bölümüne yazılır.
3. **Tutar+Tarih** (düşük güven) — belge numarası yevmiyenin hiçbir yerinde bulunamayan faturalar için:
   seçili hesaplarda eşleşmemiş belgelerden aynı tutarı (tolerans dahilinde) taşıyan ve tarihi faturadan en çok
   ±15 gün uzakta olan **tek** aday varsa eşleştirilir. Bu eşleşmeler ayrıca **Belge no uyuşmayan eşleşmeler**
   bölümünde listelenir; belge numarasının yanlış yazılması da bir bulgudur.

Seçili hesap dışı kontrolü 3. kademeden önce, tam ve seri+sıra eşleşmesiyle yapılır (belge no başka hesapta
bulunan fatura tutar+tarih ile seçili hesaptaki başka bir kayda bağlanmaz). Herhangi bir kademede eşleşen ya da
belirsiz adayı olan belge "faturası bulunmayan" sayılmaz. Mutabakat ekranının başında ve genel raporda kaç
faturanın hangi yöntemle eşleştiğini gösteren özet satırı yer alır; Excel raporunda **Eşleşme Özeti** sayfası
olarak da çıkar.

### Faturası bulunmayan yevmiye kayıtları

Eşleşmeyen her belge listelenmez; faturaya zaten dayanmayan olağan kayıtlar (bordro, amortisman, satılan malın
maliyeti, gider pusulası, mahsup fişleri …) gerçek sahte / eksik belgeli giderleri kalabalığın içinde
kaybettirmesin diye iki süzgeç uygulanır:

1. **Yalnızca borç tarafı:** belgenin seçili hesaplardaki net tutarı (Borç − Alacak) borç yönünde ve tolerans
   üstünde olmalıdır. Net alacak yönlü (ör. `153` alacak — SMM, stoktan çıkış) ya da sıfır netli belgeler
   listeye alınmaz.
   - **İşaretsiz yevmiye:** yevmiyede hiç negatif tutar yoksa (tek `Tutar` sütunuyla, alacaklar da pozitif
     yazılarak yüklenmişse) borç/alacak yönü bilinemez; bu süzgeç **uygulanmaz** ve ekranda / raporda
     "Yevmiye tutarları işaretsiz … yön filtresi uygulanmadı" notu çıkar. Çift taraflı Borç/Alacak dökümünde
     alacak satırları negatif olduğundan bu durum kendiliğinden ayırt edilir.
2. **Hariç tutulan belge önekleri:** belge no'su listedeki öneklerden biriyle başlayan kayıtlar listeye alınmaz.
   Liste firma başına saklanır, virgülle ayrılır ve mutabakat, genel rapor ve Firma ve Veri Ayarları ekranlarından
   düzenlenir (**Varsayılan** butonu varsayılana döndürür; boş bırakılırsa önek süzgeci uygulanmaz).
   Varsayılan: `BORDRO, AMORT, MAHSUP, AÇILIŞ, KAPANIŞ, DEVİR, GP` (GP = gider pusulası).
   Karşılaştırma büyük/küçük harf, Türkçe karakter ve ayraç duyarsızdır (`açılış`, `ACILIS-01`, `Açılış Fişi`
   hepsi `AÇILIŞ` önekine uyar; `GP-07001` ve `gp 12` → `GP`). Tam GİB biçimli numaralar
   (`GPS2024000000123` gibi) gerçek fatura numarası olduğundan hiçbir önekle hariç tutulmaz.

Kalan kayıtlara **Oncelik** sütunu eklenir: belge no'su GİB fatura numarasına benzeyenler (3 karakter seri +
4 hane yıl + 9 hane sıra ya da seri+sıra çözümlemesine uyan `ABC-2024-123`, `ABC123` gibi yazımlar) **Yüksek**,
diğerleri **Düşük**. Liste önce önceliğe, sonra tutara (büyükten küçüğe) göre sıralanır; düşük öncelikliler de
listede kalır.

Ekranda ve raporda kaç kaydın listelendiğini ve kaçının neden (alacak yönlü, hariç önek) listeye alınmadığını
gösteren bir özet satırı yer alır; Excel raporunda **Faturasız Kayıt Özeti** sayfası olarak da çıkar.
**Hariç tutulanları Excel'e ekle** seçiliyse listeye alınmayan kayıtlar nedenleriyle birlikte
**Faturasız Listeden Hariç Tutulanlar** bilgi sayfasına yazılır.

### Tutar farkı toleransı (TL ve dövizli faturalar)

- **TL faturalar:** fark (|yevmiye| − |fatura KDV hariç TL|) **Tolerans (TL)** değerini (varsayılan 0,01) aşarsa
  tutar farkıdır. TL faturalara yüzde tolerans **uygulanmaz**: rakam yer değiştirme hataları
  (12.345,67 → 12.354,67) yalnızca birkaç TL fark yaratır ve yüzde tolerans bunları gizlerdi.
- **Dövizli (TRY dışı) faturalar:** muhasebe çoğu zaman fatura kuru yerine ödeme / kayıt gününün kurunu kullanır.
  Fark ≤ max(TL toleransı, **Kur Toleransı (%)** × fatura TL tutarı) ise tutar farkı sayılmaz; bu faturalar
  ekranda ve raporda "Kur farkı (tolerans içi, bilgi)" sayısıyla özetlenir ve Excel'de
  **Kur Farkı (Tolerans İçi)** bilgi sayfasında (fark ve fark yüzdesiyle) listelenir. Toleransı aşan dövizli farklar
  **Tutar Farkları**'nda kalır; bölümde `Dovizli`, `Para_Birimi` ve `Kur` sütunları gösterilir.
- Kur toleransı firma başına saklanır (varsayılan **%1**; `0` → yüzde tolerans kapalı) ve **Muhasebe Mutabakatı**
  ile **Genel Denetim Raporu** ekranlarından düzenlenir.

### Maliyete eklenen vergiler (ÖTV vb.)

UBL-TR XML'de `TaxTotal` altındaki KDV (0015) dışındaki vergiler vergi türü koduyla ayrıştırılıp fatura başına
saklanır (belge toplamında yoksa satırlardan toplanır). Bayi araç alımında olduğu gibi indirilemeyen vergiler maliyete
eklenir; bu yüzden alış mutabakatında **beklenen maliyet = KDV hariç tutar + maliyete eklenen vergiler** (TL'ye fatura
kuruyla çevrilir) olur. Tutar Farkları'nda `KDV_Haric_Tutar_TL`, `Maliyete_Eklenen_Vergi_TL`, `Beklenen_Tutar_TL` ve
`Olasi_Neden` (ör. "Maliyete eklenecek vergi (ÖTV vb.) maliyete eklenmemiş", "KDV maliyete eklenmiş", "Çift kayıt")
sütunları gösterilir.

Maliyete eklenen vergi kodları firma başına düzenlenir (**Varsayılan** butonu; boş → hiçbir vergi eklenmez).
Varsayılan: `0071, 0073–0077` ÖTV (I, III, IV), `9077` ÖTV II (motorlu taşıtlar), `4080, 4081` ÖİV, `0059` konaklama
vergisi, `0021` BSMV, `4071, 8005` elektrik / havagazı tüketim vergisi, `8004` TRT payı, `8008` çevre temizlik vergisi.
Stopajlar (0003, 0011) ve KDV tevkifatı maliyet değildir.

### KDV (191) ve tevkifat (360) mutabakatı

- Alış faturasının KDV'si (TL), mutabakatın bulduğu **aynı belge** (tam, seri+sıra, tutar+tarih ya da seçili hesap dışı
  eşleşmesi) üzerinden KDV hesaplarındaki (varsayılan `191`) toplamla karşılaştırılır; muhasebeleşmemiş ve belirsiz
  faturalar bu kontrole girmez. Dövizli faturada kur toleransı uygulanır.
  - **KDV'si Kaydedilmemiş Faturalar:** faturada KDV var, belge yevmiyede var ama KDV hesabında satır yok.
    `Olasi_Neden`: "KDV maliyete eklenmiş" (seçili hesaplardaki fark KDV kadar) ya da "KDV hesabında kayıt yok";
    `Kullanilan_Hesaplar` belgenin yevmiyede geçtiği hesaplardır.
  - **KDV Farkları:** KDV hesabındaki tutar farklı. `Olasi_Neden`: çift kayıt, faturada KDV yok, "191'e tevkifat
    düşülmüş KDV yazılmış", ters yönde kayıt (yalnızca XML faturalarda; Excel'de iade tipi çoğu zaman bilinmez).
- **Tevkifatlı faturalarda 191'e TAM KDV** yazılması doğrudur (indirilecek KDV); tevkif edilen kısım
  (`WithholdingTaxTotal`, Excel'de `Tevkifat_Orani`) tevkifat hesaplarına (varsayılan `360`) **alacak** olarak
  yazılmalıdır. **Tevkifat Kaydı Eksik/Farklı**: `Durum` = "Tevkifat kaydı yok", "Borç yönlü (ters) kayıt" ya da
  "Tutar farklı". İade faturalarında yönler terstir.
- Hesap kodu kutusu boş bırakılırsa ilgili kontrol yapılmaz. 360'ta stopaj gibi başka vergiler de izleniyorsa KDV
  tevkifatı alt hesabını girin (ör. `360.02`).

### Satış faturaları

- **Yön:** Excel'de `Yon` verilmişse o; yoksa satıcı VKN'si aktif firmanın VKN'si olan faturalar **satış** (giden
  e-fatura / e-arşiv), diğerleri alış. Firmanın kestiği **iade** faturaları alış iadesidir, alış sayılır. Firma VKN'si
  girilmemişse XML satış faturaları ayırt edilemez (not düşülür).
- **Satış mutabakatı** alış mutabakatıyla aynı eşleştirmeyi kullanır; gelir hesapları (varsayılan `600, 601, 602`) ile
  KDV hariç tutar, hesaplanan KDV hesapları (varsayılan `391`) ile KDV karşılaştırılır. Satışta gelir **alacak**
  tarafındadır. Karşı taraf sütunları `Musteri` / `Musteri_VKN`'dir.
- **İhracat / istisna** (`ProfileID` IHRACAT, `InvoiceTypeCode` ISTISNA ya da IHRACKAYITLI) faturalarında KDV beklenmez;
  bu faturalara 391 kaydı yazılmışsa bulgudur. Tevkifatlı satışta 391'e KDV − alıcının tevkif ettiği KDV beklenir.
- Bölümler: **Muhasebeleşmemiş Satış Faturaları**, **Gelir Hesabı Dışına Kaydedilmiş Satış Faturaları**, **Satış Tutar
  Farkları**, **Satış Dönem Farkları**, **Satış KDV Farkları** (`Durum`: hesaplanan KDV kaydı yok / istisna faturasında
  KDV kaydı var / tutar farklı / ters yön), **Faturası Bulunmayan Gelir Kayıtları** (net alacak yönlü; hariç önek
  listesi ve öncelik alış tarafıyla aynı) ve bilgi amaçlı belge no uyuşmayan / belirsiz satış eşleşmeleri. Raporda
  **Satış Eşleşme Özeti**, **Faturasız Gelir Özeti**, **Satış Kur Farkı (Tolerans İçi)** ve (seçiliyse) **Faturasız
  Gelir Listesinden Hariç Tutulanlar** sayfaları yer alır.
- Gelir hesabı kutusu boşsa satış mutabakatı yapılmaz. Müşterinin kestiği iade (satıştan iade) faturası alış tarafında
  değerlendirilir; 610'a kaydedildiği için "seçili hesap dışı" görünebilir.

## Bulgu inceleme (denetçi iş akışı)

Fiyat Risk Analizi, Muhasebe Mutabakatı ve Genel Denetim Raporu sonuçları **Bulgular** sekmesinde satır seçilebilen
bir tabloda gösterilir (metin özeti **Özet (metin)** sekmesindedir). Akış:

1. **Kontrol** listesinden bölümü seçin (listede her kontrolün açık / sorun yok / düzeltme istendi sayıları yazar).
2. Satır(lar)ı seçin (Ctrl / Shift ile çoklu; **Tümünü Seç** ya da Ctrl+A), isterseniz **Not** yazın.
3. **✔ İncelendi – Sorun Yok**, **✎ Düzeltme İstendi** ya da **↺ Açığa Al**. Not boş bırakılırsa mevcut not korunur;
   tek satır seçildiğinde notu kutuya gelir.

- Her bulgunun durumu (Açık / İncelendi – Sorun Yok / Düzeltme İstendi), notu ve tarihi firmanın veritabanında
  kalıcıdır. **Sorun yok olanları göster** kapalıyken (varsayılan) "sorun yok" işaretli bulgular tablodan gizlenir;
  özet sayılarında görünmeye devam eder.
- **Bulgu kimliği kararlıdır:** kontrol türü + normalize tedarikçi VKN (satış bölümlerinde müşteri VKN) + normalize
  fatura no (faturasız kayıtta / faturasız gelirde
  belge no, mükerrer grubunda sıralı fatura no listesi, fiyat anomalisinde ayrıca ürün adı, hesaplama kontrolünde
  kontrol adı). Tutar, eşik, tolerans, satır sırası ya da veritabanı kimliği anahtara girmez; bu yüzden rapor yeniden
  çalıştırıldığında, veriler silinip yeniden yüklendiğinde ya da kullanılan muhasebe dökümü değiştiğinde işaretler
  korunur. Aynı fatura farklı kontrollerde (ör. tutar farkı ve dönem farkı) ayrı bulgudur.
- Excel raporunda bulgu sayfalarına `Inceleme_Durumu`, `Inceleme_Notu`, `Inceleme_Tarihi` ve `Bulgu_Anahtari`
  sütunları, başa **İnceleme Özeti** sayfası (kontrol başına açık / sorun yok / düzeltme istendi) eklenir; Excel'de
  "sorun yok" bulgular gizlenmez, durum sütunuyla süzülebilir.
- Yalnızca bilgi amaçlı listeler (yetersiz veri, analiz dışı satırlar, faturasız listeden hariç tutulanlar,
  tolerans içi kur farkları) işaretlenmez.

## Veri güvenliği

- Aynı dosya ikinci kez seçilirse uyarı verilir (yevmiyede tutarların çift sayılmasını önler).
- Aynı `Tedarikçi VKN + Fatura No` ikinci kez kaydedilmez.
- Aynı dosya farklı firmalara ayrı ayrı yüklenebilir (mükerrer dosya kontrolü firma bazındadır).
- V1.1 veritabanı (`audit_data.db`) aktarılırken kopyası yeni şemaya taşınır; eski tablolar kopyada
  `*_v1_yedek` adıyla saklanır.
- Şema sürümü 3 (madde 7): önceki sürümlerle oluşturulmuş firma veritabanları açılışta otomatik güncellenir
  (`invoice_taxes` tablosu ve faturalara tevkifat / yön / vergi sütunları eklenir; veriler korunur). Önceki sürümle
  yüklenmiş XML faturalarda ÖTV ve tevkifat bilgisi olmadığından genel raporda not düşülür; tam kontrol için verileri
  silip XML'leri yeniden yükleyin.

## Proje yapısı

```
app.py                 Arayüz (customtkinter), sütun eşleme penceresi, bulgu tablosu (ttk.Treeview)
denetim/firms.py       Firma kayıt defteri (firma başına veritabanı, son firma, eski audit_data.db aktarımı)
denetim/database.py    SQLite şeması, kayıt/okuma, eski sürüm taşıma
denetim/export.py      Çok sayfalı Excel rapor çıktısı
denetim/ubl.py         UBL-TR XML ayrıştırıcı
denetim/importers.py   Excel/XML doğrulama, başlık satırı tespiti, sütun eşleme, satır bazlı hata raporu, şablonlar
denetim/checks.py      Fiyat analizi, alış / satış mutabakatı, KDV / tevkifat / ÖTV ve tüm denetim kontrolleri
denetim/inceleme.py    Bulgu inceleme: kararlı bulgu anahtarı, durum/not kaydı, rapora uygulama ve özet
denetim/utils.py       Sayı/tarih/metin/birim normalizasyonu
tests/                 Birim testleri ve örnek XML
```
