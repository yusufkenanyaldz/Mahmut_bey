# Denetçi Simülasyonu (40 firma)

Programın gerçekçi verilerde ne kadar hata yakaladığını ve ne kadar yanlış alarm verdiğini ölçer.

```bash
python sim/generate.py   # sim/firmalar/ altına 40 firmanın XML/Excel/yevmiye verisini ve hata listesini üretir
python sim/metin.py      # her firmayı programın kodu ile denetler, sim/sonuclar.json ve sim/calisma/ raporlarını yazar
                         # (firmalar programın firma kayıt defteriyle sim/calisma/veri/ altında açılır)
```

- 5 sektör: İnşaat (8), Halı Üretimi (7), Uluslararası Taşımacılık (8), Otomotiv Satış ve Kiralama (8), Muhtelif İmalat (9)
- Her firmaya bilinen hatalar yerleştirilir (`meta.json` → `truth`): muhasebeleşmemiş fatura, yanlış hesap, tutar farkı,
  KDV'nin maliyete eklenmesi, çift kayıt, dönem kayması, faturasız gider, mükerrer fatura, fiyat şişirme,
  başka firmaya kesilmiş fatura, XML hesaplama hatası; madde 7 ile ÖTV'nin maliyete eklenmemesi, KDV farkı, KDV'nin
  hiç kaydedilmemesi, tevkifat kaydı eksik/yanlış, muhasebeleşmemiş satış, satış tutar / KDV farkı, satış dönem kayması
  ve faturasız gelir kaydı.
- Madde 7: otomotiv araç alımlarında ÖTV (9077; XML'de satır ve belge düzeyinde, Excel'de `ÖTV Tutarı` sütunu) ve
  doğru kayıtta 153/254'e matrah + ÖTV; tevkifatlı faturalarda `WithholdingTaxTotal` ve yevmiyede 360 alacak kaydı;
  her firmaya giden satış faturaları (XML firmalarında ZIP'te `giden/` ya da `xml/` klasöründe, Excel firmalarında
  `faturalar.xlsx` içinde `Yön` = Satış satırları). Halı ve imalat firmalarında satışların bir kısmı EUR ihracat
  (IHRACAT / ISTISNA, 601), taşımacılıkta uluslararası taşıma istisnası (ISTISNA). Yevmiyedeki 120/600/391 kayıtları
  eski gürültü satışlarının (`SAT…` belgeleri) yerine bu faturalarla eşleşir: aynı tarih ve tutarlar kullanılır.
- Madde 7 eklemeleri ayrı rastgele üreteçle (`random.Random(7000 + firma no)`) üretilir; gürültü satışlarının
  rastgele çekilişleri de korunur. Böylece mevcut hata türlerinin listeleri, firma bilgileri ve otomotiv dışı alış
  XML'leri madde 6 verisiyle birebir aynıdır (yalnızca otomotiv alışlarına ÖTV eklenir).
- `python sim/generate.py <klasör>` başka bir klasöre üretir (önce/sonra karşılaştırması için).
- Gerçekçi gürültü: bordro/amortisman/SMM/gider pusulası kayıtları, taşeron hakedişleri, dövizli yakıt faturaları,
  aynı gün aynı model araç alımları, farklı muhasebe programı döküm biçimleri ve belge no yazım biçimleri.
- `metin.py` yevmiye dosyasını olduğu gibi yükler; başlık satırlarını ve "Borç Tutarı" gibi sütun adlarını program
  kendisi çözer (madde 3). Belge numarası ayrı sütunda olmayan (zirve biçimi) dökümlerde program açık bir hata
  verir; `metin.py` bunu "muhasebeciden ayrı belge no sütunlu döküm istendi" (`belge_no_ayri_sutun_istendi`)
  olarak `friction`'a yazar ve Belge No sütunu eklenmiş dosyayla tekrar dener (denetçinin yeni döküm almasını
  temsil eder; program açıklamadan belge no ayıklamaz).

Rastgelelik sabit tohumla (`random.seed(2024)`) üretildiği için sonuçlar tekrarlanabilir.

## Ölçümler

`metin.py` çalışmanın sonunda kontrol bazında toplamları, belge no biçimine göre kırılımı ve eşleştirme
yöntemlerinin dağılımını yazdırır. Yalnızca bilgi amaçlı bölümler ("Belge No Uyuşmayan Eşleşmeler",
"Belirsiz Eşleşme (Birden Fazla Seri+Sıra Adayı)") bilinen bir hatayı temsil etmediği için puanlamaya
katılmaz, yanlış alarm sayılmaz; `sonuclar.json` içinde firma başına `bilgi` alanında ayrıca raporlanır.

| Kontrol | v2.0 yakalanan | v2.0 yanlış alarm | Madde 2 sonrası yakalanan | Madde 2 sonrası yanlış alarm | Madde 3 sonrası yakalanan | Madde 3 sonrası yanlış alarm | Madde 4 sonrası yakalanan | Madde 4 sonrası yanlış alarm | Madde 5 sonrası yakalanan | Madde 5 sonrası yanlış alarm | Madde 6 sonrası yakalanan | Madde 6 sonrası yanlış alarm | Madde 7 sonrası yakalanan | Madde 7 sonrası yanlış alarm |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Muhasebeleşmemiş fatura | 46/46 | 998 | 46/46 | 0 | 46/46 | 0 | 46/46 | 0 | 46/46 | 0 | 46/46 | 0 | 46/46 | 0 |
| Yanlış hesaba kayıt | 51/57 | 0 | 57/57 | 0 | 57/57 | 0 | 57/57 | 0 | 57/57 | 0 | 57/57 | 0 | 57/57 | 0 |
| Tutar farkı | 70/77 | 206 | 77/77 | 272 | 77/77 | 272 | 77/77 | 272 | 77/77 | 272 | 77/77 | **11** | 94/94 | 11 |
| Dönem kayması | 34/43 | 0 | 43/43 | 0 | 43/43 | 0 | 43/43 | 0 | 43/43 | 0 | 43/43 | 0 | 43/43 | 0 |
| Faturasız gider kaydı | 52/52 | 1.404 | 52/52 | 412 | 52/52 | 412 | 52/52 | 0 | 52/52 | 0 | 52/52 | 0 | 52/52 | 0 |
| Mükerrer fatura | 41/41 | 259 | 41/41 | 259 | 41/41 | 259 | 41/41 | 259 | 41/41 | 259 | 41/41 | 259 (incelemeden sonra açık: 0) | 41/41 | 259 (incelemeden sonra açık: 0) |
| Fiyat şişirme | 41/41 | 816 | 41/41 | 816 | 41/41 | 816 | 41/41 | 816 | 41/41 | **7** | 41/41 | 7 | 41/41 | 7 |
| Başka firmaya kesilmiş fatura | 20/20 | 0 | 20/20 | 0 | 20/20 | 0 | 20/20 | 0 | 20/20 | 0 | 20/20 | 0 | 20/20 | 0 |
| XML hesaplama hatası | 8/8 | 0 | 8/8 | 0 | 8/8 | 0 | 8/8 | 0 | 8/8 | 0 | 8/8 | 0 | 8/8 | 0 |
| **Toplam** | **363/385** | **3.683** | **385/385** | **1.759** | **385/385** | **1.759** | **385/385** | **1.347** | **385/385** | **538** | **385/385** | **277** (incelemeden sonra açık: **0**) | **385/385** (yeni hatalarla **665/665**) | **277** (incelemeden sonra açık: **0**) |

**Madde 7 sonrası** sütunu yeni veridir (satış faturaları, ÖTV, yeni hata türleri). Tutar farkındaki 94 beklenen hata
77 eski hata + 17 "ÖTV maliyete eklenmemiş" hatasıdır. Mevcut kontroller eski veride (madde 6 verisi, yeni kodla)
madde 6 sütunuyla birebir aynıdır: 385/385, 277 yanlış alarm.

Firma seçici (madde 1) sonuçları değiştirmedi; v2.0 sütunu madde 1 sonrası için de geçerlidir. Madde 6 sütunu
varsayılan %1 kur toleransıyla, inceleme işaretleri uygulanmadan (ilk çalıştırma) ölçülmüştür; "incelemeden sonra
açık" ikinci çalıştırmadır (aşağıya bakın).

**Madde 2 (akıllı belge no eşleştirme)** — belge no biçimine göre:

| Belge no biçimi | Firma | v2.0 yakalanan / yanlış alarm | Madde 2 sonrası yakalanan / yanlış alarm |
|---|---|---|---|
| tam (`ABC2024000000123`) | 28 | 273/273 / 1.063 | 273/273 / 1.063 |
| boşluklu (`ABC 2024 000000123`) | 8 | 73/73 / 499 | 73/73 / 499 |
| kısa (`ABC123`) | 4 | 17/39 / 2.121 | 39/39 / 197 |

- Eşleştirme dağılımı (40 firma): Tam 10.912, Seri+Sıra 992 (kısa biçimli 4 firmanın tüm faturaları),
  Tutar+Tarih 0, seçili hesap dışı 57, belirsiz 0, eşleşmeyen 46.
- "Belge No Uyuşmayan Eşleşmeler" bölümü bu simülasyonda boş kalıyor: kısa biçimli belge numaraları seri+sıra
  kademesinde eşleştiği için tutar+tarih yedeğine hiç kalmıyor (beklenen doluluk bu firmalarda seri+sıra
  eşleşmesi olarak görünür).
- Tutar farkındaki +66 yanlış alarm yeni değil: kısa biçimli iki taşımacılık firmasındaki dövizli yakıt
  faturalarının kur farkı (diğer taşımacılık firmalarında zaten görünen tür); bu faturalar önceden hiç
  eşleşmediği için "muhasebeleşmemiş" yanlış alarmı olarak sayılıyordu.
- Kalan faturasız kayıt yanlış alarmları bordro/amortisman/SMM/gider pusulası kayıtlarıdır (madde 4 ile giderildi).
- Tutar+tarih kademesinin ayrı stres testi: 14 firmada fatura belgelerinin %30'unun numarası tamamen bozuldu
  (ör. `FIS123456X`); 925 bozuk belgenin 822'si tutar+tarih ile eşleşti, **yanlış eşleşme 0**, gerçek
  muhasebeleşmemiş faturalardan kaçan yok. Eşleşmeyen 103 fatura, tutarı tutmadığı (tutar farkı, KDV dahil
  kayıt, dövizli kur farkı) ya da aynı tutarlı birden fazla aday olduğu (aynı gün aynı araç alımı) için
  güvenli biçimde eşleştirilmedi ve muhasebeleşmemiş olarak raporlandı (bu test kalıcı değildir, elle yapıldı).

**Madde 3 (yevmiye içe aktarma sihirbazı)** — yakalama ve yanlış alarm sayıları firma bazında birebir aynı kaldı
(385/385, 1.759; içe aktarılan yevmiye satırı ve hatalı satır sayıları da değişmedi). Değişen, elle müdahale:

| | Madde 2 sonrası | Madde 3 sonrası |
|---|---|---|
| Elle müdahale gereken yevmiye dosyası | 22/40 | 9/40 |
| Başlık satırı (üstte firma adı / rapor başlığı) | 8 (Excel'de satır silme) | 0 (program başlığı 4. satırda buluyor) |
| Sütun adı (`Borç Tutarı` / `Alacak Tutarı`) | 5 (Excel'de yeniden adlandırma) | 0 (takma ad olarak tanınıyor) |
| Belge no açıklamada (zirve biçimi) | 9 (Excel'de açıklamadan ayıklama) | 9 (program açık hata verir; muhasebeciden Belge No sütunlu döküm istenir) |
| Tahmini elle çalışma süresi | 245 dk | 180 dk |

Kalan 9 dosya bilinçli bir tercihtir: belge numarası açıklama metninden çıkarılmaz, ayrı sütunda olmalıdır.

**Madde 4 (faturası bulunmayan yevmiye kayıtları süzgeci)** — diğer kontroller birebir aynı kaldı; değişen
yalnızca "Faturası Bulunmayan Yevmiye Kayıtları":

| Faturasız kayıt kontrolü | Yakalanan | Yanlış alarm | Yüksek öncelikli kayıt |
|---|---|---|---|
| Madde 3 sonrası (süzgeç yok) | 52/52 | 412 (GP 244, AMORT 72, SMM 72, BORDRO 24) | — |
| Yalnızca borç yönü (boş önek listesi) | 52/52 | 340 (SMM'nin 72'si alacak yönlü olarak elendi) | 52 (52'si de gerçek hata) |
| Borç yönü + varsayılan önek listesi | 52/52 | **0** (önekle 340 kayıt elendi) | 52 |

- Önek listesi boşken de gerçek 52 hata listenin en üstündedir: hepsi GİB biçimli belge no'su taşıdığı için
  **Yüksek** öncelikli; kalan 340 olağan kayıt (`GP-…`, `AMORT-…`, `BORDRO-…`) **Düşük** öncelikli olarak altta.
- Simülasyondaki yevmiyelerin hepsi Borç/Alacak sütunlu (işaretli) olduğundan işaretsiz Tutar durumu burada
  oluşmuyor; birim testleriyle doğrulanıyor.
- Varsayılan önekler simülasyona göre değil, yaygın muhasebe pratiğine göre seçildi (`SMM` listede yok; SMM
  kayıtlarını borç yönü süzgeci eler). `metin.py` boş önek listesiyle ölçümü de her çalıştırmada yazdırır
  ("Faturasız kayıt, boş önek listesiyle …") ve raporlara **Faturasız Kayıt Özeti** ile
  **Faturasız Listeden Hariç Tutulanlar** sayfalarını ekler.

**Madde 5 (fiyat anomalisi analizi)** — diğer kontroller birebir aynı kaldı; değişen yalnızca "Fiyat Anomalileri"
(yanlış alarm, riskli satır içeren gerçek hatası olmayan fatura sayısıdır):

| Fiyat analizi (%15 eşik) | Yakalanan | Yanlış alarm |
|---|---|---|
| Madde 4 sonrası (dönem + ürün + birim, TL) | 41/41 | 816 (taşeron hakedişi 309, motorin 269, fason dokuma 231, diğer 7) |
| Yalnızca para birimine göre ayrı grup | 41/41 | 547 (EUR motorinin 269'u gitti) |
| + tevkifatlı faturalar hariç | 41/41 | 125 |
| + yetersiz veri (en az 3 alım) — **boş kelime listesi** | 41/41 | 125 |
| + varsayılan kelime listesi (**varsayılan kurallar**) | 41/41 | **7** |

Eşik duyarlılığı (`metin.py` her çalıştırmada yazdırır):

| Eşik | Madde 4 sonrası | Madde 5, varsayılan kelimeler | Madde 5, boş kelime listesi |
|---|---|---|---|
| %10 | 41/41, 1.041 yanlış alarm | 41/41, 133 | 41/41, 268 |
| %15 | 41/41, 816 | 41/41, **7** | 41/41, 125 |
| %25 | 39/41, 631 | 39/41, 0 | 39/41, 94 |
| %35 | 34/41, 405 | 35/41, 0 | 35/41, 56 |

- Boş kelime listesiyle kalan 125 yanlış alarmın 117'si Excel'den yüklenen firmalardaki fason dokuma (101) ve
  taşeron hakedişi (16) faturalarıdır: Excel şablonunda fatura tipi olmadığından tevkifat kuralı bu faturaları
  tanıyamaz, kelime listesi tanır. Kalan 7 (varsayılan kurallarda da kalan) yanlış alarm lateks, lastik, hazır
  beton ve boya kimyasalında %16–20'lik gerçek fiyat dalgalanmasıdır (ürünün kendisi, hizmet değil).
- Varsayılan kelimeler simülasyona göre değil, yaygın hizmet / hakediş adlandırmasına göre seçildi; simülasyonda
  `HAKEDİŞ` (taşeron), `HİZMET` (fason dokuma, nakliye), `KİRALAMA` (iş makinesi) ve `BAKIM` (tır bakım onarım)
  eşleşiyor. `Kiralık Filo Aracı` (araç alımı) `KİRALAMA` ile eşleşmez ve analizde kalır.
- **En az alım sayısı (yetersiz veri):** simülasyonda gerçek bir fiyat şişirmesinin bulunduğu en küçük grup 4 alımlı;
  en az alım 1–5 arasında yakalama ve yanlış alarm değişmiyor (41/41, 7; eşik üstü yetersiz veri satırı 0). Tek
  alımlı grupta sapma her zaman 0 olduğundan "en az 2" etkisizdir; varsayılan **3** seçildi. Seyrek veri stres
  testinde (gerçek hatalar dışındaki satırların bir kısmı rastgele atılarak, kalıcı değildir, elle yapıldı) 3'ün
  bedeli görülüyor:

  | Tutulan satır | en az 1–2 | en az 3 | en az 4 | en az 5 |
  |---|---|---|---|---|
  | %50 | 39/41, 14 | 39/41, 13 | 37/41, 12 | 36/41, 12 |
  | %20 | 34/41, 18 | 33/41, 17 | 28/41, 15 | 25/41, 13 |
  | %10 | 31/41, 22 | 25/41, 16 | 20/41, 8 | 9/41, 4 |

  İki alımlı gruplarda yakalama ve yanlış alarm aynı oranda düşüyor (iki fiyat ortalamadan simetrik sapar);
  riskli sayılmayan satırlar kaybolmaz, "Yetersiz veri (bilgi)" olarak raporda ve **Fiyat Analizi Dışı Satırlar**
  sayfasında görünür. Çok seyrek alım yapan firmalarda en az alım 2'ye (ya da kuralı kapatmak için 1'e) indirilebilir.
- Raporlara **Fiyat Analizi Özeti** ve **Fiyat Analizi Dışı Satırlar** sayfaları eklendi (40 firmada analiz dışı:
  iade 40, tevkifat 1.068, anahtar kelime 1.419 satır; incelenen 14.923 satır).

**Madde 6 (dövizli faturalarda yüzde kur toleransı + bulgu inceleme işareti)** — yakalama 385/385 kaldı; yanlış alarm
ilk çalıştırmada 538 → **277**, Metin'in incelemesinden sonraki çalıştırmada açık yanlış alarm **0**.

Tutar farkı, dövizli faturalarda yüzde kur toleransına göre (TL faturalarda her durumda yalnızca 0,01 TL tolerans):

| Kur toleransı | Tutar farkı yakalanan | Yanlış alarm | Kur farkı (tolerans içi, bilgi) |
|---|---|---|---|
| %0 (madde 5 sonrası davranış) | 77/77 | 272 | 0 |
| %0,5 | 77/77 | 62 | 210 |
| **%1 (varsayılan)** | 77/77 | **11** | 261 |
| %2 | 77/77 | 8 | 264 |
| Karşılaştırma: dövizli 8 firmada tüm faturalara 50 TL sabit tolerans, %0 kur | 12/12 (bu 8 firmada) | 182 | — |

- 8 taşımacılık firmasında 265 dövizli (EUR) fatura eşleşiyor; %1'de 261'i tolerans içi kur farkı olarak bilgiye
  alındı, 3'ü %1,02–%1,17 farkla Tutar Farkları'nda kaldı (`Dovizli` = Evet, `Kur` sütunuyla). Simülasyonda muhasebe
  kuru fatura kurundan σ ≈ %0,4 sapıyor; %1 ≈ 2,5σ.
- Kalan 8 yanlış alarm TL faturalardır (57–455 TL fark; madde 5 öncesinden beri listede, simülasyonun hata listesine
  yazmadığı farklar). TL faturalara bilinçli olarak yüzde tolerans uygulanmıyor: simülasyondaki rakam yer değiştirme
  hataları yalnızca birkaç TL fark yaratıyor (ör. 9 TL) ve %1 tolerans bunları gizlerdi. Sabit 50 TL tolerans ise
  dövizli firmalarda yanlış alarmı yalnızca 182'ye indiriyor ve küçük TL hatalarını gizleme riski taşıyor.

Bulgu inceleme senaryosu (`metin.py` her çalıştırmada yazdırır): Metin ilk çalıştırmada bilinen hata listesinde
olmayan her bulguyu "İncelendi – Sorun Yok" olarak işaretler; ardından firmanın fatura ve yevmiye verileri silinip
aynı dosyalar yeniden içe aktarılır (yeni veritabanı kimlikleri) ve rapor yeniden çalıştırılır.

| | Açık yanlış alarm | "Sorun yok" (gizli) | Gerçek hata açık |
|---|---|---|---|
| 1. çalıştırma | 277 (mükerrer 259, tutar farkı 11, fiyat 7) | 0 | 385/385 |
| 2. çalıştırma (yeniden içe aktarma sonrası) | **0** | 277 | **385/385** (hiçbiri yanlışlıkla kapanmadı) |

- 277 bulgu 277 farklı kararlı anahtarla işaretlendi; yeniden içe aktarmadan sonra hepsi aynı anahtarla bulundu.
- **Mükerrer, ardışık numara (bilgi):** yanlış alarm gruplarının 34/259'u, gerçek mükerrerlerin 2/41'i aynı serinin
  ardışık numaraları. Simülasyondaki otomotiv yanlış alarmlarının çoğu, sabit liste fiyatlı araçların farklı
  zamanlarda kesilmiş faturalarının aynı güne denk gelmesidir (numaralar ardışık değil); ardışık numaralı toplu alım
  grupları azınlıkta. Bu yüzden ardışık numara bulguyu gizlemez, yalnızca sıralar (ardışık olmayanlar önce) — mükerrer
  yanlış alarmları denetçinin "sorun yok" işaretiyle kapanır.
- Raporlara **İnceleme Özeti** ve **Kur Farkı (Tolerans İçi)** sayfaları ile bulgu sayfalarına `Inceleme_Durumu`,
  `Inceleme_Notu`, `Inceleme_Tarihi`, `Bulgu_Anahtari` sütunları eklendi (`sim/calisma/<KOD>_Denetim_Raporu.xlsx`
  ikinci çalıştırmanın raporudur).

**Madde 7 (ÖTV, KDV / tevkifat mutabakatı, satış faturaları)** — yeni veride 40 firma: 2.989 satış faturası (798'i
ihracat / istisna), ÖTV'li 2.444 araç alım faturası (toplam ÖTV 2,44 milyar TL), 589 tevkifatlı alış faturası.

Yeni kontroller (yeni veri, varsayılan ayarlar: 191, 360, gelir 600/601/602, hesaplanan KDV 391):

| Kontrol | Bilinen hata türü | Yakalanan | Yanlış alarm |
|---|---|---|---|
| Tutar farkı (ÖTV maliyete eklenmemiş) | `otv_maliyete_eklenmemis` (ÖTV 770'e yazılmış) | 17/17 (Tutar Farkları'nda; `Olasi_Neden` dolu) | 0 |
| KDV Farkları | `kdv_farki` (26) + `cift_kayit` (10) | 36/36 | 0 |
| KDV'si Kaydedilmemiş Faturalar | `kdv_kaydedilmemis` (30) + `kdv_dahil_kayit` (35) | 65/65 | 0 |
| Tevkifat Kaydı Eksik/Farklı | `tevkifat_kaydi_eksik` (360 yok ya da 5/10 ile yazılmış) | 8/8 | 0 |
| Muhasebeleşmemiş Satış Faturaları | `muhasebelesmemis_satis` | 29/29 | 0 |
| Satış Tutar Farkları | `satis_tutar_farki` | 28/28 | 0 |
| Satış Dönem Farkları | `satis_donem_kaymasi` | 31/31 | 0 |
| Satış KDV Farkları | `satis_kdv_farki` (%10 yazılmış, 391 yok, istisnaya KDV) | 28/28 | 0 |
| Faturası Bulunmayan Gelir Kayıtları | `faturasiz_gelir` | 38/38 | 0 |
| Gelir Hesabı Dışına Kaydedilmiş Satış Faturaları | — | — | 0 |
| **Toplam (tüm kontroller)** | | **665/665** | **277** (hepsi madde 6'dan kalan türler; incelemeden sonra açık **0**) |

Önce / sonra (mevcut 9 kontrolün yakalaması ve yanlış alarmı):

| | Eski veri (madde 6) | Yeni veri (madde 7) |
|---|---|---|
| Madde 6 kodu | 385/385, 277 yanlış alarm | 385/385, **9.345** yanlış alarm |
| Madde 7 kodu | 385/385, 277 yanlış alarm (yeni kontroller: KDV 45/45, 0 yanlış alarm) | 385/385 (+17 ÖTV hatası), 277 yanlış alarm |

- Madde 6 kodunun yeni verideki 9.345 yanlış alarmı = önceki 277 + 9.068 yeni: satış faturaları alış sayıldığı için
  seçili hesap dışı +2.960, alıcısı firma olmayan +2.188, fiyat anomalisi +1.490 (satış kalemleri), muhasebeleşmemiş
  +29 (kaydı olmayan satışlar); ÖTV maliyete eklenmediği için otomotiv firmalarında tutar farkı +2.401 (ÖTV'li hemen
  her araç faturası). Madde 7 kodu bunların hepsini gidermektedir.
- Madde 7 kodu eski veride: KDV Farkları 10/10 (çift kayıtlar), KDV'si Kaydedilmemiş 35/35 (KDV'nin maliyete eklendiği
  kayıtlar) — mevcut hatalar ikinci bir kontrolde de yakalanıyor. Eski veride satış faturası olmadığından satış
  mutabakatı çalışmaz. İlk ölçümde Excel firmalarındaki iade faturaları (Excel'de fatura tipi yok) "ters yönde KDV"
  olarak 20 yanlış alarm verdi; KDV yön kontrolü yalnızca XML faturalara uygulanarak giderildi.
- Satış eşleştirmesi: Tam 2.684, Seri+Sıra 276 (kısa belge no'lu firmalar), eşleşmeyen 29 (= muhasebeleşmemiş satış
  hataları). Satış KDV'si 2.960 faturada karşılaştırıldı (791'i ihracat / istisna, KDV beklenmez).
- Satış kur farkı (tolerans içi, bilgi) 183: Excel'den yüklenen EUR ihracat faturalarında birim fiyat 4 haneye
  yuvarlandığı için oluşan kuruş farkları (bulgu sayılmaz).
- Tevkifat kontrolü yalnızca XML firmalarında: Excel şablonundaki alış satırlarında `Tevkifat_Orani` verilmediği için
  (portal dökümünü temsil eder) Excel firmalarının tevkifatlı faturaları kontrol dışındadır.
- Bulgu inceleme senaryosu yeni bölümlerde de çalışıyor: 2. çalıştırmada açık yanlış alarm 0, gerçek hata açık
  665/665.
