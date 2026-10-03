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
  başka firmaya kesilmiş fatura, XML hesaplama hatası.
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

| Kontrol | v2.0 yakalanan | v2.0 yanlış alarm | Madde 2 sonrası yakalanan | Madde 2 sonrası yanlış alarm | Madde 3 sonrası yakalanan | Madde 3 sonrası yanlış alarm |
|---|---|---|---|---|---|---|
| Muhasebeleşmemiş fatura | 46/46 | 998 | 46/46 | 0 | 46/46 | 0 |
| Yanlış hesaba kayıt | 51/57 | 0 | 57/57 | 0 | 57/57 | 0 |
| Tutar farkı | 70/77 | 206 | 77/77 | 272 | 77/77 | 272 |
| Dönem kayması | 34/43 | 0 | 43/43 | 0 | 43/43 | 0 |
| Faturasız gider kaydı | 52/52 | 1.404 | 52/52 | 412 | 52/52 | 412 |
| Mükerrer fatura | 41/41 | 259 | 41/41 | 259 | 41/41 | 259 |
| Fiyat şişirme | 41/41 | 816 | 41/41 | 816 | 41/41 | 816 |
| Başka firmaya kesilmiş fatura | 20/20 | 0 | 20/20 | 0 | 20/20 | 0 |
| XML hesaplama hatası | 8/8 | 0 | 8/8 | 0 | 8/8 | 0 |
| **Toplam** | **363/385** | **3.683** | **385/385** | **1.759** | **385/385** | **1.759** |

Firma seçici (madde 1) sonuçları değiştirmedi; v2.0 sütunu madde 1 sonrası için de geçerlidir.

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
- Kalan faturasız kayıt yanlış alarmları bordro/amortisman/SMM/gider pusulası kayıtlarıdır (madde 3 kapsamı).
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
