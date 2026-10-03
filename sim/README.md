# Denetçi Simülasyonu (40 firma)

Programın gerçekçi verilerde ne kadar hata yakaladığını ve ne kadar yanlış alarm verdiğini ölçer.

```bash
python sim/generate.py   # sim/firmalar/ altına 40 firmanın XML/Excel/yevmiye verisini ve hata listesini üretir
python sim/metin.py      # her firmayı programın kodu ile denetler, sim/sonuclar.json ve sim/calisma/ raporlarını yazar
```

- 5 sektör: İnşaat (8), Halı Üretimi (7), Uluslararası Taşımacılık (8), Otomotiv Satış ve Kiralama (8), Muhtelif İmalat (9)
- Her firmaya bilinen hatalar yerleştirilir (`meta.json` → `truth`): muhasebeleşmemiş fatura, yanlış hesap, tutar farkı,
  KDV'nin maliyete eklenmesi, çift kayıt, dönem kayması, faturasız gider, mükerrer fatura, fiyat şişirme,
  başka firmaya kesilmiş fatura, XML hesaplama hatası.
- Gerçekçi gürültü: bordro/amortisman/SMM/gider pusulası kayıtları, taşeron hakedişleri, dövizli yakıt faturaları,
  aynı gün aynı model araç alımları, farklı muhasebe programı döküm biçimleri ve belge no yazım biçimleri.
- `metin.py`, yevmiye dosyası yüklenemediğinde denetçinin Excel'de elle yaptığı düzeltmeleri taklit eder ve
  bunları `friction` olarak kaydeder.

Rastgelelik sabit tohumla (`random.seed(2024)`) üretildiği için sonuçlar tekrarlanabilir.

## Başlangıç ölçümü (v2.0)

| Kontrol | Yakalanan | Yanlış alarm |
|---|---|---|
| Muhasebeleşmemiş fatura | 46/46 | 998 |
| Yanlış hesaba kayıt | 51/57 | 0 |
| Tutar farkı | 70/77 | 206 |
| Dönem kayması | 34/43 | 0 |
| Faturasız gider kaydı | 52/52 | 1.404 |
| Mükerrer fatura | 41/41 | 259 |
| Fiyat şişirme | 41/41 | 816 |
| Başka firmaya kesilmiş fatura | 20/20 | 0 |
| XML hesaplama hatası | 8/8 | 0 |
| **Toplam** | **363/385** | **3.683** |

Elle düzeltme gereken yevmiye dosyası: 22/40 (9 belge no açıklamada, 8 başlık satırı, 5 sütun adı).
