# Teminat Çözüm Raporu — Taslak Programı

Her ay teminatla KDV iadesi alan firmalar için yazılan **"KDV İadesi Teminat Çözümü Tasdik Raporu"**nun
taslağını ve bir **kontrol listesini** üretir.

- Program raporu **imzalamaz, karar vermez.** Bir önceki ayın bitmiş Word raporunu şablon olarak alır;
  dönem metinlerini ve rakamları günceller, tablolarda satır ekler/siler. Ofisin üslubu ve biçimi korunur.
- Girdi dosyalarında olmayan bilgi `[DOLDURULACAK]` yazılır ve **sarı** vurgulanır.
- Takdir gerektiren her seçim (ör. hangi tutanağın "EKLİ" olacağı) ve her tutarsızlık kontrol listesine yazılır.
- Orijinal dosyalara **asla yazılmaz**; çıktılar `CLAUDE TASLAK\` klasörüne yeni adla kaydedilir.
- İnternet gerektirmez. Mükellef verisi dışarı gönderilmez.

## Hazır program (Windows .exe)

GitHub'da her değişiklikte Windows için tek dosyalık program derlenir ve test edilir:
**Releases → "Teminat Çözüm Raporu Taslak Programı – son sürüm" → `TeminatCozum.exe`**
(depo sayfasında sağdaki "Releases" bölümü; ya da Actions → "Teminat .exe derle" → yapıt).

1. `TeminatCozum.exe`'yi bir klasöre koyun (ör. `C:\TeminatCozum\`).
2. Firma ayar dosyasını yanına koyun: `C:\TeminatCozum\firmalar\<firma>\ayar.yaml` (örnek: `ayar.yaml`).
3. Ofis raporları `.doc` ise LibreOffice kurulu olmalı.
4. `.exe`'ye çift tıklayın (ya da ayın klasörünü `.exe`'nin üzerine sürükleyin) ve pencerede:
   - **Ayar dosyası**: `firmalar\<firma>\ayar.yaml`
   - **Bu ayın klasörü**: firmanın o ayki belgelerinin bulunduğu klasör — adı ve içindeki düzen ne olursa olsun.
   - **Önceki ayın raporu**: boş bırakın; program çevredeki Word dosyalarında **içerikten** arar
     (kapaktaki dönem ve mükellef VKN'si). Bulamazsa ya da birden fazla aday varsa buradan kendiniz seçin.
   - **1) Belgeleri tanı**: klasördeki her dosyayı açıp türünü **içeriğinden** bulur ve tanıma tablosunu gösterir
     (hangi dosya beyanname, hangisi indirilecek liste, hangisi kullanılacak, hangisi eksik). Türü yanlış görünen
     satıra çift tıklayıp düzeltebilirsiniz; düzeltme sonraki aylarda da hatırlanır.
   - **2) Taslak oluştur**: tanımaya göre taslağı ve kontrol listesini üretir.
   Pencerede ayrıca "Gerçek raporla karşılaştır", "Rapor denetle" ve "Geriye dönük test" düğmeleri vardır.

`.exe` komut satırından da çalışır (`TeminatCozum.exe taslak "<ayın klasörü>" --ayar ...`); konsolu olmadığı
için çıktı yanındaki `teminat_son_calisma.log` dosyasına yazılır. Python ile çalıştırmak için aşağıdaki kurulum.

## Kurulum (Python ile)

1. Python 3.11 veya üstü (python.org). Kurulumda "Add python.exe to PATH" işaretlenmeli.
2. LibreOffice (ofisin `.doc` raporlarını `.docx`'e çevirmek için). Program `soffice.exe`'yi
   `C:\Program Files\LibreOffice\program\` altında kendisi bulur; başka yerdeyse `SOFFICE` ortam değişkenine
   tam yolunu yazın.
3. Bu klasörde bir komut penceresi açıp:

```
pip install -r requirements.txt
```

## Kullanım

Bütün komutlar bu klasörden (`teminat_cozum`) çalıştırılır. Pencereli arayüz: `python -m teminat arayuz`.

### Belgeleri tanıma ve aylık taslak

Programa yalnızca **firmanın o ayki klasörü** verilir. Dosya adlarına ve klasör düzenine bakılmaz; her belge
içeriğinden tanınır (bkz. aşağıda "Belgeler nasıl tanınır"). Önce yalnızca tanıma tablosunu görmek için:

```
python -m teminat tani "C:\...\<firmanın o ayki klasörü>" --ayar firmalar\<firma>\ayar.yaml
```

Taslak (önce aynı tanıma tablosunu yazar, sonra taslağı üretir):

```
python -m teminat taslak "C:\...\<firmanın o ayki klasörü>" --ayar firmalar\<firma>\ayar.yaml
```

Program bir belge için karar veremezse (ör. aynı dönemin iki farklı beyannamesi) adayları listeler ve durur;
hangisinin kullanılacağını `--sec ROL="dosya"` ile (pencerede açılan seçim kutusundan) belirtirsiniz. Şablonu elle
vermek için `--sablon "<önceki ayın raporu.doc>"`; klasörde birden fazla dönemin beyannamesi varsa `--donem 2026-03`.

Çıktılar (ayın klasörünün içinde, `CLAUDE TASLAK\` altında; aynı adlı dosya varsa üzerine yazılmaz, `(2)` eklenir):

| Dosya | İçerik |
|---|---|
| `MART-2026 RAPOR TASLAK.docx` | Rapor taslağı. Elle doldurulacak yerler ve kontrol edilecek satırlar sarı. |
| `MART-2026 KONTROL LİSTESİ.xlsx` | Kontrol listesi + özet + **belgeler (tanıma tablosu)** + safha (EKLİ seçimi) sayfası |
| `MART-2026 KONTROL LİSTESİ.html` | Aynı kontrol listesi, tarayıcıda açılır (durum filtresiyle) |
| `MART-2026 FARKLAR.txt` | Yalnızca ofisin bu ayki bitmiş raporu da klasördeyse: taslakla karşılaştırma |

Kontrol listesindeki durumlar:

| Durum | Anlamı |
|---|---|
| **HATA** | Taslak bu haliyle yanlış olabilir; mutlaka bakılmalı (ör. teminat dilekçesi bulunamadı, toplam tutmuyor). |
| **UYARI** | Tutarsızlık ya da takdir gerektiren seçim (ör. EKLİ firma seçimi, liste ↔ beyan farkı). |
| **ELLE** | Girdi dosyalarında olmayan bilgi; taslakta `[DOLDURULACAK]` (rapor sayısı, işçi sayısı...). |
| **BİLGİ** | Bilgi amaçlı (eklenen / silinen satırlar vb.). |
| **TAMAM** | Kontrol yapıldı, sorun yok. |

### Taslağı ofisin bitmiş raporuyla karşılaştırma (öğrenme döngüsü)

```
python -m teminat karsilastir "...\03 MART\RAPOR\MART-2026 RAPOR.doc" "...\03 MART\CLAUDE TASLAK\MART-2026 RAPOR TASLAK.docx"
```

Paragraf ve tablo farklarını listeler; ayrıca **tutar farkını** sayar (gerçek raporda olup taslakta olmayan
tutarlar). Hedef: tutar farkı 0. Her fark ya bir kurala dönüştürülür ya da firma ayar dosyasına eklenir.

### Geriye dönük test

Firma klasörünün altındaki her ayı (beyannamelerden) bulur; her ay için taslağı bir önceki ayın gerçek raporundan
üretir ve o ayın gerçek raporuyla karşılaştırır. Ay klasörlerinin adı ve düzeni önemli değildir:

```
python -m teminat geriye-donuk "C:\...\<FİRMA>" --ayar firmalar\<firma>\ayar.yaml --baslangic 2025-01 --bitis 2026-02
```

Sonuç: `<FİRMA>\CLAUDE TASLAK\GERİYE DÖNÜK TEST\GERİYE DÖNÜK TEST.xlsx` (dönem başına tutar farkı, farklı satır
sayısı, HATA/UYARI sayısı) ve her ay için `FARKLAR.txt`. Girdi klasörlerine yazılmaz.

### Bitmiş bir raporu denetleme

```
python -m teminat denetle "...\ARALIK-2025 RAPOR.doc"
```

3-4-3 safha tablosunda satır toplamı ↔ TOPLAM hücreleri, EKLİ sıra numaraları, aynı tutarın iki firmada
görünmesi (satır kayması) ve %80 oranını kontrol eder.

## Belgeler nasıl tanınır (dosya adı ve klasör düzeni önemsizdir)

Seçilen klasördeki (alt klasörler dahil) her dosyanın ilk kısmı okunur — PDF'in ilk sayfası, Excel'in ilk satırları,
Word'ün ilk paragrafları ve tabloları, `.doc` için LibreOffice — ve içindeki ifadelere göre türü belirlenir
(`teminat/tanima.py`, `KURALLAR`). Uzantısı `.xls` olup içi `.xlsx` olan dosyalar (GİB'den inenler) içeriğinden açılır.

| Tür | Tanıma (içerikte geçen) | Raporda |
|---|---|---|
| KDV1 | "KATMA DEĞER VERGİSİ BEYANNAMESİ … Gerçek Usulde" | zorunlu; dönem buradan okunur |
| LISTE_INDIRILECEK | "İndirilecek KDV listesi", "Toplam İndirilen KDV", "Alış Faturasının Tarihi … Satıcının" | zorunlu |
| TAKIP | "FİRMA … AÇIKLAMA SMMM YMM" başlıklı takip listesi | zorunlu |
| TEMINAT_DILEKCE | "teminat mektubu … kabul", "Artırımlı / İndirimli Teminat" | zorunlu |
| LISTE_YUKLENILEN | "Yüklenilen KDV listesi", "Bünyeye Giren" | beyanda 301 yüklenilen varsa |
| RAPOR | "GENEL BİLGİ … Raporun amacı" | önceki ayınki şablon; bu ayınki varsa karşılaştırma |
| KDV2, KIT, YMM_YAZISI, IMALAT, LISTE_GCB, SGK_LISTE, … (30'a yakın tür) | | şimdilik yalnızca listelenir |
| TARANMIS / BILINMEYEN / OKUNAMADI | metni olmayan PDF ya da görüntü / hiçbir kurala uymayan / açılamayan | listelenir |

Aynı türden birden fazla dosya varsa hangisinin kullanılacağı yine içerikten seçilir: indirilecek liste için dönem
sütunu ve beyanla tutan toplam; takip listesi için indirilecek listeyle tutan tutarlar; dilekçe için andığı dönem;
yüklenilen listesi için beyandaki 301 yüklenilen KDV'ye eşit toplam. Seçim ve gerekçesi kontrol listesine yazılır;
karar verilemezse kullanıcıya sorulur. Klasör adı başka bir ayı gösteriyorsa (ör. "Mart" klasöründe Şubat beyannamesi)
uyarı verilir. Zorunlu bir belge bulunamazsa tanıma tablosunda **✘ EKSİK** olarak gösterilir.

**Tür düzeltmeleri:** pencerede tanıma tablosundaki satıra çift tıklayıp türü değiştirirseniz (ya da "yok sayılsın"),
bu seçim firma ayar dosyasının yanındaki `belge_turleri.yaml`'a yazılır ve sonraki aylarda da uygulanır. Anahtar, dosyanın
ayın klasörüne göre göreli yolundan ay adları ve yıllar çıkarılmış halidir: "liste ŞUBAT 2026.xls" ile "liste MART 2026.xls"
aynı sayılır; "KDV 1.pdf" ile "KDV 2.pdf" ya da "liste.xls" ile "SİSTEM\liste.xls" farklıdır. Düzeltme başka dosyaları da
etkileyecekse önce sorulur; "Otomatik" seçilirse düzeltme kaldırılır. Ayar dosyası seçilmeden düzeltme kaydedilmez
(düzeltmeler firmaya özeldir).

Programın çıktı klasörleri (`CLAUDE TASLAK` ya da `--cikti` ile verilen klasör) bir işaret dosyası taşır ve hiçbir aramada
girdi ya da şablon sayılmaz.

Okunan dosyaların türü ve dönemi (metin değil) hız için kullanıcı klasöründe `.teminat_cozum\tanima_onbellek.json`
dosyasında saklanır; dosya değişince yeniden okunur.

## Firma ayar dosyası

`firmalar\<firma>\ayar.yaml` — örnek ve tüm alanların açıklaması: [`firmalar/ornek/ayar.yaml`](firmalar/ornek/ayar.yaml).
Ayarda olmayan her alan için programdaki varsayılan kullanılır. Tablolar sıra numarasıyla değil içerikleriyle
bulunur (araya yeni bir tablo girse de kaymaz); başlığı farklı bir rapor için `tablolar:` bölümüyle tanım değiştirilebilir.

> **Gizlilik:** Gerçek firma ayar dosyaları, `PROJE_TALIMATI.md`, girdi dosyaları ve raporlar mükellef bilgisi
> içerdiği için depoya konmaz (`.gitignore`). Bunları yalnızca ofis bilgisayarında tutun.

## Yapılan kontroller

| # | Kontrol |
|---|---|
| 1 | İndirilecek KDV listesi toplamı ↔ beyandaki yurtiçi + sorumlu + ithal |
| 2 | Liste dosyasının dönemi (sütun 15) ↔ rapor dönemi |
| 3 | Karşıt inceleme oranı ≥ %80; EKLİ toplamı ≥ %80 |
| 4 | Takip listesi tutarları ↔ indirilecek liste (firma bazında); listede olup takipte olmayan firmalar |
| 5 | Safha tablosunda satır toplamı ↔ TOPLAM hücreleri; satır kayması; sıra numaraları |
| 6 | İade türleri toplamı ↔ "İade edilmesi gereken KDV" |
| 7 | Teminat dilekçesi türler toplamı ↔ toplam; teminat ≤ iade (tür bazında da); mektup tarihi/no ve banka |
| 8 | 410: matrah × %20 × 4/10 ↔ beyandaki iade |
| 9 | Yüklenilen tutanak toplamı ↔ 301 yüklenilen (+339) |
| 10 | Önceki ay "sonraki döneme devreden" (şablondan) ↔ bu ay "önceki dönemden devreden" |
| 11 | Eski kalmış metin: önceki dönem adı, yıl ifadeleri, VUK 353 parasal had paragrafı, programın doldurmadığı tutar hücreleri |
| – | KDV 1 okuma tutarlılığı: okunan satırların toplamı ↔ matrah toplamı / hesaplanan KDV / indirimler toplamı / oran dağılımı |
| – | Paragraf güncellenemezse (kalıp tutmazsa) ve bu ay geçersiz olan bölümler (ör. 410 yokken 3.7) sarı + listede |

## Geliştirme

```
pip install -r requirements-dev.txt
python -m pytest
```

Testler tamamen **sentetik** (uydurma firma ve tutarlarla) veriyle çalışır (`tests/sentetik.py`); gerçek
mükellef verisi gerekmez. Kod yapısı:

```
teminat/
  okuyucular/   kdv1.py (KDV 1 PDF), listeler.py (indirilecek / takip / yüklenilen), teminat.py (dilekçe), girdiler.py
  rapor/        olustur.py (taslak üretimi), safha.py (EKLİ seçimi), tablolar.py (tabloyu içerikten bulma), docx_araclari.py
  kontroller/   kurallar.py (sayısal kontroller), belge.py (Word üzerindeki kontroller), cikti.py (xlsx/html), liste.py
  tanima.py     belgeleri içerikten tanıma, kullanılacak belgeyi ve şablonu seçme, tanıma tablosu
  islem.py      tanıma + seçim + şablon arama adımları (komut satırı ve pencere ortak)
  taslak.py     bir ayın taslağı + kontrol listesi
  karsilastir.py, geriye_donuk.py, klasorler.py (klasör adından ay ipucu), donusum.py (.doc → .docx), ayar.py, cli.py, arayuz.py
TeminatCozum.py   .exe giriş noktası (argümansız → pencere, argümanla → komut satırı)
tests/duman_verisi.py   .exe duman testi için gerçek PDF'li sentetik firma klasörleri (düzenli ve karışık adlı)
```
