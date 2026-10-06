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
4. `.exe`'ye çift tıklayın: ayar dosyası, firma klasörü ve dönem seçilip **Taslak oluştur**'a basılır. Pencerede
   ayrıca "Gerçek raporla karşılaştır", "Rapor denetle" ve "Geriye dönük test" düğmeleri vardır.

`.exe` komut satırından da çalışır (`TeminatCozum.exe taslak --ayar ... --kok ... --donem 2026-03`); konsolu olmadığı
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

### Aylık taslak

Firma klasörü standart düzendeyse (bkz. aşağı) dönem vermek yeterli; şablon olarak bir önceki ayın
`RAPOR\` klasöründeki rapor kullanılır:

```
python -m teminat taslak --ayar firmalar\<firma>\ayar.yaml --kok "C:\...\TEMİNAT ÇÖZÜMÜ\<FİRMA>" --donem 2026-03
```

Şablonu ve girdi klasörünü ayrı ayrı da verebilirsiniz:

```
python -m teminat taslak --ayar firmalar\<firma>\ayar.yaml --sablon "...\02 ŞUBAT\RAPOR\ŞUBAT-2026 RAPOR.doc" --girdi "...\2026\03 MART"
```

Çıktılar (girdi klasörünün içinde, `CLAUDE TASLAK\` altında; aynı adlı dosya varsa üzerine yazılmaz, `(2)` eklenir):

| Dosya | İçerik |
|---|---|
| `MART-2026 RAPOR TASLAK.docx` | Rapor taslağı. Elle doldurulacak yerler ve kontrol edilecek satırlar sarı. |
| `MART-2026 KONTROL LİSTESİ.xlsx` | Kontrol listesi + özet + safha (EKLİ seçimi) sayfası |
| `MART-2026 KONTROL LİSTESİ.html` | Aynı kontrol listesi, tarayıcıda açılır (durum filtresiyle) |

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

Firma klasöründeki her ay için taslağı bir önceki ayın gerçek raporundan üretir ve o ayın gerçek raporuyla
karşılaştırır:

```
python -m teminat geriye-donuk --ayar firmalar\<firma>\ayar.yaml --kok "C:\...\TEMİNAT ÇÖZÜMÜ\<FİRMA>" --baslangic 2025-01 --bitis 2026-02
```

Sonuç: `<FİRMA>\CLAUDE TASLAK\GERİYE DÖNÜK TEST\GERİYE DÖNÜK TEST.xlsx` (dönem başına tutar farkı, farklı satır
sayısı, HATA/UYARI sayısı) ve her ay için `FARKLAR.txt`. Girdi klasörlerine yazılmaz.

### Bitmiş bir raporu denetleme

```
python -m teminat denetle "...\ARALIK-2025 RAPOR.doc"
```

3-4-3 safha tablosunda satır toplamı ↔ TOPLAM hücreleri, EKLİ sıra numaraları, aynı tutarın iki firmada
görünmesi (satır kayması) ve %80 oranını kontrol eder.

## Girdi klasörü standardı

```
TEMİNAT ÇÖZÜMÜ\<FİRMA>\2026\02 ŞUBAT\
    KDV 1.pdf                                   ← 1 No.lu KDV beyannamesi (zorunlu)
    internetvd_kdviadesi_indirilecekkdvListesi_FORMATI.xls   (veya "İndirilecek KDV listesi ... .xls") (zorunlu)
    yüklenilen tutanak çalışması.xls
    TEMİNAT MEKTUBU KABUL DİLEKÇESİ - ŞUBAT 2026.docx
    TUTANAK ÇALIŞMASI\01 FİRMA VE MUH. BİLGİLERİ ŞUBAT 2026.xls   ← karşıt inceleme takip listesi (zorunlu)
    RAPOR\ŞUBAT-2026 RAPOR.doc                  ← ofisin bitmiş raporu (bir sonraki ayın şablonu)
```

Dosya adları küçük farklarla değişebilir: aramalar Türkçe karakter, büyük/küçük harf ve macOS/OneDrive kaynaklı
NFD farkını gözetmez. Bir dosya için birden fazla aday varsa ilki seçilir ve kontrol listesine yazılır.

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
  karsilastir.py, geriye_donuk.py, klasorler.py, donusum.py (.doc → .docx), ayar.py, cli.py, arayuz.py (pencere)
TeminatCozum.py   .exe giriş noktası (argümansız → pencere, argümanla → komut satırı)
tests/duman_verisi.py   .exe duman testi için gerçek PDF'li sentetik firma klasörü
```
