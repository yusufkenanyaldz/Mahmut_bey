"""Belge tanıma: dosya adına ve klasör düzenine bakmadan, İÇERİKTEN (PROJE_TALIMATI.md §13).

Her dosyanın ilk ~1500 karakteri okunur (PDF 1. sayfa, Excel ilk 12 satır, Word ilk paragraflar / tablolar) ve
`KURALLAR` ile türü bulunur. Kurallar `src/tani.py`'den (5 firmanın 455 dosyasında denenmiş) aynen alınmıştır.

Bir türden birden fazla dosya varsa hangisinin kullanılacağı yine içerikten seçilir (beyannamenin dönemi, listenin dönem
sütunu, tutarların beyanla tutması); içerikten karar verilemezse kullanıcıya sorulur (`BelirsizSecim`).
Kullanıcı bir dosyanın türünü elle düzeltebilir; düzeltme firma klasöründeki `belge_turleri.yaml`'a yazılır ve
sonraki aylarda (dosya adındaki ay/yıl değişse de) hatırlanır.
"""
import collections
import hashlib
import json
import os
import re
import subprocess
import tempfile
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

from .ortak import AYLAR, Donem, excel_satirlari, katla, tr

ILK = 1500
# (tür, açıklama, regex) — sıra önemli: ilk eşleşen kazanır (src/tani.py ile aynı)
KURALLAR = [
    ('KDV1', '1 No.lu KDV beyannamesi', r'KATMA DEĞER VERGİSİ BEYANNAMESİ.*Gerçek Usulde'),
    ('KDV2', '2 No.lu KDV beyannamesi (sorumlu)', r'KATMA DEĞER VERGİSİ BEYANNAMESİ.*Vergi Sorumluları'),
    ('GIB_KONTROL', 'GİB KDV İadesi Kontrol Raporu (Özet)', r'KDV Iadesi Kontrol Raporu'),
    ('LISTE_INDIRILECEK', 'İndirilecek KDV listesi', r'(?i)indirilecek kdv list|Toplam İndirilen KDV|Tevkifata Tabi Olmayan Ve Bu Dönemde İndirilen'),
    ('LISTE_YUKLENILEN', 'Yüklenilen KDV listesi / tutanak çalışması', r'(?i)yüklenilen kdv list|Bünyeye Giren'),
    ('LISTE_INDIRILECEK', 'İndirilecek KDV listesi', r'Alış Faturasının Tarihi.*Satıcının'),
    ('LISTE_TEVKIFATLI', 'Tevkifatlı satış faturaları listesi', r'KISMİ TEVKİFAT|TEVKİFAT UYGULAMASI'),
    ('LISTE_GCB', 'Gümrük çıkış beyannameleri listesi', r'GÜMRÜK ÇIKIŞ BEYANNAMELERİ LİSTESİ'),
    ('LISTE_IHRAC_KAYITLI', 'İhraç kayıtlı satış faturaları listesi',
     r'İHRAÇ KAYITLI SATIŞ FATURASI LİSTESİ|İHRAÇ KAYITLI SATIŞLAR|DİİB KAPSAMINDA SATIŞ FATURA|DAHİLDE İŞLEME İZİN BELGESİ KA'),
    ('LISTE_13MADDE', '13. madde satış listesi', r'MADDE KAPSAMINDA YAPILAN SATIŞLARA'),
    ('LISTE_SATIS', 'Satış faturası listesi', r'SATIŞ FATURASI LİSTESİ|SATIŞ FATURALARI LİSTESİ'),
    ('TAKIP', 'Karşıt inceleme takip listesi (firma/muhasebeci)', r'FİRMA\s.*(KDV\s)?AÇIKLAMA SMMM YMM'),
    ('DIIB_TAKIP', 'DİİB takip listesi', r'DAHİLDE İŞLEME İZİ'),
    ('DIIB_HESAP', 'DİİB hesap tablosu', r'KDV Ödenmeksizin Temin Edilen Mal Bedeli'),
    ('SARFIYAT', 'DİİB sarfiyat tablosu', r'Toplam İhtiyaç İthalat Talebi'),
    ('KUR_FARKI', 'Kur farkı tablosu', r'KUR FARKI|Fatura Tutar Bilgi Cari Ünvan'),
    ('SGK_LISTE', 'SGK/işçi listesi', r'(?i)\bssk\b|\bsgk\b|FİRMA ADI: .*FİRMA ADI:'),
    ('IPTAL_DILEKCE', 'İptal / istisna iptali dilekçesi', r'(?i)iptal'),
    ('IMALAT', 'İmalat tablosu', r'İMALAT TABLOSU'),
    ('MIZAN', 'KDV mizanı / muavin', r'KDV MİZAN|M U A V İ N'),
    ('RAPOR', 'Teminat çözüm / tasdik raporu', r'GENEL BİLGİ.{0,20}Raporun amacı'),
    ('KIT', 'Karşıt inceleme tutanağı', r'KARŞIT İNCELEME TUTANAĞI'),
    ('KIT_DEVAM', 'Tutanak devam sayfası (alt firmalar)', r'ALT FİRMALARIN DÖK'),
    ('YMM_YAZISI', 'YMM bilgi isteme yazısı', r'Konu\s*:?\s*Bilgi İsteme'),
    ('EKSIKLIK', 'Eksiklik yazısı', r'(?i)eksik tamamlama|eksiklik|Konu\s*:\s*Eksik'),
    ('TEMINAT_DILEKCE', 'Teminat mektubu kabul dilekçesi',
     r'(?i)teminat mektub.*(kabul|talep edilmiştir)|Artırımlı Temin|İndirimli Temin|teminat kabul'),
    ('RAPOR_KABUL', 'Rapor/ek kabul dilekçesi', r'Sayı\s*:\s*YMM.*VERGİ DAİRESİ MÜDÜRLÜĞÜ'),
    ('MAHSUP_DILEKCE', 'Mahsup dilekçesi', r'(?i)mahsup'),
    ('SOZLESME', 'Tasdik sözleşmesi', r'SÖZLEŞME|YMM VE BAĞ'),
    ('FATURA_DOKUM', 'Firma fatura dökümü', r'Faturanın Tarihi Faturanın Serisi'),
]
EK_TURLER = {
    'TARANMIS': 'Metni olmayan (taranmış) PDF / görüntü',
    'OKUNAMADI': 'Dosya açılamadı',
    'BILINMEYEN': 'Tanınamadı',
    'DESTEKLENMEYEN': 'Okunmayan dosya türü',
    'YOK_SAY': 'Kullanıcı: yok sayılsın',
}
ACIKLAMA = {**{k: a for k, a, _ in KURALLAR}, **EK_TURLER}
TUR_LISTESI = list(dict.fromkeys([k for k, _, _ in KURALLAR] + list(EK_TURLER)))
KURAL_SURUMU = hashlib.sha1(json.dumps(KURALLAR, ensure_ascii=False).encode()).hexdigest()[:10]

# Raporda kullanılan roller (sıra = tabloda gösterim sırası)
ROLLER = {
    'KDV1': ('1 No.lu KDV beyannamesi', True),
    'LISTE_INDIRILECEK': ('İndirilecek KDV listesi', True),
    'TAKIP': ('Karşıt inceleme takip listesi', True),
    'TEMINAT_DILEKCE': ('Teminat mektubu kabul dilekçesi', True),
    'LISTE_YUKLENILEN': ('Yüklenilen KDV tutanak çalışması', False),
}
BILGI_TURLERI = ['KDV2', 'LISTE_TEVKIFATLI', 'LISTE_IHRAC_KAYITLI', 'LISTE_GCB', 'IMALAT', 'DIIB_TAKIP', 'DIIB_HESAP', 'GIB_KONTROL']

GORUNTU = {'.jpg', '.jpeg', '.png', '.tif', '.tiff', '.bmp', '.gif'}
OKUNAN = {'.pdf', '.xls', '.xlsx', '.xlsm', '.docx', '.doc'}
ATLANAN_DOSYALAR = {'thumbs.db', 'desktop.ini', '.ds_store'}
ATLANAN_KLASORLER = {'CLAUDE TASLAK', 'CLAUDE OUTPUTS', '_CEVRILEN', 'GERIYE DONUK TEST', '__MACOSX'}
COK_DOSYA = 3000           # bir "ay klasörü" için makul üst sınır


class BelirsizSecim(Exception):
    """İçerikten karar verilemedi; kullanıcı seçmeli."""

    def __init__(self, rol, adaylar, aciklama):
        super().__init__(aciklama)
        self.rol, self.adaylar, self.aciklama = rol, list(adaylar), aciklama


class TanimaHatasi(Exception):
    pass


def nf(s):
    return unicodedata.normalize('NFC', s)


def atlanir(ad):
    a = ad.lower()
    return ad.startswith(('~', '._', '.~lock')) or a in ATLANAN_DOSYALAR or a.endswith('.tmp')


def dosyalari_listele(klasor, en_cok=COK_DOSYA):
    klasor = Path(klasor)
    out = []
    for kok, dirs, files in os.walk(klasor):
        dirs[:] = sorted(d for d in dirs if katla(d).upper() not in ATLANAN_KLASORLER and not d.startswith('.'))
        for f in sorted(files):
            if not atlanir(f):
                out.append(Path(kok) / f)
                if len(out) > en_cok:
                    raise TanimaHatasi(f'{klasor} içinde {en_cok}\'den fazla dosya var; bu bir ayın klasörü gibi görünmüyor. '
                                       'Firmanın o ayki klasörünü seçin.')
    return out


# ---------------------------------------------------------------- önbellek (yalnızca tür ve dönem saklanır, metin saklanmaz)
class Onbellek:
    def __init__(self, yol=None):
        self.yol = Path(yol) if yol else Path.home() / '.teminat_cozum' / 'tanima_onbellek.json'
        try:
            self.veri = json.loads(self.yol.read_text(encoding='utf-8'))
        except (OSError, ValueError):
            self.veri = {}
        self.degisti = False

    @staticmethod
    def anahtar(p, ek=''):
        try:
            st = Path(p).stat()
        except OSError:
            return None
        return f'{KURAL_SURUMU}|{ek}|{Path(p).resolve()}|{st.st_size}|{st.st_mtime_ns}'

    def al(self, p, ek=''):
        a = self.anahtar(p, ek)
        return self.veri.get(a) if a else None

    def koy(self, p, deger, ek=''):
        a = self.anahtar(p, ek)
        if a:
            self.veri[a] = deger
            self.degisti = True

    def kaydet(self):
        if not self.degisti:
            return
        try:
            self.yol.parent.mkdir(parents=True, exist_ok=True)
            if len(self.veri) > 50000:
                self.veri = dict(list(self.veri.items())[-30000:])
            self.yol.write_text(json.dumps(self.veri, ensure_ascii=False), encoding='utf-8')
        except OSError:
            pass


class _Hafiza(Onbellek):
    def __init__(self):
        self.yol, self.veri, self.degisti = None, {}, False

    def kaydet(self):
        pass


# ---------------------------------------------------------------- içerik okuma
def _excel_metni(p):
    out = []
    sayfa_sayisi = collections.Counter()
    for sayfa, r in excel_satirlari(p, en_cok=12):
        sayfa_sayisi[sayfa] += 1
        if len(sayfa_sayisi) > 3:
            break
        out.append(' '.join(str(c) for c in r if c is not None and str(c).strip()))
    return '\n'.join(out)[:ILK]


def _docx_metni(p, en_cok=ILK):
    from docx import Document
    d = Document(str(p))
    t = [x.text for x in d.paragraphs]
    for tb in d.tables[:2]:
        for r in tb.rows[:6]:
            t.append(' '.join(c.text for c in r.cells))
    return '\n'.join(t)[:en_cok]


def _pdf_metni(p):
    from .okuyucular import kdv1
    return kdv1.pdf_ilk_sayfa(p)[:ILK]


def doc_metinleri(yollar, zaman_asimi=600):
    """Eski .doc dosyalarını tek LibreOffice çağrısıyla metne çevirir: {yol: metin}."""
    from .donusum import soffice_yolu
    so = soffice_yolu()
    out = {}
    if not so or not yollar:
        return out
    kalan = list(yollar)
    while kalan:
        parti, kokler, sonraki = [], set(), []
        for p in kalan:                            # aynı adlı dosyalar ayrı partilerde (çıktılar çakışmasın)
            if p.stem in kokler or len(parti) >= 60:
                sonraki.append(p)
            else:
                parti.append(p)
                kokler.add(p.stem)
        with tempfile.TemporaryDirectory() as hedef, tempfile.TemporaryDirectory() as profil:
            komut = [so, f'-env:UserInstallation={Path(profil).as_uri()}', '--headless', '--convert-to',
                     'txt:Text (encoded):UTF8', '--outdir', hedef, *map(str, parti)]
            try:
                subprocess.run(komut, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=zaman_asimi)
            except (subprocess.TimeoutExpired, OSError):
                pass
            for p in parti:
                t = Path(hedef) / (p.stem + '.txt')
                if t.exists():
                    out[p] = t.read_text(encoding='utf-8', errors='ignore')[:ILK]
        kalan = sonraki
    return out


def siniflandir(metin):
    t = re.sub(r'\s+', ' ', metin)
    if not t.strip():
        return 'TARANMIS'
    for k, _, rx in KURALLAR:
        if re.search(rx, t):
            return k
    return 'BILINMEYEN'


# ---------------------------------------------------------------- elle tür düzeltmeleri
def ad_anahtari(ad):
    """Dosya adından ay/yıl/sayılar çıkarılmış anahtar: '01 FİRMA ... ŞUBAT 2026.xls' ile '... MART 2026.xls' aynı."""
    n = katla(nf(ad)).upper()
    for a in AYLAR:
        n = re.sub(rf'(?<![A-Z]){katla(a).upper()}(?![A-Z])', ' ', n)
    n = re.sub(r'\d+', ' ', n)
    return ' '.join(n.split())


class TurDuzeltmeleri:
    """Kullanıcının elle verdiği türler: {ad anahtarı: TÜR}. Firma ayar dosyasının yanındaki belge_turleri.yaml."""

    def __init__(self, yol=None):
        self.yol = Path(yol) if yol else None
        self.veri = {}
        if self.yol and self.yol.exists():
            import yaml
            try:
                self.veri = {str(k): str(v) for k, v in (yaml.safe_load(self.yol.read_text(encoding='utf-8')) or {}).items()}
            except Exception:  # noqa: BLE001 - bozuk dosya: düzeltmesiz devam
                self.veri = {}

    def tur(self, p):
        return self.veri.get(ad_anahtari(Path(p).name))

    def ayarla(self, p, tur):
        if tur not in TUR_LISTESI:
            raise ValueError(f'Bilinmeyen tür: {tur}')
        self.veri[ad_anahtari(Path(p).name)] = tur
        if self.yol:
            import yaml
            self.yol.parent.mkdir(parents=True, exist_ok=True)
            bas = ('# Belge türü düzeltmeleri (programın pencereden kaydettiği). Anahtar: dosya adından ay, yıl ve sayılar\n'
                   '# çıkarılmış hali; böylece sonraki aylarda da geçerli olur. Tür adları: ' + ', '.join(TUR_LISTESI) + '\n')
            self.yol.write_text(bas + yaml.safe_dump(self.veri, allow_unicode=True, sort_keys=True), encoding='utf-8')


# ---------------------------------------------------------------- tanıma
@dataclass
class Belge:
    yol: Path
    goreli: str
    tur: str
    elle: bool = False
    donem: Donem = None
    rol: str = ''            # raporda kullanılıyorsa rol adı
    notu: str = ''

    @property
    def aciklama(self):
        return ACIKLAMA.get(self.tur, '')


def tani(klasor, duzeltmeler=None, onbellek=None, ilerleme=None):
    """Klasördeki (alt klasörler dahil) bütün dosyaları içerikten tanır: [Belge]."""
    klasor = Path(klasor)
    onbellek = onbellek if onbellek is not None else Onbellek()
    duzeltmeler = duzeltmeler or TurDuzeltmeleri()
    yollar = dosyalari_listele(klasor)
    belgeler, doc_bekleyen = [], []
    for i, p in enumerate(yollar):
        if ilerleme and i % 50 == 0:
            ilerleme(f'{i}/{len(yollar)} dosya okunuyor…')
        b = Belge(p, nf(str(p.relative_to(klasor))), 'BILINMEYEN')
        belgeler.append(b)
        elle = duzeltmeler.tur(p)
        if elle:
            b.tur, b.elle = elle, True
            continue
        c = onbellek.al(p)
        if c:
            b.tur = c['tur']
            continue
        uz = p.suffix.lower()
        if uz in GORUNTU:
            b.tur = 'TARANMIS'
        elif uz not in OKUNAN:
            b.tur = 'DESTEKLENMEYEN'
        elif uz == '.doc':
            doc_bekleyen.append(b)
            continue
        else:
            try:
                if uz == '.pdf':
                    metin = _pdf_metni(p)
                elif uz in ('.xls', '.xlsx', '.xlsm'):
                    metin = _excel_metni(p)
                else:
                    metin = _docx_metni(p)
                b.tur = siniflandir(nf(metin))
            except Exception as e:  # noqa: BLE001
                b.tur, b.notu = 'OKUNAMADI', f'{type(e).__name__}: {str(e)[:80]}'
        if b.tur != 'OKUNAMADI':
            onbellek.koy(p, {'tur': b.tur})
    if doc_bekleyen:
        if ilerleme:
            ilerleme(f'{len(doc_bekleyen)} eski Word (.doc) dosyası LibreOffice ile okunuyor…')
        metinler = doc_metinleri([b.yol for b in doc_bekleyen])
        for b in doc_bekleyen:
            if b.yol in metinler:
                b.tur = siniflandir(nf(metinler[b.yol]))
                onbellek.koy(b.yol, {'tur': b.tur})
            else:
                b.tur, b.notu = 'OKUNAMADI', 'LibreOffice ile açılamadı (LibreOffice kurulu mu?)'
    onbellek.kaydet()
    return belgeler


# ---------------------------------------------------------------- dönem okuma yardımcıları
_AY_RX = '|'.join(katla(a).upper() for a in AYLAR)


def metindeki_donemler(metin):
    """Metinde geçen 'ŞUBAT/2026', 'Şubat 2026', '2026/Şubat', '02/2026' dönemleri (Counter)."""
    n = katla(metin).upper()
    c = collections.Counter()
    for a, y in re.findall(rf'(?<![A-Z])({_AY_RX})\s*[-/., ]\s*(20\d\d)(?!\d)', n):
        c[Donem(int(y), [katla(x).upper() for x in AYLAR].index(a) + 1)] += 1
    for y, a in re.findall(rf'(?<!\d)(20\d\d)\s*[-/., ]\s*({_AY_RX})(?![A-Z])', n):
        c[Donem(int(y), [katla(x).upper() for x in AYLAR].index(a) + 1)] += 1
    for a, y in re.findall(r'(?<![\d.])(0?[1-9]|1[0-2])\s*/\s*(20\d\d)(?!\d)', n):
        c[Donem(int(y), int(a))] += 1
    return c


def rapor_bilgisi(p, onbellek=None):
    """Teminat çözüm raporunun kapağından dönem ve mükellef hesap no (VKN): {'donem': Donem|None, 'vkn': str|None}."""
    onbellek = onbellek if onbellek is not None else _Hafiza()
    c = onbellek.al(p, 'rapor')
    if c is not None:
        return {'donem': Donem.coz(c['donem']) if c.get('donem') else None, 'vkn': c.get('vkn')}
    from docx import Document

    from .donusum import DonusumHatasi, docx_hazirla
    donem = vkn = None
    try:
        with tempfile.TemporaryDirectory() as tmp:
            d = Document(str(docx_hazirla(p, tmp)))
            for t in d.tables[:2]:
                for r in t.rows:
                    hucre = [c.text.strip() for c in r.cells]
                    satir = ' | '.join(hucre)
                    if donem is None and 'İncelemenin Dönemi' in satir:
                        m = re.search(r'([A-ZÇĞİÖŞÜ]+)-(\d{4})', hucre[-1])
                        if m:
                            try:
                                donem = Donem.coz(f'{m.group(1)}-{m.group(2)}')
                            except ValueError:
                                pass
                    if vkn is None and re.search(r'Hesap No|Vergi Kimlik|VKN', satir) and 'Mükellef' in satir:
                        rakam = re.sub(r'\D', '', hucre[-1])
                        if len(rakam) in (10, 11):
                            vkn = rakam
    except (DonusumHatasi, Exception):  # noqa: BLE001 - okunamayan rapor: bilgi yok
        return {'donem': None, 'vkn': None}
    onbellek.koy(p, {'donem': donem.tire if donem else None, 'vkn': vkn}, 'rapor')
    return {'donem': donem, 'vkn': vkn}


def metindeki_vknler(metin):
    """Metinde geçen 10-11 haneli numaralar (aralarında tek boşluk olabilir: '380 119 8516')."""
    return {re.sub(r'\D', '', x) for x in re.findall(r'(?<![\d])(\d(?:[ ]?\d){9,10})(?![\d])', metin or '')}


def vkn_metinde(vkn, metin):
    return bool(vkn) and vkn in metindeki_vknler(metin)


# ---------------------------------------------------------------- rol seçimi
@dataclass
class Secim:
    klasor: Path
    belgeler: list
    yollar: dict = field(default_factory=dict)        # rol → Path
    donem: Donem = None
    k: dict = None                                      # okunmuş KDV 1
    bu_ayin_raporu: Path = None
    notlar: list = field(default_factory=list)         # [(durum, konu, açıklama)]

    def eksik_zorunlu(self):
        return [r for r, (_, z) in ROLLER.items() if z and r not in self.yollar]


def _adaylar(belgeler, tur):
    return [b for b in belgeler if b.tur == tur]


def _goreli(b):
    return b.goreli


def kdv1_sec(belgeler, istenen=None, klasor_ayi=None, zorla=None):
    """Okunabilen 1 No.lu beyannameler arasından seçer: (Belge, k, notlar)."""
    from .okuyucular.kdv1 import kdv1_donem, kdv1_metinden, pdf_metni
    notlar = []
    adaylar = [b for b in _adaylar(belgeler, 'KDV1') if zorla is None or b.yol == Path(zorla)]
    if zorla is not None and not adaylar:
        adaylar = [Belge(Path(zorla), Path(zorla).name, 'KDV1', elle=True)]
    okunan = []
    for b in adaylar:
        try:
            k = kdv1_metinden(pdf_metni(b.yol), b.yol)
        except Exception as e:  # noqa: BLE001
            b.notu = f'okunamadı: {type(e).__name__}'
            continue
        yd = kdv1_donem(k)
        if yd and ('yurtici_alim' in k or 'matrah_toplam' in k):
            b.donem = Donem(*yd)
            okunan.append((b, k))
        else:
            b.notu = 'beyanname olarak tanındı ama dönem/tutarları okunamadı'
    if not okunan:
        if adaylar:
            raise TanimaHatasi('1 No.lu KDV beyannamesi olarak tanınan dosyalar okunamadı: '
                               + ', '.join(_goreli(b) for b in adaylar) + '. Beyanname PDF\'i metin içermeli (taranmış olmamalı).')
        raise TanimaHatasi('Klasörde 1 No.lu KDV beyannamesi bulunamadı (hiçbir PDF\'te "KATMA DEĞER VERGİSİ BEYANNAMESİ … '
                           'Gerçek Usulde" yazısı yok). Beyannameyi klasöre koyun ya da tanıma tablosunda türünü düzeltin.')
    if istenen:
        uyan = [x for x in okunan if x[0].donem == istenen]
        if not uyan:
            raise TanimaHatasi(f'Klasörde {istenen.tire} dönemine ait beyanname yok; bulunanlar: '
                               + ', '.join(f'{b.donem.tire} ({_goreli(b)})' for b, _ in okunan))
        okunan = uyan
    donemler = sorted({b.donem for b, _ in okunan})
    if len(donemler) > 1:
        if klasor_ayi and [d for d in donemler if d.ay == klasor_ayi[0] and (klasor_ayi[1] in (None, d.yil))]:
            secilen_d = max(d for d in donemler if d.ay == klasor_ayi[0] and (klasor_ayi[1] in (None, d.yil)))
            neden = 'klasör adındaki ayla uyan'
        else:
            secilen_d = donemler[-1]
            neden = 'en son dönem'
        notlar.append(('UYARI', 'Belge seçimi', f'Klasörde birden fazla dönemin beyannamesi var ({", ".join(d.tire for d in donemler)}); '
                       f'{neden} {secilen_d.tire} seçildi. Başka bir ay için o ayın klasörünü seçin.'))
        okunan = [x for x in okunan if x[0].donem == secilen_d]
    if len(okunan) > 1:
        imza = {(k.get('matrah_toplam'), k.get('iade_gereken'), k.get('sonraki_devreden'), k.get('duzeltme')) for _, k in okunan}
        if len(imza) > 1:
            raise BelirsizSecim('KDV1', [b.yol for b, _ in okunan],
                                f'Aynı dönemin ({okunan[0][0].donem.tire}) içeriği farklı {len(okunan)} beyannamesi var '
                                '(ör. asıl ve düzeltme beyannamesi). Hangisinin kullanılacağını seçin.')
        notlar.append(('BİLGİ', 'Belge seçimi', f'Aynı beyannamenin {len(okunan)} kopyası var; {_goreli(okunan[0][0])} kullanıldı.'))
    return okunan[0][0], okunan[0][1], notlar


def _liste_sec(adaylar, donem, base, notlar, zorla=None):
    from .okuyucular.listeler import donem_degeri, indirilecek_oku
    if zorla:
        adaylar = [b for b in adaylar if b.yol == Path(zorla)] or [Belge(Path(zorla), Path(zorla).name, 'LISTE_INDIRILECEK', True)]
    puan = []
    for b in adaylar:
        try:
            rows = indirilecek_oku(b.yol)
        except Exception as e:  # noqa: BLE001
            b.notu = f'okunamadı: {type(e).__name__}'
            continue
        if not rows:
            b.notu = 'satır okunamadı (sütun düzeni farklı olabilir)'
            continue
        donemler = collections.Counter(donem_degeri(r['donem']) for r in rows)
        uygun = donemler.get((donem.yil, donem.ay), 0) / len(rows)
        en_cok = donemler.most_common(1)[0][0]
        if en_cok:
            b.donem = Donem(*en_cok)
        top = sum(r['kdv'] for r in rows)
        b.notu = f'{len(rows)} satır, toplam {tr(top)}'
        puan.append(((round(uygun, 3), -round(abs(top - base), 2)), b, top))
    if not puan:
        return None
    puan.sort(key=lambda x: x[0], reverse=True)
    if len(puan) > 1:
        if puan[0][0] == puan[1][0] and abs(puan[0][2] - puan[1][2]) > 0.01:
            raise BelirsizSecim('LISTE_INDIRILECEK', [x[1].yol for x in puan if x[0] == puan[0][0]],
                                'Birden fazla indirilecek KDV listesi var ve dönem/toplam ile ayırt edilemedi. Hangisi bu aya ait?')
        notlar.append(('BİLGİ', 'Belge seçimi', f'{len(puan)} indirilecek KDV listesinden dönem sütunu {donem.tire} olan ve toplamı '
                       f'beyana en yakın olan seçildi: {_goreli(puan[0][1])}. Diğerleri: '
                       + ', '.join(f'{_goreli(x[1])} ({x[1].donem.tire if x[1].donem else "dönem ?"}, {tr(x[2])})' for x in puan[1:])))
    return puan[0][1]


def _takip_sec(adaylar, liste_satirlari, notlar, zorla=None):
    from .okuyucular.listeler import takip_oku
    from .rapor.safha import firma_anahtari
    if zorla:
        adaylar = [b for b in adaylar if b.yol == Path(zorla)] or [Belge(Path(zorla), Path(zorla).name, 'TAKIP', True)]
    by = collections.defaultdict(float)
    for r in liste_satirlari:
        by[firma_anahtari(r['satici'])] += r['kdv']
    tutarlar = {round(v, 2) for v in by.values()}
    puan = []
    for b in adaylar:
        try:
            t = takip_oku(b.yol)
        except Exception as e:  # noqa: BLE001
            b.notu = f'okunamadı: {type(e).__name__}'
            continue
        if not t:
            b.notu = 'firma satırı okunamadı'
            continue
        uyan = sum(1 for x in t if round(x['kdv'], 2) in tutarlar or abs(by.get(firma_anahtari(x['firma']), -1) - x['kdv']) < 0.01)
        b.notu = f'{len(t)} firma, {uyan} tanesi indirilecek listeyle tutuyor'
        puan.append((uyan / len(t), b, sum(x['kdv'] for x in t)))
    if not puan:
        return None
    puan.sort(key=lambda x: x[0], reverse=True)
    if len(puan) > 1:
        if puan[0][0] == puan[1][0] and abs(puan[0][2] - puan[1][2]) > 0.01:
            raise BelirsizSecim('TAKIP', [x[1].yol for x in puan if x[0] == puan[0][0]],
                                'Birden fazla karşıt inceleme takip listesi var ve indirilecek listeyle eşit derecede tutuyor. Hangisi bu aya ait?')
        notlar.append(('BİLGİ', 'Belge seçimi', f'{len(puan)} takip listesinden indirilecek listeyle en çok tutan seçildi: {_goreli(puan[0][1])}.'))
    return puan[0][1]


def dilekce_metni(p):
    from .donusum import docx_hazirla
    with tempfile.TemporaryDirectory() as tmp:
        return _docx_metni(docx_hazirla(p, tmp), en_cok=None)


def _dilekce_sec(adaylar, donem, notlar, zorla=None):
    from .okuyucular.teminat import dilekce_metinden
    if zorla:
        adaylar = [b for b in adaylar if b.yol == Path(zorla)] or [Belge(Path(zorla), Path(zorla).name, 'TEMINAT_DILEKCE', True)]
    puan = []
    for b in adaylar:
        try:
            metin = dilekce_metni(b.yol)
        except Exception as e:  # noqa: BLE001
            b.notu = f'okunamadı: {type(e).__name__}'
            continue
        tem = dilekce_metinden(metin)
        donemler = metindeki_donemler(metin)
        if donemler:
            b.donem = donemler.most_common(1)[0][0]
        p = (2 if donem in donemler else 0) + (1 if tem.get('tarih') else 0) + (1 if tem.get('toplam') else 0)
        puan.append((p, b, tem.get('toplam')))
    if not puan:
        return None
    puan.sort(key=lambda x: x[0], reverse=True)
    if len(puan) > 1:
        if puan[0][0] == puan[1][0] and puan[0][2] != puan[1][2]:
            raise BelirsizSecim('TEMINAT_DILEKCE', [x[1].yol for x in puan if x[0] == puan[0][0]],
                                'Birden fazla teminat dilekçesi var ve içeriğinden hangisinin bu aya ait olduğu anlaşılamadı.')
        notlar.append(('BİLGİ', 'Belge seçimi', f'{len(puan)} teminat dilekçesinden {donem.tire} dönemini anan seçildi: {_goreli(puan[0][1])}.'))
    if puan[0][0] < 2 and puan[0][1].donem and puan[0][1].donem != donem:
        notlar.append(('UYARI', 'Belge dönemi', f'Teminat dilekçesi {puan[0][1].donem.tire} dönemini anıyor, beyanname {donem.tire}: {_goreli(puan[0][1])}'))
    return puan[0][1]


def _yuklenilen_sec(adaylar, hedef, notlar, zorla=None):
    from .okuyucular.listeler import yuklenilen_oku
    if zorla:
        adaylar = [b for b in adaylar if b.yol == Path(zorla)] or [Belge(Path(zorla), Path(zorla).name, 'LISTE_YUKLENILEN', True)]
    puan = []
    for b in adaylar:
        try:
            y = yuklenilen_oku(b.yol)
        except Exception as e:  # noqa: BLE001
            b.notu = f'okunamadı: {type(e).__name__}'
            continue
        if not y:
            b.notu = 'satır okunamadı (tutanak çalışması biçiminde değil)'
            continue
        top = sum(x['kdv'] for x in y)
        b.notu = f'{len(y)} satır, toplam {tr(top)}'
        puan.append((abs(top - hedef), b, top))
    if not puan:
        return None
    puan.sort(key=lambda x: x[0])
    if puan[0][0] > 0.01 and not zorla:
        notlar.append(('UYARI', 'Belge seçimi', f'Hiçbir yüklenilen tutanak dosyasının toplamı beyandaki 301 (+339) yüklenilen KDV\'ye '
                       f'({tr(hedef)}) eşit değil; en yakını seçildi: {_goreli(puan[0][1])} ({tr(puan[0][2])}).'))
    elif len(puan) > 1:
        notlar.append(('BİLGİ', 'Belge seçimi', f'{len(puan)} yüklenilen listesinden toplamı beyana eşit olan seçildi: {_goreli(puan[0][1])}.'))
    return puan[0][1]


def secim_yap(klasor, belgeler, istenen=None, zorla=None):
    """Tanınan belgelerden raporda kullanılacakları seçer (zorla: {rol: yol} kullanıcı seçimleri)."""
    from .klasorler import ay_coz
    from .okuyucular.listeler import indirilecek_oku
    zorla = zorla or {}
    s = Secim(Path(klasor), belgeler)
    klasor_ayi = ay_coz(Path(klasor).name)
    b, k, n = kdv1_sec(belgeler, istenen, klasor_ayi, zorla.get('KDV1'))
    s.notlar += n
    b.rol, s.yollar['KDV1'], s.donem, s.k = 'KDV1', b.yol, b.donem, k
    if klasor_ayi and (klasor_ayi[0] != s.donem.ay or (klasor_ayi[1] and klasor_ayi[1] != s.donem.yil)):
        s.notlar.append(('UYARI', 'Belge dönemi', f'Klasör adı ("{Path(klasor).name}") başka bir ayı gösteriyor ama beyanname '
                         f'{s.donem.tire} dönemine ait — doğru klasörü seçtiğinizden emin olun.'))
    base = k.get('yurtici_alim', 0) + k.get('sorumlu', 0) + k.get('ithal', 0)
    lb = _liste_sec(_adaylar(belgeler, 'LISTE_INDIRILECEK'), s.donem, base, s.notlar, zorla.get('LISTE_INDIRILECEK'))
    satirlar = []
    if lb:
        lb.rol, s.yollar['LISTE_INDIRILECEK'] = 'LISTE_INDIRILECEK', lb.yol
        satirlar = indirilecek_oku(lb.yol)
    tb = _takip_sec(_adaylar(belgeler, 'TAKIP'), satirlar, s.notlar, zorla.get('TAKIP'))
    if tb:
        tb.rol, s.yollar['TAKIP'] = 'TAKIP', tb.yol
    db = _dilekce_sec(_adaylar(belgeler, 'TEMINAT_DILEKCE'), s.donem, s.notlar, zorla.get('TEMINAT_DILEKCE'))
    if db:
        db.rol, s.yollar['TEMINAT_DILEKCE'] = 'TEMINAT_DILEKCE', db.yol
    hedef = k.get('301_yuklenilen', 0) + k.get('339_iade', 0)
    yuk = _adaylar(belgeler, 'LISTE_YUKLENILEN')
    if hedef or zorla.get('LISTE_YUKLENILEN'):
        yb = _yuklenilen_sec(yuk, hedef, s.notlar, zorla.get('LISTE_YUKLENILEN'))
        if yb:
            yb.rol, s.yollar['LISTE_YUKLENILEN'] = 'LISTE_YUKLENILEN', yb.yol
    elif yuk:
        s.notlar.append(('BİLGİ', 'Belge seçimi', 'Beyanda 301 yüklenilen KDV yok; yüklenilen listeleri kullanılmadı.'))
    # bu ayın bitmiş raporu (varsa: taslak sonrası karşılaştırma için)
    for r in _adaylar(belgeler, 'RAPOR'):
        bilgi = rapor_bilgisi(r.yol)
        r.donem = bilgi['donem']
        if r.donem == s.donem and s.bu_ayin_raporu is None:
            s.bu_ayin_raporu = r.yol
            r.rol = 'BU_AYIN_RAPORU'
    for rol, (ad, zorunlu) in ROLLER.items():
        if rol not in s.yollar and zorunlu:
            s.notlar.append(('HATA', 'Eksik belge', f'{ad} bulunamadı (klasörde içeriği bu türe uyan dosya yok). '
                             'Tanıma tablosunda türü yanlış görünen bir dosya varsa türünü düzeltin.'))
    return s


# ---------------------------------------------------------------- şablon (önceki ayın raporu) arama
def sablon_bul(klasor, onceki, kdv1_metni=None, onbellek=None, en_cok=150, ilerleme=None, en_cok_klasor=4000):
    """Seçilen ay klasörünün çevresindeki (kendisi, üstü ve onun üstü; 4 seviye aşağıya kadar) Word dosyalarından kapağındaki
    dönem `onceki` olan teminat çözüm raporunu içerikten bulur. Kapaktaki mükellef VKN'si beyannamede geçmiyorsa (başka
    firmanın raporu) alınmaz. Dönüş: (yol | None, notlar). Birden çok farklı rapor bulunursa BelirsizSecim."""
    from .klasorler import ay_coz
    klasor = Path(klasor).resolve()
    onbellek = onbellek if onbellek is not None else Onbellek()
    notlar = []
    kokler = [klasor] + [q for q in klasor.parents][:2]
    gorulen, adaylar = set(), []
    gezilen = 0
    for kok in kokler:
        for d, dirs, files in os.walk(kok):
            gezilen += 1
            if gezilen > en_cok_klasor:
                notlar.append(('BİLGİ', 'Şablon arama', f'Çevrede çok fazla klasör var; ilk {en_cok_klasor} klasöre bakıldı.'))
                dirs[:] = []
                break
            dp = Path(d)
            rel = len(dp.relative_to(kok).parts)
            dirs[:] = [x for x in dirs if katla(x).upper() not in ATLANAN_KLASORLER and not x.startswith('.') and rel < 4]
            for f in files:
                p = dp / f
                if atlanir(f) or p.suffix.lower() not in ('.doc', '.docx') or p in gorulen:
                    continue
                gorulen.add(p)
                adaylar.append(p)

    def oncelik(p):
        yol = katla(str(p)).upper()
        parca = [ay_coz(x) for x in p.parent.parts[-3:]]
        ay_uyar = any(x and x[0] == onceki.ay and x[1] in (None, onceki.yil) for x in parca)
        try:
            boy = p.stat().st_size
        except OSError:
            boy = 0
        return (not ay_uyar, 'RAPOR' not in yol, p.suffix.lower() != '.docx', -boy)
    adaylar.sort(key=oncelik)
    bulunan = []
    # türü bilinmeyen .doc dosyaları tek seferde çevrilir (en fazla en_cok kadar)
    taranacak = adaylar[:en_cok]
    if len(adaylar) > en_cok:
        notlar.append(('BİLGİ', 'Şablon arama', f'Çevredeki {len(adaylar)} Word dosyasından öncelikli {en_cok} tanesine bakıldı.'))
    docs = [p for p in taranacak if p.suffix.lower() == '.doc' and onbellek.al(p) is None]
    if docs and ilerleme:
        ilerleme(f'Önceki ayın raporu aranıyor: {len(docs)} eski Word dosyası okunuyor…')
    metinler = doc_metinleri(docs) if docs else {}
    for p in taranacak:
        c = onbellek.al(p)
        if c is None:
            try:
                metin = metinler.get(p) if p.suffix.lower() == '.doc' else _docx_metni(p)
            except Exception:  # noqa: BLE001
                metin = None
            if metin is None:
                continue
            c = {'tur': siniflandir(nf(metin))}
            onbellek.koy(p, c)
        if c['tur'] != 'RAPOR':
            continue
        bilgi = rapor_bilgisi(p, onbellek)
        if bilgi['donem'] == onceki:
            bulunan.append((p, bilgi))
    onbellek.kaydet()
    if kdv1_metni and metindeki_vknler(kdv1_metni):
        ayni = [x for x in bulunan if vkn_metinde(x[1]['vkn'], kdv1_metni)]
        baska = [x for x in bulunan if x[1]['vkn'] and not vkn_metinde(x[1]['vkn'], kdv1_metni)]
        if baska:
            notlar.append(('BİLGİ', 'Şablon arama', f'{len(baska)} rapor başka mükellefe (VKN) ait olduğu için alınmadı: '
                           + ', '.join(str(p) for p, _ in baska[:5])))
        bulunan = ayni or [x for x in bulunan if not x[1]['vkn']]
    if not bulunan:
        return None, notlar
    if len(bulunan) > 1:
        imza = set()
        for p, _ in bulunan:
            try:
                imza.add(p.stat().st_size)
            except OSError:
                imza.add(str(p))
        if len(imza) > 1:
            raise BelirsizSecim('SABLON', [p for p, _ in bulunan],
                                f'{onceki.tire} dönemine ait birden fazla rapor bulundu. Şablon olarak hangisi kullanılsın '
                                '(ofisin son hali)?')
    return bulunan[0][0], notlar


# ---------------------------------------------------------------- tanıma tablosu (metin)
def tablo_metni(secim, sablon=None, sablon_notu=''):
    s = secim
    satirlar = []
    sira = {r: i for i, r in enumerate(list(ROLLER) + ['BU_AYIN_RAPORU'])}
    belgeler = sorted(s.belgeler, key=lambda b: (b.rol == '', sira.get(b.rol, 99), b.tur in EK_TURLER, b.tur, b.goreli))
    gen = max([len(b.goreli) for b in belgeler] + [20])
    gen = min(gen, 70)
    satirlar.append(f'{"TÜR":<20} {"DÖNEM":<12} {"KULLANIM":<16} DOSYA')
    gruplu = collections.Counter()
    for b in belgeler:
        if b.tur in ('BILINMEYEN', 'TARANMIS', 'DESTEKLENMEYEN') and not b.rol and not b.elle:
            gruplu[(b.tur, str(Path(b.goreli).parent))] += 1
            continue
        kullanim = '✔ KULLANILACAK' if b.rol else ''
        if b.rol == 'BU_AYIN_RAPORU':
            kullanim = '✔ karşılaştırma'
        tur = b.tur + (' (elle)' if b.elle else '')
        satirlar.append(f'{tur:<20} {(b.donem.tire if b.donem else ""):<12} {kullanim:<16} {b.goreli}'
                        + (f'   [{b.notu}]' if b.notu else ''))
    for (tur, kl), n in sorted(gruplu.items()):
        satirlar.append(f'{tur:<20} {"":<12} {"":<16} {n} dosya — {kl if kl != "." else "(ana klasör)"}')
    satirlar.append('')
    satirlar.append('--- RAPOR İÇİN DURUM ---')
    satirlar.append(f'Dönem (beyannameden): {s.donem.tire if s.donem else "?"}')
    var = collections.Counter(b.tur for b in s.belgeler)
    for rol, (ad, zorunlu) in ROLLER.items():
        if rol in s.yollar:
            ek = f'  ({var[rol]} dosya arasından seçildi)' if var[rol] > 1 else ''
            satirlar.append(f'✔ {ad}: {Path(s.yollar[rol]).name}{ek}')
        elif zorunlu:
            satirlar.append(f'✘ EKSİK (zorunlu) {ad}')
        elif var[rol]:
            satirlar.append(f'– {ad}: kullanılmadı')
    if sablon:
        satirlar.append(f'✔ Şablon (önceki ayın raporu, içerikten): {sablon}')
    else:
        satirlar.append(f'✘ Şablon (önceki ayın raporu) bulunamadı{": " + sablon_notu if sablon_notu else ""}')
    if s.bu_ayin_raporu:
        satirlar.append(f'✔ Bu ayın bitmiş raporu da var; taslak onunla karşılaştırılacak: {Path(s.bu_ayin_raporu).name}')
    for t in BILGI_TURLERI:
        if var[t]:
            satirlar.append(f'· {ACIKLAMA[t]}: {var[t]} dosya (şimdilik okunmuyor)')
    if var['TARANMIS']:
        satirlar.append(f'! {var["TARANMIS"]} taranmış PDF / görüntü var (metin yok) — gerekiyorsa elle okunmalı')
    if var['BILINMEYEN']:
        satirlar.append(f'? {var["BILINMEYEN"]} dosya tanınamadı (çoğu fatura / GÇB olabilir)')
    if var['OKUNAMADI']:
        satirlar.append(f'! {var["OKUNAMADI"]} dosya açılamadı')
    return '\n'.join(satirlar)
