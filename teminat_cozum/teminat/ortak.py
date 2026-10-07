"""Ortak yardımcılar: Türkçe sayı biçimi, dönem, dosya arama, Excel satır okuma."""
import fnmatch
import os
import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path

AYLAR = ['OCAK', 'ŞUBAT', 'MART', 'NİSAN', 'MAYIS', 'HAZİRAN', 'TEMMUZ', 'AĞUSTOS', 'EYLÜL', 'EKİM', 'KASIM', 'ARALIK']
DOLD = '[DOLDURULACAK]'
PARA = r'(-?[\d.]+,\d\d)'                       # PDF / metin içindeki tutar
PARA_RE = re.compile(r'^-?[\d.]+,\d\d$')        # tek başına tutar hücresi
PARA_BUL = re.compile(r'-?\d{1,3}(?:\.\d{3})*,\d\d(?!\d)')
# Çıktı ve dönüşüm klasörleri dosya aramalarında atlanır (kendi çıktımızı girdi sanmayalım).
HARIC_KLASORLER = ('claude taslak', 'claude outputs')


def nf(s):
    return unicodedata.normalize('NFC', s)


def katla(s):
    """Türkçe karakterleri ve NFD/NFC farkını yok sayan karşılaştırma biçimi (büyük/küçük harf korunur).

    'İNDİRİLECEK' → 'INDIRILECEK', 'Yüklenilen' → 'Yuklenilen', 'ı' → 'i'.
    """
    s = unicodedata.normalize('NFKD', str(s))
    s = ''.join(c for c in s if not unicodedata.combining(c))
    return s.replace('ı', 'i')


def num(s):
    return float(s.replace('.', '').replace(',', '.'))


def tr(x):
    s = f"{x:,.2f}"
    return s.replace(',', 'X').replace('.', ',').replace('X', '.')


def pct(x):
    return '% ' + tr(x * 100)


def pct2(x):
    return '%' + tr(x * 100)


def ay_no(ad):
    """'Şubat', 'ŞUBAT', 'subat' → 2; tanınmazsa None."""
    k = katla(ad).upper().strip()
    for i, a in enumerate(AYLAR):
        if katla(a).upper() == k:
            return i + 1
    return None


@dataclass(frozen=True, order=True)
class Donem:
    yil: int
    ay: int

    @property
    def ad(self):
        return AYLAR[self.ay - 1]

    @property
    def egik(self):
        """ŞUBAT/2026"""
        return f'{self.ad}/{self.yil}'

    @property
    def tire(self):
        """ŞUBAT-2026"""
        return f'{self.ad}-{self.yil}'

    def onceki(self):
        return Donem(self.yil - 1, 12) if self.ay == 1 else Donem(self.yil, self.ay - 1)

    def sonraki(self):
        return Donem(self.yil + 1, 1) if self.ay == 12 else Donem(self.yil, self.ay + 1)

    def __str__(self):
        return self.tire

    @classmethod
    def coz(cls, s):
        """'2026-02', '2026/2', 'ŞUBAT-2026', 'ŞUBAT/2026', '202602' biçimlerini okur."""
        s = nf(str(s)).strip()
        m = (re.fullmatch(r'(\d{4})\s*[-/.]\s*(\d{1,2})', s) or re.fullmatch(r'(\d{4})(\d{2})', s))
        if m:
            if not 1 <= int(m.group(2)) <= 12:
                raise ValueError(f'Dönemdeki ay 1-12 arasında olmalı: {s!r}')
            return cls(int(m.group(1)), int(m.group(2)))
        m = re.fullmatch(r'(\d{1,2})\s*[-/.]\s*(\d{4})', s)
        if m:
            if not 1 <= int(m.group(1)) <= 12:
                raise ValueError(f'Dönemdeki ay 1-12 arasında olmalı: {s!r}')
            return cls(int(m.group(2)), int(m.group(1)))
        m = re.fullmatch(r'(\w+)\s*[-/ ]\s*(\d{4})', s)
        if m and ay_no(m.group(1)):
            return cls(int(m.group(2)), ay_no(m.group(1)))
        raise ValueError(f'Dönem anlaşılamadı: {s!r} (örnek: 2026-02 veya ŞUBAT-2026)')


def _eslesir(ad, desen):
    return fnmatch.fnmatch(katla(ad).lower(), katla(desen).lower())


def dosyalari_bul(klasor, desen, alt_klasorler=True):
    """Desene uyan dosyaları (Türkçe karakter / NFD farkı gözetmeden) sıralı döndürür.

    Sıra: önce klasör derinliği, sonra ad. Word/Excel kilit dosyaları (~$...) ve
    'CLAUDE TASLAK' gibi çıktı klasörleri atlanır.
    """
    klasor = Path(klasor)
    if not klasor.is_dir():
        return []
    bulunan = []
    if alt_klasorler:
        for kok, dirs, files in os.walk(klasor):
            dirs[:] = [d for d in dirs if katla(d).lower() not in HARIC_KLASORLER]
            for f in files:
                if not f.startswith('~') and _eslesir(f, desen):
                    bulunan.append(Path(kok) / f)
    else:
        bulunan = [p for p in klasor.iterdir() if p.is_file() and not p.name.startswith('~') and _eslesir(p.name, desen)]
    return sorted(bulunan, key=lambda p: (len(p.relative_to(klasor).parts), katla(str(p)).lower()))


def dosya_bul(klasor, desen, alt_klasorler=True):
    c = dosyalari_bul(klasor, desen, alt_klasorler)
    return c[0] if c else None


def excel_turu(path):
    """Dosyanın ilk baytlarına göre 'xlsx' / 'xls' / None (uzantıya güvenilmez: GİB'den inen '.xls' bazen .xlsx'tir)."""
    try:
        with open(path, 'rb') as f:
            bas = f.read(8)
    except OSError:
        return None
    if bas.startswith(b'PK'):                   # .docx da zip'tir: içinde çalışma kitabı olmalı
        import zipfile
        try:
            with zipfile.ZipFile(path) as z:
                return 'xlsx' if any(a.lower().startswith('xl/workbook') for a in z.namelist()) else None
        except (OSError, zipfile.BadZipFile):
            return None
    if bas.startswith(b'\xd0\xcf\x11\xe0'):
        return 'xls'
    return None


def excel_satirlari(path, en_cok=None):
    """(sayfa adı, hücre değerleri listesi) üretir; biçim uzantıdan değil içerikten anlaşılır. en_cok: sayfa başına satır."""
    tur = excel_turu(path)
    if tur == 'xlsx':
        import io

        import openpyxl
        with open(path, 'rb') as f:
            wb = openpyxl.load_workbook(io.BytesIO(f.read()), read_only=True, data_only=True)
        try:
            for ws in wb.worksheets:
                for i, r in enumerate(ws.iter_rows(values_only=True)):
                    if en_cok is not None and i >= en_cok:
                        break
                    yield ws.title, list(r)
        finally:
            wb.close()
    elif tur == 'xls':
        import xlrd
        wb = xlrd.open_workbook(str(path))
        for sh in wb.sheets():
            for i in range(sh.nrows if en_cok is None else min(sh.nrows, en_cok)):
                yield sh.name, [c.value for c in sh.row(i)]
    else:
        raise ValueError(f'{Path(path).name}: Excel dosyası değil (içeriği .xls / .xlsx biçiminde değil)')


def xrows(path):
    """(sayfa adı, hücre değerleri listesi) üretir."""
    yield from excel_satirlari(path)


def para_bul(metin):
    """Metindeki Türkçe biçimli tutarları ('1.234,56') sırayla döndürür."""
    return PARA_BUL.findall(metin or '')
