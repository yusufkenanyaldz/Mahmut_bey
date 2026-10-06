"""Excel listeleri: indirilecek KDV listesi, karşıt inceleme takip listesi, yüklenilen tutanak çalışması."""
import datetime
import re
from pathlib import Path

from ..ortak import dosyalari_bul, katla, xrows

ITHALAT_VKN = '1111111111'


def indirilecek_bul(klasor):
    """Önce klasörün kendisinde 'İndirilecek KDV listesi*.xls', yoksa alt klasörlerle birlikte
    '*indirilecekkdvListesi*' aranır."""
    c = [p for p in dosyalari_bul(klasor, '*.xls', alt_klasorler=False)
         if 'indirilecek kdv listesi' in katla(p.name).lower()]
    return c or dosyalari_bul(klasor, '*indirilecekkdvListesi*')


def _vkn(v):
    return str(v).split('.')[0].strip()


def indirilecek_oku(path):
    """Satırlar: sütun 5 satıcı, 6 VKN, 9 bedel, 10 KDV, 11 tevkifatsız indirilen, 12 2 no.lu ödenen,
    13 toplam indirilen, 15 indirim dönemi."""
    out = []
    for _, r in xrows(path):
        if len(r) > 13 and isinstance(r[1], (int, float)) and r[1] >= 1 and isinstance(r[10], (int, float)) and r[5]:
            out.append({
                'satici': str(r[5]).strip(),
                'vkn': _vkn(r[6]),
                'bedel': r[9] if isinstance(r[9], (int, float)) else None,
                'kdv': float(r[10]),
                'donem': r[15] if len(r) > 15 else None,
                'ham': r,
            })
    return out


def donem_degeri(v):
    """İndirim dönemi hücresi → (yıl, ay) ya da None. 202602, '2026/02', '02/2026', '2026-02' biçimleri."""
    if v is None or v == '':
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    m = re.fullmatch(r'(\d{4})(\d{2})', s) or re.fullmatch(r'(\d{4})\s*[-/.]\s*(\d{1,2})', s)
    if m:
        y, a = int(m.group(1)), int(m.group(2))
    else:
        m = re.fullmatch(r'(\d{1,2})\s*[-/.]\s*(\d{4})', s)
        if not m:
            return None
        a, y = int(m.group(1)), int(m.group(2))
    return (y, a) if 1 <= a <= 12 else None


def takip_bul(klasor):
    return dosyalari_bul(klasor, '01 FİRMA*')


def takip_oku(path):
    """Karşıt inceleme takip listesi: sütun 2 firma, 3 KDV, 4 açıklama/e-posta, 5 SMMM, 6 YMM."""
    out = []
    for _, r in xrows(path):
        if len(r) > 3 and isinstance(r[1], (int, float)) and str(r[2]).strip() and isinstance(r[3], (int, float)):
            def s(i):
                return str(r[i]).strip() if len(r) > i and r[i] is not None else ''
            out.append({'firma': str(r[2]).strip(), 'kdv': float(r[3]), 'aciklama': s(4), 'smmm': s(5), 'ymm': s(6)})
    return out


def yuklenilen_bul(klasor):
    return dosyalari_bul(klasor, 'yüklenilen tutanak*.xls*')


def yuklenilen_oku(path):
    """Yüklenilen KDV dökümü: sütun 2 fatura tarihi (Excel seri), 4 satıcı, 5 VKN, 10 bünyeye giren KDV,
    12 kod, 15 dönem (yyyymm; yoksa fatura tarihinden)."""
    out = []
    for _, r in xrows(path):
        if len(r) > 10 and isinstance(r[1], float) and isinstance(r[2], float) and isinstance(r[10], float):
            d = datetime.date(1899, 12, 30) + datetime.timedelta(days=int(r[2]))
            per = None
            if len(r) > 15 and str(r[15]).strip():
                try:
                    v = int(float(r[15]))
                    per = (v // 100, v % 100)
                except (TypeError, ValueError):
                    pass
            if not per:
                per = (d.year, d.month)
            out.append({'firma': str(r[4]).strip(), 'vkn': str(r[5]).strip(), 'kdv': r[10], 'per': per,
                        'kod': str(r[12]).strip() if len(r) > 12 else ''})
    return out


def dosya_adi(p):
    return Path(p).name if p else '-'
