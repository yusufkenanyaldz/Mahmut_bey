"""1 No.lu KDV beyannamesi (PDF) okuyucusu.

Desenler 2025-2026 beyannamelerinin pdfplumber metnine göre yazılmıştır. Bilinen tuzak: başlıklar
bozuk çıkabiliyor (ör. 'TUTİNARDIİNRIİNM OLREARNLARA GÖRE DAĞILIMI'); bu yüzden desenler başlığın sağlam
kalan kısmına bağlanır.
"""
import re
from pathlib import Path

from ..ortak import PARA as M, ay_no, katla, num


def pdf_metni(path):
    import pdfplumber
    with pdfplumber.open(str(path)) as pdf:
        return '\n'.join((p.extract_text() or '') for p in pdf.pages)


def pdf_ilk_sayfa(path):
    """Belge tanıma için yalnızca ilk sayfanın metni."""
    import pdfplumber
    with pdfplumber.open(str(path)) as pdf:
        return (pdf.pages[0].extract_text() or '') if pdf.pages else ''


def kdv1_bul(klasor):
    """Klasörün kendisindeki 1 No.lu beyanname PDF'i ('KDV 1.pdf', 'KDV1.pdf', 'KDV..pdf')."""
    klasor = Path(klasor)
    if not klasor.is_dir():
        return []
    aday = []
    for p in sorted(klasor.iterdir()):
        if not p.is_file() or p.suffix.lower() != '.pdf' or p.name.startswith('~'):
            continue
        ad = katla(p.name).lower()
        if re.search(r'kdv\s*1(?!\d)', ad) or ad == 'kdv..pdf':
            aday.append(p)
    return aday


def kdv1_metinden(t, dosya=None):
    d = {'_file': str(dosya) if dosya else None}

    def g(k, pat):
        m = re.search(pat, t)
        if m:
            d[k] = num(m.group(1))

    g('matrah_toplam', r'Matrah Toplamı ' + M)
    g('hesaplanan', r'Hesaplanan Katma Değer Vergisi ' + M)
    g('ilave', r'Daha Önce İndirim Konusu Yapılan KDV’nin İlavesi ' + M)
    g('toplam_kdv', r'Toplam Katma Değer Vergisi ' + M)
    g('yurtici_alim', r'Yurtiçi Alımlara İlişkin KDV ' + M)
    g('sorumlu', r'Sorumlu Sıfatıyla Beyan Edilerek Ödenen KDV ' + M)
    g('ithal', r'İthalde Ödenen KDV ' + M)
    g('satis_iade', r'Satıştan İade Edilen.*? ' + M)
    g('devreden_onceki', r'Önceki Dönemden Devreden ' + M)
    g('indirim_toplam', r'İndirimler Toplamı ' + M)
    g('iade_edilebilir', r'İade Edilebilir KDV ' + M)
    g('iade_gereken', r'İade Edilmesi Gereken Katma Değer Vergisi ' + M)
    g('sonraki_devreden', r'Sonraki Döneme Devreden Katma Değer Vergisi ' + M)
    g('aylik_bedel', r'Bedel \(aylık\) ' + M)
    g('ihrac_kayitli_bedel', r'İhraç Kaydıyla Teslim Bedeli Toplamı ' + M)
    g('ihrac_kayitli_iade', r'İhracatın Gerçekleştiği Dönemde İade Edilecek KDV ' + M)
    m = re.search(r'301 - Mal İhracatı ' + M + ' ' + M + ' ' + M, t)
    if m:
        d['301_tutar'], d['301_odenmeksizin'], d['301_yuklenilen'] = map(num, m.groups())
    m = re.search(r'339 - .*?' + M + r'\s+' + M + r'(?:\s+' + M + ')?', t, re.S)
    if m:
        d['339_tutar'] = num(m.group(1))
        d['339_iade'] = num(m.group(m.lastindex))
    m = re.search(r'410 - Yapım İşleri İle Bu İşlerle Birlikte İfa Edilen ' + M + ' ' + M, t)
    if m:
        d['410_tutar'], d['410_iade'] = map(num, m.groups())
    m = re.search(r'Düzeltme Nedeni : (.*)', t)
    d['duzeltme'] = m.group(1).strip() if m else ''
    m = re.search(r'Yıl (\d{4}).*?Ay (\w+)', t, re.S)
    d['donem'] = m.groups() if m else None
    # oran dağılımı
    sec = t.split('GÖRE DAĞILIMI')[1][:400] if 'GÖRE DAĞILIMI' in t else ''
    d['oranlar'] = [(int(a), num(b), num(c)) for a, b, c in re.findall(r'^(\d{1,2}) ' + M + ' ' + M, sec, re.M)]

    def g2(pat):
        m = re.search(pat, t, re.S)
        return tuple(num(x) if re.fullmatch(r'-?[\d.]+,\d\d', x) else x for x in m.groups()) if m else None

    d['r701'] = g2(r'11/1-c Maddesi Kapsamında Teslimi ' + M + r' (\d+) ' + M)
    d['yurtici'] = [(num(a), int(b), num(c)) for a, b, c in re.findall(r'Yurtiçi Teslim ve Hizmetler\s*\n' + M + r' (\d+) ' + M, t)]
    d['k410'] = g2(r'Yapım İşleri ile Bu İşlerle Birlikte İfa Edilen ' + M + r' (\d+) (\d+/\d+) ' + M)
    d['k448'] = g2(r'Demir-Çelik Ürünlerinin Teslimi \[KDVGUT- ' + M + r' (\d+) (\d+/\d+) ' + M)
    d['d_iade'] = g2(r'Alınan Malların İadesi,\s*\n' + M + ' ' + M)
    d['d_kur'] = g2(r'Kur Farkı / Yuvarlama Farkı\s*\n' + M + ' ' + M)
    d['d_diger'] = g2(r'\nDiğerleri ' + M + ' ' + M)
    d['d_amort'] = g2(r'Demirbaş, ' + M + ' ' + M)
    d['i318'] = g2(r'318 - 3996[^\n]*? ' + M + ' ' + M + ' ' + M)
    d['i350'] = g2(r'350 - [^\n]*? ' + M + ' ' + M + ' ' + M)
    d['i448'] = g2(r'448 - Demir Çelik[^\n]*? ' + M + ' ' + M)
    d['kumulatif'] = (g2(r'Bedel \(kümülatif\) ' + M) or (None,))[0]
    d['tecil_edilecek'] = (g2(r'Tecil Edilecek Katma Değer Vergisi ' + M) or (0,))[0]
    d['odenmesi_gereken'] = (g2(r'\nÖdenmesi Gereken Katma Değer Vergisi ' + M) or (0,))[0]
    d['_text'] = t
    return d


def kdv1_donem(k):
    """Beyannamenin dönemi (yıl, ay) ya da None."""
    if not k or not k.get('donem'):
        return None
    yil, ay = k['donem']
    a = ay_no(ay)
    return (int(yil), a) if a else None


# İade türleri: kod → (iş hacmi tablosundaki varsayılan etiket, tutar, iade)
IADE_ETIKETLERI = {
    '301': '301 – Mal İhracatı',
    '318': '318 – 3996 Sayılı Kanuna Göre Yap İşlet Devret Modeli Çer.',
    '339': '339 – İmalat Sanayii ile Turizme Yönelik Yatırım Teşvik Belgesi Kapsam',
    '350': '350 – Diğerleri',
    '410': '410 – Yapım İşleri ile Bu İşlerle Birlikte İfa',
    '448': '448 – Demir Çelik Ürünlerinin Teslimi',
}


def iade_kalemleri(k):
    """Beyandaki iade türleri: {kod: (teslim tutarı, iadeye konu KDV)} (beyanda olmayan tür yazılmaz)."""
    out = {}
    if k.get('301_tutar') is not None:
        out['301'] = (k['301_tutar'], k.get('301_yuklenilen', 0.0))
    if k.get('i318'):
        out['318'] = (k['i318'][0], k['i318'][2])
    if k.get('339_tutar'):
        out['339'] = (k['339_tutar'], k.get('339_iade', 0.0))
    if k.get('i350'):
        out['350'] = (k['i350'][0], k['i350'][2])
    if k.get('410_tutar') is not None:
        out['410'] = (k['410_tutar'], k.get('410_iade', 0.0))
    if k.get('i448'):
        out['448'] = (k['i448'][0], k['i448'][1])
    return out
