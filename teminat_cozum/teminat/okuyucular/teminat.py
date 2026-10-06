"""Teminat mektubu kabul dilekçesi (.docx): tür bazında teminat tutarları, toplam, mektup tarihi/no, banka."""
import re

from ..ortak import dosyalari_bul, num

KODLAR = ['301', '410', '339', '318', '701', '448']


def dilekce_bul(klasor):
    return dosyalari_bul(klasor, 'TEMİNAT MEKTUBU KABUL*.docx') or dosyalari_bul(klasor, 'TEMİNAT MEKTUBU KABUL*.doc')


def dilekce_metni(path):
    from docx import Document
    d = Document(str(path))
    return '\n'.join(p.text for p in d.paragraphs)


def dilekce_metinden(txt):
    tem = {'_text': txt}
    for kod in KODLAR:
        m = re.search(kod + r'[-– ].*?([\d.]+,\d\d) TL', txt)
        if m:
            tem[kod] = num(m.group(1))
    m = re.search(r'([\d.]+,\d\d) TL talep', txt) or re.search(r'toplam ([\d.]+,\d\d)', txt)
    if m:
        tem['toplam'] = num(m.group(1))
    m = re.search(r'(\d\d)[/.](\d\d)[/.](\d{4}) tarih ve (\d+) nolu', txt)
    if m:
        tem['tarih'] = f'{m.group(1)}.{m.group(2)}.{m.group(3)}'
        tem['no'] = m.group(4)
    m = re.search(r'iadesi (.*?)\s*’nden alınan', txt) or re.search(r'KDV iadesi (.*?) ’nden', txt)
    if m:
        tem['banka'] = m.group(1).strip()
    return tem


def tur_kodlari(tem):
    return [k for k in KODLAR if k in tem]
