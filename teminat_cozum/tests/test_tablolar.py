import copy

from docx import Document

import sentetik as S
from teminat.ortak import Donem
from teminat.rapor.tablolar import tablolari_bul, tanimlari_al
from teminat.ayar import FirmaAyari


def test_tum_tablolar_bulunur(tmp_path):
    p = S.sablon_docx(tmp_path / 's.docx', S.Senaryo(Donem(2026, 2)))
    d = Document(str(p))
    T, sorunlar = tablolari_bul(d)
    assert sorunlar == []
    assert set(T) == {'kapak', 'isci', 'matrah', 'indirim', 'is_hacmi', 'karsit', 'safha', 'yuklenilen', 'tevkifat', 'dokum'}
    assert 'İNDİRİLECEK KDV TUTANAKLARI' in T['safha'].rows[0].cells[0].text


def test_araya_tablo_girince_index_kaymasindan_etkilenmez(tmp_path):
    """Aralık-2025 raporunda 339 tablosu araya girince sıra numaraları kaymıştı."""
    p = S.sablon_docx(tmp_path / 's.docx', S.Senaryo(Donem(2026, 2)))
    d = Document(str(p))
    T, _ = tablolari_bul(d)
    once = T['karsit']._tbl
    ek = copy.deepcopy(T['tevkifat']._tbl)
    for t in ek.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
        t.text = 'İMALAT SANAYİİ 339' if t.text else t.text
    once.addprevious(ek)
    T2, sorunlar = tablolari_bul(d)
    assert sorunlar == []
    assert T2['safha']._tbl is T['safha']._tbl and T2['karsit']._tbl is once


def test_ayar_ile_tanim_degisir():
    t = tanimlari_al(FirmaAyari(tablolar={'safha': {'icerik': ['TUTANAK LİSTESİ']}, 'yeni': {'icerik': ['X']}}))
    assert t['safha'].icerik == ['TUTANAK LİSTESİ'] and t['yeni'].icerik == ['X']
