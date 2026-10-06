import collections

from docx import Document

import sentetik as S
from teminat.kontroller import kurallar
from teminat.kontroller.belge import safha_tablosu_denetle
from teminat.kontroller.liste import KontrolListesi
from teminat.ortak import Donem
from teminat.rapor.docx_araclari import set_cell, uniq_cells
from teminat.rapor.tablolar import tablolari_bul


def _safha(tmp_path):
    p = S.sablon_docx(tmp_path / 's.docx', S.Senaryo(Donem(2026, 2)))
    d = Document(str(p))
    return d, tablolari_bul(d)[0]['safha']


def _durum(kl, konu):
    return [k.durum for k in kl if k.konu == konu]


def test_safha_tablosu_tutarli(tmp_path):
    _, t = _safha(tmp_path)          # sentetik şablon: EKLİ 1+2+3=6, muhafaza 4+5=9, genel 15
    kl = KontrolListesi()
    safha_tablosu_denetle(kl, t)
    assert _durum(kl, 'Safha tablosu') == ['TAMAM'], [(k.konu, k.aciklama) for k in kl]


def test_safha_toplami_tutmazsa_ve_satir_kaymasi(tmp_path):
    _, t = _safha(tmp_path)
    rows = list(t.rows)
    set_cell(uniq_cells(rows[2])[-2], '1,00')   # ikinci EKLİ satırı birinciyle aynı tutar → kayma + toplam hatası
    kl = KontrolListesi()
    safha_tablosu_denetle(kl, t, base=100.0)
    assert 'HATA' in _durum(kl, 'Safha toplamı')
    assert 'UYARI' in _durum(kl, 'Satır kayması')
    assert 'HATA' in _durum(kl, 'Karşıt inceleme oranı')   # 15 < 100 × %80


def test_kurallar():
    kl = KontrolListesi()
    kurallar.liste_toplami(kl, 100.0, 100.5)
    kurallar.liste_toplami(kl, 90.0, 100.0)
    assert _durum(kl, 'İndirilecek liste ↔ beyan') == ['TAMAM', 'UYARI']
    kurallar.liste_donemi(kl, collections.Counter({202603.0: 5}), Donem(2026, 3))
    kurallar.liste_donemi(kl, collections.Counter({202603.0: 5, 202602.0: 2}), Donem(2026, 3))
    assert _durum(kl, 'Liste dönemi') == ['TAMAM', 'HATA']
    kurallar.iade_turleri(kl, {'301': (1, 10.0), '410': (1, 5.0), '318': (1, 0.0)}, 15.0)
    assert _durum(kl, 'İade türleri toplamı') == ['TAMAM']
    kurallar.teminat(kl, {'301': 9.0, '410': 6.0, 'toplam': 15.0, 'tarih': 'x', 'no': 'y'}, {'301': (1, 10.0), '410': (1, 5.0)}, 15.0)
    assert 'HATA' in _durum(kl, 'Teminat ≤ iade')          # 410 teminatı 410 iadesinden büyük
    kurallar.tevkifat_410(kl, 1000.0, 20, '4/10', 80.0)
    kurallar.tevkifat_410(kl, 1000.0, 20, '4/10', 81.5)
    assert _durum(kl, '410 tevkifat hesabı') == ['TAMAM', 'UYARI']
    kurallar.devreden(kl, 100.0, 100.0)
    kurallar.devreden(kl, 100.0, 90.0)
    kurallar.devreden(kl, None, 90.0)
    assert _durum(kl, 'Devreden KDV') == ['TAMAM', 'HATA', 'BİLGİ']


def test_beyan_tutarliligi_eksik_satiri_yakalar():
    s = S.Senaryo(Donem(2026, 3))
    from teminat.okuyucular.kdv1 import kdv1_metinden
    k = kdv1_metinden(S.kdv1_metni(s))
    kl = KontrolListesi()
    kurallar.beyan_tutarliligi(kl, k)
    assert set(_durum(kl, 'KDV 1 okuma')) == {'TAMAM'}
    k['yurtici'] = k['yurtici'][:1]          # bir satır okunamamış gibi
    kl = KontrolListesi()
    kurallar.beyan_tutarliligi(kl, k)
    assert 'HATA' in _durum(kl, 'KDV 1 okuma')
