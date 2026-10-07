"""Klasör bulma: ofisteki farklı klasör düzenlerinde bu ayın klasörü ve önceki ayın raporu bulunmalı."""
import unicodedata

import pytest

import sentetik as S
from teminat.cli import main
from teminat.klasorler import KlasorHatasi, ay_coz, donem_haritasi, girdi_ve_sablon_bul
from teminat.ortak import Donem

SUBAT, MART = Donem(2026, 2), Donem(2026, 3)


@pytest.mark.parametrize('ad,beklenen', [
    ('02 ŞUBAT', (2, None)), ('02 ŞUBAT (GİRDİ)', (2, None)), ('ESKA 12- ARALIK', (12, None)), ('Şubat 2026', (2, 2026)),
    ('SUBAT', (2, None)), ('2026-02', (2, 2026)), ('02.2026', (2, 2026)), ('3- GİRDİ', (3, None)), ('03', (3, None)),
    ('2026 02 ŞUBAT', (2, 2026)), ('ARALIK-2025', (12, 2025)), ('11 KASIM 2025', (11, 2025)),
    (unicodedata.normalize('NFD', '02 ŞUBAT'), (2, None)),
    ('TUTANAK ÇALIŞMASI', None), ('301', None), ('ithalat', None), ('SİSTEM', None), ('RAPOR', None), ('2026', None),
    ('OCAK-ŞUBAT', None), ('MARTI YAPI A.Ş.', None), ('13 NOTLAR', None), ('ESKA', None),
])
def test_ay_coz(ad, beklenen):
    assert ay_coz(ad) == beklenen


def _firma(tmp_path, subat_ad='2026/02 ŞUBAT', mart_ad='2026/03 MART', **kw):
    kok = tmp_path / 'ESKA'
    kok.mkdir()
    sub = S.Senaryo(SUBAT)
    mar = S.Senaryo(MART, devreden_onceki=sub.sonraki_devreden)
    (kok / subat_ad).mkdir(parents=True)
    S.girdi_klasoru_yol(kok / subat_ad, sub)
    r = kok / subat_ad / kw.get('rapor_alt', 'RAPOR') if kw.get('rapor_alt', 'RAPOR') else kok / subat_ad
    r.mkdir(parents=True, exist_ok=True)
    S.sablon_docx(r / kw.get('rapor_adi', 'ŞUBAT-2026 RAPOR.docx'), sub)
    (kok / mart_ad).mkdir(parents=True)
    S.girdi_klasoru_yol(kok / mart_ad, mar)
    return kok


@pytest.mark.parametrize('subat_ad,mart_ad', [
    ('2026/02 ŞUBAT', '2026/03 MART'),                     # önerilen düzen
    ('2026/02 ŞUBAT (GİRDİ)', '2026/03 MART (GİRDİ)'),
    ('ESKA 02- ŞUBAT', 'ESKA 03- MART'),                   # yıl klasörü yok, firma adıyla başlayan ay klasörleri
    ('ŞUBAT 2026', 'MART 2026'),
    ('2026-02', '2026-03'),
    ('2026 YILI/02-ŞUBAT', '2026 YILI/03-MART'),
    ('KDV İADE/2026/2', 'KDV İADE/2026/3'),               # bir ara klasör + yalnızca ay numarası
])
def test_farkli_duzenler_firma_klasoru_ve_donem(tmp_path, subat_ad, mart_ad):
    kok = _firma(tmp_path, subat_ad, mart_ad)
    b = girdi_ve_sablon_bul(kok, MART)
    assert b.girdi == kok / mart_ad
    assert b.sablon == kok / subat_ad / 'RAPOR' / 'ŞUBAT-2026 RAPOR.docx'
    # doğrudan ayın klasörü seçilirse dönem KDV 1'den okunur
    b2 = girdi_ve_sablon_bul(kok / mart_ad, None)
    assert (b2.girdi, b2.donem, b2.sablon) == (kok / mart_ad, MART, b.sablon)


def test_yil_klasoru_secilirse(tmp_path):
    kok = _firma(tmp_path)
    b = girdi_ve_sablon_bul(kok / '2026', MART)
    assert b.girdi == kok / '2026/03 MART' and b.sablon is not None


def test_rapor_ay_klasorunde_alt_klasorsuz(tmp_path):
    kok = _firma(tmp_path, rapor_alt=None, rapor_adi='ŞUBAT-2026 YMM RAPOR.doc')
    assert girdi_ve_sablon_bul(kok, MART).sablon.name == 'ŞUBAT-2026 YMM RAPOR.doc'


def test_rapor_klasoru_farkli_adla(tmp_path):
    kok = _firma(tmp_path, rapor_alt='RAPORLAR', rapor_adi='son hali.docx')
    assert girdi_ve_sablon_bul(kok, MART).sablon.name == 'son hali.docx'


def test_aralik_raporundan_ocak(tmp_path):
    kok = tmp_path / 'ESKA'
    ara, oca = S.Senaryo(Donem(2025, 12)), S.Senaryo(Donem(2026, 1))
    (kok / '2025/12 ARALIK/RAPOR').mkdir(parents=True)
    S.sablon_docx(kok / '2025/12 ARALIK/RAPOR/ARALIK-2025 RAPOR.doc', ara)
    (kok / '2026/01 OCAK').mkdir(parents=True)
    S.girdi_klasoru_yol(kok / '2026/01 OCAK', oca)
    for secilen, d in ((kok, Donem(2026, 1)), (kok / '2026/01 OCAK', None), (kok / '2026', Donem(2026, 1))):
        b = girdi_ve_sablon_bul(secilen, d)
        assert b.sablon.name == 'ARALIK-2025 RAPOR.doc', secilen


def test_ic_ice_girdi_klasoru_secilirse(tmp_path):
    kok = tmp_path / 'ESKA'
    sub, mar = S.Senaryo(SUBAT), S.Senaryo(MART)
    (kok / '2026/02 ŞUBAT/RAPOR').mkdir(parents=True)
    S.sablon_docx(kok / '2026/02 ŞUBAT/RAPOR/r.docx', sub)
    (kok / '2026/03 MART/SİSTEM').mkdir(parents=True)
    S.girdi_klasoru_yol(kok / '2026/03 MART/SİSTEM', mar)
    b = girdi_ve_sablon_bul(kok / '2026/03 MART/SİSTEM', None)
    assert b.sablon is not None and b.donem == MART
    assert girdi_ve_sablon_bul(kok, MART).girdi == kok / '2026/03 MART/SİSTEM'
    assert girdi_ve_sablon_bul(kok / '2026/03 MART', None).girdi == kok / '2026/03 MART/SİSTEM'


def test_bulunamazsa_gorulen_klasorleri_listeler(tmp_path):
    kok = _firma(tmp_path)
    with pytest.raises(KlasorHatasi) as e:
        girdi_ve_sablon_bul(kok, Donem(2026, 5))
    assert 'MAYIS-2026 için ay klasörü bulunamadı' in str(e.value)
    assert 'ŞUBAT-2026 → 2026/02 ŞUBAT' in str(e.value) and 'MART-2026 → 2026/03 MART' in str(e.value)


def test_donem_yazilmazsa_ve_firma_klasoru_secilirse(tmp_path):
    with pytest.raises(KlasorHatasi, match='Dönemi yazın'):
        girdi_ve_sablon_bul(_firma(tmp_path), None)


def test_donem_kdv1_ile_celisirse(tmp_path):
    kok = _firma(tmp_path)
    with pytest.raises(KlasorHatasi, match='ŞUBAT-2026 dönemine ait'):
        girdi_ve_sablon_bul(kok / '2026/02 ŞUBAT', MART)


def test_bircok_firmanin_ustu_secilirse(tmp_path):
    _firma(tmp_path)
    ikinci = tmp_path / 'MUTAS'
    (ikinci / '2026/03 MART').mkdir(parents=True)
    S.girdi_klasoru_yol(ikinci / '2026/03 MART', S.Senaryo(MART))
    with pytest.raises(KlasorHatasi, match='birden fazla klasörde KDV 1.pdf'):
        girdi_ve_sablon_bul(tmp_path, MART)


def test_onceki_ay_yoksa_sablon_bos_ve_not(tmp_path):
    kok = tmp_path / 'ESKA'
    (kok / '2026/03 MART').mkdir(parents=True)
    S.girdi_klasoru_yol(kok / '2026/03 MART', S.Senaryo(MART))
    b = girdi_ve_sablon_bul(kok, MART)
    assert b.sablon is None and any('ŞUBAT-2026 klasörü bulunamadı' in n for n in b.notlar)


def test_cli_ay_klasoru_ile_taslak_ve_sablonsuz_hata(tmp_path, capsys):
    kok = _firma(tmp_path, 'ESKA 02- ŞUBAT', 'ESKA 03- MART')
    assert main(['taslak', '--kok', str(kok / 'ESKA 03- MART')]) == 0, capsys.readouterr()
    assert (kok / 'ESKA 03- MART' / 'CLAUDE TASLAK' / 'MART-2026 RAPOR TASLAK.docx').exists()
    # rapor yoksa anlaşılır hata, --sablon ile çalışır
    rapor = kok / 'ESKA 02- ŞUBAT' / 'RAPOR' / 'ŞUBAT-2026 RAPOR.docx'
    tasindi = tmp_path / 'r.docx'
    rapor.rename(tasindi)
    assert main(['taslak', '--kok', str(kok / 'ESKA 03- MART')]) == 2
    assert 'bitmiş raporu bulunamadı' in capsys.readouterr().err
    assert main(['taslak', '--kok', str(kok / 'ESKA 03- MART'), '--sablon', str(tasindi)]) == 0


def test_cli_klasor_komutu(tmp_path, capsys):
    kok = _firma(tmp_path)
    assert main(['klasor', '--kok', str(kok), '--donem', '2026-03']) == 0
    out = capsys.readouterr().out
    assert 'MART-2026' in out and 'ŞUBAT-2026 RAPOR.docx' in out


def test_donem_haritasi_yilsiz_klasorler_kdv1den(tmp_path):
    kok = _firma(tmp_path, 'ESKA 02- ŞUBAT', 'ESKA 03- MART')
    h, atlanan = donem_haritasi(kok)
    assert set(h) == {SUBAT, MART} and not atlanan
    assert h[SUBAT]['rapor'] and h[MART]['girdi'] == [kok / 'ESKA 03- MART']
