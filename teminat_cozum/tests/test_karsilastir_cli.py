import shutil

import pytest
from docx import Document

import sentetik as S
from teminat.cli import main
from teminat.geriye_donuk import geriye_donuk
from teminat.karsilastir import karsilastir
from teminat.klasorler import KlasorHatasi, ay_klasorleri, rapor_dosyasi
from teminat.ortak import Donem
from teminat.rapor.docx_araclari import uniq_cells
from teminat.rapor.tablolar import tablolari_bul


def test_karsilastir_tutar_farki(tmp_path):
    a = S.sablon_docx(tmp_path / 'a.docx', S.Senaryo(Donem(2026, 2)))
    b = tmp_path / 'b.docx'
    shutil.copy(a, b)
    assert karsilastir(a, b).tutar_farki == 0
    d = Document(str(b))
    c = uniq_cells(tablolari_bul(d)[0]['safha'].rows[1])[-2]
    c.paragraphs[0].runs[0].text = '99,99'
    d.save(str(b))
    k = karsilastir(a, b)
    assert k.tutar_farki == 1 and k.farkli_satir == 1
    assert '99,99' in k.metin()


def test_ay_klasorleri_ve_rapor_secimi(tmp_path):
    kok = tmp_path / 'FIRMA'
    for ad in ('2026/01 OCAK', '2026/02 ŞUBAT (GİRDİ)', '2025/12 ARALIK', '2025/notlar'):
        (kok / ad / 'RAPOR').mkdir(parents=True)
    a = ay_klasorleri(kok)
    assert sorted(a) == [Donem(2025, 12), Donem(2026, 1), Donem(2026, 2)]
    r = kok / '2026/01 OCAK/RAPOR'
    (r / 'OCAK-2026 RAPOR.doc').write_text('x')
    (r / 'OCAK-2026 RAPOR ORJ.doc').write_text('x')
    (r / '~$OCAK-2026 RAPOR.doc').write_text('x')
    assert rapor_dosyasi(kok / '2026/01 OCAK').name == 'OCAK-2026 RAPOR.doc'
    (r / 'OCAK-2026 RAPOR.docx').write_text('x')
    assert rapor_dosyasi(kok / '2026/01 OCAK').name == 'OCAK-2026 RAPOR.docx'
    (r / 'başka RAPOR.docx').write_text('x')
    with pytest.raises(KlasorHatasi):
        rapor_dosyasi(kok / '2026/01 OCAK')


def _firma_klasoru(kok):
    """ŞUBAT (rapor var) + MART (girdi + rapor) içeren firma klasörü."""
    eski = S.Senaryo(Donem(2026, 2))
    yeni = S.Senaryo(Donem(2026, 3), devreden_onceki=eski.sonraki_devreden)
    S.girdi_klasoru(kok, eski)
    (kok / '2026/02 ŞUBAT/RAPOR').mkdir()
    S.sablon_docx(kok / '2026/02 ŞUBAT/RAPOR/ŞUBAT-2026 RAPOR.docx', eski)
    S.girdi_klasoru(kok, yeni)
    return kok


def test_cli_taslak_kok_ve_donem(tmp_path, capsys):
    kok = _firma_klasoru(tmp_path / 'ORNEK')
    ayar = tmp_path / 'ayar.yaml'
    ayar.write_text('kisa_ad: ÖRNEK\nymm_no: "00000000"\nrapor_referanslari:\n  OCAK/2026: {tarih: "15.04.2026", sayi: "2026-50"}\n',
                    encoding='utf-8')
    kod = main(['taslak', '--ayar', str(ayar), '--kok', str(kok), '--donem', '2026-03'])
    cikti = capsys.readouterr().out
    assert kod == 0, cikti
    t = kok / '2026/03 MART/CLAUDE TASLAK'
    assert (t / 'MART-2026 RAPOR TASLAK.docx').exists() and (t / 'MART-2026 KONTROL LİSTESİ.xlsx').exists()
    assert 'HATA: 0' in cikti


def test_cli_hatali_girdi_kodu(tmp_path, capsys):
    assert main(['taslak', '--sablon', str(tmp_path / 'yok.docx'), '--girdi', str(tmp_path)]) == 2
    assert 'HATA' in capsys.readouterr().err


def test_cli_ayar_bilinmeyen_alan(tmp_path, capsys):
    ayar = tmp_path / 'ayar.yaml'
    ayar.write_text('yanlis_alan: 1\n', encoding='utf-8')
    assert main(['taslak', '--ayar', str(ayar), '--sablon', 'x', '--girdi', 'y']) == 2
    assert 'bilinmeyen ayar' in capsys.readouterr().err


def test_cli_denetle(tmp_path, capsys):
    p = S.sablon_docx(tmp_path / 'r.docx', S.Senaryo(Donem(2026, 2)))
    # sentetik şablonda safha genel toplamı (15) karşıt tablodaki indirilecek KDV'nin (30) %80'inin altında
    assert main(['denetle', str(p)]) == 1
    assert 'Karşıt inceleme oranı' in capsys.readouterr().out


def test_geriye_donuk(tmp_path):
    from teminat.ayar import FirmaAyari
    kok = _firma_klasoru(tmp_path / 'ORNEK')
    # MART raporu olarak programın kendi taslağını koy → tutar farkı 0 beklenir
    from teminat.taslak import taslak_uret
    tc = taslak_uret(FirmaAyari(ymm_no='00000000'), kok / '2026/02 ŞUBAT/RAPOR/ŞUBAT-2026 RAPOR.docx', kok / '2026/03 MART',
                     tmp_path / 'gecici')
    (kok / '2026/03 MART/RAPOR').mkdir()
    shutil.copy(tc.docx, kok / '2026/03 MART/RAPOR/MART-2026 RAPOR.docx')
    s = geriye_donuk(FirmaAyari(ymm_no='00000000'), kok, tmp_path / 'GDT')
    d = {x.donem: x for x in s}
    assert d[Donem(2026, 2)].durum == 'atlandı'
    assert d[Donem(2026, 3)].durum == 'tamam' and d[Donem(2026, 3)].tutar_farki == 0
    assert (tmp_path / 'GDT' / 'GERİYE DÖNÜK TEST.xlsx').exists()
    assert not (kok / '2026/03 MART/CLAUDE TASLAK').exists()     # girdi klasörüne yazılmadı
