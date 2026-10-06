import unicodedata

import pytest

import sentetik as S
from teminat.okuyucular import kdv1, listeler, teminat
from teminat.okuyucular.girdiler import GirdiHatasi, girdileri_oku
from teminat.ortak import Donem, dosyalari_bul, katla, num, pct, pct2, tr


def test_turkce_sayi_bicimi():
    assert tr(1234567.891) == '1.234.567,89'
    assert tr(-0.5) == '-0,50'
    assert num('1.234.567,89') == 1234567.89
    assert pct(0.2336) == '% 23,36'
    assert pct2(0.9649) == '%96,49'


def test_donem():
    assert Donem.coz('2026-02') == Donem(2026, 2)
    assert Donem.coz('ŞUBAT-2026') == Donem(2026, 2)
    assert Donem.coz('subat/2026') == Donem(2026, 2)
    assert Donem.coz('202612') == Donem(2026, 12)
    assert Donem(2026, 1).onceki() == Donem(2025, 12)
    assert Donem(2025, 12).sonraki() == Donem(2026, 1)
    assert Donem(2026, 2).egik == 'ŞUBAT/2026' and Donem(2026, 2).tire == 'ŞUBAT-2026'
    with pytest.raises(ValueError):
        Donem.coz('2026')


def test_katla_ve_nfd_dosya_adlari(tmp_path):
    assert katla('İNDİRİLECEK yüklenilen ışık') == 'INDIRILECEK yuklenilen isik'
    nfd = unicodedata.normalize('NFD', 'yüklenilen tutanak çalışması.xls')
    (tmp_path / nfd).write_text('x')
    (tmp_path / '~$yüklenilen tutanak.xls').write_text('x')        # Excel kilit dosyası atlanır
    (tmp_path / 'CLAUDE TASLAK').mkdir()
    (tmp_path / 'CLAUDE TASLAK' / 'yüklenilen tutanak eski.xls').write_text('x')   # çıktı klasörü atlanır
    bulunan = dosyalari_bul(tmp_path, 'YÜKLENİLEN TUTANAK*.xls*')
    assert len(bulunan) == 1


def test_kdv1_metinden_tum_alanlar():
    s = S.Senaryo(Donem(2026, 3), k448=(300_000.0, '20', '5/10', 30_000.0), i318=(6_000_000.0, 0.0, 0.0), i448=(300_000.0, 0.0),
                  diger={'iade': (100_000.0, 18_000.0), 'amort': (600_000.0, 6_000.0)}, duzeltme='Hatalı beyan')
    k = kdv1.kdv1_metinden(S.kdv1_metni(s))
    assert kdv1.kdv1_donem(k) == (2026, 3)
    assert k['yurtici_alim'] == s.yurtici_alim and k['sorumlu'] == s.sorumlu and k['ithal'] == s.ithal
    assert k['matrah_toplam'] == s.matrah_toplam and k['hesaplanan'] == s.hesaplanan
    assert k['devreden_onceki'] == s.devreden_onceki and k['indirim_toplam'] == s.indirim_toplam
    assert k['iade_gereken'] == s.iade_gereken and k['sonraki_devreden'] == s.sonraki_devreden
    assert k['r701'] == (5_000_000.0, '20', 1_000_000.0)
    assert k['yurtici'] == [(10_000.0, 1, 100.0), (60_000_000.0, 20, 12_000_000.0)]
    assert k['k410'] == (40_000_000.0, '20', '4/10', 4_800_000.0)
    assert k['k448'] == (300_000.0, '20', '5/10', 30_000.0)
    assert k['d_iade'] == (100_000.0, 18_000.0) and k['d_amort'] == (600_000.0, 6_000.0) and k['d_kur'] is None
    assert [o for o, _, _ in k['oranlar']] == [1, 10, 20]
    assert k['301_yuklenilen'] == 400_000.0 and k['410_iade'] == 3_200_000.0
    assert k['i318'] == (6_000_000.0, 0.0, 0.0) and k['i448'] == (300_000.0, 0.0)
    assert k['duzeltme'] == 'Hatalı beyan'
    kal = kdv1.iade_kalemleri(k)
    assert kal == {'301': (2_000_000.0, 400_000.0), '318': (6_000_000.0, 0.0), '410': (40_000_000.0, 3_200_000.0),
                   '448': (300_000.0, 0.0)}


def test_kdv1_dosya_adi(tmp_path):
    for ad in ('KDV 1.pdf', 'KDV 2.pdf', 'KDV 10.pdf', 'kdv1 beyanname.pdf', 'notlar.pdf'):
        (tmp_path / ad).write_text('x')
    assert sorted(p.name for p in kdv1.kdv1_bul(tmp_path)) == ['KDV 1.pdf', 'kdv1 beyanname.pdf']


def test_listeler(tmp_path):
    s = S.Senaryo(Donem(2026, 3))
    g = S.girdi_klasoru(tmp_path, s)
    ind = listeler.indirilecek_oku(listeler.indirilecek_bul(g)[0])
    assert round(sum(r['kdv'] for r in ind), 2) == round(s.base, 2)
    assert {r['vkn'] for r in ind if r['satici'] == 'PI TRADING GMBH'} == {'1111111111'}
    assert listeler.donem_degeri(ind[0]['donem']) == (2026, 3)
    takip = listeler.takip_oku(listeler.takip_bul(g)[0])
    assert takip[0]['firma'] == 'ALFA ÇELİK SAN. VE TİC. A.Ş.' and takip[0]['ymm'] == 'YMM ALİ VELİ'
    assert all(t['firma'] != 'SİGMA KIRTASİYE' for t in takip)
    yuk = listeler.yuklenilen_oku(listeler.yuklenilen_bul(g)[0])
    assert [(y['firma'], y['per']) for y in yuk] == [('GAMA DEMİR SAN. A.Ş.', (2026, 3)), ('ALFA ÇELİK SAN. VE TİC. A.Ş.', (2026, 2))]


@pytest.mark.parametrize('deger,beklenen', [(202602.0, (2026, 2)), ('2026/02', (2026, 2)), ('02/2026', (2026, 2)),
                                            ('2026-2', (2026, 2)), ('', None), ('abc', None), (202613, None)])
def test_donem_degeri(deger, beklenen):
    assert listeler.donem_degeri(deger) == beklenen


def test_teminat_dilekcesi(tmp_path):
    s = S.Senaryo(Donem(2026, 3))
    S.dilekce_docx(tmp_path / 'TEMİNAT MEKTUBU KABUL DİLEKÇESİ - MART 2026.docx', s)
    p = teminat.dilekce_bul(tmp_path)[0]
    t = teminat.dilekce_metinden(teminat.dilekce_metni(p))
    assert t['301'] == 360_000.0 and t['410'] == 2_880_000.0 and t['toplam'] == 3_240_000.0
    assert t['tarih'] == '26.03.2026' and t['no'] == '1234567'
    assert t['banka'] == 'ÖRNEK KATILIM BANKASI A.Ş. Gaziantep Şubesi'


def test_girdiler_zorunlu_dosyalar(tmp_path):
    g = S.girdi_klasoru(tmp_path, S.Senaryo(Donem(2026, 3)))
    assert girdileri_oku(g).donem == Donem(2026, 3)
    (g / 'KDV 1.pdf').unlink()
    with pytest.raises(GirdiHatasi, match='KDV 1'):
        girdileri_oku(g)


def test_girdiler_birden_fazla_aday_uyarir(tmp_path):
    g = S.girdi_klasoru(tmp_path, S.Senaryo(Donem(2026, 3)))
    (g / 'yüklenilen tutanak çalışması - Kopya.xls').write_bytes((g / 'yüklenilen tutanak çalışması.xls').read_bytes())
    notlar = girdileri_oku(g).notlar
    assert any('birden fazla' in n[2] for n in notlar)


def test_gercek_pdf_okunur(tmp_path):
    """pdfplumber ile gerçek bir PDF'ten (reportlab ile üretilmiş) beyanname okunur."""
    pytest.importorskip('reportlab')
    if not S.yazi_tipi():
        pytest.skip('Türkçe karakterli yazı tipi yok')
    from conftest import GERCEK_PDF_METNI
    s = S.Senaryo(Donem(2026, 3), k448=(300_000.0, '20', '5/10', 30_000.0), i318=(6_000_000.0, 0.0, 0.0))
    p = S.kdv1_pdf(tmp_path / 'KDV 1.pdf', s)
    k = kdv1.kdv1_metinden(GERCEK_PDF_METNI(p))
    beklenen = kdv1.kdv1_metinden(S.kdv1_metni(s))
    for alan in ('matrah_toplam', 'hesaplanan', 'yurtici_alim', 'sorumlu', 'ithal', 'devreden_onceki', 'indirim_toplam',
                 'iade_gereken', 'sonraki_devreden', 'r701', 'yurtici', 'k410', 'k448', 'oranlar', 'i318', 'donem'):
        assert k[alan] == beklenen[alan], alan
