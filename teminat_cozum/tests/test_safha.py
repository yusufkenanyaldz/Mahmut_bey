from teminat.ayar import EkliKurali, FirmaAyari
from teminat.rapor.safha import BILGI_ISTEME, ITHALAT, KIT, firma_anahtari, safha_hesapla


def _takip(*satirlar):
    return [{'firma': f, 'kdv': k, 'aciklama': '', 'smmm': '', 'ymm': y} for f, k, y in satirlar]


def _liste(*satirlar):
    return [{'satici': f, 'vkn': v, 'kdv': k} for f, v, k in satirlar]


def test_firma_anahtari_turkce_duyarsiz():
    assert firma_anahtari('DELTA GALVANİZLİ YAPI') == firma_anahtari('Delta Galvanizli Yapı Elem.')


def test_ekli_secimi_ilk_n_ve_ithalat_sonda():
    ayar = FirmaAyari(ekli_kurali=EkliKurali(yurtici_adet=2))
    takip = _takip(('A FİRMA', 50.0, ''), ('B FİRMA', 30.0, 'YMM X'), ('C FİRMA', 10.0, ''), ('İTH GMBH', 5.0, ''))
    liste = _liste(('A FİRMA', '1', 50.0), ('B FİRMA', '2', 30.0), ('C FİRMA', '3', 10.0), ('İTH GMBH', '1111111111', 5.0))
    s = safha_hesapla(takip, liste, 100.0, ayar)
    assert [(r.firma, r.sekil) for r in s.ekli] == [('A FİRMA', KIT), ('B FİRMA', BILGI_ISTEME), ('İTH GMBH', ITHALAT)]
    assert [r.firma for r in s.muhafaza] == ['C FİRMA']
    assert (s.tot_e, s.tot_m, s.tot) == (85.0, 10.0, 95.0)


def test_ekli_sayisi_yuzde_80_icin_artar():
    ayar = FirmaAyari(ekli_kurali=EkliKurali(yurtici_adet=2))
    takip = _takip(*[(f'F{i} ŞİRKET', 20.0, '') for i in range(5)])
    liste = _liste(*[(f'F{i} ŞİRKET', str(i), 20.0) for i in range(5)])
    s = safha_hesapla(takip, liste, 100.0, ayar)
    assert s.ekli_adet == 4 and s.tot_e == 80.0
    assert any('EKLİ tutanak sayısı 4' in n[2] for n in s.notlar)


def test_kdvsiz_firma_alinmaz_ve_fark_uyarilir():
    takip = _takip(('A FİRMA', 50.0, ''), ('B FİRMA', 0.0, ''), ('C FİRMA', 10.0, ''))
    liste = _liste(('A FİRMA', '1', 49.0), ('C FİRMA', '3', 10.0))
    s = safha_hesapla(takip, liste, 60.0, FirmaAyari())
    assert [r.firma for r in s.satirlar] == ['A FİRMA', 'C FİRMA']
    metin = ' '.join(n[2] for n in s.notlar)
    assert "B FİRMA KDV'siz" in metin
    assert 'takip listesinde 50,00, indirilecek listede 49,00' in metin


def test_ozel_inceleme_sekli_ve_osb_listeden():
    ayar = FirmaAyari(ekli_kurali=EkliKurali(yurtici_adet=1, osb_listeden=True))
    takip = _takip(('A FİRMA', 50.0, ''), ('TÜRK STANDARDLARI ENSTİTÜSÜ', 3.0, 'YMM X'))
    liste = _liste(('A FİRMA', '1', 50.0), ('TÜRK STANDARDLARI ENSTİTÜSÜ', '2', 3.0),
                   ('GAZİANTEP ORGANİZE SANAYİ BÖLGESİ', '3', 4.0))
    s = safha_hesapla(takip, liste, 57.0, ayar)
    sek = {r.firma: r.sekil for r in s.satirlar}
    assert sek['TÜRK STANDARDLARI ENSTİTÜSÜ'] == 'TSE'
    assert sek['GAZİANTEP ORGANİZE SANAYİ BÖLGESİ'] == 'OSB-SU'
    # listeden eklenen OSB EKLİ olmaz, muhafazanın sonuna yazılır
    assert s.muhafaza[-1].firma == 'GAZİANTEP ORGANİZE SANAYİ BÖLGESİ'


def test_ayni_tutar_iki_firmada_satir_kaymasi_uyarisi():
    takip = _takip(('A FİRMA', 50.0, ''), ('B FİRMA', 50.0, ''))
    liste = _liste(('A FİRMA', '1', 50.0), ('B FİRMA', '2', 50.0))
    s = safha_hesapla(takip, liste, 100.0, FirmaAyari())
    assert any(n[1] == 'Satır kayması' for n in s.notlar)


def test_firma_adi_duzeltmesi():
    ayar = FirmaAyari(firma_adi_duzeltmeleri={'A FİRMA': 'A FİRMASI SAN. A.Ş.'})
    s = safha_hesapla(_takip(('A FİRMA', 50.0, '')), _liste(('A FİRMA', '1', 50.0)), 50.0, ayar)
    assert s.ekli[0].firma == 'A FİRMASI SAN. A.Ş.'
