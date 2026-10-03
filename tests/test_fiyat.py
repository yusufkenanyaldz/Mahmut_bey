"""Fiyat anomalisi analizi: tevkifat / anahtar kelime hariç tutma, para birimine göre gruplama, yetersiz veri."""
import pandas as pd
import pytest

from denetim import checks
from denetim.database import DatabaseManager

_id = iter(range(1, 10_000))


def satir(no, urun, miktar, fiyat, tarih="2024-07-10", tip="SATIS", pb="TRY", kur=1.0, birim="C62"):
    return {"id": next(_id), "invoice_no": no, "issue_date": tarih, "supplier_vkn": "111", "supplier_name": "A",
            "item_name": urun, "item_norm": urun.lower(), "quantity": miktar, "uom": birim,
            "line_net": miktar * fiyat, "currency": pb, "exchange_rate": kur, "invoice_type": tip}


def df(*rows):
    return pd.DataFrame(list(rows))


def riskli_nolar(res):
    return sorted(res.riskli["Fatura_No"])


# ------------------------------------------------------------------ anahtar kelime listesi
def test_parse_kelime_listesi_turkce_duyarsiz():
    assert checks.parse_kelime_listesi("Hakediş, işçilik; HİZMET\n fason  dokuma ,, hakedis") == \
        ["HAKEDIS", "ISCILIK", "HIZMET", "FASON DOKUMA"]
    assert checks.parse_kelime_listesi("") == [] and checks.parse_kelime_listesi(None) == []


@pytest.mark.parametrize("urun,beklenen", [
    ("Taşeron İşçilik Hakedişi", "HAKEDIS"), ("TAŞERON İŞÇİLİK", "ISCILIK"), ("taşeron işçilik", "ISCILIK"),
    ("Fason Dokuma Hizmeti", "HIZMET"), ("Tır Bakım Onarım", "BAKIM"), ("İş Makinesi Kiralama", "KIRALAMA"),
    ("Mali Danışmanlık Ücreti", "DANISMANLIK"),
    ("Motorin", None), ("Kiralık Filo Aracı Model A", None), ("Çimento CEM I 42.5", None), (None, None)])
def test_haric_kelime_varsayilanlar(urun, beklenen):
    kelimeler = checks.parse_kelime_listesi(checks.VARSAYILAN_FIYAT_HARIC_KELIMELER)
    assert checks.haric_kelime(urun, kelimeler) == beklenen


def test_fatura_tipi_ve_para_birimi():
    assert checks.tevkifatli_fatura("TEVKIFAT") and checks.tevkifatli_fatura("tevkifat")
    assert not checks.tevkifatli_fatura("TEVKIFATIADE") and not checks.tevkifatli_fatura("SATIS")
    assert checks.iade_faturasi("IADE") and checks.iade_faturasi("TEVKIFATIADE") and not checks.iade_faturasi(None)
    assert checks.para_birimi(None) == "TRY" and checks.para_birimi("TL") == "TRY"
    assert checks.para_birimi(" eur ") == "EUR"


# ------------------------------------------------------------------ hariç tutma kuralları
def _hizmet_senaryosu():
    """Çimento normal fiyatlı, hakediş her ay farklı tutarlı (tevkifatlı), kiralama tevkifatsız ama değişken."""
    return df(
        satir("C1", "Çimento", 10, 100), satir("C2", "Çimento", 10, 102), satir("C3", "Çimento", 10, 98),
        satir("H1", "Taşeron İşçilik Hakedişi", 1, 200_000, tip="TEVKIFAT"),
        satir("H2", "Taşeron İşçilik Hakedişi", 1, 600_000, tip="TEVKIFAT"),
        satir("H3", "Taşeron İşçilik Hakedişi", 1, 400_000, tip="TEVKIFAT"),
        satir("K1", "Vinç Kiralama", 1, 5_000), satir("K2", "Vinç Kiralama", 1, 15_000),
        satir("K3", "Vinç Kiralama", 1, 10_000),
        satir("I1", "Çimento", 2, 300, tip="IADE"),
    )


def test_varsayilan_kurallar_hizmet_kalemlerini_haric_tutar():
    res = checks.fiyat_analizi(_hizmet_senaryosu(), "Aylık", 15)
    assert res.riskli.empty
    assert sorted(res.satirlar["Fatura_No"]) == ["C1", "C2", "C3"]
    o = res.ozet
    assert (o["incelenen"], o["iade"], o["tevkifat"], o["kelime"], o["riskli"]) == (3, 1, 3, 3, 0)
    nedenler = dict(zip(res.haric["Fatura_No"], res.haric["Analiz_Disi_Nedeni"]))
    assert nedenler["H1"] == checks.FIYAT_NEDEN_TEVKIFAT  # tevkifat kuralı kelimeden önce gelir
    assert nedenler["K1"] == "Anahtar kelime (KIRALAMA)"
    assert nedenler["I1"] == checks.FIYAT_NEDEN_IADE
    assert len(res.haric) == 7


def test_tevkifat_kurali_kapatilabilir():
    kurallar = checks.FiyatKurallari(tevkifat_haric=False, kelimeler=checks.parse_kelime_listesi("KİRALAMA"))
    res = checks.fiyat_analizi(_hizmet_senaryosu(), "Aylık", 15, kurallar)
    assert riskli_nolar(res) == ["H1", "H2"]  # hakediş artık analizde: 200 bin ve 600 bin ortalamadan sapıyor
    assert res.ozet["tevkifat"] == 0 and res.ozet["kelime"] == 3


def test_kelime_kurali_kapatilabilir_ve_bos_liste():
    for kurallar in (checks.FiyatKurallari(kelime_haric=False), checks.FiyatKurallari(kelimeler=[])):
        res = checks.fiyat_analizi(_hizmet_senaryosu(), "Aylık", 15, kurallar)
        assert riskli_nolar(res) == ["K1", "K2"] and res.ozet["kelime"] == 0 and res.ozet["tevkifat"] == 3


def test_ozel_kelime_listesi():
    kurallar = checks.FiyatKurallari(kelimeler=checks.parse_kelime_listesi("vinç"))
    res = checks.fiyat_analizi(_hizmet_senaryosu(), "Aylık", 15, kurallar)
    assert res.riskli.empty and res.ozet["kelime"] == 3


def test_iki_kural_da_kapali_eski_davranis():
    kurallar = checks.FiyatKurallari(tevkifat_haric=False, kelime_haric=False)
    res = checks.fiyat_analizi(_hizmet_senaryosu(), "Aylık", 15, kurallar)
    assert riskli_nolar(res) == ["H1", "H2", "K1", "K2"]
    assert res.ozet["iade"] == 1 and len(res.haric) == 1  # iade her zaman analiz dışı


# ------------------------------------------------------------------ para birimi
def test_para_birimine_gore_ayri_grup_ve_belge_pb_ile_sapma():
    lines = df(
        satir("T1", "Motorin", 1000, 43, birim="LTR"), satir("T2", "Motorin", 1000, 44, birim="LTR"),
        satir("T3", "Motorin", 1000, 42, birim="LTR"),
        # EUR motorin: TL karşılığı (1,70 × 37 ≈ 63 TL) TL motorinden çok farklı ama kendi içinde normal
        satir("E1", "Motorin", 500, 1.70, pb="EUR", kur=36.5, birim="LTR"),
        satir("E2", "Motorin", 500, 1.72, pb="EUR", kur=37.5, birim="LTR"),
        satir("E3", "Motorin", 500, 1.68, pb="EUR", kur=38.5, birim="LTR"),
    )
    res = checks.fiyat_analizi(lines, "Aylık", 15)
    assert res.riskli.empty
    s = res.satirlar.set_index("Fatura_No")
    assert s.at["E1", "Para_Birimi"] == "EUR" and s.at["T1", "Para_Birimi"] == "TRY"
    assert s.at["E1", "Birim_Fiyat"] == pytest.approx(1.70)
    assert s.at["E1", "Birim_Fiyat_TL"] == pytest.approx(62.05)
    assert s.at["E1", "AOBF"] == pytest.approx(1.70)
    assert s.at["E3", "Fark_Yuzdesi"] == pytest.approx(-1.18, abs=0.01)  # kur artışı sapma yaratmaz
    assert s.at["T2", "AOBF"] == pytest.approx(43)


def test_para_birimi_icinde_sisirme_yakalanir():
    lines = df(*(satir(f"E{i}", "Motorin", 500, 1.70, pb="EUR", kur=37) for i in range(3)),
               satir("EX", "Motorin", 500, 2.60, pb="EUR", kur=37))
    assert riskli_nolar(checks.fiyat_analizi(lines, "Aylık", 15)) == ["EX"]


# ------------------------------------------------------------------ yetersiz veri
def test_yetersiz_veri_riskli_sayilmaz_bilgi_olarak_kalir():
    lines = df(satir("A1", "Rulman", 10, 100), satir("A2", "Rulman", 10, 160),
               satir("B1", "Sac", 10, 100), satir("B2", "Sac", 10, 101), satir("B3", "Sac", 10, 99),
               satir("BX", "Sac", 10, 150))
    res = checks.fiyat_analizi(lines, "Aylık", 15, checks.FiyatKurallari(min_alim=3))
    assert riskli_nolar(res) == ["BX"]
    s = res.satirlar.set_index("Fatura_No")
    assert s.at["A2", "Risk_Durumu"] == checks.RISK_YETERSIZ and s.at["A1", "Risk_Durumu"] == checks.RISK_YETERSIZ
    assert s.at["A2", "Donemdeki_Alim_Sayisi"] == 2
    assert res.ozet["yetersiz_veri"] == 2
    bilgi = res.haric[res.haric["Analiz_Disi_Nedeni"].str.startswith(checks.FIYAT_NEDEN_YETERSIZ)]
    assert sorted(bilgi["Fatura_No"]) == ["A1", "A2"] and "2 alım < 3" in bilgi["Analiz_Disi_Nedeni"].iloc[0]
    # min_alim=1 (kural kapalı) → iki alımlı grupta da riskli
    assert riskli_nolar(checks.fiyat_analizi(lines, "Aylık", 15, checks.FiyatKurallari(min_alim=1))) == \
        ["A1", "A2", "BX"]


def test_yetersiz_veri_donem_bazinda():
    """Aynı ürünün alımları farklı aylardaysa aylık dönemde yetersiz, çeyreklikte yeterli."""
    lines = df(satir("M1", "Sac", 10, 100, "2024-07-05"), satir("M2", "Sac", 10, 100, "2024-08-05"),
               satir("M3", "Sac", 10, 160, "2024-08-20"), satir("M4", "Sac", 10, 100, "2024-09-05"))
    assert checks.fiyat_analizi(lines, "Aylık", 15).riskli.empty
    assert riskli_nolar(checks.fiyat_analizi(lines, "Çeyreklik", 15)) == ["M3"]


def test_miktar_sifir_ve_bos_veri():
    res = checks.fiyat_analizi(df(satir("Z1", "Sac", 0, 0)), "Aylık", 15)
    assert res.satirlar.empty and res.ozet["miktar"] == 1 and list(res.haric["Fatura_No"]) == ["Z1"]
    bos = checks.fiyat_analizi(pd.DataFrame(), "Aylık", 15)
    assert bos.satirlar.empty and bos.haric.empty and bos.ozet["toplam"] == 0
    assert list(bos.satirlar.columns) == checks.FIYAT_COLS


# ------------------------------------------------------------------ özet ve genel rapor
def test_ozet_metni_ve_tablosu():
    res = checks.fiyat_analizi(_hizmet_senaryosu(), "Aylık", 15)
    metin = checks.fiyat_ozeti_metni(res.ozet)
    assert "3 satır incelendi, 0 riskli" in metin
    assert "iade 1" in metin and "tevkifat 3" in metin and "anahtar kelime 3" in metin and "Yetersiz veri" in metin
    tablo = checks.fiyat_ozeti_df(res.ozet)
    assert dict(zip(tablo["Kalem"], tablo["Satir_Sayisi"]))["Analiz dışı: tevkifat"] == 3
    kapali = checks.fiyat_analizi(_hizmet_senaryosu(), "Aylık", 15,
                                  checks.FiyatKurallari(tevkifat_haric=False, kelime_haric=False))
    assert "tevkifat (kural kapalı)" in checks.fiyat_ozeti_metni(kapali.ozet)
    assert "anahtar kelime (kural kapalı)" in checks.fiyat_ozeti_metni(kapali.ozet)
    assert checks.fiyat_ozeti_metni({}) == "" and checks.fiyat_ozeti_df({}).empty


def test_price_anomalies_geriye_uyumlu():
    out = checks.price_anomalies(_hizmet_senaryosu(), "Aylık", 15)
    assert list(out.columns) == checks.FIYAT_COLS and len(out) == 3


def test_full_audit_fiyat_kurallari(tmp_path):
    db = DatabaseManager(str(tmp_path / "f.db"))
    invs = []
    for no, tip, urun, fiyat in [("C1", "SATIS", "Çimento", 100), ("C2", "SATIS", "Çimento", 100),
                                 ("C3", "SATIS", "Çimento", 100), ("CX", "SATIS", "Çimento", 150),
                                 ("H1", "TEVKIFAT", "Taşeron Hakedişi", 1000), ("H2", "TEVKIFAT", "Taşeron Hakedişi", 3000),
                                 ("H3", "TEVKIFAT", "Taşeron Hakedişi", 2000)]:
        invs.append(({"invoice_no": no, "issue_date": "2024-07-10", "supplier_vkn": "1111111111",
                      "supplier_name": "A", "total_amount": 10 * fiyat, "invoice_type": tip, "source": "EXCEL"},
                     [{"item_name": urun, "quantity": 10, "line_net": 10 * fiyat}]))
    db.save_invoices(invs, "FATURA_EXCEL", "f.xlsx", "x")
    _, sections, _, sayilar = checks.run_full_audit(db, "Aylık", 15, [], 0.01, "")
    assert list(sections["Fiyat Anomalileri (±%15)"]["Fatura_No"]) == ["CX"]
    assert sayilar["fiyat_ozeti"]["tevkifat"] == 3 and len(sayilar["fiyat_haric"]) == 3
    _, sections, _, _ = checks.run_full_audit(db, "Aylık", 15, [], 0.01, "",
                                              fiyat_kurallari=checks.FiyatKurallari(False, False))
    assert set(sections["Fiyat Anomalileri (±%15)"]["Fatura_No"]) == {"CX", "H1", "H2"}
