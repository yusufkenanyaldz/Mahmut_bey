"""Madde 6: dövizli faturalarda yüzde kur toleransı, mükerrerde ardışık numara ve bulgu inceleme kayıtları."""
import pandas as pd
import pytest

from denetim import checks, inceleme
from denetim.database import DatabaseManager
from denetim.inceleme import DURUM_ACIK, DURUM_DUZELTME, DURUM_SORUN_YOK


@pytest.fixture
def db(tmp_path):
    return DatabaseManager(str(tmp_path / "firma.db"))


def fatura(no, tutar, para="TRY", kur=1.0, vkn="1111111111", tarih="2024-07-10", tip="SATIS"):
    return ({"invoice_no": no, "issue_date": tarih, "supplier_vkn": vkn, "supplier_name": f"Tedarikçi {vkn[:2]}",
             "customer_vkn": "5555555555", "total_amount": tutar, "currency": para, "exchange_rate": kur,
             "invoice_type": tip, "source": "EXCEL"},
            [{"item_name": "Motorin", "quantity": 1, "line_net": tutar}])


def yevmiye(belge, tutar, hesap="740.01", tarih="2024-07-12"):
    return [{"entry_date": tarih, "document_no": belge, "account_code": hesap, "amount": tutar},
            {"entry_date": tarih, "document_no": belge, "account_code": "320.01", "amount": -tutar}]


def yukle(db, faturalar, kayitlar):
    db.save_invoices(faturalar, "FATURA_EXCEL", "f.xlsx", str(len(faturalar)) + str(len(kayitlar)))
    db.save_journal([r for k in kayitlar for r in k], "y.xlsx", "y" + str(len(kayitlar)))


# ------------------------------------------------------------------ yüzde kur toleransı
def test_izin_verilen_fark():
    assert checks.izin_verilen_fark(100000, False, 0.01, 1) == 0.01  # TL: yalnızca sabit tolerans
    assert checks.izin_verilen_fark(100000, True, 0.01, 1) == pytest.approx(1000)
    assert checks.izin_verilen_fark(0.5, True, 0.01, 1) == 0.01  # küçük tutarda TL toleransı alt sınır
    assert checks.izin_verilen_fark(100000, True, 0.01, 0) == 0.01  # %0 → kapalı


def _kur_senaryosu(db):
    yukle(db, [
        fatura("TLA2024000000001", 12345.67),                       # TL, rakam yer değiştirme (9 TL fark)
        fatura("EUR2024000000001", 1000, "EUR", 36.5),              # 36.500 TL; muhasebe %0,5 farklı kur
        fatura("EUR2024000000002", 1000, "EUR", 36.5),              # %3 fark → tolerans dışı
        fatura("EUR2024000000003", 1000, "EUR", 36.5),              # tam tutar
    ], [yevmiye("TLA2024000000001", 12354.67), yevmiye("EUR2024000000001", 36682.5),
        yevmiye("EUR2024000000002", 37595.0), yevmiye("EUR2024000000003", 36500.0)])


def test_kur_toleransi_dovizli_kucuk_farki_bilgiye_alir_tl_farki_kalir(db):
    _kur_senaryosu(db)
    res = checks.reconcile(db.get_invoices_df(), db.get_journal_df(), ["740"], 0.01)
    fark = res["Tutar Farkları"]
    assert list(fark.columns) == checks.TUTAR_FARKI_COLS
    assert set(fark["Fatura_No"]) == {"TLA2024000000001", "EUR2024000000002"}
    tl = fark[fark["Fatura_No"] == "TLA2024000000001"].iloc[0]
    assert tl["Dovizli"] == "Hayır" and tl["Fark_TL"] == pytest.approx(9.0) and tl["Para_Birimi"] == "TRY"
    eur = fark[fark["Fatura_No"] == "EUR2024000000002"].iloc[0]
    assert eur["Dovizli"] == "Evet" and eur["Para_Birimi"] == "EUR" and eur["Kur"] == pytest.approx(36.5)
    assert list(res.kur_farki["Fatura_No"]) == ["EUR2024000000001"]
    assert res.kur_farki.iloc[0]["Fark_Yuzdesi"] == pytest.approx(0.5)
    assert res.kur_ozeti == {"kur_toleransi": 1.0, "dovizli_fatura": 3, "tolerans_ici": 1, "asan": 1}
    assert "1 dövizli fatura" in checks.kur_ozeti_metni(res.kur_ozeti)


def test_kur_toleransi_kapali_ve_genis(db):
    _kur_senaryosu(db)
    inv, jou = db.get_invoices_df(), db.get_journal_df()
    kapali = checks.reconcile(inv, jou, ["740"], 0.01, kur_toleransi=0)
    assert len(kapali["Tutar Farkları"]) == 3 and kapali.kur_farki.empty
    assert "kapalı" in checks.kur_ozeti_metni(kapali.kur_ozeti)
    genis = checks.reconcile(inv, jou, ["740"], 0.01, kur_toleransi=5)
    assert list(genis["Tutar Farkları"]["Fatura_No"]) == ["TLA2024000000001"]  # TL hiçbir yüzdeyle gizlenmez
    assert len(genis.kur_farki) == 2


def test_full_audit_kur_toleransi(db):
    _kur_senaryosu(db)
    _, sections, _, sayilar = checks.run_full_audit(db, "Aylık", 15, ["740"], 0.01, "", kur_toleransi=1)
    assert len(sections["Tutar Farkları"]) == 2 and sayilar["kur_ozeti"]["tolerans_ici"] == 1
    assert len(sayilar["kur_farki"]) == 1
    _, sections, _, sayilar = checks.run_full_audit(db, "Aylık", 15, [], 0.01, "")
    assert sayilar["kur_ozeti"] is None and sayilar["kur_farki"] is None


# ------------------------------------------------------------------ mükerrer: ardışık numara
@pytest.mark.parametrize("nolar,beklenen", [
    (["ABC2024000000017", "ABC2024000000018"], True),
    (["ABC2024000000019", "ABC2024000000017", "ABC2024000000018"], True),
    (["ABC2024000000017", "ABC2024000000019"], False),
    (["ABC2024000000017", "ABD2024000000018"], False),   # farklı seri
    (["ABC2023000000017", "ABC2024000000018"], False),   # farklı yıl
    (["F1", "F2"], False),                                # GİB biçimi değil
])
def test_ardisik_numaralar(nolar, beklenen):
    assert checks.ardisik_numaralar(nolar) is beklenen


def test_mukerrer_ardisik_sutunu_ve_sirasi():
    inv = pd.DataFrame([
        {"id": 1, "invoice_no": "ABC2024000000005", "issue_date": "2024-07-01", "supplier_vkn": "1", "supplier_name": "A",
         "total_amount": 100.0},
        {"id": 2, "invoice_no": "ABC2024000000006", "issue_date": "2024-07-01", "supplier_vkn": "1", "supplier_name": "A",
         "total_amount": 100.0},
        {"id": 3, "invoice_no": "ABC2024000000010", "issue_date": "2024-07-02", "supplier_vkn": "1", "supplier_name": "A",
         "total_amount": 50.0},
        {"id": 4, "invoice_no": "ABC2024000000090", "issue_date": "2024-07-02", "supplier_vkn": "1", "supplier_name": "A",
         "total_amount": 50.0},
    ])
    out = checks.duplicate_suspects(inv)
    assert list(out["Ardisik_Numara"]) == ["Hayır", "Evet"]  # ardışık olmayan (daha şüpheli) önce
    assert out.iloc[1]["Fatura_No"] == "ABC2024000000005, ABC2024000000006"


# ------------------------------------------------------------------ bulgu anahtarı
def test_kontrol_kodu():
    assert inceleme.kontrol_kodu("Fiyat Anomalileri (±%15)") == "FIYAT"
    assert inceleme.kontrol_kodu("Tutar Farkları") == "TUTAR_FARKI"
    assert inceleme.kontrol_kodu("Belirsiz Eşleşme (Aynı No Farklı Tedarikçi)") == "BELIRSIZ_AYNI_NO"
    assert inceleme.kontrol_kodu("Belirsiz Eşleşme (Birden Fazla Seri+Sıra Adayı)") == "BELIRSIZ_SERI_SIRA"
    assert inceleme.kontrol_kodu("Özet") is None and inceleme.kontrol_kodu("Eşleşme Özeti") is None


def test_bulgu_anahtari_normalize_ve_kararli():
    a = inceleme.bulgu_anahtari("TUTAR_FARKI", {"Tedarikci_VKN": "111 111 1111", "Fatura_No": "abc 2024000000001",
                                                "Fark_TL": 5})
    b = inceleme.bulgu_anahtari("TUTAR_FARKI", {"Tedarikci_VKN": 1111111111.0, "Fatura_No": "ABC2024000000001",
                                                "Fark_TL": 7})
    assert a == b == "TUTAR_FARKI|1111111111|ABC2024000000001"
    m1 = inceleme.bulgu_anahtari("MUKERRER", {"Tedarikci_VKN": "1", "Fatura_No": "B2, a1"})
    m2 = inceleme.bulgu_anahtari("MUKERRER", {"Tedarikci_VKN": "1", "Fatura_No": "A1, B2"})
    assert m1 == m2 == "MUKERRER|1|A1+B2"
    assert inceleme.bulgu_anahtari("FATURASIZ", {"Yevmiye_Belge_No": "gp 01"}) == "FATURASIZ|GP01"
    f = inceleme.bulgu_anahtari("FIYAT", pd.Series({"Tedarikci_VKN": "1", "Fatura_No": "X1", "Urun_Adi": "  ÇİMENTO "}))
    assert f == "FIYAT|1|X1|çimento"
    # Aynı fatura farklı kontrollerde farklı bulgudur
    assert inceleme.bulgu_anahtari("DONEM_FARKI", {"Tedarikci_VKN": "1", "Fatura_No": "X1"}) != \
        inceleme.bulgu_anahtari("TUTAR_FARKI", {"Tedarikci_VKN": "1", "Fatura_No": "X1"})


# ------------------------------------------------------------------ kayıt ve uygulama
def test_isaretle_oku_not_korunur(db):
    assert inceleme.incelemeleri_oku(db) == {}
    n = inceleme.isaretle(db, ["TUTAR_FARKI|1|A", "TUTAR_FARKI|1|A", "MUKERRER|1|A+B"], DURUM_SORUN_YOK,
                          "Kur farkı, sorun yok", tarih="2024-10-01 10:00")
    assert n == 2
    k = inceleme.incelemeleri_oku(db)
    assert k["TUTAR_FARKI|1|A"] == {"kontrol": "TUTAR_FARKI", "durum": DURUM_SORUN_YOK,
                                    "aciklama": "Kur farkı, sorun yok", "tarih": "2024-10-01 10:00"}
    inceleme.isaretle(db, ["TUTAR_FARKI|1|A"], DURUM_ACIK)  # not verilmezse korunur
    k = inceleme.incelemeleri_oku(db)["TUTAR_FARKI|1|A"]
    assert k["durum"] == DURUM_ACIK and k["aciklama"] == "Kur farkı, sorun yok" and k["tarih"] != "2024-10-01 10:00"
    inceleme.isaretle(db, ["TUTAR_FARKI|1|A"], DURUM_DUZELTME, "")
    assert inceleme.incelemeleri_oku(db)["TUTAR_FARKI|1|A"]["aciklama"] == ""
    with pytest.raises(ValueError):
        inceleme.isaretle(db, ["X|1"], "Bilinmeyen")


def test_durum_uygula_ozet_ve_gizle(db):
    df = pd.DataFrame({"Fatura_No": ["A", "B", "C"], "Tedarikci_VKN": ["1", "1", "2"], "Fark_TL": [1, 2, 3]})
    inceleme.isaretle(db, ["TUTAR_FARKI|1|A"], DURUM_SORUN_YOK, "ok")
    inceleme.isaretle(db, ["TUTAR_FARKI|2|C"], DURUM_DUZELTME, "muhasebeciye yazıldı")
    sections = inceleme.bolumlere_uygula({"Özet": pd.DataFrame({"x": [1]}), "Tutar Farkları": df},
                                         inceleme.incelemeleri_oku(db))
    assert list(sections["Özet"].columns) == ["x"]  # incelenebilir olmayan bölüm aynen kalır
    out = sections["Tutar Farkları"]
    assert list(out.columns) == ["Fatura_No", "Tedarikci_VKN", "Fark_TL"] + inceleme.INCELEME_COLS
    assert list(out[inceleme.DURUM_COL]) == [DURUM_SORUN_YOK, DURUM_ACIK, DURUM_DUZELTME]
    assert list(out[inceleme.NOT_COL]) == ["ok", "", "muhasebeciye yazıldı"]
    assert list(inceleme.sorun_yok_gizle(out)["Fatura_No"]) == ["B", "C"]
    ozet = inceleme.inceleme_ozeti(sections)
    assert ozet.to_dict("records") == [{"Kontrol": "Tutar Farkları", "Bulgu_Sayisi": 3, "Acik": 1, "Sorun_Yok": 1,
                                        "Duzeltme_Istendi": 1}]
    assert inceleme.ozet_satiri(ozet.iloc[0]) == "açık 1 | sorun yok 1 | düzeltme istendi 1"
    # Tekrar uygulamak sütunları çoğaltmaz
    again = inceleme.durum_uygula("Tutar Farkları", out, inceleme.incelemeleri_oku(db))
    assert list(again.columns) == list(out.columns)


def test_inceleme_yeniden_ice_aktarma_ve_yeniden_calistirmada_kalir(db):
    _kur_senaryosu(db)
    _, sections, _, _ = checks.run_full_audit(db, "Aylık", 15, ["740"], 0.01, "")
    sections = inceleme.bolumlere_uygula(sections, inceleme.incelemeleri_oku(db))
    fark = sections["Tutar Farkları"]
    eur = fark[fark["Fatura_No"] == "EUR2024000000002"][inceleme.ANAHTAR_COL]
    inceleme.isaretle(db, eur, DURUM_SORUN_YOK, "Kur farkı, muhasebe ödeme günü kurunu kullanmış")
    # Veriler silinip aynı dosyalar yeniden yüklenir (farklı veritabanı kimlikleri), rapor yeniden çalıştırılır
    db.clear_data()
    yukle(db, [fatura("TLA2024000000001", 12345.67), fatura("EUR2024000000002", 1000, "EUR", 36.5)],
          [yevmiye("TLA2024000000001", 12354.67), yevmiye("EUR2024000000002", 37595.0)])
    _, sections, _, _ = checks.run_full_audit(db, "Aylık", 15, ["740"], 0.01, "")
    sections = inceleme.bolumlere_uygula(sections, inceleme.incelemeleri_oku(db))
    fark = sections["Tutar Farkları"].set_index("Fatura_No")
    assert fark.loc["EUR2024000000002", inceleme.DURUM_COL] == DURUM_SORUN_YOK
    assert fark.loc["TLA2024000000001", inceleme.DURUM_COL] == DURUM_ACIK
    assert list(inceleme.sorun_yok_gizle(sections["Tutar Farkları"])["Fatura_No"]) == ["TLA2024000000001"]


def test_hesaplama_ve_alici_bolumlerinde_vkn_var():
    inv = pd.DataFrame([{"id": 1, "invoice_no": "X1", "issue_date": "2024-07-01", "supplier_vkn": "1",
                         "supplier_name": "A", "customer_vkn": "9", "customer_name": "B", "source": "XML",
                         "line_extension_amount": 100.0, "allowance_total": 0.0, "total_amount": 90.0,
                         "vat_amount": 18.0}])
    lines = pd.DataFrame([{"invoice_id": 1, "line_net": 100.0, "vat_amount": 18.0}])
    calc = checks.calculation_errors(inv, lines)
    assert set(calc["Tedarikci_VKN"]) == {"1"}
    assert inceleme.bolum_anahtarlari("Fatura Hesaplama Tutarsızlıkları", calc)[0].startswith("HESAPLAMA|1|X1|")
    assert list(checks.customer_mismatch(inv, "5")["Tedarikci_VKN"]) == ["1"]
