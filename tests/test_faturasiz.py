"""Faturası bulunmayan yevmiye kayıtları: borç yönü filtresi, hariç önekler, işaretsiz Tutar ve öncelik."""
import pandas as pd
import pytest

from denetim import checks
from denetim.database import DatabaseManager


def adaylar(*rows):
    cols = ["Yevmiye_Belge_No", "Yevmiye_Tarihi", "Yevmiye_Tutari", "Hesaplar"]
    return pd.DataFrame(columns=cols) if not rows else pd.DataFrame([{"Yevmiye_Belge_No": no, "Yevmiye_Tarihi": "2024-07-31", "Yevmiye_Tutari": tutar,
                          "Hesaplar": "153.01"} for no, tutar in rows])


def kayit(belge, hesap, tutar, tarih="2024-07-15"):
    return {"entry_date": tarih, "document_no": belge, "document_no_norm": belge.replace(" ", "").upper(),
            "account_code": hesap, "account_norm": hesap.replace(".", ""), "amount": tutar}


FATURA = pd.DataFrame([{"invoice_no": "ABC2024000000001", "invoice_no_norm": "ABC2024000000001",
                        "issue_date": "2024-07-15", "supplier_vkn": "111", "supplier_name": "A",
                        "total_amount": 1000, "exchange_rate": 1.0, "invoice_type": "SATIS"}])


# ------------------------------------------------------------------ önek listesi
def test_parse_onek_listesi_turkce_ve_ayrac():
    assert checks.parse_onek_listesi("Bordro, amort; Açılış\nkapanış ,, GP-, devir") == \
        ["BORDRO", "AMORT", "ACILIS", "KAPANIS", "GP", "DEVIR"]
    assert checks.parse_onek_listesi("") == [] and checks.parse_onek_listesi(None) == []
    assert checks.parse_onek_listesi("GP, gp, G.P.") == ["GP"]  # tekrarsız


@pytest.mark.parametrize("belge,beklenen", [
    ("BORDRO-2024-07", "BORDRO"),
    ("bordro 07", "BORDRO"),
    ("AÇILIŞ-2024", "ACILIS"),
    ("ACILIS FİŞİ", "ACILIS"),
    ("açılış", "ACILIS"),
    ("Kapanış/2024", "KAPANIS"),
    ("DEVİR-01", "DEVIR"),
    ("devir", "DEVIR"),
    ("GP-07001", "GP"),
    ("gp 12", "GP"),
    ("MAHSUP-15", "MAHSUP"),
    ("ABC2024000000001", None),
    ("FIS-778", None),
    ("GPS2024000000123", None),    # tam GİB biçimli fatura no hiçbir önekle hariç tutulmaz
    ("GPS 2024 000000123", None),
    ("", None),
    (None, None),
])
def test_haric_onek_eslesmesi(belge, beklenen):
    onekler = checks.parse_onek_listesi(checks.VARSAYILAN_HARIC_ONEKLER)
    assert checks.haric_onek(belge, onekler) == beklenen


def test_haric_onek_kullanici_listesi():
    assert checks.haric_onek("SMM-07", checks.parse_onek_listesi("smm")) == "SMM"
    assert checks.haric_onek("BORDRO-07", []) is None
    assert checks.haric_onek("İADE-5", checks.parse_onek_listesi("iade")) == "IADE"  # İ/i duyarsız
    assert checks.haric_onek("IADE-5", checks.parse_onek_listesi("İade")) == "IADE"


# ------------------------------------------------------------------ öncelik
@pytest.mark.parametrize("belge,oncelik", [
    ("ABC2024000000123", "Yüksek"),
    ("ABC 2024 000000123", "Yüksek"),
    ("ABC-2024-123", "Yüksek"),
    ("ABC123", "Yüksek"),            # seri+sıra çözümlemesine uyuyor
    ("BORDRO-2024-07", "Düşük"),
    ("FIS-778X", "Düşük"),
    ("SMM-07", "Yüksek"),            # 3 harf + sayı: seri+sıra çözümlemesine uyar (hariç önekle ayıklanabilir)
    ("MASRAF-1", "Düşük"),
    ("GP-07001", "Düşük"),
    ("12345", "Düşük"),
])
def test_belge_onceligi(belge, oncelik):
    assert checks.belge_onceligi(belge) == oncelik


# ------------------------------------------------------------------ ayırma
def test_borc_yonu_filtresi_ve_oncelik_siralamasi():
    df = adaylar(("SMM-07", -50000.0),            # 153 alacak: stoktan çıkış
                 ("XYZ2024000000005", 0.0),       # net sıfır
                 ("XYZ2024000000006", 0.005),     # tolerans içinde
                 ("MASRAF-1", 9000.0),               # düşük öncelik, büyük tutar
                 ("ABC2024000000009", 100.0),
                 ("ABC2024000000010", 5000.0),
                 ("MASRAF-2", 300.0),
                 ("GP-07001", 2500.0))            # hariç önek
    liste, haric, ozet = checks.faturasiz_kayitlari_ayir(df, None, 0.01, isaretli=True)
    assert list(liste["Yevmiye_Belge_No"]) == ["ABC2024000000010", "ABC2024000000009", "MASRAF-1", "MASRAF-2"]
    assert list(liste["Oncelik"]) == ["Yüksek", "Yüksek", "Düşük", "Düşük"]
    assert list(liste.columns) == checks.FATURASIZ_COLS
    nedenler = dict(zip(haric["Yevmiye_Belge_No"], haric["Haric_Tutulma_Nedeni"]))
    assert nedenler == {"SMM-07": "Alacak yönlü", "XYZ2024000000005": "Alacak yönlü",
                        "XYZ2024000000006": "Alacak yönlü", "GP-07001": "Hariç önek (GP)"}
    assert ozet["listelenen"] == 4 and ozet["yuksek"] == 2 and ozet["dusuk"] == 2
    assert ozet["alacak_yonlu"] == 3 and ozet["haric_onek"] == 1 and ozet["isaretli"]
    metin = checks.faturasiz_ozeti_metni(ozet)
    assert "4 listelendi" in metin and "alacak yönlü 3" in metin and "hariç önek 1" in metin


def test_bos_onek_listesi_yalnizca_yon_filtresi():
    df = adaylar(("BORDRO-2024-07", 1000.0), ("SMM-07", -10.0))
    liste, haric, ozet = checks.faturasiz_kayitlari_ayir(df, [], 0.01, isaretli=True)
    assert list(liste["Yevmiye_Belge_No"]) == ["BORDRO-2024-07"]
    assert list(haric["Yevmiye_Belge_No"]) == ["SMM-07"] and ozet["haric_onek"] == 0


def test_isaretsiz_tutar_yon_filtresi_uygulanmaz():
    df = adaylar(("SMM-07", 50000.0), ("ABC2024000000001", 100.0), ("AMORT-07", 300.0))
    liste, haric, ozet = checks.faturasiz_kayitlari_ayir(df, None, 0.01, isaretli=False)
    assert list(liste["Yevmiye_Belge_No"]) == ["SMM-07", "ABC2024000000001"]  # ikisi de Yüksek; tutara göre
    assert list(haric["Haric_Tutulma_Nedeni"]) == ["Hariç önek (AMORT)"]
    assert ozet["alacak_yonlu"] == 0 and not ozet["isaretli"]
    assert "yön filtresi uygulanmadı" in checks.faturasiz_ozeti_metni(ozet)
    assert checks.faturasiz_ozeti_df(ozet)["Kayit_Sayisi"].isna().sum() == 1


def test_bos_aday_listesi():
    liste, haric, ozet = checks.faturasiz_kayitlari_ayir(adaylar(), None)
    assert liste.empty and haric.empty and ozet["listelenen"] == 0
    assert list(liste.columns) == checks.FATURASIZ_COLS and list(haric.columns) == checks.HARIC_COLS


def test_yevmiye_isaretli():
    assert checks.yevmiye_isaretli(pd.DataFrame([kayit("A", "153", 10), kayit("A", "320", -10)]))
    assert not checks.yevmiye_isaretli(pd.DataFrame([kayit("A", "153", 10), kayit("A", "320", 10)]))


# ------------------------------------------------------------------ mutabakat entegrasyonu
def _isaretli_yevmiye():
    return pd.DataFrame([
        kayit("ABC2024000000001", "153.01", 1000), kayit("ABC2024000000001", "320.01", -1200),
        kayit("XYZ2024000000777", "153.01", 4000), kayit("XYZ2024000000777", "320.01", -4800),  # sahte fatura
        kayit("SMM-07", "621.01", 3000, "2024-07-31"), kayit("SMM-07", "153.01", -3000, "2024-07-31"),
        kayit("Bordro-07", "770.01", 9000, "2024-07-31"), kayit("Bordro-07", "335.01", -9000, "2024-07-31"),
        kayit("Açılış", "153.01", 20000, "2024-01-01"), kayit("Açılış", "500.01", -20000, "2024-01-01"),
        kayit("MASRAF-55", "770.01", 750), kayit("MASRAF-55", "100.01", -750),
    ])


def test_reconcile_faturasiz_isaretli():
    res = checks.reconcile(FATURA, _isaretli_yevmiye(), ["153", "770"])
    liste = res["Faturası Bulunmayan Yevmiye Kayıtları"]
    assert list(liste["Yevmiye_Belge_No"]) == ["XYZ2024000000777", "MASRAF-55"]
    assert list(liste["Oncelik"]) == ["Yüksek", "Düşük"]
    nedenler = dict(zip(res.faturasiz_haric["Yevmiye_Belge_No"], res.faturasiz_haric["Haric_Tutulma_Nedeni"]))
    assert nedenler == {"SMM-07": "Alacak yönlü", "Bordro-07": "Hariç önek (BORDRO)",
                        "Açılış": "Hariç önek (ACILIS)"}
    assert res.faturasiz_ozeti["alacak_yonlu"] == 1 and res.faturasiz_ozeti["haric_onek"] == 2
    # Önek listesi boşken yalnızca yön filtresi
    res2 = checks.reconcile(FATURA, _isaretli_yevmiye(), ["153", "770"], haric_onekler=[])
    assert set(res2["Faturası Bulunmayan Yevmiye Kayıtları"]["Yevmiye_Belge_No"]) == \
        {"XYZ2024000000777", "MASRAF-55", "Bordro-07", "Açılış"}


def test_reconcile_faturasiz_isaretsiz_tutar():
    yev = _isaretli_yevmiye()
    yev["amount"] = yev["amount"].abs()  # tek "Tutar" sütunu, işaretsiz
    res = checks.reconcile(FATURA, yev, ["153"])
    assert not res.faturasiz_ozeti["isaretli"]
    # SMM-07'nin 153 satırı (alacak) yön bilinmediği için listede kalır
    assert "SMM-07" in set(res["Faturası Bulunmayan Yevmiye Kayıtları"]["Yevmiye_Belge_No"])


def test_full_audit_faturasiz_ozeti_ve_not(tmp_path):
    db = DatabaseManager(str(tmp_path / "f.db"))
    db.save_invoices([({"invoice_no": "ABC2024000000001", "issue_date": "2024-07-15", "supplier_vkn": "111",
                        "supplier_name": "A", "total_amount": 1000, "source": "EXCEL"},
                       [{"item_name": "Vida", "quantity": 1, "line_net": 1000}])], "FATURA_EXCEL", "f", "h")
    rows = [dict(r, description="") for r in _isaretli_yevmiye().to_dict("records")]
    db.save_journal(rows, "y.xlsx", "y")
    _, sections, notes, sayilar = checks.run_full_audit(db, "Aylık", 15, ["153", "770"], 0.01, "",
                                                        checks.parse_onek_listesi("bordro"))
    liste = sections["Faturası Bulunmayan Yevmiye Kayıtları"]
    assert list(liste["Yevmiye_Belge_No"]) == ["XYZ2024000000777", "Açılış", "MASRAF-55"]  # Yüksek önce, sonra tutar
    assert sayilar["faturasiz_ozeti"]["haric_onek"] == 1 and len(sayilar["faturasiz_haric"]) == 2
    assert not any("işaretsiz" in n for n in notes)
