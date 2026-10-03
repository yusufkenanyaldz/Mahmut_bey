"""Madde 7: ÖTV / tevkifat ayrıştırma, beklenen maliyet, KDV ve tevkifat mutabakatı, satış faturaları."""
import io
import sqlite3

import pandas as pd
import pytest

from denetim import checks, importers
from denetim.database import DatabaseManager
from denetim.ubl import parse_ubl

FIRMA_VKN = "1111111111"
TEDARIKCI_VKN = "2222222222"

# Madde 7 öncesi (şema sürümü 2) invoices tablosu: göç testi için
ESKI_SEMA = """
CREATE TABLE settings (key TEXT PRIMARY KEY, value TEXT);
CREATE TABLE invoices (
    id INTEGER PRIMARY KEY AUTOINCREMENT, invoice_no TEXT NOT NULL, invoice_no_norm TEXT NOT NULL,
    issue_date TEXT NOT NULL, supplier_vkn TEXT NOT NULL, supplier_name TEXT, customer_vkn TEXT, customer_name TEXT,
    invoice_type TEXT, profile TEXT, currency TEXT DEFAULT 'TRY', exchange_rate REAL DEFAULT 1,
    line_extension_amount REAL, allowance_total REAL DEFAULT 0, total_amount REAL NOT NULL,
    vat_amount REAL DEFAULT 0, payable_amount REAL, source TEXT, source_file TEXT, import_id INTEGER,
    UNIQUE(supplier_vkn, invoice_no_norm));
CREATE TABLE invoice_lines (id INTEGER PRIMARY KEY AUTOINCREMENT, invoice_id INTEGER NOT NULL, line_no INTEGER,
    item_name TEXT, item_norm TEXT, quantity REAL, uom TEXT, unit_price REAL, line_net REAL, unit_price_net REAL,
    vat_rate REAL, vat_amount REAL);
CREATE TABLE journal_entries (id INTEGER PRIMARY KEY AUTOINCREMENT, entry_date TEXT, document_no TEXT,
    document_no_norm TEXT, account_code TEXT, account_norm TEXT, amount REAL, description TEXT, source_row INTEGER,
    import_id INTEGER);
INSERT INTO settings VALUES ('schema_version', '2');
INSERT INTO invoices (invoice_no, invoice_no_norm, issue_date, supplier_vkn, supplier_name, total_amount, vat_amount,
    source) VALUES ('ESK2024000000001', 'ESK2024000000001', '2024-07-01', '3333333333', 'Eski', 100, 20, 'XML');
"""


def ubl(no="ABC2024000000001", satici=TEDARIKCI_VKN, alici=FIRMA_VKN, net=1000.0, kdv_orani=20, otv=0.0,
        otv_kodu="9077", tevkifat=0.0, tevkifat_orani=None, tip="SATIS", profil="TICARIFATURA", doviz="TRY", kur=None,
        satir_otv=True, satir_tevkifat=False):
    """Test için UBL-TR faturası üretir (ÖTV KDV matrahına dahil)."""
    kdv = round((net + otv) * kdv_orani / 100, 2)
    otv_satir = (f"""<cac:TaxSubtotal><cbc:TaxableAmount>{net:.2f}</cbc:TaxableAmount><cbc:TaxAmount>{otv:.2f}</cbc:TaxAmount>
        <cbc:Percent>80</cbc:Percent><cac:TaxCategory><cac:TaxScheme><cbc:Name>ÖTV</cbc:Name>
        <cbc:TaxTypeCode>{otv_kodu}</cbc:TaxTypeCode></cac:TaxScheme></cac:TaxCategory></cac:TaxSubtotal>""" if otv else "")
    belge_otv = otv_satir if not satir_otv else ""
    wh = (f"""<cac:WithholdingTaxTotal><cbc:TaxAmount>{tevkifat:.2f}</cbc:TaxAmount><cac:TaxSubtotal>
        <cbc:TaxAmount>{tevkifat:.2f}</cbc:TaxAmount><cbc:Percent>{tevkifat_orani}</cbc:Percent><cac:TaxCategory>
        <cac:TaxScheme><cbc:TaxTypeCode>624</cbc:TaxTypeCode></cac:TaxScheme></cac:TaxCategory></cac:TaxSubtotal>
        </cac:WithholdingTaxTotal>""" if tevkifat else "")
    kur_xml = f"<cac:PricingExchangeRate><cbc:CalculationRate>{kur}</cbc:CalculationRate></cac:PricingExchangeRate>" \
        if kur else ""
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<Invoice xmlns="urn:oasis:names:specification:ubl:schema:xsd:Invoice-2"
 xmlns:cac="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"
 xmlns:cbc="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2">
 <cbc:ProfileID>{profil}</cbc:ProfileID><cbc:ID>{no}</cbc:ID><cbc:IssueDate>2024-07-10</cbc:IssueDate>
 <cbc:InvoiceTypeCode>{tip}</cbc:InvoiceTypeCode><cbc:DocumentCurrencyCode>{doviz}</cbc:DocumentCurrencyCode>{kur_xml}
 <cac:AccountingSupplierParty><cac:Party><cac:PartyIdentification><cbc:ID schemeID="VKN">{satici}</cbc:ID>
  </cac:PartyIdentification><cac:PartyName><cbc:Name>Satıcı</cbc:Name></cac:PartyName></cac:Party></cac:AccountingSupplierParty>
 <cac:AccountingCustomerParty><cac:Party><cac:PartyIdentification><cbc:ID schemeID="VKN">{alici}</cbc:ID>
  </cac:PartyIdentification><cac:PartyName><cbc:Name>Alıcı</cbc:Name></cac:PartyName></cac:Party></cac:AccountingCustomerParty>
 <cac:TaxTotal><cbc:TaxAmount>{kdv + otv:.2f}</cbc:TaxAmount>{belge_otv}
  <cac:TaxSubtotal><cbc:TaxableAmount>{net + otv:.2f}</cbc:TaxableAmount><cbc:TaxAmount>{kdv:.2f}</cbc:TaxAmount>
  <cbc:Percent>{kdv_orani}</cbc:Percent><cac:TaxCategory><cac:TaxScheme><cbc:Name>KDV</cbc:Name>
  <cbc:TaxTypeCode>0015</cbc:TaxTypeCode></cac:TaxScheme></cac:TaxCategory></cac:TaxSubtotal></cac:TaxTotal>
 {wh if not satir_tevkifat else ""}
 <cac:LegalMonetaryTotal><cbc:LineExtensionAmount>{net:.2f}</cbc:LineExtensionAmount>
  <cbc:TaxExclusiveAmount>{net:.2f}</cbc:TaxExclusiveAmount><cbc:TaxInclusiveAmount>{net + otv + kdv:.2f}</cbc:TaxInclusiveAmount>
  <cbc:PayableAmount>{net + otv + kdv - tevkifat:.2f}</cbc:PayableAmount></cac:LegalMonetaryTotal>
 <cac:InvoiceLine><cbc:ID>1</cbc:ID><cbc:InvoicedQuantity unitCode="C62">1</cbc:InvoicedQuantity>
  <cbc:LineExtensionAmount>{net:.2f}</cbc:LineExtensionAmount>
  <cac:TaxTotal><cbc:TaxAmount>{kdv + otv:.2f}</cbc:TaxAmount>{otv_satir if satir_otv else ""}
   <cac:TaxSubtotal><cbc:TaxAmount>{kdv:.2f}</cbc:TaxAmount><cbc:Percent>{kdv_orani}</cbc:Percent>
   <cac:TaxCategory><cac:TaxScheme><cbc:TaxTypeCode>0015</cbc:TaxTypeCode></cac:TaxScheme></cac:TaxCategory>
   </cac:TaxSubtotal></cac:TaxTotal>{wh if satir_tevkifat else ""}
  <cac:Item><cbc:Name>Binek Otomobil</cbc:Name></cac:Item><cac:Price><cbc:PriceAmount>{net:.2f}</cbc:PriceAmount></cac:Price>
 </cac:InvoiceLine></Invoice>""".encode("utf-8")


def yev(belge, hesap, borc=0.0, alacak=0.0, tarih="2024-07-10"):
    return {"entry_date": tarih, "document_no": belge, "account_code": hesap, "amount": borc - alacak}


def xlsx(df):
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue()


@pytest.fixture
def db(tmp_path):
    return DatabaseManager(str(tmp_path / "firma.db"))


def yukle(db, xml_listesi, yevmiye):
    db.save_invoices([parse_ubl(x)[:2] for x in xml_listesi], "XML", "test", "h1")
    db.save_journal(yevmiye, "yev.xlsx", "h2")


# ------------------------------------------------------------------ 7.1 ÖTV
def test_otv_ayristirma_satir_ve_belge_duzeyi():
    for satir in (True, False):
        h, _, _ = parse_ubl(ubl(net=1_000_000, otv=800_000, satir_otv=satir))
        assert h["vat_amount"] == pytest.approx(360_000)
        assert [(t["code"], t["amount"]) for t in h["taxes"]] == [("9077", 800_000)]
        assert h["taxes"][0]["name"].startswith("ÖTV II") and h["other_tax_amount"] == 800_000
        assert h["total_amount"] == 1_000_000 and h["tax_detail"] == 1


def test_otvsiz_faturada_vergi_yok():
    h, _, _ = parse_ubl(ubl())
    assert h["taxes"] == [] and h["other_tax_amount"] == 0 and h["withholding_amount"] == 0


def test_parse_vergi_kodlari():
    assert checks.parse_vergi_kodlari("9077, 71; 4080 9077") == ["9077", "0071", "4080"]
    assert "9077" in checks.parse_vergi_kodlari(checks.VARSAYILAN_MALIYET_VERGI_KODLARI)
    assert checks.parse_vergi_kodlari("") == []


def test_otv_maliyete_eklenir_ve_kod_listesi_firma_ayari(db):
    yukle(db, [ubl("OTO2024000000001", net=1_000_000, otv=800_000), ubl("OTO2024000000002", net=1_000_000, otv=800_000)],
          [yev("OTO2024000000001", "153.01", 1_800_000), yev("OTO2024000000001", "191.01", 360_000),
           yev("OTO2024000000001", "320.01", alacak=2_160_000),
           # ÖTV giderlere yazılmış: 153'te yalnızca matrah
           yev("OTO2024000000002", "153.01", 1_000_000), yev("OTO2024000000002", "770.01", 800_000),
           yev("OTO2024000000002", "191.01", 360_000), yev("OTO2024000000002", "320.01", alacak=2_160_000)])
    inv = checks.maliyet_vergisi_ekle(db.get_invoices_df(), db.get_invoice_taxes_df())
    res = checks.reconcile(inv, db.get_journal_df(), ["153"])
    fark = res["Tutar Farkları"]
    assert list(fark["Fatura_No"]) == ["OTO2024000000002"]
    assert fark.iloc[0]["Beklenen_Tutar_TL"] == 1_800_000 and fark.iloc[0]["Fark_TL"] == -800_000
    assert fark.iloc[0]["Olasi_Neden"] == checks.NEDEN_VERGI_EKSIK
    # Kod listesinde ÖTV yoksa beklenen maliyet matrahtır: doğru kayıt fark çıkar, hatalı kayıt çıkmaz
    inv = checks.maliyet_vergisi_ekle(db.get_invoices_df(), db.get_invoice_taxes_df(), ["4080"])
    res = checks.reconcile(inv, db.get_journal_df(), ["153"])
    assert list(res["Tutar Farkları"]["Fatura_No"]) == ["OTO2024000000001"]
    assert checks.maliyet_vergisi_ozeti(checks.maliyet_vergisi_ekle(db.get_invoices_df(),
                                                                     db.get_invoice_taxes_df())) == \
        {"fatura": 2, "toplam_tl": 1_600_000.0}


def test_otv_dovizli_fatura_kurla_cevrilir(db):
    yukle(db, [ubl("EUR2024000000001", net=10_000, otv=5_000, doviz="EUR", kur=35.0)],
          [yev("EUR2024000000001", "153.01", 525_000), yev("EUR2024000000001", "320.01", alacak=525_000)])
    inv = checks.maliyet_vergisi_ekle(db.get_invoices_df(), db.get_invoice_taxes_df())
    res = checks.reconcile(inv, db.get_journal_df(), ["153"])
    assert res["Tutar Farkları"].empty and res.kur_farki.empty


def test_excel_otv_tutari(db):
    df = pd.DataFrame([{"Fatura_No": "X1", "Tarih": "10.07.2024", "Tedarikci_VKN": TEDARIKCI_VKN,
                        "Tedarikci_Ad": "Bayi", "Urun_Adi": "Otomobil", "Miktar": 1, "Fiyat": 1000, "KDV_Orani": 20,
                        "ÖTV Tutarı": 800}])
    res = importers.read_invoice_excel(xlsx(df))
    assert not res.errors
    h, lines = res.items[0]
    assert h["taxes"][0]["code"] == importers.EXCEL_OTV_KODU and h["other_tax_amount"] == 800
    assert h["vat_amount"] == pytest.approx(360) and "otv" not in lines[0]
    db.save_invoices(res.items, "FATURA_EXCEL", "f.xlsx", "h")
    assert db.get_invoice_taxes_df()["amount"].tolist() == [800]


def test_eski_veritabani_gocu(tmp_path):
    path = str(tmp_path / "eski.db")
    conn = sqlite3.connect(path)
    conn.executescript(ESKI_SEMA)
    conn.commit()
    conn.close()
    db = DatabaseManager(path)
    inv = db.get_invoices_df()
    assert inv.loc[0, "invoice_no"] == "ESK2024000000001"
    for col in ("withholding_amount", "withholding_rate", "direction", "other_tax_amount", "tax_detail"):
        assert col in inv.columns
    assert pd.isna(inv.loc[0, "tax_detail"]) and db.get_invoice_taxes_df().empty
    assert db.get_setting("schema_version") == "3"
    yukle(db, [ubl(net=1000, otv=100, tevkifat=48, tevkifat_orani=40)], [yev("ABC2024000000001", "153.01", 1100)])
    assert len(db.get_invoice_taxes_df()) == 1
    DatabaseManager(path)  # İkinci açılışta göç tekrar çalışmaz / hata vermez
    db.clear_data()
    assert db.get_invoice_taxes_df().empty


def test_eski_xml_faturalar_icin_not(tmp_path):
    path = str(tmp_path / "eski.db")
    conn = sqlite3.connect(path)
    conn.executescript(ESKI_SEMA)
    conn.close()
    db = DatabaseManager(path)
    _, _, notes, _ = checks.run_full_audit(db, "Aylık", 15, [], 0.01, FIRMA_VKN)
    assert any("önceki bir sürümle" in n for n in notes)


# ------------------------------------------------------------------ 7.3 tevkifat ayrıştırma
@pytest.mark.parametrize("satir", [False, True])
def test_tevkifat_ayristirma(satir):
    h, _, _ = parse_ubl(ubl(tip="TEVKIFAT", net=10_000, tevkifat=800, tevkifat_orani=40, satir_tevkifat=satir))
    assert h["withholding_amount"] == 800 and h["withholding_rate"] == 40 and h["withholding_code"] == "624"
    assert h["vat_amount"] == 2000 and h["payable_amount"] == 11_200


@pytest.mark.parametrize("raw,expected", [("4/10", 40), ("%50", 50), ("0,7", 70), (90, 90), ("9/10", 90)])
def test_tevkifat_orani_excel(raw, expected):
    assert importers.parse_tevkifat_orani(raw) == pytest.approx(expected)


def test_excel_tevkifat_ve_fatura_tipi():
    df = pd.DataFrame([
        {"Fatura_No": "T1", "Tarih": "10.07.2024", "Tedarikci_VKN": TEDARIKCI_VKN, "Tedarikci_Ad": "Taşeron",
         "Urun_Adi": "Hakediş", "Miktar": 1, "Fiyat": 10_000, "KDV_Orani": 20, "Tevkifat Oranı": "4/10"},
        {"Fatura_No": "T2", "Tarih": "10.07.2024", "Tedarikci_VKN": TEDARIKCI_VKN, "Tedarikci_Ad": "X",
         "Urun_Adi": "Mal", "Miktar": 1, "Fiyat": 100, "KDV_Orani": 20, "Fatura Tipi": "İade"},
        {"Fatura_No": "T3", "Tarih": "10.07.2024", "Tedarikci_VKN": TEDARIKCI_VKN, "Tedarikci_Ad": "X",
         "Urun_Adi": "Mal", "Miktar": 1, "Fiyat": 100, "KDV_Orani": 20, "Fatura Tipi": "Bilinmeyen"}])
    res = importers.read_invoice_excel(xlsx(df))
    h = {x["invoice_no"]: x for x, _ in res.items}
    assert h["T1"]["invoice_type"] == "TEVKIFAT" and h["T1"]["withholding_amount"] == 800
    assert h["T1"]["withholding_rate"] == 40 and h["T1"]["payable_amount"] == pytest.approx(11_200)
    assert h["T2"]["invoice_type"] == "IADE" and h["T2"]["withholding_amount"] == 0
    assert h["T3"]["invoice_type"] == "SATIS" and any("Fatura_Tipi" in w for w in res.warnings)


# ------------------------------------------------------------------ 7.2 KDV / 7.3 tevkifat mutabakatı
def _alis_senaryosu(db):
    x = [ubl("AAA2024000000001", net=1000),                       # doğru
         ubl("AAA2024000000002", net=1000),                       # KDV yanlış tutar (191 = 100)
         ubl("AAA2024000000003", net=1000),                       # KDV maliyete eklenmiş
         ubl("AAA2024000000004", net=1000),                       # KDV hiç kaydedilmemiş
         ubl("AAA2024000000005", tip="TEVKIFAT", net=10_000, tevkifat=800, tevkifat_orani=40),  # doğru
         ubl("AAA2024000000006", tip="TEVKIFAT", net=10_000, tevkifat=800, tevkifat_orani=40),  # 360 yok
         ubl("AAA2024000000007", tip="TEVKIFAT", net=10_000, tevkifat=800, tevkifat_orani=40),  # 191'e net KDV
         ubl("AAA2024000000008", tip="TEVKIFAT", net=10_000, tevkifat=800, tevkifat_orani=40),  # 360 borç
         ubl("AAA2024000000009", net=1000),                       # yanlış hesap (770) ama 191 doğru
         ubl("AAA2024000000010", net=1000)]                       # muhasebeleşmemiş: KDV kontrolüne girmez
    j = [yev("AAA2024000000001", "153.01", 1000), yev("AAA2024000000001", "191.01", 200),
         yev("AAA2024000000001", "320.01", alacak=1200),
         yev("AAA2024000000002", "153.01", 1000), yev("AAA2024000000002", "191.01", 100),
         yev("AAA2024000000002", "320.01", alacak=1100),
         yev("AAA2024000000003", "153.01", 1200), yev("AAA2024000000003", "320.01", alacak=1200),
         yev("AAA2024000000004", "153.01", 1000), yev("AAA2024000000004", "320.01", alacak=1000),
         yev("AAA2024000000005", "153.01", 10_000), yev("AAA2024000000005", "191.01", 2000),
         yev("AAA2024000000005", "360.02", alacak=800), yev("AAA2024000000005", "320.01", alacak=11_200),
         yev("AAA2024000000006", "153.01", 10_000), yev("AAA2024000000006", "191.01", 2000),
         yev("AAA2024000000006", "320.01", alacak=12_000),
         yev("AAA2024000000007", "153.01", 10_000), yev("AAA2024000000007", "191.01", 1200),
         yev("AAA2024000000007", "360.02", alacak=800), yev("AAA2024000000007", "320.01", alacak=10_400),
         yev("AAA2024000000008", "153.01", 10_000), yev("AAA2024000000008", "191.01", 2000),
         yev("AAA2024000000008", "360.02", 800), yev("AAA2024000000008", "320.01", alacak=12_800),
         yev("AAA2024000000009", "770.01", 1000), yev("AAA2024000000009", "191.01", 200),
         yev("AAA2024000000009", "320.01", alacak=1200)]
    yukle(db, x, j)
    return checks.reconcile(db.get_invoices_df(), db.get_journal_df(), ["153"], kdv_hesaplari=["191"],
                            tevkifat_hesaplari=["360"])


def test_kdv_mutabakati(db):
    res = _alis_senaryosu(db)
    fark = res["KDV Farkları"].set_index("Fatura_No")
    assert sorted(fark.index) == ["AAA2024000000002", "AAA2024000000007"]
    assert fark.loc["AAA2024000000002", "Fark_TL"] == -100
    assert fark.loc["AAA2024000000007", "Olasi_Neden"] == checks.NEDEN_KDV_TEVKIFAT_DUSULMUS
    yok = res["KDV'si Kaydedilmemiş Faturalar"].set_index("Fatura_No")
    assert sorted(yok.index) == ["AAA2024000000003", "AAA2024000000004"]
    assert yok.loc["AAA2024000000003", "Olasi_Neden"] == checks.NEDEN_KDV_MALIYETTE
    assert yok.loc["AAA2024000000004", "Olasi_Neden"] == checks.NEDEN_KDV_YOK
    # Tevkifatlı faturalarda 191'e tam KDV yazılması doğrudur (AAA...05 bulgu değil)
    assert "AAA2024000000005" not in fark.index
    assert res.vergi_ozeti["kdv_karsilastirilan"] == 9  # muhasebeleşmemiş fatura karşılaştırılmaz
    assert list(res["Muhasebeleşmemiş Faturalar"]["Fatura_No"]) == ["AAA2024000000010"]


def test_tevkifat_mutabakati(db):
    tev = _alis_senaryosu(db)["Tevkifat Kaydı Eksik/Farklı"].set_index("Fatura_No")
    assert sorted(tev.index) == ["AAA2024000000006", "AAA2024000000008"]
    assert tev.loc["AAA2024000000006", "Durum"] == checks.TEVKIFAT_YOK
    assert tev.loc["AAA2024000000008", "Durum"] == checks.TEVKIFAT_TERS
    assert tev.loc["AAA2024000000006", "Tevkifat_Orani"] == 40


def test_vergi_kontrolleri_kapali_ve_iade(db):
    yukle(db, [ubl("IAD2024000000001", tip="IADE", net=1000)],
          [yev("IAD2024000000001", "320.01", 1200), yev("IAD2024000000001", "153.01", alacak=1000),
           yev("IAD2024000000001", "191.01", alacak=200)])
    res = checks.reconcile(db.get_invoices_df(), db.get_journal_df(), ["153"])
    assert "KDV Farkları" not in res and "Tevkifat Kaydı Eksik/Farklı" not in res
    res = checks.reconcile(db.get_invoices_df(), db.get_journal_df(), ["153"], kdv_hesaplari=["191"])
    assert res["KDV Farkları"].empty and res["KDV'si Kaydedilmemiş Faturalar"].empty


def test_genel_rapor_vergi_bolumleri_ve_inceleme(db):
    from denetim import inceleme
    _alis_senaryosu(db)
    _, sections, _, counts = checks.run_full_audit(db, "Aylık", 15, ["153"], 0.01, FIRMA_VKN)
    for baslik in ("KDV Farkları", "KDV'si Kaydedilmemiş Faturalar", "Tevkifat Kaydı Eksik/Farklı"):
        assert baslik in sections and inceleme.kontrol_kodu(baslik) is not None
    assert counts["vergi_ozeti"]["tevkifat_bulgu"] == 2
    assert "Vergi mutabakatı" in checks.vergi_ozeti_metni(counts["vergi_ozeti"])
    anahtar = inceleme.bolum_anahtarlari("KDV Farkları", sections["KDV Farkları"])[0]
    assert anahtar.startswith("KDV_FARKI|" + TEDARIKCI_VKN)
    _, sections, _, _ = checks.run_full_audit(db, "Aylık", 15, ["153"], 0.01, FIRMA_VKN,
                                              vergi_ayarlari=checks.VergiAyarlari(kdv_hesaplari=[],
                                                                                  tevkifat_hesaplari=[]))
    assert "KDV Farkları" not in sections


# ------------------------------------------------------------------ 7.4 satış faturaları
MUSTERI_VKN = "4444444444"


def test_satis_yonu_tespiti():
    inv = pd.DataFrame({"supplier_vkn": [FIRMA_VKN, TEDARIKCI_VKN, FIRMA_VKN, TEDARIKCI_VKN],
                        "invoice_type": ["SATIS", "SATIS", "IADE", "SATIS"],
                        "direction": [None, None, None, "SATIS"]})
    assert list(checks.fatura_yonleri(inv, FIRMA_VKN)) == ["SATIS", "ALIS", "ALIS", "SATIS"]
    # Firma VKN'si yoksa yalnızca açıkça verilen yön satış sayılır
    assert list(checks.fatura_yonleri(inv, "")) == ["ALIS", "ALIS", "ALIS", "SATIS"]
    assert len(checks.satis_faturalari(inv, FIRMA_VKN)) == 2


def test_kdv_istisna():
    assert checks.kdv_istisna("IHRACAT", "ISTISNA") and checks.kdv_istisna("TEMELFATURA", "ISTISNA")
    assert checks.kdv_istisna(None, "IHRACKAYITLI") and not checks.kdv_istisna("TICARIFATURA", "SATIS")


def test_excel_satis_yonu(db):
    df = pd.DataFrame([
        {"Fatura_No": "S1", "Tarih": "10.07.2024", "Tedarikci_VKN": MUSTERI_VKN, "Tedarikci_Ad": "Müşteri",
         "Urun_Adi": "Halı", "Miktar": 1, "Fiyat": 1000, "KDV_Orani": 20, "Yön": "Satış"},
        {"Fatura_No": "S2", "Tarih": "10.07.2024", "Tedarikci_VKN": MUSTERI_VKN, "Tedarikci_Ad": "Yabancı",
         "Urun_Adi": "Halı", "Miktar": 1, "Fiyat": 1000, "KDV_Orani": 0, "Yön": "GİDEN", "Fatura Tipi": "İhracat"},
        {"Fatura_No": "A1", "Tarih": "10.07.2024", "Tedarikci_VKN": TEDARIKCI_VKN, "Tedarikci_Ad": "Tedarikçi",
         "Urun_Adi": "İplik", "Miktar": 1, "Fiyat": 1000, "KDV_Orani": 20, "Yön": "Alış"},
        {"Fatura_No": "X1", "Tarih": "10.07.2024", "Tedarikci_VKN": TEDARIKCI_VKN, "Tedarikci_Ad": "?",
         "Urun_Adi": "İplik", "Miktar": 1, "Fiyat": 1000, "Yön": "Belirsiz"}])
    res = importers.read_invoice_excel(xlsx(df), company_vkn=FIRMA_VKN)
    h = {x["invoice_no"]: x for x, _ in res.items}
    assert set(h) == {"S1", "S2", "A1"} and any("Yon" in e for e in res.errors)
    assert h["S1"]["direction"] == "SATIS" and h["S1"]["supplier_vkn"] == FIRMA_VKN
    assert h["S1"]["customer_vkn"] == MUSTERI_VKN and h["S1"]["customer_name"] == "Müşteri"
    assert h["S2"]["invoice_type"] == "ISTISNA" and h["S2"]["profile"] == "IHRACAT"
    assert h["A1"]["direction"] == "ALIS" and h["A1"]["supplier_vkn"] == TEDARIKCI_VKN
    db.save_invoices(res.items, "FATURA_EXCEL", "f.xlsx", "h")
    yon = checks.fatura_yonleri(db.get_invoices_df(), "")  # açık yön firma VKN'si olmadan da geçerli
    assert sorted(yon) == ["ALIS", "SATIS", "SATIS"]


def _satis_senaryosu(db):
    s = lambda no, **kw: ubl(no, satici=FIRMA_VKN, alici=MUSTERI_VKN, **kw)  # noqa: E731
    x = [s("SAT2024000000001", net=1000),                                      # doğru
         s("SAT2024000000002", net=1000),                                      # muhasebeleşmemiş
         s("SAT2024000000003", net=1000),                                      # tutar farkı (600 = 900)
         s("SAT2024000000004", net=1000),                                      # KDV farkı (391 = 100)
         s("SAT2024000000005", net=1000),                                      # 391 yok
         s("SAT2024000000006", net=5000, kdv_orani=0, tip="ISTISNA", profil="IHRACAT", doviz="EUR", kur=2),  # ihracat
         s("SAT2024000000007", net=5000, kdv_orani=0, tip="ISTISNA", profil="IHRACAT"),  # ihracata 391 yazılmış
         s("SAT2024000000008", net=10_000, tip="TEVKIFAT", tevkifat=1000, tevkifat_orani=50),  # tevkifatlı satış
         ubl("ALI2024000000001", net=500)]                                   # alış faturası
    j = []
    for no, gelir, kdv in (("SAT2024000000001", 1000, 200), ("SAT2024000000003", 900, 200),
                           ("SAT2024000000004", 1000, 100), ("SAT2024000000005", 1000, 0),
                           ("SAT2024000000006", 10_000, 0), ("SAT2024000000007", 5000, 1000),
                           ("SAT2024000000008", 10_000, 1000)):
        j += [yev(no, "120.01", gelir + kdv), yev(no, "600.01" if "06" not in no else "601.01", alacak=gelir)]
        if kdv:
            j.append(yev(no, "391.01", alacak=kdv))
    j += [yev("SAT2024000000999", "120.01", 3000), yev("SAT2024000000999", "600.01", alacak=3000),  # faturasız gelir
          yev("MAHSUP-07", "600.01", alacak=50), yev("MAHSUP-07", "100.01", 50),                    # hariç önek
          yev("DUZ2024000000001", "600.01", 400), yev("DUZ2024000000001", "120.01", alacak=400),    # borç yönlü
          yev("ALI2024000000001", "153.01", 500), yev("ALI2024000000001", "191.01", 100),
          yev("ALI2024000000001", "320.01", alacak=600)]
    yukle(db, x, j)


def test_satis_mutabakati(db):
    _satis_senaryosu(db)
    inv = db.get_invoices_df()
    satis = checks.satis_faturalari(inv, FIRMA_VKN)
    assert len(satis) == 8
    res = checks.satis_mutabakati(satis, db.get_journal_df(), ["600", "601", "602"], ["391"])
    assert list(res["Muhasebeleşmemiş Satış Faturaları"]["Fatura_No"]) == ["SAT2024000000002"]
    fark = res["Satış Tutar Farkları"]
    assert list(fark["Fatura_No"]) == ["SAT2024000000003"] and "Musteri_VKN" in fark
    assert fark.iloc[0]["Musteri_VKN"] == MUSTERI_VKN and "Beklenen_Tutar_TL" not in fark
    kdv = res["Satış KDV Farkları"].set_index("Fatura_No")
    assert sorted(kdv.index) == ["SAT2024000000004", "SAT2024000000005", "SAT2024000000007"]
    assert kdv.loc["SAT2024000000005", "Durum"] == checks.SATIS_KDV_YOK
    assert kdv.loc["SAT2024000000007", "Durum"] == checks.SATIS_KDV_ISTISNA
    assert kdv.loc["SAT2024000000004", "Durum"] == checks.SATIS_KDV_FARKLI
    # İhracat faturasında KDV beklenmez; tevkifatlı satışta 391 = KDV − tevkif edilen
    assert res.vergi_ozeti["istisna"] == 2
    assert list(res["Faturası Bulunmayan Gelir Kayıtları"]["Yevmiye_Belge_No"]) == ["SAT2024000000999"]
    assert res.faturasiz_ozeti["borc_yonlu"] == 1 and res.faturasiz_ozeti["haric_onek"] == 1
    assert "Faturasız gelir kaydı" in checks.faturasiz_ozeti_metni(res.faturasiz_ozeti)
    assert res["Satış Dönem Farkları"].empty


def test_genel_rapor_satis_ve_alis_ayrimi(db):
    from denetim import inceleme
    _satis_senaryosu(db)
    _, sections, notes, counts = checks.run_full_audit(db, "Aylık", 15, ["153"], 0.01, FIRMA_VKN)
    assert counts["alis_fatura"] == 1 and counts["satis_fatura"] == 8
    # Satışlar alış mutabakatına, fiyat analizine, mükerrer ve alıcı VKN kontrolüne girmez
    assert sections["Muhasebeleşmemiş Faturalar"].empty
    assert sections["Alıcısı Firma Olmayan Faturalar"].empty
    assert counts["fiyat_ozeti"]["toplam"] == 1
    assert list(sections["Muhasebeleşmemiş Satış Faturaları"]["Fatura_No"]) == ["SAT2024000000002"]
    assert counts["satis"]["vergi_ozeti"]["kdv_farki"] == 3
    for baslik in checks.SATIS_BASLIKLARI.values():
        assert baslik in sections, baslik
        assert inceleme.kontrol_kodu(baslik) is not None, baslik
    key = inceleme.bolum_anahtarlari("Satış Tutar Farkları", sections["Satış Tutar Farkları"])[0]
    assert key == f"SATIS_TUTAR_FARKI|{MUSTERI_VKN}|SAT2024000000003"
    view = inceleme.bolumlere_uygula(sections, {key: {"durum": inceleme.DURUM_SORUN_YOK}})
    assert view["Satış Tutar Farkları"][inceleme.DURUM_COL].tolist() == [inceleme.DURUM_SORUN_YOK]
    # Gelir hesabı boşsa satış mutabakatı atlanır
    _, sections, notes, counts = checks.run_full_audit(db, "Aylık", 15, ["153"], 0.01, FIRMA_VKN,
                                                       vergi_ayarlari=checks.VergiAyarlari(gelir_hesaplari=[]))
    assert counts["satis"] is None and "Satış Tutar Farkları" not in sections
    assert any("satış mutabakatı atlandı" in n for n in notes)


def test_firma_kestigi_iade_alis_sayilir_alici_kontrolu_yok(db):
    yukle(db, [ubl("IAD2024000000001", satici=FIRMA_VKN, alici=TEDARIKCI_VKN, tip="IADE", net=100)],
          [yev("IAD2024000000001", "320.01", 120), yev("IAD2024000000001", "153.01", alacak=100),
           yev("IAD2024000000001", "191.01", alacak=20)])
    _, sections, _, counts = checks.run_full_audit(db, "Aylık", 15, ["153"], 0.01, FIRMA_VKN)
    assert counts["alis_fatura"] == 1 and counts["satis"] is None
    assert sections["Alıcısı Firma Olmayan Faturalar"].empty and sections["Tutar Farkları"].empty
    assert sections["KDV Farkları"].empty


def test_kdv_yonu_yalnizca_xml_faturada(db):
    # Excel faturada tip bilinmediğinden iade faturası SATIS görünür: alacak yönlü 191 ters yön sayılmaz
    df = pd.DataFrame([{"Fatura_No": "EXC2024000000001", "Tarih": "10.07.2024", "Tedarikci_VKN": TEDARIKCI_VKN,
                        "Tedarikci_Ad": "X", "Urun_Adi": "Mal", "Miktar": 1, "Fiyat": 1000, "KDV_Orani": 20}])
    db.save_invoices(importers.read_invoice_excel(xlsx(df)).items, "FATURA_EXCEL", "f.xlsx", "h0")
    yukle(db, [ubl("XML2024000000001", net=1000)],
          [yev("EXC2024000000001", "320.01", 1200), yev("EXC2024000000001", "153.01", alacak=1000),
           yev("EXC2024000000001", "191.01", alacak=200),
           yev("XML2024000000001", "320.01", 1200), yev("XML2024000000001", "153.01", alacak=1000),
           yev("XML2024000000001", "191.01", alacak=200)])
    res = checks.reconcile(db.get_invoices_df(), db.get_journal_df(), ["153"], kdv_hesaplari=["191"])
    assert list(res["KDV Farkları"]["Fatura_No"]) == ["XML2024000000001"]
    assert res["KDV Farkları"].iloc[0]["Olasi_Neden"] == checks.NEDEN_TERS_YON
