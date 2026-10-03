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
