import io
import os
import sqlite3

import pandas as pd
import pytest

from denetim import checks, importers
from denetim.database import DatabaseManager
from denetim.ubl import UBLParseError, parse_ubl
from denetim.utils import parse_account_list, parse_date, parse_number, period_of

DATA = os.path.join(os.path.dirname(__file__), "data")


def xlsx(df):
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue()


@pytest.fixture
def db(tmp_path):
    return DatabaseManager(str(tmp_path / "test.db"))


# ------------------------------------------------------------------ utils
@pytest.mark.parametrize("raw,expected", [
    ("1.234,56", 1234.56), ("1,234.56", 1234.56), ("1234,5", 1234.5), (12, 12.0), ("100 TL", 100.0)])
def test_parse_number(raw, expected):
    assert parse_number(raw) == pytest.approx(expected)


def test_parse_number_invalid():
    with pytest.raises(ValueError):
        parse_number("abc")


def test_parse_date():
    assert parse_date("05.03.2024") == "2024-03-05"
    assert parse_date(pd.Timestamp("2024-03-05 00:00:00")) == "2024-03-05"
    assert parse_date("2024-03-05 00:00:00") == "2024-03-05"


def test_period_and_accounts():
    assert period_of("2024-05-10", "Aylık") == "2024-05"
    assert period_of("2024-05-10", "Çeyreklik") == "2024-Ç2"
    assert period_of("2024-05-10", "Yıllık") == "2024"
    assert parse_account_list("153, 770.01; 760") == ["153", "77001", "760"]


# ------------------------------------------------------------------ UBL
def test_parse_ubl():
    with open(os.path.join(DATA, "ornek_fatura.xml"), "rb") as fh:
        header, lines, warnings = parse_ubl(fh.read())
    assert header["invoice_no"] == "ABC2024000000123"
    assert header["supplier_vkn"] == "1234567890"
    assert header["customer_vkn"] == "9876543210"
    assert header["total_amount"] == 1900.0
    assert header["vat_amount"] == 380.0
    assert len(lines) == 2
    assert lines[0]["line_net"] == 900.0 and lines[0]["uom"] == "C62"
    assert warnings == []


def test_parse_ubl_invalid():
    with pytest.raises(UBLParseError):
        parse_ubl(b"<foo/>")


# ------------------------------------------------------------------ Excel fatura
def test_invoice_excel_multiline_and_errors():
    df = pd.DataFrame([
        {"Fatura_No": "F1", "Tarih": "01.01.2024", "Tedarikçi VKN": 1234567890, "Tedarikci_Ad": "A",
         "Urun_Adi": "Vida", "Miktar": 10, "Fiyat": "1,5"},
        {"Fatura_No": "F1", "Tarih": "01.01.2024", "Tedarikçi VKN": 1234567890, "Tedarikci_Ad": "A",
         "Urun_Adi": "Somun", "Miktar": 2, "Fiyat": 3},
        {"Fatura_No": "F2", "Tarih": "02.01.2024", "Tedarikçi VKN": 1234567890, "Tedarikci_Ad": "A",
         "Urun_Adi": "Vida", "Miktar": "abc", "Fiyat": 3},
        {"Fatura_No": "F2", "Tarih": "02.01.2024", "Tedarikçi VKN": 1234567890, "Tedarikci_Ad": "A",
         "Urun_Adi": "Pul", "Miktar": 1, "Fiyat": 3},
    ])
    res = importers.read_invoice_excel(xlsx(df))
    assert len(res.items) == 1
    header, lines = res.items[0]
    assert header["invoice_no"] == "F1" and len(lines) == 2
    assert header["total_amount"] == pytest.approx(21.0)
    assert any("Satır 4" in e and "Miktar" in e for e in res.errors)
    assert any("F2" in w for w in res.warnings)


def test_invoice_excel_missing_column():
    res = importers.read_invoice_excel(xlsx(pd.DataFrame([{"Fatura_No": "F1"}])))
    assert res.items == [] and "Tedarikci_VKN" in res.errors[0]


def test_save_invoices_duplicate(db):
    header = {"invoice_no": "F1", "issue_date": "2024-01-01", "supplier_vkn": "1234567890", "total_amount": 10}
    line = {"item_name": "X", "quantity": 1, "line_net": 10}
    assert db.save_invoices([(header, [line])], "XML", "a.xml", "h1") == (1, [])
    added, dups = db.save_invoices([(dict(header, invoice_no="f1 "), [line])], "XML", "b.xml", "h2")
    assert added == 0 and len(dups) == 1
    assert db.find_import("XML", "h1")[0] == "a.xml"


# ------------------------------------------------------------------ Yevmiye
def test_journal_debit_credit():
    df = importers.journal_template()
    df.loc[3] = ["bozuk", "X", "153", 1, 0, ""]
    res = importers.read_journal_excel(xlsx(df))
    assert [r["amount"] for r in res.items] == [1000.0, 200.0, -1200.0]
    assert len(res.errors) == 1 and "Satır 5" in res.errors[0]


# ------------------------------------------------------------------ Kontroller
def _load_scenario(db):
    invs = []
    for no, date, vkn, qty, price in [
        ("F1", "2024-01-10", "1111111111", 10, 100),   # yevmiye doğru
        ("F2", "2024-01-20", "1111111111", 10, 130),   # fiyat anomali + tutar farkı
        ("F3", "2024-01-25", "2222222222", 10, 100),   # muhasebeleşmemiş
        ("F4", "2024-01-31", "2222222222", 10, 100),   # dönem farkı
        ("F5", "2024-01-31", "2222222222", 1, 50),     # 760 hesabına kaydedilmiş + fiyat anomali
        ("F6", "2024-01-31", "2222222222", 10, 100),   # F4 ile olası mükerrer
    ]:
        invs.append(({"invoice_no": no, "issue_date": date, "supplier_vkn": vkn, "supplier_name": vkn,
                      "total_amount": qty * price, "source": "EXCEL"},
                     [{"item_name": "Vida", "quantity": qty, "line_net": qty * price}]))
    db.save_invoices(invs, "FATURA_EXCEL", "f.xlsx", "x")
    db.save_journal([
        {"entry_date": "2024-01-10", "document_no": "F1", "account_code": "153.01", "amount": 1000},
        {"entry_date": "2024-01-10", "document_no": "F1", "account_code": "320.01", "amount": -1200},
        {"entry_date": "2024-01-20", "document_no": "F2", "account_code": "153.01", "amount": 1200},
        {"entry_date": "2024-02-01", "document_no": "F4", "account_code": "153.01", "amount": 1000},
        {"entry_date": "2024-01-31", "document_no": "F5", "account_code": "760.01", "amount": 50},
        {"entry_date": "2024-01-31", "document_no": "F6", "account_code": "153.01", "amount": 1000},
        {"entry_date": "2024-01-31", "document_no": "F99", "account_code": "153.01", "amount": 500},
    ], "y.xlsx", "y")


def test_reconcile(db):
    _load_scenario(db)
    res = checks.reconcile(db.get_invoices_df(), db.get_journal_df(), ["153"])
    assert list(res["Muhasebeleşmemiş Faturalar"]["Fatura_No"]) == ["F3"]
    assert list(res["Seçili Hesap Dışına Kaydedilmiş Faturalar"]["Fatura_No"]) == ["F5"]
    diff = res["Tutar Farkları"]
    assert list(diff["Fatura_No"]) == ["F2"] and diff["Fark_TL"].iloc[0] == pytest.approx(-100)
    assert list(res["Dönem Farkları"]["Fatura_No"]) == ["F4"]
    assert list(res["Faturası Bulunmayan Yevmiye Kayıtları"]["Yevmiye_Belge_No"]) == ["F99"]


# ------------------------------------------------------------------ Belge no eşleştirme
@pytest.mark.parametrize("raw,expected", [
    ("ABC2024000000123", [("ABC", 2024, 123)]),
    ("abc 2024 000000123", [("ABC", 2024, 123)]),
    ("ABC-2024-123", [("ABC", 2024, 123)]),
    ("ABC 2024 123", [("ABC", 2024, 123)]),
    ("ABC123", [("ABC", None, 123)]),
    ("ABC-123", [("ABC", None, 123)]),
    ("ABC2024123", [("ABC", 2024, 123), ("ABC", None, 2024123)]),
    ("BORDRO-2024-07", []), ("GP-07001", []), ("F1", []), ("", []), (None, []),
])
def test_belge_anahtarlari(raw, expected):
    assert checks.belge_anahtarlari(raw) == expected


def test_anahtar_uyumlu():
    assert checks.anahtar_uyumlu(("ABC", 2024, 123), ("ABC", None, 123))
    assert checks.anahtar_uyumlu(("ABC", 2024, 123), ("ABC", 2024, 123))
    assert not checks.anahtar_uyumlu(("ABC", 2024, 123), ("ABC", 2023, 123))
    assert not checks.anahtar_uyumlu(("ABC", 2024, 123), ("ABD", 2024, 123))


@pytest.mark.parametrize("belge", ["ABC123", "ABC-123", "ABC 2024 123", "ABC2024123", "abc-2024-000000123"])
def test_seri_sira_tek_aday(belge):
    eslesen, belirsiz = checks.seri_sira_eslestir({1: "ABC2024000000123", 2: "XYZ2024000000123"},
                                                  {"B": belge, "C": "BORDRO-2024-07"})
    assert eslesen == {1: "B"} and belirsiz == {}


def test_seri_sira_ayni_seri_farkli_sira_ve_farkli_yil():
    eslesen, belirsiz = checks.seri_sira_eslestir({1: "ABC2024000000123"},
                                                  {"B": "ABC124", "C": "ABC-2023-123", "D": "ABD123"})
    assert eslesen == {} and belirsiz == {}


def test_seri_sira_belirsiz():
    # Fatura için iki aday belge
    eslesen, belirsiz = checks.seri_sira_eslestir({1: "ABC2024000000123"}, {"B": "ABC123", "C": "ABC-123 "})
    assert eslesen == {} and belirsiz == {1: ["B", "C"]}
    # Yılsız belge iki farklı yılın faturasına uyuyor
    eslesen, belirsiz = checks.seri_sira_eslestir({1: "ABC2023000000123", 2: "ABC2024000000123"}, {"B": "ABC123"})
    assert eslesen == {} and belirsiz == {1: ["B"], 2: ["B"]}


def test_tutar_tarih_eslestir():
    faturalar = {1: ("2024-07-10", 1000.0), 2: ("2024-07-10", 500.0), 3: ("2024-07-10", 750.0)}
    belgeler = {"A": ("2024-07-20", -1000.004),   # tutar (mutlak) ve tarih tutuyor → tek aday
                "B": ("2024-08-01", 500.0),      # 22 gün → pencere dışı
                "C": ("2024-07-12", 750.0), "D": ("2024-07-14", 750.0)}  # iki aday → belirsiz
    eslesen, belirsiz = checks.tutar_tarih_eslestir(faturalar, belgeler, tolerance=0.01)
    assert eslesen == {1: "A"} and belirsiz == {3: ["C", "D"]}
    assert checks.tutar_tarih_eslestir({1: ("2024-07-10", 1000.0)}, {"A": ("2024-07-10", 1000.5)}, 0.01) == ({}, {})


def _gib_scenario():
    def fatura(no, tarih, vkn, tutar):
        return {"invoice_no": no, "invoice_no_norm": no.upper(), "issue_date": tarih, "supplier_vkn": vkn,
                "supplier_name": vkn, "total_amount": tutar, "exchange_rate": 1.0, "invoice_type": "SATIS"}

    def kayit(tarih, belge, hesap, tutar):
        return {"entry_date": tarih, "document_no": belge, "document_no_norm": belge.replace(" ", "").upper(),
                "account_code": hesap, "account_norm": hesap.replace(".", ""), "amount": tutar}

    invoices = pd.DataFrame([
        fatura("ABC2024000000001", "2024-07-01", "111", 1000),   # tam
        fatura("ABC2024000000002", "2024-07-02", "111", 2000),   # seri+sıra (ABC2)
        fatura("ABC2024000000003", "2024-07-03", "111", 3000),   # seri+sıra (ABC-2024-3) + tutar farkı
        fatura("XYZ2024000000010", "2024-07-04", "222", 4000),   # tutar+tarih (belge no tamamen farklı)
        fatura("XYZ2024000000011", "2024-07-05", "222", 5000),   # seri+sıra ile başka hesapta
        fatura("XYZ2024000000012", "2024-07-06", "222", 6000),   # muhasebeleşmemiş
        fatura("KLM2024000000005", "2024-07-07", "333", 7000),   # iki aday → belirsiz
    ])
    journal = pd.DataFrame([
        kayit("2024-07-01", "ABC2024000000001", "153.01", 1000),
        kayit("2024-07-02", "ABC2", "153.01", 2000),
        kayit("2024-07-03", "ABC-2024-3", "153.01", 3100),
        kayit("2024-07-10", "FIS-778", "153.01", 4000),
        kayit("2024-07-05", "XYZ11", "689.01", 5000),
        kayit("2024-07-07", "KLM5", "153.01", 7000),
        kayit("2024-07-07", "KLM 2024 5", "153.01", 7000),
        kayit("2024-07-31", "BORDRO-2024-07", "153.01", 9000),
    ])
    return invoices, journal


def test_reconcile_kademeli_eslestirme():
    invoices, journal = _gib_scenario()
    res = checks.reconcile(invoices, journal, ["153"])
    assert list(res["Muhasebeleşmemiş Faturalar"]["Fatura_No"]) == ["XYZ2024000000012"]
    other = res["Seçili Hesap Dışına Kaydedilmiş Faturalar"]
    assert list(other["Fatura_No"]) == ["XYZ2024000000011"]
    assert other["Yevmiye_Belge_No"].iloc[0] == "XYZ11" and other["Eslesme_Yontemi"].iloc[0] == "Seri+Sıra"
    assert other["Kullanilan_Hesaplar"].iloc[0] == "689.01"
    diff = res["Tutar Farkları"]
    assert list(diff["Fatura_No"]) == ["ABC2024000000003"] and diff["Eslesme_Yontemi"].iloc[0] == "Seri+Sıra"
    weak = res["Belge No Uyuşmayan Eşleşmeler (Kontrol Edin)"]
    assert list(weak["Fatura_No"]) == ["XYZ2024000000010"] and weak["Yevmiye_Belge_No"].iloc[0] == "FIS-778"
    assert list(weak["Eslesme_Yontemi"]) == ["Tutar+Tarih"]
    multi = res["Belirsiz Eşleşme (Birden Fazla Seri+Sıra Adayı)"]
    assert list(multi["Fatura_No"]) == ["KLM2024000000005"] and multi["Aday_Belgeler"].iloc[0] == "KLM 2024 5, KLM5"
    # Herhangi bir yöntemle eşleşen ya da belirsiz adayı olan belge faturasız sayılmaz
    assert list(res["Faturası Bulunmayan Yevmiye Kayıtları"]["Yevmiye_Belge_No"]) == ["BORDRO-2024-07"]
    assert dict(res.eslesme_ozeti) == {"Tam": 1, "Seri+Sıra": 2, "Tutar+Tarih": 1, "Seçili hesap dışı": 1,
                                       "Belirsiz": 1, "Eşleşmeyen": 1}
    assert "Seri+Sıra: 2" in checks.eslesme_ozeti_metni(res.eslesme_ozeti)
    assert list(checks.eslesme_ozeti_df(res.eslesme_ozeti)["Fatura_Sayisi"]) == [1, 2, 1, 1, 1, 1]


def test_reconcile_tutar_tarih_yanlis_hesabi_gizlemez():
    """Belge no'su başka hesapta olan fatura, seçili hesapta aynı tutarlı kayıt varsa da yanlış hesaptadır."""
    invoices, journal = _gib_scenario()
    extra = journal.iloc[[0]].assign(document_no="ZZZ-1", document_no_norm="ZZZ-1", entry_date="2024-07-06",
                                     amount=5000)
    res = checks.reconcile(invoices, pd.concat([journal, extra], ignore_index=True), ["153"])
    assert "XYZ2024000000011" in list(res["Seçili Hesap Dışına Kaydedilmiş Faturalar"]["Fatura_No"])
    assert "ZZZ-1" in list(res["Faturası Bulunmayan Yevmiye Kayıtları"]["Yevmiye_Belge_No"])


def test_reconcile_tutar_tarih_belirsiz_aday():
    invoices, journal = _gib_scenario()
    extra = journal.iloc[[3]].assign(document_no="FIS-779", document_no_norm="FIS-779")
    res = checks.reconcile(invoices, pd.concat([journal, extra], ignore_index=True), ["153"])
    assert res["Belge No Uyuşmayan Eşleşmeler (Kontrol Edin)"].empty
    assert "XYZ2024000000010" in list(res["Muhasebeleşmemiş Faturalar"]["Fatura_No"])
    assert {"FIS-778", "FIS-779"} <= set(res["Faturası Bulunmayan Yevmiye Kayıtları"]["Yevmiye_Belge_No"])


def test_price_anomalies_threshold_and_period(db):
    _load_scenario(db)
    lines = db.get_lines_df()
    out = checks.price_anomalies(lines, "Aylık", 15)
    risky = out[out["Risk_Durumu"] == "YÜKSEK RİSK"]
    assert list(risky["Fatura_No"]) == ["F2", "F5"]
    assert checks.price_anomalies(lines, "Aylık", 60)["Risk_Durumu"].eq("Normal").all()


def test_full_audit(db):
    _load_scenario(db)
    summary, sections, notes, sayilar = checks.run_full_audit(db, "Aylık", 15, ["153"], 0.01, "")
    counts = dict(zip(summary["Kontrol"], summary["Bulgu_Sayisi"]))
    assert counts["Olası Mükerrer Faturalar"] == 1
    assert counts["Muhasebeleşmemiş Faturalar"] == 1
    assert any("Firma VKN" in n for n in notes)
    assert counts["Belge No Uyuşmayan Eşleşmeler (Kontrol Edin)"] == 0
    assert sayilar["eslesme"]["Tam"] == 4 and sayilar["eslesme"]["Eşleşmeyen"] == 1


def test_xml_import_and_customer_check(db):
    with open(os.path.join(DATA, "ornek_fatura.xml"), "rb") as fh:
        res = importers.read_xml(fh.read(), "ornek_fatura.xml")
    db.save_invoices(res.items, "XML", "ornek_fatura.xml", "h")
    inv, lines = db.get_invoices_df(), db.get_lines_df()
    assert checks.calculation_errors(inv, lines).empty
    assert len(checks.customer_mismatch(inv, "1111111111")) == 1
    assert checks.customer_mismatch(inv, "9876543210").empty


def test_legacy_migration(tmp_path):
    path = str(tmp_path / "old.db")
    conn = sqlite3.connect(path)
    conn.executescript("""
        CREATE TABLE invoices (id INTEGER PRIMARY KEY AUTOINCREMENT, invoice_no TEXT, issue_date DATE,
            supplier_vkn TEXT, supplier_name TEXT, currency TEXT, total_amount REAL, UNIQUE(supplier_vkn, invoice_no));
        CREATE TABLE invoice_lines (id INTEGER PRIMARY KEY AUTOINCREMENT, invoice_id INTEGER, item_name TEXT,
            quantity REAL, uom TEXT, unit_price_net REAL);
        CREATE TABLE journal_entries (id INTEGER PRIMARY KEY AUTOINCREMENT, entry_date DATE, document_no TEXT,
            account_code TEXT, amount REAL);
        INSERT INTO invoices VALUES (1, 'F1', '2024-01-01 00:00:00', '11111111111', 'A', 'TRY', 100);
        INSERT INTO invoice_lines VALUES (1, 1, 'Vida', 10, 'ADET', 10);
        INSERT INTO journal_entries VALUES (1, '2024-01-01', 'F1', '153', 100);
    """)
    conn.commit()
    conn.close()
    db = DatabaseManager(path)
    assert db.counts()["fatura"] == 1 and db.counts()["yevmiye_satiri"] == 1
    assert db.get_lines_df()["line_net"].iloc[0] == 100


def test_uom_and_upper():
    from denetim.utils import normalize_uom, tr_upper
    assert normalize_uom("C62") == normalize_uom("Adet") == normalize_uom(None) == "ADET"
    assert normalize_uom("KGM") == normalize_uom("kg") == "KG"
    assert tr_upper("Seçili faturalar") == "SEÇİLİ FATURALAR"
