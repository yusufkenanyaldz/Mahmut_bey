import os
import sqlite3

import pytest

from denetim import checks
from denetim.database import DatabaseManager
from denetim.firms import (DEFAULT_LEGACY_TITLE, REGISTRY_FILE, FirmError, FirmRegistry, code_to_filename,
                           validate_vkn)


@pytest.fixture
def reg(tmp_path):
    return FirmRegistry(str(tmp_path / "veri"))


def header(no, vkn="1111111111", customer_vkn="2222222222"):
    return {"invoice_no": no, "issue_date": "2024-01-10", "supplier_vkn": vkn, "supplier_name": "Tedarikçi",
            "customer_vkn": customer_vkn, "total_amount": 100.0, "source": "XML"}


def line():
    return {"item_name": "Çimento", "quantity": 1.0, "line_net": 100.0}


def test_create_and_list(reg):
    a = reg.create("ABC İnşaat", "ABC İnşaat A.Ş.", "123 456 7890", "İnşaat")
    reg.create("hali", "Halı Ltd.")
    assert a.vkn == "1234567890"
    assert a.db_file == "ABC_INSAAT.db"
    assert os.path.isfile(reg.db_path(a))
    assert [f.code for f in reg.list_firms()] == ["ABC İnşaat", "hali"]
    assert reg.get("HALI").title == "Halı Ltd."  # Kod büyük/küçük harf duyarsız (ASCII)
    assert os.path.isfile(os.path.join(reg.data_dir, REGISTRY_FILE))


@pytest.mark.parametrize("kwargs,msg", [
    ({"code": "", "title": "X"}, "boş"),
    ({"code": "A", "title": "  "}, "unvan"),
    ({"code": "A", "title": "X", "vkn": "12345"}, "10"),
    ({"code": "A" * 31, "title": "X"}, "en fazla"),
])
def test_create_validation(reg, kwargs, msg):
    with pytest.raises(FirmError, match=msg):
        reg.create(**kwargs)


def test_duplicate_code_and_unique_file(reg):
    reg.create("A-1", "Bir")
    with pytest.raises(FirmError, match="zaten var"):
        reg.create("a-1", "İki")
    b = reg.create("A 1", "Üç")  # Aynı dosya adı köküne düşer
    assert b.db_file == "A_1.db"
    c = reg.create("A/1", "Dört")
    assert c.db_file == "A_1_2.db"
    assert code_to_filename("çğış") == "CGIS"
    assert code_to_filename("///") == "FIRMA"


def test_validate_vkn():
    assert validate_vkn("") == ""
    assert validate_vkn(12345678901.0) == "12345678901"
    with pytest.raises(FirmError):
        validate_vkn("123")


def test_firm_databases_are_isolated(reg):
    a, b = reg.create("A", "Firma A"), reg.create("B", "Firma B")
    db_a, db_b = reg.open_db(a.code), reg.open_db(b.code)
    db_a.save_invoices([(header("F1"), [line()])], "XML", "f.xml", "h1")
    db_a.set_setting("threshold", "20")
    assert db_a.counts()["fatura"] == 1
    assert db_b.counts()["fatura"] == 0
    assert db_b.get_setting("threshold") is None
    # Aynı dosya diğer firmada "daha önce yüklenmiş" sayılmaz
    assert db_b.find_import("XML", "h1") is None
    # Veri silme yalnızca aktif firmayı etkiler
    db_b.save_invoices([(header("F2"), [line()])], "XML", "g.xml", "h2")
    db_a.clear_data()
    assert db_a.counts()["fatura"] == 0
    assert db_b.counts()["fatura"] == 1
    assert db_a.get_setting("threshold") == "20"  # Analiz ayarları korunur


def test_update(reg):
    reg.create("A", "Eski Unvan", "1234567890")
    reg.set_last_firm("A")
    f = reg.update("A", new_code="A2", title="Yeni Unvan", vkn="", sector="Otomotiv")
    assert (f.code, f.title, f.vkn, f.sector, f.db_file) == ("A2", "Yeni Unvan", "", "Otomotiv", "A.db")
    assert reg.get("A") is None
    assert reg.get_last_firm().code == "A2"
    reg.create("B", "B")
    with pytest.raises(FirmError, match="zaten var"):
        reg.update("A2", new_code="b")
    with pytest.raises(FirmError, match="10"):
        reg.update("A2", vkn="1")
    with pytest.raises(FirmError, match="bulunamadı"):
        reg.update("YOK", title="x")
    assert reg.update("A2", title="Sadece unvan").sector == "Otomotiv"


def test_delete(reg):
    a, b = reg.create("A", "A"), reg.create("B", "B")
    reg.open_db("A").set_setting("x", "1")
    reg.set_last_firm("A")
    path = reg.db_path(a)
    reg.delete("A")
    assert not os.path.exists(path)
    assert reg.get("A") is None
    assert reg.get_last_firm() is None
    assert os.path.exists(reg.db_path(b))
    with pytest.raises(FirmError):
        reg.delete("A")
    with pytest.raises(FirmError):
        reg.open_db("A")


def test_last_firm_persists(tmp_path):
    reg = FirmRegistry(str(tmp_path / "veri"))
    reg.create("A", "A")
    reg.create("B", "B")
    assert reg.get_last_firm() is None
    reg.set_last_firm("B")
    assert FirmRegistry(str(tmp_path / "veri")).get_last_firm().code == "B"


def test_company_vkn_from_firm_record(reg):
    f = reg.create("A", "A", "2222222222")
    db = reg.open_db(f.code)
    db.save_invoices([(header("F1"), [line()]), (header("F2", customer_vkn="9999999999"), [line()])],
                     "XML", "f.xml", "h")
    _, sections, notes, _ = checks.run_full_audit(db, "Aylık", 15, [], 0.01, f.vkn)
    assert list(sections["Alıcısı Firma Olmayan Faturalar"]["Fatura_No"]) == ["F2"]


# ------------------------------------------------------------------ eski audit_data.db aktarımı
def make_legacy(folder, title=None, vkn=None):
    path = os.path.join(folder, "audit_data.db")
    db = DatabaseManager(path)
    db.save_invoices([(header("ESKI1"), [line()])], "XML", "e.xml", "eh")
    db.set_setting("threshold", "25")
    if title:
        db.set_setting("company_title", title)
    if vkn:
        db.set_setting("company_vkn", vkn)
    return path


def test_legacy_import(tmp_path, reg):
    legacy = make_legacy(str(tmp_path), "Eski Firma A.Ş.", "2222222222")
    before = open(legacy, "rb").read()
    assert FirmRegistry.find_legacy_db(str(tmp_path / "yok"), str(tmp_path)) == legacy
    assert reg.legacy_pending(legacy)
    f = reg.import_legacy(legacy)
    assert (f.code, f.title, f.vkn) == ("AKTARILAN", "Eski Firma A.Ş.", "2222222222")
    db = reg.open_db(f.code)
    assert db.counts()["fatura"] == 1
    assert db.get_setting("threshold") == "25"  # Analiz ayarları taşınır
    assert os.path.exists(legacy) and open(legacy, "rb").read() == before  # Eski dosya olduğu gibi kalır
    assert not reg.legacy_pending(legacy)
    # Silinse bile tekrar aktarılmaz
    reg.delete(f.code)
    assert not reg.legacy_pending(legacy)


def test_legacy_import_default_title_and_code_clash(tmp_path, reg):
    legacy = make_legacy(str(tmp_path))
    reg.create("AKTARILAN", "Mevcut")
    f = reg.import_legacy(legacy)
    assert (f.code, f.title, f.vkn) == ("AKTARILAN_2", DEFAULT_LEGACY_TITLE, "")
    assert reg.update(f.code, title="Düzenlendi").title == "Düzenlendi"
    g = FirmRegistry(str(tmp_path / "veri2")).import_legacy(legacy, title="Kullanıcının Girdiği")
    assert g.title == "Kullanıcının Girdiği"


def test_legacy_v1_schema_migrated_on_copy(tmp_path, reg):
    legacy = str(tmp_path / "audit_data.db")
    conn = sqlite3.connect(legacy)
    conn.executescript("""
        CREATE TABLE invoices (id INTEGER PRIMARY KEY, invoice_no TEXT, issue_date TEXT, supplier_vkn TEXT,
                               supplier_name TEXT, currency TEXT, total_amount REAL);
        CREATE TABLE invoice_lines (id INTEGER PRIMARY KEY, invoice_id INTEGER, item_name TEXT, quantity REAL,
                                    uom TEXT, unit_price_net REAL);
        INSERT INTO invoices VALUES (1, 'A1', '2024-01-01', '1111111111', 'T', 'TRY', 100);
        INSERT INTO invoice_lines VALUES (1, 1, 'Ürün', 2, 'ADET', 50);
    """)
    conn.close()
    f = reg.import_legacy(legacy)
    assert reg.open_db(f.code).counts()["fatura"] == 1
    cols = {r[1] for r in sqlite3.connect(legacy).execute("PRAGMA table_info(invoices)")}
    assert "invoice_no_norm" not in cols  # Eski dosyaya dokunulmadı


def test_legacy_pending_missing_file(reg, tmp_path):
    assert not reg.legacy_pending(None)
    assert not reg.legacy_pending(str(tmp_path / "audit_data.db"))
