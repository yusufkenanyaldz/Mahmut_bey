"""Yevmiye / fatura Excel içe aktarma sihirbazı: başlık satırı tespiti, takma adlar, eşleme ve kayıtlı eşlemeler."""
import io
from datetime import datetime

import pytest
from openpyxl import Workbook

from denetim import importers
from denetim.database import DatabaseManager
from denetim.importers import KIND_FATURA, KIND_YEVMIYE, ColumnMapping


def sheet(rows):
    """Satır listelerinden xlsx üretir (None = boş hücre, [] = boş satır)."""
    wb = Workbook()
    ws = wb.active
    for r, values in enumerate(rows, start=1):
        for c, value in enumerate(values, start=1):
            if value is not None:
                ws.cell(row=r, column=c, value=value)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


D = datetime(2024, 7, 1)
LOGO_HEADER = ["Tarih", "Fiş No", "Belge No", "Hesap Kodu", "Hesap Adı", "Borç Tutarı", "Alacak Tutarı", "Açıklama"]
LOGO_ROWS = [
    [D, 1, "ABC2024000000001", "153.01", "", 1000, 0, "Mal alışı"],
    [D, 1, "ABC2024000000001", "320.01", "", 0, 1000, "Satıcı"],
]


@pytest.fixture
def db(tmp_path):
    return DatabaseManager(str(tmp_path / "firma.db"))


# ------------------------------------------------------------------ başlık satırı
def test_baslik_satiri_ustteki_satirlar_atlanir():
    data = sheet([["Yılmaz İnşaat A.Ş."], ["YEVMİYE DEFTERİ DÖKÜMÜ 01.07.2024 - 30.09.2024"], [],
                  ["Tarih", "Belge No", "Hesap Kodu", "Borç", "Alacak", "Açıklama"],
                  [D, "F1", "153", 100, 0, "a"], [], [D, "F1", "320", 0, 100, "b"],
                  ["bozuk", "F2", "153", 5, 0, ""]])
    res = importers.read_journal_excel(data)
    assert [r["amount"] for r in res.items] == [100.0, -100.0]
    assert [r["source_row"] for r in res.items] == [5, 7]  # Excel satır numaraları
    assert res.mapping.header_row == 3
    assert len(res.errors) == 1 and "Satır 8" in res.errors[0]
    assert any("4. satırda" in i and "3 satır" in i for i in res.infos)


def test_detect_header_row_en_cok_eslesen_ve_bos_sutun():
    raw = importers.read_raw_excel(sheet([["Rapor"], ["Tarih: 01.07.2024"],
                                          [None, "Tarih", "Evrak No", "Hesap No", None, "Borç TL", "Alacak TL"],
                                          [None, D, "F1", "153", "x", 10, 0]]))
    assert importers.detect_header_row(raw, importers.JOURNAL_SCHEMA) == 2
    layout = importers.analyze_layout(raw, importers.JOURNAL_SCHEMA)
    assert layout.columns[0] == "Sütun A" and layout.columns[4] == "Sütun E"
    assert layout.mapping == {"Tarih": "Tarih", "Belge_No": "Evrak No", "Hesap_Kodu": "Hesap No",
                              "Borc": "Borç TL", "Alacak": "Alacak TL"}
    assert list(layout.preview().index) == [4]
    assert layout.preview().loc[4, "Tarih"] == "01.07.2024"


def test_detect_header_row_eslesme_yoksa_en_genis_metin_satiri():
    raw = importers.read_raw_excel(sheet([[], ["a", "b"], ["1", "2"]]))
    assert importers.detect_header_row(raw, importers.JOURNAL_SCHEMA) == 1
    raw = importers.read_raw_excel(sheet([["ABC Ltd."], ["YEVMİYE DÖKÜMÜ"], [], ["Kayıt Günü", "Ref", "Hsp", "Meblağ"],
                                          [D, "F1", "153", 100]]))
    assert importers.detect_header_row(raw, importers.JOURNAL_SCHEMA) == 3


# ------------------------------------------------------------------ takma adlar
def test_logo_borc_tutari_ve_fis_no_belge_no_sayilmaz():
    res = importers.read_journal_excel(sheet([LOGO_HEADER] + LOGO_ROWS))
    assert not res.errors
    assert res.mapping.columns["Belge_No"] == "Belge No"
    assert res.mapping.columns["Borc"] == "Borç Tutarı" and res.mapping.columns["Alacak"] == "Alacak Tutarı"
    assert [r["amount"] for r in res.items] == [1000.0, -1000.0]
    assert res.items[0]["document_no"] == "ABC2024000000001"


@pytest.mark.parametrize("header,std", [
    ("Fiş Tarihi", "Tarih"), ("Kayıt Tarihi", "Tarih"), ("Evrak No", "Belge_No"), ("Belge Numarası", "Belge_No"),
    ("Hesap No", "Hesap_Kodu"), ("Borç TL", "Borc"), ("Alacak (TL)", "Alacak"), ("Fiş Açıklaması", "Aciklama"),
])
def test_yevmiye_takma_adlari(header, std):
    assert importers.match_columns([header], importers.JOURNAL_SCHEMA) == {std: header}


def test_fis_no_ve_yevmiye_no_takma_ad_degil():
    assert importers.match_columns(["Fiş No", "Yevmiye No"], importers.JOURNAL_SCHEMA) == {}


# ------------------------------------------------------------------ belge no ayrı sütunda olmalı
def test_belge_no_aciklamada_ise_aciklayici_hata():
    data = sheet([["Yevmiye No", "Tarih", "Hesap Kodu", "Açıklama", "Borç", "Alacak"],
                  [1, D, "153", "ABC2024000000001 - Mal alışı", 100, 0]])
    res = importers.read_journal_excel(data)
    assert res.items == [] and res.needs_mapping and res.missing == ["Belge_No"]
    assert "ayrı bir sütunda" in res.errors[0] and "Evrak No / Belge No" in res.errors[0]
    assert "Belge_No" not in res.mapping.columns  # Açıklamadan belge no ayıklanmaz


def test_validate_mapping():
    ok = {"Tarih": "T", "Belge_No": "B", "Hesap_Kodu": "H", "Borc": "Bo"}
    assert importers.validate_mapping(ok, KIND_YEVMIYE) == []
    errs = importers.validate_mapping({"Tarih": "T", "Hesap_Kodu": "H", "Tutar": "X"}, KIND_YEVMIYE)
    assert len(errs) == 1 and "ayrı bir sütunda" in errs[0]
    errs = importers.validate_mapping({"Tarih": "T"}, KIND_YEVMIYE)  # Yalnız belge no eksik değilse genel mesaj
    assert "Sütunları Eşle" in errs[0] and "ayrı bir sütunda" not in errs[0]
    errs = importers.validate_mapping(dict(ok, Aciklama="B"), KIND_YEVMIYE)
    assert len(errs) == 1 and "'B'" in errs[0]
    assert "Tutar (veya Borc/Alacak)" in importers.validate_mapping(
        {"Tarih": "T", "Belge_No": "B", "Hesap_Kodu": "H"}, KIND_YEVMIYE)[0]
    assert "Miktar" in importers.validate_mapping({"Fatura_No": "F"}, KIND_FATURA)[0]


# ------------------------------------------------------------------ özet / tekrarlanan başlık satırları
def test_toplam_satirlari_veri_sayilmaz():
    header = ["Tarih", "Belge No", "Hesap Kodu", "Borç", "Alacak", "Açıklama"]
    data = sheet([header,
                  [D, "F1", "153", 100, 0, "Toplam ödeme"],    # Tarihi olan gerçek kayıt: korunur
                  [D, "F1", "320", 0, 100, ""],
                  header,                                       # Sayfa sonunda tekrarlanan başlık
                  [D, "F2", "153", 50, 0, ""],
                  [None, None, "Ara Toplam", 150, 100, None],
                  ["GENEL TOPLAM", None, None, 150, 100, None]])
    res = importers.read_journal_excel(data)
    assert not res.errors
    assert [r["document_no"] for r in res.items] == ["F1", "F1", "F2"]
    info = " | ".join(res.infos)
    assert "Özet satırı veri sayılmadı: satır 6 (Ara Toplam), 7 (GENEL TOPLAM)" in info
    assert "Tekrarlanan başlık satırı atlandı: satır 4" in info


# ------------------------------------------------------------------ eşleme parametresi
def test_elle_eslesme_ile_okuma():
    data = sheet([["Döküm"], ["Kayıt Günü", "Ref", "Hsp", "Meblağ", "Not"],
                  [D, "F1", "153", "1.250,50", "x"], [D, "F1", "320", "-1.250,50", ""]])
    res = importers.read_journal_excel(data)
    assert res.needs_mapping and "Tarih" in res.missing
    mapping = ColumnMapping(1, {"Tarih": "Kayıt Günü", "Belge_No": "Ref", "Hesap_Kodu": "Hsp", "Tutar": "Meblağ",
                                "Aciklama": "Not"})
    res = importers.read_journal_excel(data, mapping=mapping)
    assert not res.errors
    assert [r["amount"] for r in res.items] == [1250.5, -1250.5]
    assert res.items[0]["description"] == "x" and res.items[1]["description"] is None
    assert res.mapping.signature == importers.row_signature(importers.read_raw_excel(data), 1)


def test_eslemedeki_sutun_dosyada_yoksa_hata():
    mapping = ColumnMapping(0, {"Tarih": "Tarih", "Belge_No": "Yok", "Hesap_Kodu": "Hesap Kodu",
                                "Borc": "Borç Tutarı"})
    res = importers.read_journal_excel(sheet([LOGO_HEADER] + LOGO_ROWS), mapping=mapping)
    assert res.items == [] and res.needs_mapping and "Yok" in res.errors[0]


# ------------------------------------------------------------------ kayıtlı eşleme
def test_kayitli_eslesme_ayni_bicimde_uygulanir(db):
    rows = [["Kayıt Günü", "Ref", "Hsp", "Meblağ"], [D, "F1", "153", 10]]
    data = sheet(rows)
    mapping = ColumnMapping(0, {"Tarih": "Kayıt Günü", "Belge_No": "Ref", "Hesap_Kodu": "Hsp", "Tutar": "Meblağ"})
    res = importers.read_journal_excel(data, mapping=mapping)
    db.save_column_mapping(KIND_YEVMIYE, res.mapping.signature, res.mapping.to_json())
    db.save_column_mapping(KIND_FATURA, "baska", '{"header_row": 0, "columns": {}, "signature": "baska"}')
    saved = importers.parse_saved_mappings(db.get_column_mappings(KIND_YEVMIYE) + ["bozuk json"])
    assert len(saved) == 1 and saved[0].saved

    # Aynı firmadan aynı biçimde yeni dosya (üstte başlık satırları farklı olsa bile sütun adları aynı)
    yeni = sheet(rows + [[D, "F2", "153", 20]])
    res2 = importers.read_journal_excel(yeni, saved_mappings=saved)
    assert not res2.errors and res2.mapping.saved and len(res2.items) == 2
    assert any("Kayıtlı eşleme kullanıldı" in i for i in res2.infos)

    # Başka biçimdeki dosyada kayıtlı eşleme uygulanmaz
    res3 = importers.read_journal_excel(sheet([LOGO_HEADER] + LOGO_ROWS), saved_mappings=saved)
    assert not res3.mapping.saved and not res3.errors


def test_imza_buyuk_kucuk_harf_ve_turkce_karakter_duyarsiz():
    a = importers.read_raw_excel(sheet([["Tarih", "Belge No", "Borç"]]))
    b = importers.read_raw_excel(sheet([["TARİH", "belge no", "Borc", None]]))
    c = importers.read_raw_excel(sheet([["Tarih", "Borç", "Belge No"]]))
    assert importers.row_signature(a, 0) == importers.row_signature(b, 0) != importers.row_signature(c, 0)


def test_kayitli_eslesme_yeniden_kaydedilince_en_yeni_once(db):
    db.save_column_mapping(KIND_YEVMIYE, "s1", ColumnMapping(0, {"Tarih": "A"}, "s1").to_json())
    db.save_column_mapping(KIND_YEVMIYE, "s2", ColumnMapping(0, {"Tarih": "B"}, "s2").to_json())
    db.save_column_mapping(KIND_YEVMIYE, "s1", ColumnMapping(1, {"Tarih": "C"}, "s1").to_json())
    saved = importers.parse_saved_mappings(db.get_column_mappings(KIND_YEVMIYE))
    assert [(m.signature, m.header_row) for m in saved] == [("s1", 1), ("s2", 0)]
    db.clear_data()  # Veri silme eşlemeleri silmez
    assert len(db.get_column_mappings(KIND_YEVMIYE)) == 2


# ------------------------------------------------------------------ fatura excel
def test_fatura_excel_baslik_satiri_ve_takma_adlar():
    data = sheet([["e-Fatura Listesi"], [],
                  ["Fatura No", "Fatura Tarihi", "Satıcı VKN/TCKN", "Satıcı Unvanı", "Mal/Hizmet", "Miktarı",
                   "Birim Fiyatı", "KDV %"],
                  ["F1", "01.07.2024", "1234567890", "A Ltd.", "Vida", 10, "2,5", 20],
                  ["F1", "01.07.2024", "1234567890", "A Ltd.", "Somun", 4, 5, 20],
                  ["Toplam", None, None, None, None, 14, None, None]])
    res = importers.read_invoice_excel(data, source_file="liste.xlsx")
    assert not res.errors and len(res.items) == 1
    header, lines = res.items[0]
    assert header["total_amount"] == pytest.approx(45.0) and len(lines) == 2
    assert any("3. satırda" in i for i in res.infos) and any("Toplam" in i for i in res.infos)


def test_fatura_excel_eslesme_parametresi():
    data = sheet([["No", "Gün", "VN", "Firma", "Kalem", "Adet", "BF"],
                  ["F9", D, "1234567890", "A", "Vida", 2, 3]])
    mapping = ColumnMapping(0, {"Fatura_No": "No", "Tarih": "Gün", "Tedarikci_VKN": "VN", "Tedarikci_Ad": "Firma",
                                "Urun_Adi": "Kalem", "Miktar": "Adet", "Fiyat": "BF"})
    assert importers.read_invoice_excel(data).needs_mapping
    res = importers.read_invoice_excel(data, mapping=mapping)
    assert not res.errors and res.items[0][0]["invoice_no"] == "F9"
