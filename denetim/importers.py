"""Excel ve XML dosyalarını doğrulayıp veritabanına yazılacak yapılara çevirir."""
import hashlib
import io
import os
import zipfile
from dataclasses import dataclass, field

import pandas as pd

from .ubl import UBLParseError, parse_ubl
from .utils import map_columns, normalize_doc_no, normalize_vkn, parse_date, parse_number

INVOICE_SCHEMA = {
    "Fatura_No": ["Fatura Numarası", "Fatura Numarasi"],
    "Tarih": ["Fatura_Tarihi", "Fatura Tarihi"],
    "Tedarikci_VKN": ["Tedarikçi VKN", "VKN", "TCKN", "VKN_TCKN", "Satici_VKN"],
    "Tedarikci_Ad": ["Tedarikçi", "Tedarikci", "Tedarikçi Adı", "Unvan", "Satici"],
    "Urun_Adi": ["Ürün Adı", "Urun", "Ürün", "Mal_Hizmet"],
    "Miktar": [],
    "Birim": ["Olcu_Birimi", "Ölçü Birimi"],
    "Fiyat": ["Birim_Fiyat", "Birim Fiyat"],
    "Iskonto": ["Iskonto_Tutari", "İskonto Tutarı"],
    "KDV_Orani": ["KDV", "KDV Oranı"],
    "Para_Birimi": ["Doviz", "Döviz"],
    "Kur": ["Doviz_Kuru", "Döviz Kuru"],
}
INVOICE_REQUIRED = ["Fatura_No", "Tarih", "Tedarikci_VKN", "Tedarikci_Ad", "Urun_Adi", "Miktar", "Fiyat"]

JOURNAL_SCHEMA = {
    "Tarih": ["Yevmiye_Tarihi", "Fis_Tarihi", "Fiş Tarihi"],
    "Belge_No": ["Belge No", "Evrak_No", "Evrak No", "Fatura_No"],
    "Hesap_Kodu": ["Hesap Kodu", "Hesap", "Hesap_No"],
    "Tutar": [],
    "Borc": ["Borç"],
    "Alacak": [],
    "Aciklama": ["Açıklama"],
}
JOURNAL_REQUIRED = ["Tarih", "Belge_No", "Hesap_Kodu"]


@dataclass
class ImportResult:
    items: list = field(default_factory=list)
    errors: list = field(default_factory=list)
    warnings: list = field(default_factory=list)


def file_hash(data):
    return hashlib.sha256(data).hexdigest()


def _excel_row(idx):
    """DataFrame indeksini Excel satır numarasına çevirir (başlık 1. satır)."""
    return int(idx) + 2


def _is_blank(value):
    return value is None or (isinstance(value, float) and pd.isna(value)) or str(value).strip() == ""


def _read_excel(data):
    return pd.read_excel(io.BytesIO(data), dtype=object)


# ---------------------------------------------------------------------- fatura excel
def read_invoice_excel(data, source_file=None):
    """Fatura Excel'ini okur. Aynı Fatura_No + VKN'ye sahip satırlar tek faturanın kalemleri sayılır."""
    result = ImportResult()
    df = _read_excel(data)
    df, found = map_columns(df, INVOICE_SCHEMA)
    missing = [c for c in INVOICE_REQUIRED if c not in found]
    if missing:
        result.errors.append(f"Eksik sütun(lar): {', '.join(missing)}. Gerekli: {', '.join(INVOICE_REQUIRED)}")
        return result

    groups = {}  # (vkn, no_norm) -> {"header": ..., "lines": [...], "rows": [...], "bad": bool}
    order = []
    for idx, row in df.iterrows():
        if all(_is_blank(row.get(c)) for c in INVOICE_REQUIRED):
            continue  # tamamen boş satır
        excel_row = _excel_row(idx)
        row_errors = []

        invoice_no = "" if _is_blank(row["Fatura_No"]) else str(row["Fatura_No"]).strip()
        if isinstance(row["Fatura_No"], float) and row["Fatura_No"].is_integer():
            invoice_no = str(int(row["Fatura_No"]))
        if not invoice_no:
            row_errors.append("Fatura_No boş")
        vkn = normalize_vkn(row["Tedarikci_VKN"])
        if len(vkn) not in (10, 11):
            row_errors.append(f"Tedarikci_VKN 10 (VKN) veya 11 (TCKN) haneli olmalı ({row['Tedarikci_VKN']})")
        try:
            issue_date = parse_date(row["Tarih"])
        except ValueError as e:
            row_errors.append(f"Tarih: {e}")
            issue_date = None
        name = "" if _is_blank(row["Urun_Adi"]) else str(row["Urun_Adi"]).strip()
        if not name:
            row_errors.append("Urun_Adi boş")

        def num(col, default=None, required=True):
            if col not in found or _is_blank(row.get(col)):
                if required:
                    row_errors.append(f"{col} boş")
                return default
            try:
                return parse_number(row[col])
            except ValueError as e:
                row_errors.append(f"{col}: {e}")
                return default

        qty = num("Miktar", 0.0)
        price = num("Fiyat", 0.0)
        discount = num("Iskonto", 0.0, required=False)
        vat_rate = num("KDV_Orani", None, required=False)
        rate = num("Kur", 1.0, required=False)
        if qty is not None and qty <= 0 and "Miktar boş" not in row_errors:
            row_errors.append(f"Miktar sıfırdan büyük olmalı ({qty})")
        if price is not None and price < 0:
            row_errors.append(f"Fiyat negatif olamaz ({price})")
        currency = "TRY" if "Para_Birimi" not in found or _is_blank(row.get("Para_Birimi")) \
            else str(row["Para_Birimi"]).strip().upper().replace("TL", "TRY")

        key = (vkn, normalize_doc_no(invoice_no))
        if key not in groups:
            groups[key] = {"header": None, "lines": [], "rows": [], "bad": False}
            order.append(key)
        grp = groups[key]
        grp["rows"].append(excel_row)

        if row_errors:
            grp["bad"] = True
            result.errors.append(f"Satır {excel_row}: " + "; ".join(row_errors))
            continue

        line_net = qty * price - (discount or 0.0)
        line = {
            "line_no": len(grp["lines"]) + 1,
            "item_name": name,
            "quantity": qty,
            "uom": "ADET" if "Birim" not in found or _is_blank(row.get("Birim")) else str(row["Birim"]).strip().upper(),
            "unit_price": price,
            "line_net": line_net,
            "vat_rate": vat_rate,
            "vat_amount": line_net * vat_rate / 100 if vat_rate is not None else None,
        }
        header = {
            "invoice_no": invoice_no,
            "issue_date": issue_date,
            "supplier_vkn": vkn,
            "supplier_name": str(row["Tedarikci_Ad"]).strip() if not _is_blank(row["Tedarikci_Ad"]) else "",
            "invoice_type": "SATIS",
            "currency": currency,
            "exchange_rate": rate or 1.0,
            "source": "EXCEL",
            "source_file": source_file,
        }
        if grp["header"] is None:
            grp["header"] = header
        else:
            first = grp["header"]
            for fld, label in (("issue_date", "Tarih"), ("currency", "Para_Birimi")):
                if first[fld] != header[fld]:
                    grp["bad"] = True
                    result.errors.append(
                        f"Satır {excel_row}: {invoice_no} nolu faturanın {label} değeri önceki satırlarla farklı "
                        f"({first[fld]} ≠ {header[fld]})")
        grp["lines"].append(line)

    for key in order:
        grp = groups[key]
        if grp["bad"] or grp["header"] is None:
            if grp["header"] is not None or key[1]:
                result.warnings.append(
                    f"Fatura {key[1] or '?'} (VKN {key[0] or '?'}, satır {', '.join(map(str, grp['rows']))}) "
                    f"hatalı satır içerdiği için tamamen atlandı")
            continue
        header = grp["header"]
        net = sum(ln["line_net"] for ln in grp["lines"])
        vat_values = [ln["vat_amount"] for ln in grp["lines"] if ln["vat_amount"] is not None]
        header.update({
            "line_extension_amount": net,
            "allowance_total": 0.0,
            "total_amount": net,
            "vat_amount": sum(vat_values) if vat_values else 0.0,
            "payable_amount": net + sum(vat_values) if vat_values else None,
        })
        result.items.append((header, grp["lines"]))
    return result


# ---------------------------------------------------------------------- yevmiye excel
def read_journal_excel(data):
    result = ImportResult()
    df = _read_excel(data)
    df, found = map_columns(df, JOURNAL_SCHEMA)
    missing = [c for c in JOURNAL_REQUIRED if c not in found]
    has_amount = "Tutar" in found or "Borc" in found or "Alacak" in found
    if missing or not has_amount:
        if not has_amount:
            missing.append("Tutar (veya Borc/Alacak)")
        result.errors.append(f"Eksik sütun(lar): {', '.join(missing)}")
        return result
    use_debit_credit = "Borc" in found or "Alacak" in found
    if use_debit_credit and "Tutar" in found:
        result.warnings.append("Hem Tutar hem Borç/Alacak sütunu var; Borç - Alacak kullanıldı")

    for idx, row in df.iterrows():
        cols = JOURNAL_REQUIRED + [c for c in ("Tutar", "Borc", "Alacak") if c in found]
        if all(_is_blank(row.get(c)) for c in cols):
            continue
        excel_row = _excel_row(idx)
        errs = []
        try:
            entry_date = parse_date(row["Tarih"])
        except ValueError as e:
            errs.append(f"Tarih: {e}")
            entry_date = None
        doc = row["Belge_No"]
        if isinstance(doc, float) and not pd.isna(doc) and doc.is_integer():
            doc = int(doc)
        doc = "" if _is_blank(doc) else str(doc).strip()
        if not doc:
            errs.append("Belge_No boş")
        account = row["Hesap_Kodu"]
        if isinstance(account, float) and not pd.isna(account) and account.is_integer():
            account = int(account)
        account = "" if _is_blank(account) else str(account).strip()
        if not account:
            errs.append("Hesap_Kodu boş")

        amount = None
        try:
            if use_debit_credit:
                debit = 0.0 if "Borc" not in found or _is_blank(row.get("Borc")) else parse_number(row["Borc"])
                credit = 0.0 if "Alacak" not in found or _is_blank(row.get("Alacak")) else parse_number(row["Alacak"])
                amount = debit - credit
            else:
                amount = parse_number(row["Tutar"])
        except ValueError as e:
            errs.append(f"Tutar: {e}")

        if errs:
            result.errors.append(f"Satır {excel_row}: " + "; ".join(errs))
            continue
        result.items.append({
            "entry_date": entry_date,
            "document_no": doc,
            "account_code": account,
            "amount": amount,
            "description": None if "Aciklama" not in found or _is_blank(row.get("Aciklama")) else str(row["Aciklama"]),
            "source_row": excel_row,
        })
    return result


# ---------------------------------------------------------------------- xml
def iter_xml_payloads(path):
    """Bir .xml dosyası ya da içinde XML'ler bulunan .zip için (görünen_ad, bytes) üretir."""
    base = os.path.basename(path)
    if path.lower().endswith(".zip"):
        with zipfile.ZipFile(path) as zf:
            for info in zf.infolist():
                if not info.is_dir() and info.filename.lower().endswith(".xml"):
                    yield f"{base}/{info.filename}", zf.read(info)
    else:
        with open(path, "rb") as fh:
            yield base, fh.read()


def read_xml(data, display_name):
    """Tek XML içeriğini ayrıştırır. Dönüş: ImportResult (items: [(header, lines)])"""
    result = ImportResult()
    try:
        header, lines, warnings = parse_ubl(data, source_file=display_name)
    except UBLParseError as e:
        result.errors.append(f"{display_name}: {e}")
        return result
    result.items.append((header, lines))
    result.warnings.extend(f"{display_name}: {w}" for w in warnings)
    return result


# ---------------------------------------------------------------------- şablonlar
def invoice_template():
    return pd.DataFrame([{
        "Fatura_No": "ABC2024000000001", "Tarih": "15.01.2024", "Tedarikci_VKN": "1234567890",
        "Tedarikci_Ad": "Örnek Tedarikçi A.Ş.", "Urun_Adi": "Örnek Ürün", "Miktar": 10, "Birim": "ADET",
        "Fiyat": 100.0, "Iskonto": 0, "KDV_Orani": 20, "Para_Birimi": "TRY", "Kur": 1,
    }])


def journal_template():
    return pd.DataFrame([
        {"Tarih": "15.01.2024", "Belge_No": "ABC2024000000001", "Hesap_Kodu": "153.01", "Borc": 1000.0,
         "Alacak": 0, "Aciklama": "Ticari mal alışı"},
        {"Tarih": "15.01.2024", "Belge_No": "ABC2024000000001", "Hesap_Kodu": "191.01", "Borc": 200.0,
         "Alacak": 0, "Aciklama": "İndirilecek KDV"},
        {"Tarih": "15.01.2024", "Belge_No": "ABC2024000000001", "Hesap_Kodu": "320.01", "Borc": 0,
         "Alacak": 1200.0, "Aciklama": "Satıcılar"},
    ])
