"""Excel ve XML dosyalarını doğrulayıp veritabanına yazılacak yapılara çevirir."""
import hashlib
import io
import json
import os
import re
import zipfile
from dataclasses import dataclass, field

import pandas as pd

from .ubl import UBLParseError, parse_ubl
from .utils import (column_key, match_columns, normalize_doc_no, normalize_vkn, parse_date,
                    parse_number)

INVOICE_SCHEMA = {
    "Fatura_No": ["Fatura Numarası", "Fatura Numarasi", "Fatura No", "Fatura Nosu", "Fatura Seri No",
                  "Belge No", "Belge Numarası", "Evrak No", "Evrak Numarası"],
    "Tarih": ["Fatura_Tarihi", "Fatura Tarihi", "Belge Tarihi", "Evrak Tarihi", "Düzenleme Tarihi"],
    "Tedarikci_VKN": ["Tedarikçi VKN", "VKN", "TCKN", "VKN_TCKN", "VKN/TCKN", "Satici_VKN", "Satıcı VKN/TCKN",
                      "Gönderici VKN", "Gönderici VKN/TCKN", "Vergi No", "Vergi Numarası", "Vergi Kimlik No"],
    "Tedarikci_Ad": ["Tedarikçi", "Tedarikci", "Tedarikçi Adı", "Tedarikçi Unvanı", "Unvan", "Unvanı", "Satici",
                     "Satıcı Unvanı", "Satıcı Adı", "Gönderici", "Gönderici Unvanı", "Firma Adı", "Firma Unvanı"],
    "Urun_Adi": ["Ürün Adı", "Urun", "Ürün", "Mal_Hizmet", "Mal/Hizmet", "Mal Hizmet Adı", "Mal/Hizmet Adı",
                 "Ürün/Hizmet", "Ürün/Hizmet Adı", "Stok Adı", "Malzeme Adı", "Hizmet Adı"],
    "Miktar": ["Miktarı", "Adet"],
    "Birim": ["Olcu_Birimi", "Ölçü Birimi", "Birimi", "Miktar Birimi"],
    "Fiyat": ["Birim_Fiyat", "Birim Fiyat", "Birim Fiyatı", "B.Fiyat", "Birim Fiyat TL", "Fiyatı"],
    "Iskonto": ["Iskonto_Tutari", "İskonto Tutarı", "İskonto TL", "İndirim Tutarı"],
    "KDV_Orani": ["KDV", "KDV Oranı", "KDV %", "KDV Yüzdesi"],
    "Para_Birimi": ["Doviz", "Döviz", "Para Birimi", "Döviz Cinsi", "Döviz Türü"],
    "Kur": ["Doviz_Kuru", "Döviz Kuru", "Kuru"],
}
INVOICE_REQUIRED = ["Fatura_No", "Tarih", "Tedarikci_VKN", "Tedarikci_Ad", "Urun_Adi", "Miktar", "Fiyat"]

# "Fiş No" / "Yevmiye No" bilerek takma ad değildir: fiş/madde sıra numarasıdır, belge (fatura) numarası değil.
JOURNAL_SCHEMA = {
    "Tarih": ["Yevmiye_Tarihi", "Yevmiye Tarihi", "Fis_Tarihi", "Fiş Tarihi", "Kayıt Tarihi", "İşlem Tarihi",
              "Evrak Tarihi", "Belge Tarihi", "Fiş Tar."],
    "Belge_No": ["Belge No", "Belge Numarası", "Belge Nosu", "Belge Seri No", "Belge Seri Sıra No", "Evrak_No",
                 "Evrak No", "Evrak Numarası", "Evrak Nosu", "Fatura_No", "Fatura No", "Fatura Numarası"],
    "Hesap_Kodu": ["Hesap Kodu", "Hesap", "Hesap_No", "Hesap No", "Hesap Numarası", "Hesap Kod", "Hes. Kodu",
                   "Muhasebe Hesabı", "Muhasebe Hesap Kodu", "Hesap Plan Kodu"],
    "Tutar": ["Tutar TL", "Tutarı", "İşlem Tutarı"],
    "Borc": ["Borç", "Borç Tutarı", "Borç Tutar", "Borç TL", "Borç (TL)", "Borçlu Tutar", "Borç Tutarı TL"],
    "Alacak": ["Alacak Tutarı", "Alacak Tutar", "Alacak TL", "Alacak (TL)", "Alacaklı Tutar", "Alacak Tutarı TL"],
    "Aciklama": ["Açıklama", "Açıklaması", "Fiş Açıklaması", "Satır Açıklaması", "Madde Açıklaması",
                 "Detay Açıklama"],
}
JOURNAL_REQUIRED = ["Tarih", "Belge_No", "Hesap_Kodu"]
JOURNAL_AMOUNT = ["Tutar", "Borc", "Alacak"]

# Arayüzde gösterilen alan adları
FIELD_LABELS = {
    "Tarih": "Tarih", "Belge_No": "Belge No", "Hesap_Kodu": "Hesap Kodu", "Borc": "Borç", "Alacak": "Alacak",
    "Tutar": "Tutar (Borç − Alacak yerine)", "Aciklama": "Açıklama", "Fatura_No": "Fatura No",
    "Tedarikci_VKN": "Tedarikçi VKN/TCKN", "Tedarikci_Ad": "Tedarikçi Adı", "Urun_Adi": "Ürün Adı",
    "Miktar": "Miktar", "Birim": "Birim", "Fiyat": "Birim Fiyat (KDV hariç)", "Iskonto": "İskonto",
    "KDV_Orani": "KDV Oranı", "Para_Birimi": "Para Birimi", "Kur": "Kur",
}
# Eşleme türleri: (şema, zorunlu alanlar). Kayıtlı eşlemeler bu anahtarlarla saklanır.
KIND_YEVMIYE, KIND_FATURA = "YEVMIYE", "FATURA"
SCHEMAS = {KIND_YEVMIYE: JOURNAL_SCHEMA, KIND_FATURA: INVOICE_SCHEMA}

HEADER_SCAN_ROWS = 20   # Başlık satırı aranırken taranan ilk satır sayısı
PREVIEW_ROWS = 10       # Eşleme penceresindeki önizleme satırı sayısı
BELGE_NO_HINT = ("Belge numarası ayrı bir sütunda bulunamadı. Program belge numarasını yalnızca ayrı bir sütundan "
                 "okur; açıklama metninin içinden ayıklamaz. Lütfen muhasebe programından belge numarasının ayrı "
                 "bir sütunda olduğu bir döküm alın (ör. döküme Evrak No / Belge No sütununu ekleyerek) ya da "
                 "dosyada böyle bir sütun varsa 'Sütunları Eşle' ile seçin.")
# Döküm sonundaki özet satırları: "Toplam", "Genel Toplam", "Ara Toplam", "Nakli Yekün", "Devreden" ...
_SUMMARY_RE = re.compile(r"^(genel|ara|donem|sayfa|ay|hesap|fis|kumulatif)?(toplam|toplami|yekun)|^nakliyekun|^devreden")


@dataclass
class ImportResult:
    items: list = field(default_factory=list)
    errors: list = field(default_factory=list)
    warnings: list = field(default_factory=list)
    infos: list = field(default_factory=list)       # Bilgi: atlanan başlık/özet satırları, kullanılan eşleme
    missing: list = field(default_factory=list)     # Bulunamayan zorunlu alanlar (doluysa eşleme gerekir)
    layout: "TableLayout" = None                    # Dosyanın algılanan düzeni (eşleme penceresi için)
    mapping: "ColumnMapping" = None                 # Yüklemede kullanılan eşleme

    @property
    def needs_mapping(self):
        return bool(self.missing)


@dataclass
class ColumnMapping:
    """Dosya sütunlarının standart alanlara eşlemesi.

    header_row: başlık satırı (0 tabanlı; Excel satır no = header_row + 1)
    columns: {standart_alan: dosyadaki sütun adı}
    """
    header_row: int
    columns: dict
    signature: str = ""
    saved: bool = False  # Firmanın kayıtlı eşlemesinden mi geldi

    def to_json(self):
        return json.dumps({"header_row": self.header_row, "columns": self.columns, "signature": self.signature},
                          ensure_ascii=False)

    @classmethod
    def from_json(cls, text, saved=True):
        d = json.loads(text)
        return cls(int(d["header_row"]), dict(d["columns"]), d.get("signature", ""), saved)


@dataclass
class TableLayout:
    """Excel sayfasının ham hali, başlık satırı ve otomatik sütun eşlemesi."""
    raw: pd.DataFrame
    header_row: int
    columns: list           # Başlık satırındaki sütun adları (boşlar "Sütun C" gibi adlandırılır)
    mapping: dict           # Otomatik eşleme {standart_alan: sütun adı}
    signature: str

    def preview(self, n=PREVIEW_ROWS):
        """Başlık satırının altındaki ilk n dolu satır (sütun adlarıyla)."""
        body = self.raw.iloc[self.header_row + 1:]
        body = body[~body.apply(lambda r: all(_is_blank(v) for v in r), axis=1)].head(n).copy()
        body.columns = self.columns
        body.index = [_excel_row(i) for i in body.index]
        return body.apply(lambda col: col.map(_preview_text))


def file_hash(data):
    return hashlib.sha256(data).hexdigest()


def _excel_row(idx):
    """Ham (başlıksız okunmuş) sayfanın satır indeksini Excel satır numarasına çevirir."""
    return int(idx) + 1


def _is_blank(value):
    return value is None or (isinstance(value, float) and pd.isna(value)) or str(value).strip() == ""


def _cell_text(value):
    if _is_blank(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip()


def _preview_text(value):
    if hasattr(value, "strftime") and not _is_blank(value) and not pd.isna(value):
        return value.strftime("%d.%m.%Y")
    return _cell_text(value)


def _col_letter(i):
    letters = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        letters = chr(65 + r) + letters
    return letters


def read_raw_excel(data):
    """İlk sayfayı başlıksız okur: satır indeksi + 1 = Excel satır numarası."""
    raw = pd.read_excel(io.BytesIO(data), header=None, dtype=object)
    raw.columns = range(raw.shape[1])
    return raw


def header_columns(raw, header_row):
    """Başlık satırındaki hücrelerden benzersiz sütun adları üretir."""
    names, seen = [], {}
    values = raw.iloc[header_row].tolist() if 0 <= header_row < len(raw) else [None] * raw.shape[1]
    for j, value in enumerate(values):
        name = _cell_text(value) or f"Sütun {_col_letter(j)}"
        if name in seen:
            seen[name] += 1
            name = f"{name} ({seen[name]})"
        else:
            seen[name] = 1
        names.append(name)
    return names


def row_signature(raw, header_row):
    """Başlık satırının imzası: dolu hücrelerin normalize adları (sıralı). Aynı biçimdeki dosyalar aynı imzayı verir."""
    if not 0 <= header_row < len(raw):
        return ""
    keys = [column_key(_cell_text(v)) for v in raw.iloc[header_row].tolist() if _cell_text(v)]
    return hashlib.sha1("|".join(keys).encode("utf-8")).hexdigest()[:16]


def detect_header_row(raw, schema, max_rows=HEADER_SCAN_ROWS):
    """İlk max_rows satır içinde bilinen sütun adlarıyla en çok eşleşen satırı başlık kabul eder.

    Eşitlikte (hiç eşleşme yoksa da) en çok metin hücresi olan satır, o da eşitse üstteki satır seçilir; böylece
    tek hücreli firma adı / rapor başlığı satırları başlık sanılmaz.
    """
    best, best_score = 0, None
    for i in range(min(max_rows, len(raw))):
        values = raw.iloc[i].tolist()
        cells = [_cell_text(v) for v in values]
        if not any(cells):
            continue
        texts = sum(1 for v in values if isinstance(v, str) and v.strip())
        score = (len(match_columns([c for c in cells if c], schema)), texts)
        if best_score is None or score > best_score:
            best, best_score = i, score
    return best


def analyze_layout(data, schema, header_row=None):
    """Dosyanın düzenini çıkarır: başlık satırı (verilmezse otomatik), sütun adları, otomatik eşleme, imza."""
    raw = data if isinstance(data, pd.DataFrame) else read_raw_excel(data)
    if header_row is None:
        header_row = detect_header_row(raw, schema)
    columns = header_columns(raw, header_row)
    real = [c for c, v in zip(columns, raw.iloc[header_row].tolist() if header_row < len(raw) else [])
            if _cell_text(v)]
    return TableLayout(raw, header_row, columns, match_columns(real, schema), row_signature(raw, header_row))


def find_saved_mapping(raw, saved_mappings):
    """Kayıtlı eşlemelerden bu dosyanın başlık imzasına uyan ilkini döndürür (yoksa None).

    saved_mappings: [ColumnMapping, ...] (en yeni önce). İmza, eşlemenin kendi başlık satırında kontrol edilir;
    böylece başlık satırı elle seçilmiş eşlemeler de tanınır.
    """
    for m in saved_mappings or []:
        if m.signature and row_signature(raw, m.header_row) == m.signature:
            columns = header_columns(raw, m.header_row)
            if all(c in columns for c in m.columns.values()):
                return ColumnMapping(m.header_row, dict(m.columns), m.signature, saved=True)
    return None


def missing_fields(columns_map, kind):
    """Eşlemede eksik zorunlu alanlar."""
    if kind == KIND_YEVMIYE:
        missing = [c for c in JOURNAL_REQUIRED if c not in columns_map]
        if not any(c in columns_map for c in JOURNAL_AMOUNT):
            missing.append("Tutar (veya Borc/Alacak)")
        return missing
    return [c for c in INVOICE_REQUIRED if c not in columns_map]


def missing_message(missing, kind):
    msg = f"Eksik sütun(lar): {', '.join(missing)}."
    if kind == KIND_FATURA:
        msg += f" Gerekli: {', '.join(INVOICE_REQUIRED)}."
    if kind == KIND_YEVMIYE and missing == ["Belge_No"]:
        return msg + " " + BELGE_NO_HINT
    return msg + " Dosyadaki sütun adları farklıysa 'Sütunları Eşle' ile eşleyebilirsiniz."


def validate_mapping(columns_map, kind):
    """Eşleme penceresinden gelen eşlemeyi doğrular. Dönüş: hata mesajları listesi (boşsa geçerli)."""
    errors = []
    missing = missing_fields(columns_map, kind)
    if missing:
        errors.append(missing_message(missing, kind))
    used = {}
    for std, col in columns_map.items():
        if col in used:
            errors.append(f"'{col}' sütunu hem {FIELD_LABELS.get(used[col], used[col])} hem "
                          f"{FIELD_LABELS.get(std, std)} için seçilmiş; her sütun tek alana eşlenebilir.")
        used[col] = std
    return errors


def _summary_label(cells):
    for text in cells:
        if text and _SUMMARY_RE.match(column_key(text)):
            return text
    return None


def _load_table(data, kind, mapping=None, saved_mappings=None):
    """Excel'i başlık tespiti + eşleme ile standart sütunlu bir tabloya çevirir.

    Dönüş: (df ya da None, ImportResult). df'nin indeksi ham satır indeksidir (Excel satırı = indeks + 1);
    üstteki başlık satırları, boş satırlar, tekrarlanan başlıklar ve özet (Toplam) satırları çıkarılmıştır.
    """
    schema = SCHEMAS[kind]
    result = ImportResult()
    raw = data if isinstance(data, pd.DataFrame) else read_raw_excel(data)
    if mapping is None and saved_mappings:
        mapping = find_saved_mapping(raw, saved_mappings)
    layout = analyze_layout(raw, schema, None if mapping is None else mapping.header_row)
    if mapping is None:
        mapping = ColumnMapping(layout.header_row, dict(layout.mapping), layout.signature)
    elif not mapping.signature:
        mapping.signature = layout.signature
    result.layout, result.mapping = layout, mapping

    unknown = [c for c in mapping.columns.values() if c not in layout.columns]
    if unknown:
        result.errors.append(f"Eşlemedeki sütun(lar) dosyada yok: {', '.join(unknown)}")
        result.missing = [s for s, c in mapping.columns.items() if c in unknown]
        return None, result
    result.missing = missing_fields(mapping.columns, kind)
    if result.missing:
        result.errors.append(missing_message(result.missing, kind))
        return None, result

    if mapping.saved:
        result.infos.append("Kayıtlı eşleme kullanıldı (firmanın aynı sütun başlıklı dosyası için daha önce "
                            "onaylanan eşleme)")
    if mapping.header_row > 0:
        result.infos.append(f"Sütun başlıkları {_excel_row(mapping.header_row)}. satırda bulundu; üstteki "
                            f"{mapping.header_row} satır (firma adı / rapor başlığı / boş satır) atlandı")
    renamed = [f"{c} → {s}" for s, c in mapping.columns.items() if c != s]
    if renamed:
        result.infos.append("Sütun eşlemesi: " + ", ".join(renamed))

    body = raw.iloc[mapping.header_row + 1:]
    pos = {c: j for j, c in enumerate(layout.columns)}
    date_col = pos[mapping.columns["Tarih"]]
    keep, summary, repeated = [], [], []
    for idx, values in zip(body.index, body.itertuples(index=False, name=None)):
        cells = [_cell_text(v) for v in values]
        if not any(cells):
            continue
        if all(column_key(cells[pos[c]]) == column_key(c) for c in mapping.columns.values()):
            repeated.append(_excel_row(idx))
            continue
        label = _summary_label(cells)
        if label is not None:
            try:
                parse_date(values[date_col])
            except ValueError:
                summary.append(f"{_excel_row(idx)} ({label})")
                continue
        keep.append(idx)
    if summary:
        result.infos.append(f"Özet satırı veri sayılmadı: satır {', '.join(summary)}")
    if repeated:
        result.infos.append(f"Tekrarlanan başlık satırı atlandı: satır {', '.join(map(str, repeated))}")
    df = pd.DataFrame({std: body.loc[keep, pos[col]] for std, col in mapping.columns.items()}, index=keep)
    return df, result


def parse_saved_mappings(items):
    """Ayar tablosundan gelen JSON metinlerini ColumnMapping listesine çevirir (bozuk kayıtlar atlanır)."""
    out = []
    for text in items:
        try:
            out.append(ColumnMapping.from_json(text))
        except (ValueError, KeyError, TypeError):
            continue
    return out


# ---------------------------------------------------------------------- fatura excel
def read_invoice_excel(data, source_file=None, mapping=None, saved_mappings=None):
    """Fatura Excel'ini okur. Aynı Fatura_No + VKN'ye sahip satırlar tek faturanın kalemleri sayılır.

    data: dosya içeriği (bytes) ya da read_raw_excel çıktısı. mapping: ColumnMapping (verilmezse başlık satırı ve
    sütunlar otomatik bulunur). saved_mappings: firmanın kayıtlı eşlemeleri; dosyanın imzası uyuyorsa kullanılır.
    """
    df, result = _load_table(data, KIND_FATURA, mapping, saved_mappings)
    if df is None:
        return result
    found = set(df.columns)

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
def read_journal_excel(data, mapping=None, saved_mappings=None):
    """Yevmiye Excel'ini okur. Parametreler read_invoice_excel ile aynıdır.

    Belge numarası mutlaka ayrı bir sütunda olmalıdır; açıklama metninden belge no ayıklanmaz.
    """
    df, result = _load_table(data, KIND_YEVMIYE, mapping, saved_mappings)
    if df is None:
        return result
    found = set(df.columns)
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
