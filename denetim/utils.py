"""Ortak yardımcı fonksiyonlar: sayı/tarih ayrıştırma, metin normalizasyonu, sütun eşleme."""
import math
import re
from datetime import date, datetime

import pandas as pd

_TR_UPPER_TO_LOWER = str.maketrans({"I": "ı", "İ": "i"})
_TR_ASCII = str.maketrans("çğıöşüÇĞİÖŞÜ", "cgiosuCGIOSU")


def tr_lower(text):
    """Türkçe kurallara uygun küçük harfe çevirme (I→ı, İ→i)."""
    return str(text).translate(_TR_UPPER_TO_LOWER).lower()


def normalize_text(text):
    """Ürün adı gibi serbest metinleri karşılaştırma için normalize eder."""
    if text is None or (isinstance(text, float) and math.isnan(text)):
        return ""
    return re.sub(r"\s+", " ", tr_lower(text)).strip()


def normalize_doc_no(value):
    """Fatura / belge numarasını eşleştirme için normalize eder (boşluk yok, büyük harf)."""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return re.sub(r"\s+", "", str(value)).upper()


def normalize_vkn(value):
    """VKN/TCKN değerini rakamlardan oluşan metne çevirir (Excel'in 1.23E+9 vb. bozmalarına karşı)."""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return re.sub(r"\D", "", str(value))


def normalize_account(code):
    """Hesap kodunu ayraçlarından arındırır: '153.01.001' → '15301001'."""
    if code is None or (isinstance(code, float) and math.isnan(code)):
        return ""
    if isinstance(code, float) and code.is_integer():
        code = int(code)
    return re.sub(r"[\s.\-_/]", "", str(code))


def parse_number(value):
    """Sayıyı float'a çevirir. '1.234,56', '1,234.56', '1234,5' gibi biçimleri destekler.

    Boş / geçersiz değerde ValueError fırlatır.
    """
    if value is None:
        raise ValueError("boş değer")
    if isinstance(value, bool):
        raise ValueError(f"geçersiz sayı: {value}")
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            raise ValueError("boş değer")
        return float(value)
    text = str(value).strip().replace(" ", "").replace(" ", "")
    text = re.sub(r"(TL|TRY|₺)$", "", text, flags=re.IGNORECASE)
    if not text:
        raise ValueError("boş değer")
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        raise ValueError(f"geçersiz sayı: {value}") from None


def parse_date(value):
    """Tarihi 'YYYY-MM-DD' metnine çevirir. Geçersizse ValueError fırlatır."""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        raise ValueError("boş tarih")
    if isinstance(value, (datetime, pd.Timestamp)):
        if pd.isna(value):
            raise ValueError("boş tarih")
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    text = str(value).strip()
    if not text:
        raise ValueError("boş tarih")
    if re.match(r"^\d{4}-\d{1,2}-\d{1,2}", text):
        parsed = pd.to_datetime(text[:10], format="%Y-%m-%d", errors="coerce")
    else:
        parsed = pd.to_datetime(text, dayfirst=True, errors="coerce")
    if pd.isna(parsed):
        raise ValueError(f"geçersiz tarih: {value}")
    return parsed.strftime("%Y-%m-%d")


def column_key(name):
    """Sütun başlığını karşılaştırma anahtarına çevirir: 'Tedarikçi VKN' → 'tedarikcivkn'."""
    text = str(name).translate(_TR_ASCII).lower()
    return re.sub(r"[^a-z0-9]", "", text)


def match_columns(columns, schema):
    """Sütun başlıklarını şemadaki standart adlara eşler.

    schema: {standart_ad: [takma_ad, ...]}
    Dönüş: {standart_ad: dosyadaki_sütun_adı} (her standart ad ve her sütun en çok bir kez kullanılır)
    """
    lookup = {}
    for standard, aliases in schema.items():
        for alias in [standard] + list(aliases):
            lookup.setdefault(column_key(alias), standard)
    mapping = {}
    for col in columns:
        std = lookup.get(column_key(col))
        if std and std not in mapping:
            mapping[std] = col
    return mapping


def map_columns(df, schema):
    """Excel sütunlarını şemadaki standart adlara eşler.

    schema: {standart_ad: [takma_ad, ...]}
    Dönüş: (yeniden adlandırılmış df, bulunan standart adlar kümesi)
    """
    mapping = match_columns(df.columns, schema)
    return df.rename(columns={col: std for std, col in mapping.items()}), set(mapping)


def period_of(iso_date, period_type):
    """'YYYY-MM-DD' tarihinden dönem etiketi üretir: Aylık → 2024-03, Çeyreklik → 2024-Ç1, Yıllık → 2024."""
    year, month = iso_date[:4], int(iso_date[5:7])
    if period_type == "Yıllık":
        return year
    if period_type == "Çeyreklik":
        return f"{year}-Ç{(month - 1) // 3 + 1}"
    return f"{year}-{month:02d}"


def parse_account_list(text):
    """'153, 770.01; 760' → ['153', '77001', '760']"""
    parts = re.split(r"[,;\s]+", str(text or ""))
    return [normalize_account(p) for p in parts if normalize_account(p)]


def tr_upper(text):
    """Türkçe kurallara uygun büyük harfe çevirme (i→İ, ı→I)."""
    return str(text).replace("i", "İ").replace("ı", "I").upper()


def tr_ascii_upper(text):
    """Büyük/küçük harf ve Türkçe karakter duyarsız karşılaştırma anahtarı: 'Açılış' → 'ACILIS'."""
    return tr_upper(text).translate(_TR_ASCII)


# UBL-TR (UN/ECE Rec. 20) birim kodları ve yaygın Türkçe yazımlar → ortak birim
_UOM_MAP = {
    "C62": "ADET", "NIU": "ADET", "EA": "ADET", "AD": "ADET", "ADT": "ADET", "ADET": "ADET", "PCS": "ADET",
    "KGM": "KG", "KG": "KG", "KILO": "KG", "KİLOGRAM": "KG", "KILOGRAM": "KG",
    "GRM": "GR", "GR": "GR", "GRAM": "GR",
    "TNE": "TON", "TON": "TON",
    "LTR": "LT", "LT": "LT", "L": "LT", "LITRE": "LT", "LİTRE": "LT",
    "MTR": "M", "M": "M", "METRE": "M", "MT": "M",
    "MTK": "M2", "M2": "M2", "M²": "M2",
    "MTQ": "M3", "M3": "M3", "M³": "M3",
    "PA": "PAKET", "PK": "PAKET", "PAKET": "PAKET",
    "BX": "KUTU", "KUTU": "KUTU", "KOLİ": "KOLI", "KOLI": "KOLI", "CT": "KOLI",
    "SET": "SET", "PR": "CIFT", "ÇİFT": "CIFT", "CIFT": "CIFT",
    "HUR": "SAAT", "SAAT": "SAAT", "DAY": "GUN", "GÜN": "GUN", "GUN": "GUN", "MON": "AY", "AY": "AY",
    "KWH": "KWH",
}


def normalize_uom(uom):
    """Birimi karşılaştırma için ortak forma çevirir: 'C62' ve 'Adet' → 'ADET'."""
    if uom is None or (isinstance(uom, float) and math.isnan(uom)) or not str(uom).strip():
        return "ADET"
    key = tr_upper(str(uom).strip())
    return _UOM_MAP.get(key, key)
