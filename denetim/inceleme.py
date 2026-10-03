"""Bulgu inceleme kayıtları (GUI'den bağımsız): denetçinin her bulguya verdiği durum, not ve tarih.

Durumlar: Açık / İncelendi – Sorun Yok / Düzeltme İstendi. Kayıtlar firmanın veritabanında `bulgu_inceleme`
tablosunda saklanır.

Bulgu anahtarı yeniden içe aktarmalardan ve raporun yeniden çalıştırılmasından etkilenmez: kontrol kodu + normalize
tedarikçi VKN + normalize fatura / belge no (mükerrer gruplarında sıralı fatura no listesi, fiyat anomalisinde ayrıca
ürün adı, hesaplama kontrolünde kontrol adı). Satır sırası, veritabanı kimliği, tutar, eşik ya da tolerans gibi
değişebilen değerler anahtara girmez; böylece "sorun yok" denen bir kur farkı tutar değişmeden tekrar raporlansa da
işaretli kalır.
"""
from collections import OrderedDict
from datetime import datetime

import pandas as pd

from .utils import normalize_doc_no, normalize_text, normalize_vkn

DURUM_ACIK = "Açık"
DURUM_SORUN_YOK = "İncelendi – Sorun Yok"
DURUM_DUZELTME = "Düzeltme İstendi"
DURUMLAR = (DURUM_ACIK, DURUM_SORUN_YOK, DURUM_DUZELTME)

DURUM_COL = "Inceleme_Durumu"
NOT_COL = "Inceleme_Notu"
TARIH_COL = "Inceleme_Tarihi"
ANAHTAR_COL = "Bulgu_Anahtari"
INCELEME_COLS = [DURUM_COL, NOT_COL, TARIH_COL, ANAHTAR_COL]
OZET_COLS = ["Kontrol", "Bulgu_Sayisi", "Acik", "Sorun_Yok", "Duzeltme_Istendi"]

_VKN, _NO, _BELGE, _URUN, _KONTROL, _NOLAR = "vkn", "no", "belge", "urun", "kontrol", "nolar"
_ALANLAR = {_VKN: ("Tedarikci_VKN", normalize_vkn), _NO: ("Fatura_No", normalize_doc_no),
            _BELGE: ("Yevmiye_Belge_No", normalize_doc_no), _URUN: ("Urun_Adi", normalize_text),
            _KONTROL: ("Kontrol", normalize_text)}

# (bölüm başlığı öneki, kontrol kodu, anahtar alanları). Başlık önekle eşleşir: "Fiyat Anomalileri (±%15)".
KONTROLLER = [
    ("Fiyat Anomalileri", "FIYAT", (_VKN, _NO, _URUN)),
    ("Muhasebeleşmemiş Faturalar", "MUHASEBESIZ", (_VKN, _NO)),
    ("Seçili Hesap Dışına Kaydedilmiş Faturalar", "HESAP_DISI", (_VKN, _NO)),
    ("Tutar Farkları", "TUTAR_FARKI", (_VKN, _NO)),
    ("Dönem Farkları", "DONEM_FARKI", (_VKN, _NO)),
    ("Belge No Uyuşmayan Eşleşmeler", "BELGE_NO_UYUSMAYAN", (_VKN, _NO)),
    ("Faturası Bulunmayan Yevmiye Kayıtları", "FATURASIZ", (_BELGE,)),
    ("Belirsiz Eşleşme (Aynı No", "BELIRSIZ_AYNI_NO", (_VKN, _NO)),
    ("Belirsiz Eşleşme (Birden Fazla", "BELIRSIZ_SERI_SIRA", (_VKN, _NO)),
    ("Olası Mükerrer Faturalar", "MUKERRER", (_VKN, _NOLAR)),
    ("Fatura Hesaplama Tutarsızlıkları", "HESAPLAMA", (_VKN, _NO, _KONTROL)),
    ("Alıcısı Firma Olmayan Faturalar", "ALICI", (_VKN, _NO)),
    ("KDV Farkları", "KDV_FARKI", (_VKN, _NO)),
    ("KDV'si Kaydedilmemiş Faturalar", "KDV_KAYITSIZ", (_VKN, _NO)),
    ("Tevkifat Kaydı Eksik", "TEVKIFAT", (_VKN, _NO)),
]


def kontrol_kodu(baslik):
    """Bölüm başlığının kontrol kodu; incelenebilir bir bulgu bölümü değilse (özet sayfaları vb.) None."""
    baslik = str(baslik)
    return next((kod for onek, kod, _ in KONTROLLER if baslik.startswith(onek)), None)


def _alanlar(kod):
    return next(alanlar for _, k, alanlar in KONTROLLER if k == kod)


def bulgu_anahtari(kod, satir):
    """Bulgunun kararlı anahtarı: 'TUTAR_FARKI|1234567890|ABC2024000000123'.

    satir: bölüm satırı (dict ya da pandas Series). Mükerrer gruplarında Fatura_No virgülle ayrılmış listedir;
    numaralar normalize edilip sıralanarak '+' ile birleştirilir.
    """
    parcalar = [kod]
    for alan in _alanlar(kod):
        if alan == _NOLAR:
            nolar = sorted(normalize_doc_no(x) for x in str(satir["Fatura_No"]).split(",") if x.strip())
            parcalar.append("+".join(nolar))
        else:
            sutun, normalize = _ALANLAR[alan]
            parcalar.append(normalize(satir.get(sutun)))
    return "|".join(parcalar)


def bolum_anahtarlari(baslik, df):
    """Bölümdeki her satırın bulgu anahtarı (liste); bölüm incelenebilir değilse None."""
    kod = kontrol_kodu(baslik)
    if kod is None or df is None:
        return None
    return [bulgu_anahtari(kod, r) for _, r in df.iterrows()]


# ---------------------------------------------------------------------- kayıt
def incelemeleri_oku(db):
    """{anahtar: {"kontrol", "durum", "aciklama", "tarih"}}"""
    with db.connection() as conn:
        rows = conn.execute("SELECT anahtar, kontrol, durum, aciklama, tarih FROM bulgu_inceleme").fetchall()
    return {a: {"kontrol": k, "durum": d, "aciklama": n or "", "tarih": t or ""} for a, k, d, n, t in rows}


def isaretle(db, anahtarlar, durum, aciklama=None, tarih=None):
    """Bulguların inceleme durumunu kaydeder (varsa günceller).

    aciklama None ise mevcut not korunur; "" verilirse not silinir. Kontrol kodu anahtarın ilk parçasıdır.
    Dönüş: kaydedilen bulgu sayısı.
    """
    if durum not in DURUMLAR:
        raise ValueError(f"Geçersiz inceleme durumu: {durum}")
    tarih = tarih or datetime.now().strftime("%Y-%m-%d %H:%M")
    anahtarlar = list(dict.fromkeys(a for a in anahtarlar if a))
    with db.connection() as conn:
        for a in anahtarlar:
            onceki = conn.execute("SELECT aciklama FROM bulgu_inceleme WHERE anahtar = ?", (a,)).fetchone()
            not_ = (onceki[0] if onceki else "") if aciklama is None else str(aciklama).strip()
            conn.execute("INSERT OR REPLACE INTO bulgu_inceleme (anahtar, kontrol, durum, aciklama, tarih) "
                         "VALUES (?,?,?,?,?)", (a, a.split("|", 1)[0], durum, not_ or "", tarih))
    return len(anahtarlar)


# ---------------------------------------------------------------------- uygulama
def durum_uygula(baslik, df, kayitlar):
    """Bölüme inceleme sütunlarını (durum, not, tarih, anahtar) ekler. İncelenebilir değilse df aynen döner.

    Kaydı olmayan bulgular "Açık"tır.
    """
    anahtarlar = bolum_anahtarlari(baslik, df)
    if anahtarlar is None:
        return df
    out = df.drop(columns=[c for c in INCELEME_COLS if c in df.columns]).copy()
    kayit = [kayitlar.get(a) or {} for a in anahtarlar]
    out[DURUM_COL] = [k.get("durum") or DURUM_ACIK for k in kayit]
    out[NOT_COL] = [k.get("aciklama", "") for k in kayit]
    out[TARIH_COL] = [k.get("tarih", "") for k in kayit]
    out[ANAHTAR_COL] = anahtarlar
    return out


def bolumlere_uygula(sections, kayitlar):
    """OrderedDict(başlık → DataFrame) içindeki tüm bulgu bölümlerine inceleme sütunlarını ekler."""
    return OrderedDict((t, durum_uygula(t, df, kayitlar)) for t, df in sections.items())


def sorun_yok_gizle(df):
    """"İncelendi – Sorun Yok" işaretli bulguları çıkarır (inceleme sütunu yoksa df aynen döner)."""
    if df is None or DURUM_COL not in df.columns:
        return df
    return df[df[DURUM_COL] != DURUM_SORUN_YOK]


def inceleme_ozeti(sections):
    """Her bulgu bölümü için açık / sorun yok / düzeltme istendi sayıları (bölümlere_uygula çıktısı üzerinde)."""
    rows = []
    for baslik, df in sections.items():
        if kontrol_kodu(baslik) is None or df is None or DURUM_COL not in df.columns:
            continue
        d = df[DURUM_COL]
        rows.append({"Kontrol": baslik, "Bulgu_Sayisi": len(df), "Acik": int((d == DURUM_ACIK).sum()),
                     "Sorun_Yok": int((d == DURUM_SORUN_YOK).sum()),
                     "Duzeltme_Istendi": int((d == DURUM_DUZELTME).sum())})
    return pd.DataFrame(rows, columns=OZET_COLS)


def ozet_satiri(satir):
    """İnceleme özeti satırını kısa metne çevirir: 'açık 12 | sorun yok 3 | düzeltme 1'."""
    return (f"açık {satir['Acik']} | sorun yok {satir['Sorun_Yok']} | düzeltme istendi "
            f"{satir['Duzeltme_Istendi']}")
