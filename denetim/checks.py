"""Denetim kontrolleri: fiyat anomalisi, mutabakat ve tutarlılık kontrolleri."""
import math
import re
from bisect import bisect_left, bisect_right
from collections import Counter, OrderedDict, defaultdict
from datetime import date

import pandas as pd

from .utils import normalize_uom, normalize_vkn, parse_account_list, period_of, tr_ascii_upper

TOLERANCE_DEFAULT = 0.01
# Dövizli (TRY dışı) faturalarda muhasebe fatura kuru yerine başka günün kurunu kullanabilir; TL tutarın bu
# yüzdesine kadar olan fark tutar farkı sayılmaz (firma başına değiştirilebilir). TL faturalara uygulanmaz:
# rakam yer değiştirme hataları (12.345,67 → 12.354,67) yalnızca birkaç TL fark yaratır.
KUR_TOLERANSI_VARSAYILAN = 1.0

# ---------------------------------------------------------------------- maliyete eklenen vergiler
# Alış faturasındaki bu vergiler indirilemez, maliyete eklenir (153/254/770 ... hesaplarına matrahla birlikte yazılır):
# ÖTV (0071 I, 0073–0076 III, 0077 IV, 9077 II motorlu taşıtlar), ÖİV (4080, 4081), konaklama vergisi (0059),
# BSMV (0021), elektrik / havagazı tüketim vergisi (4071, 8005), TRT payı (8004), çevre temizlik vergisi (8008).
# Stopajlar (0003, 0011) ve KDV tevkifatı maliyet değildir. Firma başına değiştirilebilir.
VARSAYILAN_MALIYET_VERGI_KODLARI = "0071, 0073, 0074, 0075, 0076, 0077, 9077, 4080, 4081, 0059, 0021, 4071, 8004, " \
                                   "8005, 8008"


def parse_vergi_kodlari(text):
    """'9077, 0071; 4080' → ['9077', '0071', '4080'] (tekrarsız). Kodlar 4 haneye tamamlanır ('71' → '0071')."""
    out = []
    for parca in re.split(r"[,;\s]+", str(text or "")):
        kod = re.sub(r"\D", "", parca)
        if kod:
            kod = kod.zfill(4)
            if kod not in out:
                out.append(kod)
    return out


class VergiAyarlari:
    """Vergi mutabakatı ayarları (firma başına saklanır).

    maliyet_kodlari: maliyete eklenen vergi türü kodları (parse_vergi_kodlari çıktısı; None → varsayılan)
    kdv_hesaplari: indirilecek KDV hesapları (parse_account_list çıktısı; None → 191, [] → KDV kontrolü kapalı)
    tevkifat_hesaplari: tevkif edilen KDV hesapları (None → 360, [] → kapalı)
    gelir_hesaplari: satış mutabakatının gelir hesapları (None → 600, 601, 602; [] → satış mutabakatı kapalı)
    satis_kdv_hesaplari: hesaplanan KDV hesapları (None → 391, [] → satış KDV kontrolü kapalı)
    """

    def __init__(self, maliyet_kodlari=None, kdv_hesaplari=None, tevkifat_hesaplari=None, gelir_hesaplari=None,
                 satis_kdv_hesaplari=None):
        def varsayilan(deger, metin, cevir):
            return cevir(metin) if deger is None else list(deger)
        self.maliyet_kodlari = varsayilan(maliyet_kodlari, VARSAYILAN_MALIYET_VERGI_KODLARI, parse_vergi_kodlari)
        self.kdv_hesaplari = varsayilan(kdv_hesaplari, VARSAYILAN_KDV_HESAPLARI, parse_account_list)
        self.tevkifat_hesaplari = varsayilan(tevkifat_hesaplari, VARSAYILAN_TEVKIFAT_HESAPLARI, parse_account_list)
        self.gelir_hesaplari = varsayilan(gelir_hesaplari, VARSAYILAN_GELIR_HESAPLARI, parse_account_list)
        self.satis_kdv_hesaplari = varsayilan(satis_kdv_hesaplari, VARSAYILAN_SATIS_KDV_HESAPLARI, parse_account_list)

    def __repr__(self):
        return (f"VergiAyarlari(maliyet_kodlari={self.maliyet_kodlari}, kdv_hesaplari={self.kdv_hesaplari}, "
                f"tevkifat_hesaplari={self.tevkifat_hesaplari}, gelir_hesaplari={self.gelir_hesaplari}, "
                f"satis_kdv_hesaplari={self.satis_kdv_hesaplari})")


def maliyet_vergisi_ozeti(invoices):
    """maliyet_vergisi_ekle() çıktısından özet: kaç faturada maliyete eklenen vergi var, toplam (TL)."""
    if invoices is None or invoices.empty or "maliyet_vergisi" not in invoices:
        return {"fatura": 0, "toplam_tl": 0.0}
    mv = invoices["maliyet_vergisi"].astype(float) * invoices["exchange_rate"].fillna(1.0).astype(float)
    return {"fatura": int((mv > 0).sum()), "toplam_tl": round(float(mv.sum()), 2)}


def maliyet_vergisi_metni(ozet, kodlar):
    if not ozet:
        return ""
    tutar = f"{ozet['toplam_tl']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return (f"Maliyete eklenen vergiler ({', '.join(kodlar) or 'kod yok'}): {ozet['fatura']} faturada {tutar} TL; "
            f"bu faturalarda beklenen maliyet = KDV hariç tutar + vergi")


def maliyet_vergisi_ekle(invoices, taxes, kodlar=None):
    """Faturalara maliyete eklenen vergi tutarını (belge para biriminde) ve vergi adlarını ekler.

    taxes: DatabaseManager.get_invoice_taxes_df() (invoice_id, tax_code, tax_name, amount)
    kodlar: parse_vergi_kodlari() çıktısı; None → VARSAYILAN_MALIYET_VERGI_KODLARI
    Dönüş: invoices kopyası + maliyet_vergisi (float), maliyet_vergi_adlari (metin)
    """
    kodlar = parse_vergi_kodlari(VARSAYILAN_MALIYET_VERGI_KODLARI) if kodlar is None else list(kodlar)
    inv = invoices.copy()
    inv["maliyet_vergisi"] = 0.0
    inv["maliyet_vergi_adlari"] = ""
    if inv.empty or taxes is None or taxes.empty or "id" not in inv:
        return inv
    t = taxes[taxes["tax_code"].astype(str).isin(kodlar)]
    if t.empty:
        return inv
    tutar = t.groupby("invoice_id")["amount"].sum()
    adlar = t.groupby("invoice_id").apply(
        lambda g: ", ".join(f"{k} {a}" for k, a in sorted(set(zip(g["tax_code"], g["tax_name"].fillna(""))))))
    inv["maliyet_vergisi"] = inv["id"].map(tutar).fillna(0.0).astype(float)
    inv["maliyet_vergi_adlari"] = inv["id"].map(adlar).fillna("")
    return inv


# ---------------------------------------------------------------------- fiyat analizi
# Hizmet / hakediş kalemleri her ay farklı tutarda faturalanır; birim fiyat karşılaştırması anlamsızdır.
# Ürün adında geçen kelimeler (firma başına değiştirilebilir)
VARSAYILAN_FIYAT_HARIC_KELIMELER = "HAKEDİŞ, İŞÇİLİK, HİZMET, FASON, KİRALAMA, BAKIM, ONARIM, DANIŞMANLIK"
# Grupta (dönem + ürün + birim + para birimi) bundan az alım varsa sapma riskli işaretlenmez (bilgi olarak kalır)
FIYAT_MIN_ALIM = 3
RISK_YUKSEK = "YÜKSEK RİSK"
RISK_NORMAL = "Normal"
RISK_YETERSIZ = "Yetersiz veri (bilgi)"
FIYAT_NEDEN_IADE = "İade faturası"
FIYAT_NEDEN_MIKTAR = "Miktar sıfır / boş"
FIYAT_NEDEN_TEVKIFAT = "Tevkifatlı fatura"
FIYAT_NEDEN_KELIME = "Anahtar kelime"
FIYAT_NEDEN_YETERSIZ = "Yetersiz veri"
FIYAT_COLS = ["Donem", "Tarih", "Fatura_No", "Tedarikci", "Tedarikci_VKN", "Urun_Adi", "Birim", "Para_Birimi",
              "Miktar", "Birim_Fiyat", "Kur", "Birim_Fiyat_TL", "AOBF", "Fark_Yuzdesi", "Donemdeki_Alim_Sayisi",
              "Risk_Durumu"]
FIYAT_HARIC_COLS = ["Tarih", "Fatura_No", "Tedarikci", "Fatura_Tipi", "Urun_Adi", "Birim", "Para_Birimi", "Miktar",
                    "Birim_Fiyat", "Birim_Fiyat_TL", "Analiz_Disi_Nedeni"]


class FiyatKurallari:
    """Fiyat analizinin hariç tutma kuralları (firma başına saklanır).

    tevkifat_haric: tevkifatlı (InvoiceTypeCode TEVKIFAT) faturaların satırları analize alınmaz
    kelime_haric: ürün adında hariç kelimelerden biri geçen satırlar analize alınmaz
    kelimeler: parse_kelime_listesi() çıktısı; None → VARSAYILAN_FIYAT_HARIC_KELIMELER
    min_alim: grupta bundan az alım varsa sapma riskli işaretlenmez (1 → kural kapalı)
    """

    def __init__(self, tevkifat_haric=True, kelime_haric=True, kelimeler=None, min_alim=FIYAT_MIN_ALIM):
        self.tevkifat_haric = bool(tevkifat_haric)
        self.kelime_haric = bool(kelime_haric)
        self.kelimeler = parse_kelime_listesi(VARSAYILAN_FIYAT_HARIC_KELIMELER) if kelimeler is None \
            else list(kelimeler)
        self.min_alim = max(1, int(min_alim))

    def __repr__(self):
        return (f"FiyatKurallari(tevkifat_haric={self.tevkifat_haric}, kelime_haric={self.kelime_haric}, "
                f"kelimeler={self.kelimeler}, min_alim={self.min_alim})")


def _kelime_anahtari(value):
    """Kelime karşılaştırma anahtarı: büyük/küçük harf ve Türkçe karakter duyarsız, tek boşluklu."""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    return re.sub(r"\s+", " ", tr_ascii_upper(str(value))).strip()


def parse_kelime_listesi(text):
    """'Hakediş, işçilik; HİZMET' → ['HAKEDIS', 'ISCILIK', 'HIZMET'] (karşılaştırma anahtarları, tekrarsız)."""
    out = []
    for parca in re.split(r"[,;\n]+", str(text or "")):
        anahtar = _kelime_anahtari(parca)
        if anahtar and anahtar not in out:
            out.append(anahtar)
    return out


def haric_kelime(urun_adi, kelimeler):
    """Ürün adında hariç kelimelerden biri geçiyorsa o kelimeyi, geçmiyorsa None döndürür.

    kelimeler: parse_kelime_listesi() çıktısı ('Taşeron İşçilik Hakedişi' → 'ISCILIK').
    """
    ad = _kelime_anahtari(urun_adi)
    if not ad:
        return None
    return next((k for k in kelimeler if k in ad), None)


def _fatura_tipi(value):
    return "" if value is None or (isinstance(value, float) and math.isnan(value)) else str(value).strip().upper()


def iade_faturasi(invoice_type):
    """IADE ve TEVKIFATIADE iade faturasıdır."""
    return "IADE" in _fatura_tipi(invoice_type)


def tevkifatli_fatura(invoice_type):
    """TEVKIFAT (iade olmayan) tevkifatlı faturadır."""
    tip = _fatura_tipi(invoice_type)
    return tip.startswith("TEVKIFAT") and "IADE" not in tip


def para_birimi(value):
    """Belge para birimi kodu: boş / TL → TRY."""
    pb = _fatura_tipi(value)
    return "TRY" if pb in ("", "TL", "YTL") else pb


# ---------------------------------------------------------------------- fatura yönü (alış / satış)
YON_ALIS = "ALIS"
YON_SATIS = "SATIS"


def fatura_yonleri(invoices, company_vkn):
    """Her faturanın yönü (ALIS / SATIS).

    Sırayla: Excel'de açıkça verilen yön (direction sütunu); satıcı VKN'si firma VKN'si olan faturalar SATIS
    (giden e-fatura / e-arşiv); diğerleri ALIS. Firmanın kestiği İADE faturaları alış iadesidir, ALIS sayılır.
    invoices: faturalar ya da fatura satırları (supplier_vkn, direction, invoice_type sütunları)
    """
    if invoices is None or invoices.empty:
        return pd.Series([], dtype=object)
    firma = normalize_vkn(company_vkn)
    acik = invoices["direction"].map(_fatura_tipi) if "direction" in invoices else pd.Series("", index=invoices.index)
    tip = invoices["invoice_type"] if "invoice_type" in invoices else pd.Series("", index=invoices.index)
    satici_firma = invoices["supplier_vkn"].map(normalize_vkn) == firma if firma else \
        pd.Series(False, index=invoices.index)
    yon = [a if a in (YON_ALIS, YON_SATIS) else (YON_SATIS if sf and not iade_faturasi(t) else YON_ALIS)
           for a, sf, t in zip(acik, satici_firma, tip)]
    return pd.Series(yon, index=invoices.index, dtype=object)


def alis_faturalari(invoices, company_vkn):
    """Yalnızca alış faturaları (ya da alış faturalarının satırları)."""
    if invoices is None or invoices.empty:
        return invoices
    return invoices[fatura_yonleri(invoices, company_vkn) == YON_ALIS]


def satis_faturalari(invoices, company_vkn):
    """Yalnızca satış faturaları."""
    if invoices is None or invoices.empty:
        return invoices
    return invoices[fatura_yonleri(invoices, company_vkn) == YON_SATIS]


def kdv_istisna(profile, invoice_type):
    """İhracat / istisna faturası mı? (KDV hesaplanmaz ya da 391'e yazılmaz: IHRACAT senaryosu, ISTISNA ve
    IHRACKAYITLI tipleri)"""
    return _fatura_tipi(profile) == "IHRACAT" or _fatura_tipi(invoice_type) in ("ISTISNA", "IHRACKAYITLI")


def fiyat_analiz_disi_nedeni(invoice_type, quantity, urun_adi, kurallar):
    """Satırın fiyat analizine neden alınmadığını döndürür (alınıyorsa None). Sıra: iade, miktar, tevkifat,
    anahtar kelime."""
    if iade_faturasi(invoice_type):
        return FIYAT_NEDEN_IADE
    if quantity is None or pd.isna(quantity) or float(quantity) <= 0:
        return FIYAT_NEDEN_MIKTAR
    if kurallar.tevkifat_haric and tevkifatli_fatura(invoice_type):
        return FIYAT_NEDEN_TEVKIFAT
    if kurallar.kelime_haric:
        k = haric_kelime(urun_adi, kurallar.kelimeler)
        if k:
            return f"{FIYAT_NEDEN_KELIME} ({k})"
    return None


class FiyatAnalizSonucu:
    """fiyat_analizi() sonucu: satirlar (analiz edilen tüm satırlar, Risk_Durumu sütunuyla), haric (analiz dışı
    satırlar ve eşik üstü olduğu hâlde yetersiz veri nedeniyle riskli işaretlenmeyenler) ve ozet (dict)."""

    def __init__(self, satirlar, haric, ozet):
        self.satirlar, self.haric, self.ozet = satirlar, haric, ozet

    @property
    def riskli(self):
        return self.satirlar[self.satirlar["Risk_Durumu"] == RISK_YUKSEK].reset_index(drop=True)


def fiyat_analizi(lines, period_type="Aylık", threshold=15.0, kurallar=None):
    """Dönem + ürün + birim + para birimi bazında ağırlıklı ortalama birim fiyattan (AOBF) sapmaları hesaplar.

    Birim fiyat ve AOBF belge para birimindedir; sapma belge para birimindeki fiyat üzerinden hesaplanır
    (kur dalgalanması sapma yaratmaz, EUR motorin TL motorinle karşılaştırılmaz). Birim_Fiyat_TL bilgi içindir.
    Analize alınmayanlar (bkz. fiyat_analiz_disi_nedeni): iade faturaları, miktarı 0 olan satırlar ve
    kurallar açıksa tevkifatlı faturalar ile ürün adında hariç kelime geçen satırlar.
    Grupta kurallar.min_alim'den az alım varsa eşik aşılsa da Risk_Durumu RISK_YETERSIZ olur (riskli sayılmaz).
    Dönüş: FiyatAnalizSonucu
    """
    kurallar = kurallar or FiyatKurallari()
    ozet = {"toplam": len(lines), "incelenen": 0, "riskli": 0, "iade": 0, "miktar": 0, "tevkifat": 0,
            "kelime": 0, "yetersiz_veri": 0, "esik": threshold, "min_alim": kurallar.min_alim,
            "tevkifat_haric": kurallar.tevkifat_haric, "kelime_haric": kurallar.kelime_haric,
            "kelimeler": list(kurallar.kelimeler)}
    if lines.empty:
        return FiyatAnalizSonucu(pd.DataFrame(columns=FIYAT_COLS), pd.DataFrame(columns=FIYAT_HARIC_COLS), ozet)
    df = lines.copy()
    df["Para_Birimi"] = df["currency"].map(para_birimi) if "currency" in df else "TRY"
    df["Kur"] = df["exchange_rate"].fillna(1.0).astype(float)
    df["uom_key"] = df["uom"].map(normalize_uom)
    miktar = pd.to_numeric(df["quantity"], errors="coerce")
    df["Birim_Fiyat"] = df["line_net"].astype(float) / miktar.where(miktar > 0)
    df["Birim_Fiyat_TL"] = df["Birim_Fiyat"] * df["Kur"]
    df["neden"] = [fiyat_analiz_disi_nedeni(t, q, a, kurallar)
                   for t, q, a in zip(df["invoice_type"], df["quantity"], df["item_name"])]

    def haric_tablosu(d, neden):
        return d.assign(Analiz_Disi_Nedeni=neden).rename(columns={
            "issue_date": "Tarih", "invoice_no": "Fatura_No", "supplier_name": "Tedarikci", "invoice_type": "Fatura_Tipi",
            "item_name": "Urun_Adi", "uom_key": "Birim", "quantity": "Miktar"})[FIYAT_HARIC_COLS]

    disarida = df[df["neden"].notna()]
    haric = haric_tablosu(disarida, disarida["neden"])
    nedenler = disarida["neden"].astype(str)
    ozet.update(iade=int((nedenler == FIYAT_NEDEN_IADE).sum()), miktar=int((nedenler == FIYAT_NEDEN_MIKTAR).sum()),
                tevkifat=int((nedenler == FIYAT_NEDEN_TEVKIFAT).sum()),
                kelime=int(nedenler.str.startswith(FIYAT_NEDEN_KELIME).sum()))

    df = df[df["neden"].isna()].copy()
    if df.empty:
        return FiyatAnalizSonucu(pd.DataFrame(columns=FIYAT_COLS), haric.reset_index(drop=True), ozet)
    df["Donem"] = df["issue_date"].map(lambda d: period_of(d, period_type))
    keys = ["Donem", "item_norm", "uom_key", "Para_Birimi"]
    grp = df.groupby(keys).agg(top_tutar=("line_net", "sum"), top_miktar=("quantity", "sum"),
                               Donemdeki_Alim_Sayisi=("id", "count")).reset_index()
    grp["AOBF"] = grp["top_tutar"] / grp["top_miktar"]
    df = df.merge(grp[keys + ["AOBF", "Donemdeki_Alim_Sayisi"]], on=keys)
    df["Fark_Yuzdesi"] = (df["Birim_Fiyat"] - df["AOBF"]) / df["AOBF"] * 100
    df.loc[df["AOBF"] == 0, "Fark_Yuzdesi"] = 0.0
    esik_ustu = df["Fark_Yuzdesi"].abs().gt(threshold)
    yetersiz = df["Donemdeki_Alim_Sayisi"] < kurallar.min_alim
    df["Risk_Durumu"] = RISK_NORMAL
    df.loc[esik_ustu & ~yetersiz, "Risk_Durumu"] = RISK_YUKSEK
    df.loc[esik_ustu & yetersiz, "Risk_Durumu"] = RISK_YETERSIZ

    bilgi = df[df["Risk_Durumu"] == RISK_YETERSIZ]
    if not bilgi.empty:
        neden = [f"{FIYAT_NEDEN_YETERSIZ} (dönemde {n} alım < {kurallar.min_alim}; sapma %{f:.1f})"
                 for n, f in zip(bilgi["Donemdeki_Alim_Sayisi"], bilgi["Fark_Yuzdesi"])]
        haric = pd.concat([haric, haric_tablosu(bilgi, neden)], ignore_index=True)

    out = df.rename(columns={"issue_date": "Tarih", "invoice_no": "Fatura_No", "supplier_name": "Tedarikci",
                             "supplier_vkn": "Tedarikci_VKN", "item_name": "Urun_Adi", "uom_key": "Birim",
                             "quantity": "Miktar"})[FIYAT_COLS]
    for c in ("Birim_Fiyat", "Birim_Fiyat_TL", "AOBF", "Fark_Yuzdesi"):
        out[c] = out[c].round(4 if c == "Birim_Fiyat" else 2)
    for c in ("Birim_Fiyat", "Birim_Fiyat_TL"):
        haric[c] = haric[c].astype(float).round(4 if c == "Birim_Fiyat" else 2)
    ozet.update(incelenen=len(out), riskli=int((out["Risk_Durumu"] == RISK_YUKSEK).sum()),
                yetersiz_veri=int((out["Risk_Durumu"] == RISK_YETERSIZ).sum()))
    out = out.sort_values(["Donem", "Urun_Adi", "Para_Birimi", "Tarih"]).reset_index(drop=True)
    haric = haric.sort_values(["Analiz_Disi_Nedeni", "Tarih", "Fatura_No"], kind="mergesort").reset_index(drop=True)
    return FiyatAnalizSonucu(out, haric, ozet)


def price_anomalies(lines, period_type="Aylık", threshold=15.0, kurallar=None):
    """fiyat_analizi() ile aynı; yalnızca analiz edilen satırları (Risk_Durumu sütunuyla) döndürür."""
    return fiyat_analizi(lines, period_type, threshold, kurallar).satirlar


def fiyat_ozeti_metni(ozet):
    """Fiyat analizinin özetini (kaç satır neden analiz dışı kaldı) tek satırlık metne çevirir."""
    if not ozet:
        return ""
    fmt = lambda n: f"{n:,}".replace(",", ".")  # noqa: E731
    disarida = [f"iade {fmt(ozet['iade'])}"]
    if ozet["miktar"]:
        disarida.append(f"miktar sıfır {fmt(ozet['miktar'])}")
    disarida.append(f"tevkifat {fmt(ozet['tevkifat'])}" if ozet["tevkifat_haric"] else "tevkifat (kural kapalı)")
    disarida.append(f"anahtar kelime {fmt(ozet['kelime'])}" if ozet["kelime_haric"] and ozet["kelimeler"]
                    else "anahtar kelime (kural kapalı)")
    return (f"Fiyat analizi: {fmt(ozet['incelenen'])} satır incelendi, {fmt(ozet['riskli'])} riskli | "
            f"Analiz dışı: {', '.join(disarida)} | Yetersiz veri (dönemde < {ozet['min_alim']} alım, riskli "
            f"sayılmadı): {fmt(ozet['yetersiz_veri'])}")


def fiyat_ozeti_df(ozet):
    """Fiyat analizinin özetini Excel raporu için tabloya çevirir."""
    cols = ["Kalem", "Satir_Sayisi", "Aciklama"]
    if not ozet:
        return pd.DataFrame(columns=cols)
    rows = [
        ("İncelenen", ozet["incelenen"], "Dönem + ürün + birim + para birimi bazında AOBF ile karşılaştırılan satırlar"),
        ("Riskli", ozet["riskli"], f"AOBF'den sapması ±%{ozet['esik']:g} eşiğini aşan satırlar"),
        ("Analiz dışı: iade", ozet["iade"], "İade faturası satırları"),
        ("Analiz dışı: miktar sıfır", ozet["miktar"], "Miktarı 0 ya da boş olan satırlar"),
        ("Analiz dışı: tevkifat", ozet["tevkifat"] if ozet["tevkifat_haric"] else None,
         "Tevkifatlı faturaların satırları (hizmet / hakediş)" if ozet["tevkifat_haric"] else "Kural kapalı"),
        ("Analiz dışı: anahtar kelime", ozet["kelime"] if ozet["kelime_haric"] else None,
         "Ürün adında geçen kelimeler: " + (", ".join(ozet["kelimeler"]) or "(yok)") if ozet["kelime_haric"]
         else "Kural kapalı"),
        ("Yetersiz veri (bilgi)", ozet["yetersiz_veri"],
         f"Eşik üstü sapma, ama grupta dönem içinde {ozet['min_alim']} alımdan az var; riskli sayılmadı"),
    ]
    return pd.DataFrame(rows, columns=cols)


# ---------------------------------------------------------------------- belge no eşleştirme
YONTEM_TAM = "Tam"
YONTEM_SERI_SIRA = "Seri+Sıra"
YONTEM_TUTAR_TARIH = "Tutar+Tarih"
# Tutar + tarih yedek eşleşmesinde fatura tarihi ile yevmiye tarihi arasında izin verilen en büyük fark (gün)
TARIH_PENCERESI_GUN = 15

_SERI = r"[A-ZÇĞİÖŞÜ][A-Z0-9ÇĞİÖŞÜ]{2}"
_GIB_NO = re.compile(rf"({_SERI})(\d{{4}})(\d{{9}})")
_KISA_NO = re.compile(rf"({_SERI})(\d{{1,13}})")
_SERI_TEK = re.compile(_SERI)
_YIL = re.compile(r"20\d\d")
_AYRAC = re.compile(r"[\s\-_/.\\]+")


def belge_anahtarlari(value):
    """Fatura / belge numarasından olası (seri, yıl, sıra) anahtarlarını çıkarır.

    GİB fatura numarası 3 karakter seri + 4 hane yıl + 9 hane sıra numarasıdır:
      ABC2024000000123          → [('ABC', 2024, 123)]
      ABC-2024-123, ABC 2024 123 → [('ABC', 2024, 123)]
      ABC123, ABC-123           → [('ABC', None, 123)]   (yıl yazılmamış)
      ABC2024123                → [('ABC', 2024, 123), ('ABC', None, 2024123)]   (yıl/sıra sınırı belirsiz)
    Bu biçimlere uymayan numaralarda (BORDRO-2024-07, F1 ...) boş liste döner.
    """
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return []
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    parcalar = [p for p in _AYRAC.split(str(value).upper()) if p]
    if not parcalar:
        return []
    if len(parcalar) > 1 and len(parcalar[0]) < 3:  # seri ayraçla bölünmez (GP-07001 gider pusulası vb.)
        return []
    birlesik = "".join(parcalar)
    m = _GIB_NO.fullmatch(birlesik)
    if m:
        return [(m[1], int(m[2]), int(m[3]))]
    if (len(parcalar) == 3 and _SERI_TEK.fullmatch(parcalar[0]) and _YIL.fullmatch(parcalar[1])
            and parcalar[2].isdigit() and len(parcalar[2]) <= 9):
        return [(parcalar[0], int(parcalar[1]), int(parcalar[2]))]
    m = _KISA_NO.fullmatch(birlesik)
    if not m:
        return []
    seri, rakam = m[1], m[2]
    anahtarlar = []
    if len(rakam) > 4 and _YIL.fullmatch(rakam[:4]) and len(rakam) - 4 <= 9:
        anahtarlar.append((seri, int(rakam[:4]), int(rakam[4:])))
    if len(rakam) <= 9:
        anahtarlar.append((seri, None, int(rakam)))
    return anahtarlar


def anahtar_uyumlu(a, b):
    """İki (seri, yıl, sıra) anahtarı aynı faturayı gösterebilir mi?

    Seri ve sıra aynı olmalı; yıl yalnızca iki tarafta da yazılmışsa karşılaştırılır.
    """
    return a[0] == b[0] and a[2] == b[2] and (a[1] is None or b[1] is None or a[1] == b[1])


def seri_sira_adaylari(faturalar, belgeler):
    """Her fatura için seri + sıra (ve varsa yıl) anahtarı uyuşan belgeleri bulur.

    faturalar: {fatura_anahtari: fatura_no}, belgeler: {belge_anahtari: belge_no}
    Dönüş: {fatura_anahtari: {belge_anahtari, ...}} (adayı olmayan faturalar yer almaz)
    """
    dizin = defaultdict(list)
    for bkey, no in belgeler.items():
        for seri, yil, sira in belge_anahtarlari(no):
            dizin[(seri, sira)].append((yil, bkey))
    adaylar = {}
    for fkey, no in faturalar.items():
        bulunan = set()
        for seri, yil, sira in belge_anahtarlari(no):
            for b_yil, bkey in dizin.get((seri, sira), ()):
                if yil is None or b_yil is None or yil == b_yil:
                    bulunan.add(bkey)
        if bulunan:
            adaylar[fkey] = bulunan
    return adaylar


def tutar_tarih_adaylari(faturalar, belgeler, tolerance=TOLERANCE_DEFAULT, pencere=TARIH_PENCERESI_GUN):
    """Her fatura için aynı tutarı (mutlak değer, tolerans dahilinde) taşıyan ve tarihi ±pencere gün içinde
    olan belgeleri bulur.

    faturalar / belgeler: {anahtar: (tarih 'YYYY-MM-DD', tutar)}
    Dönüş: {fatura_anahtari: {belge_anahtari, ...}}
    """
    sirali = []
    for bkey, (tarih, tutar) in belgeler.items():
        gun = _tarih(tarih)
        if gun is None or tutar is None or pd.isna(tutar) or round(abs(float(tutar)), 2) == 0:
            continue
        sirali.append((round(abs(float(tutar)), 2), gun, bkey))
    sirali.sort(key=lambda x: x[0])
    tutarlar = [x[0] for x in sirali]
    adaylar = {}
    for fkey, (tarih, tutar) in faturalar.items():
        gun = _tarih(tarih)
        if gun is None or tutar is None or pd.isna(tutar):
            continue
        hedef = round(abs(float(tutar)), 2)
        if hedef == 0:
            continue
        lo = bisect_left(tutarlar, hedef - tolerance - 0.005)
        hi = bisect_right(tutarlar, hedef + tolerance + 0.005)
        bulunan = {bkey for b_tutar, b_gun, bkey in sirali[lo:hi]
                   if round(abs(b_tutar - hedef), 2) <= tolerance and abs((b_gun - gun).days) <= pencere}
        if bulunan:
            adaylar[fkey] = bulunan
    return adaylar


def tekil_eslesmeler(adaylar):
    """Aday kümelerinden yalnızca iki yönde de tek aday olan çiftleri eşleştirir.

    Dönüş: (eşleşmeler {fatura: belge}, belirsizler {fatura: [belge, ...]})
    Bir faturanın birden fazla adayı varsa ya da adayı başka bir faturanın da adayıysa belirsiz sayılır.
    """
    ters = defaultdict(set)
    for fkey, bkeys in adaylar.items():
        for bkey in bkeys:
            ters[bkey].add(fkey)
    eslesen, belirsiz = {}, {}
    for fkey, bkeys in adaylar.items():
        if len(bkeys) == 1 and len(ters[next(iter(bkeys))]) == 1:
            eslesen[fkey] = next(iter(bkeys))
        else:
            belirsiz[fkey] = sorted(bkeys)
    return eslesen, belirsiz


def seri_sira_eslestir(faturalar, belgeler):
    """Seri + sıra eşleşmesi (2. kademe). Dönüş: (eşleşmeler, belirsizler) — bkz. tekil_eslesmeler."""
    return tekil_eslesmeler(seri_sira_adaylari(faturalar, belgeler))


def tutar_tarih_eslestir(faturalar, belgeler, tolerance=TOLERANCE_DEFAULT, pencere=TARIH_PENCERESI_GUN):
    """Tutar + tarih yedek eşleşmesi (3. kademe, düşük güven). Dönüş: (eşleşmeler, belirsizler)."""
    return tekil_eslesmeler(tutar_tarih_adaylari(faturalar, belgeler, tolerance, pencere))


def _tarih(value):
    try:
        return date.fromisoformat(str(value)[:10])
    except ValueError:
        return None


class MutabakatSonucu(OrderedDict):
    """reconcile() sonucu: OrderedDict(başlık → DataFrame) + eslesme_ozeti (kategori → fatura sayısı),
    faturasiz_ozeti (bkz. faturasiz_kayitlari_ayir), faturasiz_haric (listeye alınmayan belgeler), kur_ozeti
    (bkz. kur_farki_ayir) ve kur_farki (yüzde kur toleransı içinde kaldığı için tutar farkı sayılmayan dövizli
    faturalar, bilgi)."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.eslesme_ozeti = OrderedDict()
        self.faturasiz_ozeti = {}
        self.faturasiz_haric = pd.DataFrame(columns=HARIC_COLS)
        self.kur_ozeti = {}
        self.kur_farki = pd.DataFrame(columns=KUR_FARKI_COLS)
        self.vergi_ozeti = {}


_OZET_ACIKLAMA = OrderedDict([
    (YONTEM_TAM, "Fatura no ile belge no birebir aynı"),
    (YONTEM_SERI_SIRA, "Seri + sıra numarası (ve varsa yıl) aynı; belge no kısaltılarak yazılmış"),
    (YONTEM_TUTAR_TARIH,
     f"Belge no uyuşmuyor; tek aday aynı tutar ve ±{TARIH_PENCERESI_GUN} gün (düşük güven)"),
    ("Seçili hesap dışı", "Belge yevmiyede var ama seçili hesaplarda değil"),
    ("Belirsiz", "Birden fazla aday belge ya da aynı no farklı tedarikçi"),
    ("Eşleşmeyen", "Muhasebeleşmemiş"),
])


def eslesme_ozeti_metni(ozet):
    """Eşleştirme özetini tek satırlık metne çevirir."""
    if not ozet:
        return ""
    toplam = sum(ozet.values())
    parcalar = [f"{k}: {v:,}".replace(",", ".") for k, v in ozet.items()]
    return f"Eşleştirme ({toplam:,} fatura) — ".replace(",", ".") + " | ".join(parcalar)


def eslesme_ozeti_df(ozet):
    """Eşleştirme özetini Excel raporu için tabloya çevirir."""
    return pd.DataFrame([{"Eslesme_Yontemi": k, "Fatura_Sayisi": v, "Aciklama": _OZET_ACIKLAMA.get(k, "")}
                         for k, v in (ozet or {}).items()], columns=["Eslesme_Yontemi", "Fatura_Sayisi", "Aciklama"])

# ---------------------------------------------------------------------- faturasız yevmiye kayıtları
# Faturaya dayanmayan olağan kayıtların yaygın belge no önekleri (firma başına değiştirilebilir)
VARSAYILAN_HARIC_ONEKLER = "BORDRO, AMORT, MAHSUP, AÇILIŞ, KAPANIŞ, DEVİR, GP"
ONCELIK_YUKSEK = "Yüksek"
ONCELIK_DUSUK = "Düşük"
NEDEN_ALACAK = "Alacak yönlü"
NEDEN_BORC = "Borç yönlü"
NEDEN_ONEK = "Hariç önek"
FATURASIZ_COLS = ["Oncelik", "Yevmiye_Belge_No", "Yevmiye_Tarihi", "Yevmiye_Tutari", "Hesaplar"]
HARIC_COLS = ["Yevmiye_Belge_No", "Yevmiye_Tarihi", "Yevmiye_Tutari", "Hesaplar", "Haric_Tutulma_Nedeni"]


def _onek_anahtari(value):
    """Önek karşılaştırma anahtarı: büyük/küçük harf, Türkçe karakter ve ayraç (boşluk, -, /, ., _) duyarsız."""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    return _AYRAC.sub("", tr_ascii_upper(str(value).strip()))


def parse_onek_listesi(text):
    """'BORDRO, Amort; açılış' → ['BORDRO', 'AMORT', 'ACILIS'] (karşılaştırma anahtarları, tekrarsız)."""
    out = []
    for parca in re.split(r"[,;\n]+", str(text or "")):
        anahtar = _onek_anahtari(parca)
        if anahtar and anahtar not in out:
            out.append(anahtar)
    return out


def gib_bicimli(belge_no):
    """Belge no tam GİB fatura numarası biçiminde mi? (3 karakter seri + 4 hane yıl + 9 hane sıra)"""
    return bool(_GIB_NO.fullmatch(_onek_anahtari(belge_no)))


def haric_onek(belge_no, onekler):
    """Belge no hariç tutulan öneklerden biriyle başlıyorsa o öneki, değilse None döndürür.

    onekler: parse_onek_listesi() çıktısı. Tam GİB biçimli numaralar (ör. GPS2024000000001) gerçek fatura
    numarası olduğundan hiçbir önekle hariç tutulmaz.
    """
    anahtar = _onek_anahtari(belge_no)
    if not anahtar or gib_bicimli(belge_no):
        return None
    return next((p for p in onekler if anahtar.startswith(p)), None)


def belge_onceligi(belge_no):
    """GİB fatura numarasına benzeyen (tam biçim ya da seri+sıra çözümlemesine uyan) belge no → Yüksek."""
    return ONCELIK_YUKSEK if belge_anahtarlari(belge_no) else ONCELIK_DUSUK


def yevmiye_isaretli(journal):
    """Yevmiye tutarları işaretli mi (Borç − Alacak)? Hiç negatif tutar yoksa tek 'Tutar' sütunuyla işaretsiz
    yüklenmiş kabul edilir; çift taraflı bir yevmiyede alacak satırları negatif olur."""
    if journal is None or journal.empty:
        return True
    return bool((pd.to_numeric(journal["amount"], errors="coerce").fillna(0) < 0).any())


def faturasiz_kayitlari_ayir(adaylar, onekler=None, tolerance=TOLERANCE_DEFAULT, isaretli=True, yon="borc"):
    """Hiçbir faturayla eşleşmeyen seçili hesap belgelerinden faturasız gider/alış adaylarını ayırır.

    adaylar: Yevmiye_Belge_No, Yevmiye_Tarihi, Yevmiye_Tutari (seçili hesaplarda Borç − Alacak), Hesaplar
    onekler: parse_onek_listesi() çıktısı; None → VARSAYILAN_HARIC_ONEKLER, [] → önek filtresi yok
    isaretli: False ise (işaretsiz Tutar) yön bilinmediğinden borç/alacak filtresi uygulanmaz.
    yon: "borc" (alış: gider/maliyet borç yönlü) ya da "alacak" (satış: gelir hesabında alacak yönlü belgeler).
    Sırayla:
      1. Alacak yönlü: net tutar borç yönünde tolerans üstünde değil (ör. 153 alacak — SMM, stoktan çıkış)
      2. Hariç önek: belge no hariç tutulan öneklerden biriyle başlıyor (bordro, amortisman, mahsup ...)
    Kalanlara Oncelik (Yüksek: GİB fatura no biçimi, Düşük: diğer) verilir; önceliğe ve tutara göre sıralanır.
    Dönüş: (listelenen DataFrame, hariç tutulan DataFrame, özet dict)
    """
    if onekler is None:
        onekler = parse_onek_listesi(VARSAYILAN_HARIC_ONEKLER)
    df = adaylar.reset_index(drop=True).copy()
    tutar = pd.to_numeric(df["Yevmiye_Tutari"], errors="coerce").fillna(0.0)
    alacak = yon == "alacak"
    neden = pd.Series([None] * len(df), index=df.index, dtype=object)
    ters = NEDEN_BORC if alacak else NEDEN_ALACAK
    if isaretli:
        neden[(-tutar if alacak else tutar) <= tolerance] = ters
    for idx in df.index[neden.isna()]:
        p = haric_onek(df.at[idx, "Yevmiye_Belge_No"], onekler)
        if p:
            neden[idx] = f"{NEDEN_ONEK} ({p})"

    liste = df[neden.isna()].copy()
    liste["Oncelik"] = liste["Yevmiye_Belge_No"].map(belge_onceligi)
    liste["_sira"] = (liste["Oncelik"] != ONCELIK_YUKSEK).astype(int)
    liste["_tutar"] = pd.to_numeric(liste["Yevmiye_Tutari"], errors="coerce").fillna(0.0).abs()
    liste = liste.sort_values(["_sira", "_tutar", "Yevmiye_Belge_No"], ascending=[True, False, True],
                              kind="mergesort")[FATURASIZ_COLS].reset_index(drop=True)

    haric = df[neden.notna()].copy()
    haric["Haric_Tutulma_Nedeni"] = neden[neden.notna()]
    haric = haric[HARIC_COLS].reset_index(drop=True)

    ozet = {
        "listelenen": len(liste),
        "yuksek": int((liste["Oncelik"] == ONCELIK_YUKSEK).sum()),
        "dusuk": int((liste["Oncelik"] == ONCELIK_DUSUK).sum()),
        "alacak_yonlu": int((haric["Haric_Tutulma_Nedeni"] == NEDEN_ALACAK).sum()),
        "borc_yonlu": int((haric["Haric_Tutulma_Nedeni"] == NEDEN_BORC).sum()),
        "yon": "alacak" if alacak else "borc",
        "haric_onek": int(haric["Haric_Tutulma_Nedeni"].str.startswith(NEDEN_ONEK).sum()),
        "isaretli": bool(isaretli),
        "onekler": list(onekler),
    }
    return liste, haric, ozet


def faturasiz_ozeti_metni(ozet):
    """Faturasız kayıt filtresinin özetini tek satırlık metne çevirir."""
    if not ozet:
        return ""
    fmt = lambda n: f"{n:,}".replace(",", ".")  # noqa: E731
    satis = ozet.get("yon") == "alacak"
    metin = (f"{'Faturasız gelir kaydı' if satis else 'Faturasız kayıt'}: {fmt(ozet['listelenen'])} listelendi "
             f"(Yüksek öncelik: {fmt(ozet['yuksek'])}, Düşük: {fmt(ozet['dusuk'])}) | Listeye alınmayan: ")
    if ozet["isaretli"]:
        metin += (f"borç yönlü {fmt(ozet.get('borc_yonlu', 0))}, " if satis
                  else f"alacak yönlü {fmt(ozet['alacak_yonlu'])}, ")
    metin += f"hariç önek {fmt(ozet['haric_onek'])}"
    if not ozet["isaretli"]:
        metin += (" | Yevmiye tutarları işaretsiz (negatif tutar yok): borç/alacak yönü bilinmediği için "
                  "yön filtresi uygulanmadı")
    return metin


def faturasiz_ozeti_df(ozet):
    """Faturasız kayıt filtresinin özetini Excel raporu için tabloya çevirir."""
    cols = ["Kalem", "Kayit_Sayisi", "Aciklama"]
    if not ozet:
        return pd.DataFrame(columns=cols)
    satis = ozet.get("yon") == "alacak"
    if not ozet["isaretli"]:
        yon = "Uygulanmadı: yevmiye tutarları işaretsiz (negatif tutar yok), yön bilinmiyor"
    elif satis:
        yon = "Gelir hesaplarındaki net tutar (Borç − Alacak) alacak yönünde değil (ör. satıştan iade, düzeltme)"
    else:
        yon = "Seçili hesaplardaki net tutar (Borç − Alacak) borç yönünde değil (ör. stoktan çıkış, SMM)"
    rows = [
        ("Listelenen (Yüksek öncelik)", ozet["yuksek"], "Belge no GİB fatura numarası biçiminde"),
        ("Listelenen (Düşük öncelik)", ozet["dusuk"], "Belge no fatura numarasına benzemiyor"),
        ("Hariç: borç yönlü" if satis else "Hariç: alacak yönlü",
         (ozet.get("borc_yonlu", 0) if satis else ozet["alacak_yonlu"]) if ozet["isaretli"] else None, yon),
        ("Hariç: önek", ozet["haric_onek"],
         "Hariç tutulan önekler: " + (", ".join(ozet["onekler"]) or "(yok)")),
    ]
    return pd.DataFrame(rows, columns=cols)


# ---------------------------------------------------------------------- kur farkı toleransı
TUTAR_FARKI_COLS = ["Fatura_No", "Fatura_Tarihi", "Tedarikci", "Tedarikci_VKN", "KDV_Haric_Tutar_TL",
                    "Maliyete_Eklenen_Vergi_TL", "Beklenen_Tutar_TL", "Yevmiye_Tutari_TL", "Fark_TL", "Olasi_Neden",
                    "Dovizli", "Para_Birimi", "Kur", "Hesaplar", "Yevmiye_Belge_No", "Eslesme_Yontemi"]
NEDEN_VERGI_EKSIK = "Maliyete eklenecek vergi (ÖTV vb.) maliyete eklenmemiş"
NEDEN_KDV_MALIYETTE = "KDV maliyete eklenmiş"
NEDEN_CIFT_KAYIT = "Çift kayıt"
KUR_FARKI_COLS = ["Fatura_No", "Fatura_Tarihi", "Tedarikci", "Tedarikci_VKN", "Para_Birimi", "Kur",
                  "KDV_Haric_Tutar_TL", "Yevmiye_Tutari_TL", "Fark_TL", "Fark_Yuzdesi", "Yevmiye_Belge_No"]


def izin_verilen_fark(tutar_tl, dovizli, tolerance=TOLERANCE_DEFAULT, kur_toleransi=KUR_TOLERANSI_VARSAYILAN):
    """Tutar farkı sayılmayacak en büyük fark (TL).

    TL faturada sabit TL toleransı; dövizli faturada max(TL toleransı, kur_toleransi % × TL tutar).
    """
    if not dovizli or not kur_toleransi:
        return tolerance
    return max(tolerance, abs(float(tutar_tl or 0)) * float(kur_toleransi) / 100)


def _yakin(a, b, sinir):
    return abs(float(a) - float(b)) <= sinir + 0.005


def tutar_farki_nedeni(r, tolerance=TOLERANCE_DEFAULT, kur_toleransi=KUR_TOLERANSI_VARSAYILAN):
    """Tutar farkının olası nedeni (bilgi): maliyete eklenecek vergi eksik, KDV maliyette ya da çift kayıt.

    r: net_tl, beklenen_tl, maliyet_vergisi_tl, kdv_tl, Yevmiye_Tutari, Fark, Para_Birimi alanları
    """
    sinir = izin_verilen_fark(r["beklenen_tl"], r.get("Para_Birimi", "TRY") != "TRY", tolerance, kur_toleransi)
    fark, mv, kdv = float(r["Fark"]), float(r.get("maliyet_vergisi_tl") or 0), float(r.get("kdv_tl") or 0)
    if mv > sinir and _yakin(fark, -mv, sinir):
        return NEDEN_VERGI_EKSIK
    if kdv > sinir and _yakin(fark, kdv, sinir):
        return NEDEN_KDV_MALIYETTE
    if abs(float(r["beklenen_tl"])) > sinir and _yakin(abs(float(r["Yevmiye_Tutari"])), 2 * abs(float(r["beklenen_tl"])),
                                                     2 * sinir):
        return NEDEN_CIFT_KAYIT
    return ""


def kur_farki_ayir(eslesen, tolerance=TOLERANCE_DEFAULT, kur_toleransi=KUR_TOLERANSI_VARSAYILAN):
    """Eşleşen faturaları (Fark sütunu TL tolerans üstü olanlar) tutar farkı ve tolerans içi kur farkı olarak ayırır.

    eslesen: reconcile() içindeki eşleşen faturalar (beklenen_tl ya da net_tl, Fark, Para_Birimi sütunlarıyla)
    Dönüş: (tutar farkı satırları, kur farkı (tolerans içi) satırları, özet dict)
    """
    fark = eslesen[eslesen["Fark"].abs() > tolerance]
    dovizli = fark["Para_Birimi"] != "TRY"
    tutar = fark["beklenen_tl"] if "beklenen_tl" in fark else fark["net_tl"]
    sinir = [izin_verilen_fark(t, d, tolerance, kur_toleransi) for t, d in zip(tutar, dovizli)]
    icinde = dovizli & (fark["Fark"].abs() <= pd.Series(sinir, index=fark.index, dtype=float) + 0.005)
    ozet = {"kur_toleransi": float(kur_toleransi or 0), "dovizli_fatura": int((eslesen["Para_Birimi"] != "TRY").sum()),
            "tolerans_ici": int(icinde.sum()), "asan": int((dovizli & ~icinde).sum())}
    return fark[~icinde], fark[icinde], ozet


def kur_ozeti_metni(ozet):
    """Kur farkı toleransının özetini tek satırlık metne çevirir."""
    if not ozet:
        return ""
    if not ozet["kur_toleransi"]:
        return f"Kur farkı toleransı kapalı (%0): {ozet['asan']} dövizli fatura tutar farklarında"
    return (f"Kur farkı (tolerans içi, ±%{ozet['kur_toleransi']:g}, bilgi): {ozet['tolerans_ici']} dövizli fatura "
            f"tutar farkı sayılmadı | Toleransı aşan dövizli fark: {ozet['asan']} "
            f"(dövizli eşleşen fatura: {ozet['dovizli_fatura']})")


# ---------------------------------------------------------------------- KDV ve tevkifat mutabakatı
KDV_FARKI_COLS = ["Fatura_No", "Fatura_Tarihi", "Tedarikci", "Tedarikci_VKN", "KDV_Tutari_TL", "Tevkifat_Tutari_TL",
                  "Yevmiye_KDV_TL", "Fark_TL", "Olasi_Neden", "Dovizli", "Para_Birimi", "Kur", "Yevmiye_Belge_No",
                  "Eslesme_Yontemi"]
KDV_YOK_COLS = ["Fatura_No", "Fatura_Tarihi", "Tedarikci", "Tedarikci_VKN", "KDV_Tutari_TL", "Olasi_Neden",
                "Kullanilan_Hesaplar", "Yevmiye_Belge_No", "Eslesme_Yontemi"]
TEVKIFAT_COLS = ["Fatura_No", "Fatura_Tarihi", "Tedarikci", "Tedarikci_VKN", "KDV_Tutari_TL", "Tevkifat_Orani",
                 "Tevkifat_Tutari_TL", "Yevmiye_Tevkifat_TL", "Fark_TL", "Durum", "Yevmiye_Belge_No",
                 "Eslesme_Yontemi"]
NEDEN_KDV_YOK = "KDV hesabında kayıt yok"
NEDEN_KDV_FATURADA_YOK = "Faturada KDV yok"
NEDEN_KDV_TEVKIFAT_DUSULMUS = "191'e tevkifat düşülmüş KDV yazılmış (tam KDV yazılmalı)"
NEDEN_TERS_YON = "Ters yönde kayıt"
TEVKIFAT_YOK = "Tevkifat kaydı yok"
TEVKIFAT_TERS = "Borç yönlü (ters) kayıt"
TEVKIFAT_FARKLI = "Tutar farklı"
VARSAYILAN_KDV_HESAPLARI = "191"
VARSAYILAN_TEVKIFAT_HESAPLARI = "360"


def _hesap_maskesi(jou, hesaplar):
    return jou["account_norm"].map(lambda a: any(a.startswith(p) for p in hesaplar)).astype(bool)


def _belge_toplamlari(jou, hesaplar):
    """Hesap önekleriyle eşleşen yevmiye satırlarının belge bazında toplamı ve satır sayısı."""
    if not hesaplar or jou.empty:
        return {}
    g = jou[_hesap_maskesi(jou, hesaplar)].groupby("document_no_norm")["amount"].agg(["sum", "count"])
    return {d: (float(r["sum"]), int(r["count"])) for d, r in g.iterrows()}


def _topla(toplamlar, belgeler):
    tutar = sum(toplamlar.get(d, (0.0, 0))[0] for d in belgeler)
    adet = sum(toplamlar.get(d, (0.0, 0))[1] for d in belgeler)
    return round(tutar, 2), adet


def alis_vergi_bulgulari(inv, belgeler, yontemler, jou, kdv_hesaplari, tevkifat_hesaplari,
                         tolerance=TOLERANCE_DEFAULT, kur_toleransi=KUR_TOLERANSI_VARSAYILAN, isaretli=True,
                         maliyet_farki=None):
    """Alış faturalarının KDV'sini (191) ve tevkifatını (360) yevmiyedeki aynı belgeyle karşılaştırır.

    inv: reconcile() içindeki fatura tablosu (kdv_tl, tevkifat_tl, Para_Birimi, Kur, invoice_type ...)
    belgeler: {fatura index: [document_no_norm, ...]} (madde 2 eşleştirmesinin bulduğu belgeler)
    maliyet_farki: {fatura index: seçili hesaplardaki tutar farkı} (KDV maliyete eklenmiş tespiti için)
    Tevkifatlı faturada 191'e TAM KDV yazılır (indirilecek KDV); tevkif edilen kısım 360'a ALACAK yazılır.
    İade faturalarında yönler terstir. Dönüş: (KDV farkları, KDV'si kaydedilmemiş, tevkifat bulguları, özet)
    """
    kdv_top = _belge_toplamlari(jou, kdv_hesaplari)
    tev_top = _belge_toplamlari(jou, tevkifat_hesaplari)
    belge_adi = jou.groupby("document_no_norm")["document_no"].first().to_dict() if not jou.empty else {}
    hesap_adi = jou.groupby("document_no_norm")["account_code"].agg(lambda s: sorted(set(map(str, s)))).to_dict() \
        if not jou.empty else {}
    maliyet_farki = maliyet_farki or {}
    fark_rows, yok_rows, tev_rows = [], [], []
    ozet = {"kdv_karsilastirilan": 0, "kdv_farki": 0, "kdv_yok": 0, "tevkifatli": 0, "tevkifat_bulgu": 0,
            "kdv_hesaplari": list(kdv_hesaplari or []), "tevkifat_hesaplari": list(tevkifat_hesaplari or [])}
    for idx, docs in belgeler.items():
        r = inv.loc[idx]
        dovizli = r["Para_Birimi"] != "TRY"
        isaret = -1 if iade_faturasi(r.get("invoice_type")) else 1  # 191 borç yönlü (iadede alacak)
        temel = {"Fatura_No": r["invoice_no"], "Fatura_Tarihi": r["issue_date"], "Tedarikci": r["supplier_name"],
                 "Tedarikci_VKN": r["supplier_vkn"], "KDV_Tutari_TL": r["kdv_tl"],
                 "Yevmiye_Belge_No": ", ".join(str(belge_adi.get(d, d)) for d in docs),
                 "Eslesme_Yontemi": yontemler.get(idx, "")}
        kdv, tev = float(r["kdv_tl"]), float(r["tevkifat_tl"])
        if kdv_hesaplari:
            ozet["kdv_karsilastirilan"] += 1
            tutar, adet = _topla(kdv_top, docs)
            sinir = izin_verilen_fark(kdv, dovizli, tolerance, kur_toleransi)
            if adet == 0:
                if kdv > sinir:
                    mf = maliyet_farki.get(idx)
                    neden = NEDEN_KDV_MALIYETTE if mf is not None and _yakin(mf, kdv, sinir) else NEDEN_KDV_YOK
                    yok_rows.append(dict(temel, Olasi_Neden=neden, Kullanilan_Hesaplar=", ".join(
                        sorted(set().union(*(hesap_adi.get(d, []) for d in docs))))))
            else:
                fark = round(abs(tutar) - kdv, 2)
                ters = isaretli and kdv > sinir and tutar * isaret < 0
                if abs(fark) > sinir or ters:
                    if ters:
                        neden = NEDEN_TERS_YON
                    elif kdv <= tolerance:
                        neden = NEDEN_KDV_FATURADA_YOK
                    elif _yakin(abs(tutar), 2 * kdv, 2 * sinir):
                        neden = NEDEN_CIFT_KAYIT
                    elif tev > sinir and _yakin(abs(tutar), kdv - tev, sinir):
                        neden = NEDEN_KDV_TEVKIFAT_DUSULMUS
                    else:
                        neden = ""
                    fark_rows.append(dict(temel, Tevkifat_Tutari_TL=tev, Yevmiye_KDV_TL=tutar, Fark_TL=fark,
                                          Olasi_Neden=neden, Dovizli="Evet" if dovizli else "Hayır",
                                          Para_Birimi=r["Para_Birimi"], Kur=r["Kur"]))
        if tevkifat_hesaplari and tev > tolerance:
            ozet["tevkifatli"] += 1
            tutar, adet = _topla(tev_top, docs)
            sinir = izin_verilen_fark(tev, dovizli, tolerance, kur_toleransi)
            beklenen = -isaret * tev  # 360 alacak yönlü (iadede borç)
            if adet == 0:
                durum, fark = TEVKIFAT_YOK, round(-tev, 2)
            else:
                fark = round(abs(tutar) - tev, 2)
                if isaretli and tutar * beklenen < 0:
                    durum = TEVKIFAT_TERS
                elif abs(fark) > sinir:
                    durum = TEVKIFAT_FARKLI
                else:
                    continue
            oran = r.get("withholding_rate")
            tev_rows.append(dict(temel, Tevkifat_Orani=None if oran is None or pd.isna(oran) else float(oran),
                                 Tevkifat_Tutari_TL=tev, Yevmiye_Tevkifat_TL=tutar, Fark_TL=fark, Durum=durum))
    ozet.update(kdv_farki=len(fark_rows), kdv_yok=len(yok_rows), tevkifat_bulgu=len(tev_rows))
    return (pd.DataFrame(fark_rows, columns=KDV_FARKI_COLS), pd.DataFrame(yok_rows, columns=KDV_YOK_COLS),
            pd.DataFrame(tev_rows, columns=TEVKIFAT_COLS), ozet)


SATIS_KDV_COLS = ["Fatura_No", "Fatura_Tarihi", "Musteri", "Musteri_VKN", "Fatura_Tipi", "KDV_Istisna",
                  "Fatura_KDV_TL", "Beklenen_KDV_TL", "Yevmiye_KDV_TL", "Fark_TL", "Durum", "Dovizli", "Para_Birimi",
                  "Kur", "Yevmiye_Belge_No", "Eslesme_Yontemi"]
SATIS_KDV_YOK = "Hesaplanan KDV kaydı yok"
SATIS_KDV_ISTISNA = "İstisna / ihracat faturasında KDV kaydı var"
SATIS_KDV_FATURADA_YOK = "Faturada KDV yok, KDV kaydı var"
SATIS_KDV_TERS = "Borç yönlü (ters) kayıt"
SATIS_KDV_FARKLI = "Tutar farklı"
VARSAYILAN_GELIR_HESAPLARI = "600, 601, 602"
VARSAYILAN_SATIS_KDV_HESAPLARI = "391"


def satis_kdv_bulgulari(inv, belgeler, yontemler, jou, kdv_hesaplari, tolerance=TOLERANCE_DEFAULT,
                        kur_toleransi=KUR_TOLERANSI_VARSAYILAN, isaretli=True):
    """Satış faturalarının KDV'sini hesaplanan KDV hesaplarındaki (varsayılan 391) ALACAK kaydıyla karşılaştırır.

    Beklenen KDV: ihracat / istisna faturasında (bkz. kdv_istisna) 0; tevkifatlı satışta faturadaki KDV − alıcının
    tevkif ettiği KDV (satıcı yalnızca tahsil ettiği kısmı 391'e yazar); diğerlerinde faturadaki KDV.
    Dönüş: (bulgular DataFrame, özet dict)
    """
    top = _belge_toplamlari(jou, kdv_hesaplari)
    belge_adi = jou.groupby("document_no_norm")["document_no"].first().to_dict() if not jou.empty else {}
    rows = []
    ozet = {"kdv_karsilastirilan": 0, "kdv_farki": 0, "istisna": 0, "kdv_hesaplari": list(kdv_hesaplari or [])}
    for idx, docs in belgeler.items():
        r = inv.loc[idx]
        dovizli = r["Para_Birimi"] != "TRY"
        istisna = kdv_istisna(r.get("profile"), r.get("invoice_type"))
        ozet["kdv_karsilastirilan"] += 1
        ozet["istisna"] += int(istisna)
        beklenen = 0.0 if istisna else round(float(r["kdv_tl"]) - float(r["tevkifat_tl"]), 2)
        isaret = 1 if iade_faturasi(r.get("invoice_type")) else -1  # 391 alacak yönlü
        tutar, adet = _topla(top, docs)
        sinir = izin_verilen_fark(beklenen, dovizli, tolerance, kur_toleransi)
        fark = round(abs(tutar) - beklenen, 2)
        if beklenen > sinir and adet == 0:
            durum = SATIS_KDV_YOK
        elif beklenen <= tolerance and abs(tutar) > sinir:
            durum = SATIS_KDV_ISTISNA if istisna else SATIS_KDV_FATURADA_YOK
        elif isaretli and beklenen > sinir and tutar * isaret < 0:
            durum = SATIS_KDV_TERS
        elif abs(fark) > sinir:
            durum = SATIS_KDV_FARKLI
        else:
            continue
        rows.append({"Fatura_No": r["invoice_no"], "Fatura_Tarihi": r["issue_date"], "Musteri": r["supplier_name"],
                     "Musteri_VKN": r["supplier_vkn"], "Fatura_Tipi": r.get("invoice_type"),
                     "KDV_Istisna": "Evet" if istisna else "Hayır", "Fatura_KDV_TL": r["kdv_tl"],
                     "Beklenen_KDV_TL": beklenen, "Yevmiye_KDV_TL": tutar, "Fark_TL": fark, "Durum": durum,
                     "Dovizli": "Evet" if dovizli else "Hayır", "Para_Birimi": r["Para_Birimi"], "Kur": r["Kur"],
                     "Yevmiye_Belge_No": ", ".join(str(belge_adi.get(d, d)) for d in docs),
                     "Eslesme_Yontemi": yontemler.get(idx, "")})
    ozet["kdv_farki"] = len(rows)
    return pd.DataFrame(rows, columns=SATIS_KDV_COLS), ozet


def satis_kdv_ozeti_metni(ozet):
    if not ozet:
        return ""
    if not ozet["kdv_hesaplari"]:
        return "Satış KDV kontrolü kapalı (hesap kodu yok)"
    return (f"Satış KDV ({', '.join(ozet['kdv_hesaplari'])}): {ozet['kdv_karsilastirilan']} fatura karşılaştırıldı "
            f"(istisna / ihracat {ozet['istisna']}, KDV beklenmez), fark {ozet['kdv_farki']}")


def vergi_ozeti_metni(ozet):
    """KDV / tevkifat mutabakatı özetini tek satırlık metne çevirir."""
    if not ozet:
        return ""
    parcalar = []
    if ozet["kdv_hesaplari"]:
        parcalar.append(f"KDV ({', '.join(ozet['kdv_hesaplari'])}): {ozet['kdv_karsilastirilan']} fatura "
                        f"karşılaştırıldı, KDV farkı {ozet['kdv_farki']}, KDV'si kaydedilmemiş {ozet['kdv_yok']}")
    else:
        parcalar.append("KDV kontrolü kapalı (hesap kodu yok)")
    if ozet["tevkifat_hesaplari"]:
        parcalar.append(f"Tevkifat ({', '.join(ozet['tevkifat_hesaplari'])}): {ozet['tevkifatli']} tevkifatlı fatura, "
                        f"eksik/farklı {ozet['tevkifat_bulgu']}")
    else:
        parcalar.append("Tevkifat kontrolü kapalı (hesap kodu yok)")
    return "Vergi mutabakatı — " + " | ".join(parcalar)


# ---------------------------------------------------------------------- mutabakat
def reconcile(invoices, journal, accounts, tolerance=TOLERANCE_DEFAULT, period_type="Aylık", haric_onekler=None,
              kur_toleransi=KUR_TOLERANSI_VARSAYILAN, kdv_hesaplari=None, tevkifat_hesaplari=None, yon=YON_ALIS):
    """Faturaları (KDV hariç, TL) seçili hesap kodlarındaki yevmiye kayıtlarıyla karşılaştırır.

    yon: YON_ALIS (varsayılan; maliyet / stok hesapları, borç yönlü) ya da YON_SATIS (bkz. satis_mutabakati:
    gelir hesapları, alacak yönlü; bölüm başlıkları SATIS_BASLIKLARI, karşı taraf Musteri / Musteri_VKN).

    accounts: normalize edilmiş hesap kodu önekleri listesi (ör. ['153', '770']).
    Fatura no ↔ yevmiye belge no kademeli eşleştirilir:
      1. Tam: normalize edilmiş Fatura_No = Belge_No
      2. Seri+Sıra: GİB numarasının seri + sıra (+ varsa yıl) anahtarı aynı, tek aday
      3. Tutar+Tarih: belge no uyuşmayan, aynı tutar ve ±TARIH_PENCERESI_GUN gün içinde tek aday (düşük güven)
    Seçili hesap dışı kontrolü 3. kademeden önce (tam + seri/sıra ile) yapılır.
    Eşleşmeyen belgeler faturasiz_kayitlari_ayir() ile süzülür (haric_onekler: parse_onek_listesi() çıktısı,
    None → varsayılan önekler).
    Tutar farkı: TL faturada fark > tolerance (TL); dövizli faturada fark > max(tolerance, kur_toleransi % × TL tutar).
    Dövizli faturalarda bu yüzde içinde kalan farklar tutar farkı sayılmaz, .kur_farki / .kur_ozeti'de bilgi olarak
    raporlanır (kur_toleransi=0 → yüzde tolerans yok).
    kdv_hesaplari / tevkifat_hesaplari verilirse (ör. ['191'] / ['360']) aynı eşleştirmeyle bulunan belgelerde KDV ve
    tevkifat kayıtları da karşılaştırılır (bkz. alis_vergi_bulgulari): "KDV Farkları", "KDV'si Kaydedilmemiş
    Faturalar", "Tevkifat Kaydı Eksik/Farklı" bölümleri ve .vergi_ozeti.
    Dönüş: MutabakatSonucu (OrderedDict başlık → DataFrame, .eslesme_ozeti, .faturasiz_ozeti, .faturasiz_haric)
    """
    res = MutabakatSonucu()
    if not accounts:
        raise ValueError("En az bir hesap kodu girilmelidir")

    inv = invoices.copy()
    if inv.empty:
        inv = pd.DataFrame(columns=["invoice_no", "invoice_no_norm", "issue_date", "supplier_vkn", "supplier_name",
                                    "total_amount", "exchange_rate", "invoice_type", "currency"])
    if "currency" not in inv:
        inv["currency"] = "TRY"
    inv["Para_Birimi"] = inv["currency"].map(para_birimi).astype(object)
    inv["Kur"] = inv["exchange_rate"].fillna(1.0).astype(float)
    inv["net_tl"] = (inv["total_amount"].astype(float) * inv["exchange_rate"].fillna(1.0).astype(float)).round(2)
    # Beklenen maliyet = KDV hariç tutar + maliyete eklenen vergiler (ÖTV vb.; bkz. maliyet_vergisi_ekle)
    mv = inv["maliyet_vergisi"] if "maliyet_vergisi" in inv else pd.Series(0.0, index=inv.index)
    inv["maliyet_vergisi_tl"] = (pd.to_numeric(mv, errors="coerce").fillna(0.0) * inv["Kur"]).round(2)
    inv["beklenen_tl"] = (inv["net_tl"] + inv["maliyet_vergisi_tl"]).round(2)
    if "vat_amount" not in inv:
        inv["vat_amount"] = 0.0
    inv["kdv_tl"] = (pd.to_numeric(inv["vat_amount"], errors="coerce").fillna(0.0) * inv["Kur"]).round(2)
    wh = inv["withholding_amount"] if "withholding_amount" in inv else pd.Series(0.0, index=inv.index)
    inv["tevkifat_tl"] = (pd.to_numeric(wh, errors="coerce").fillna(0.0) * inv["Kur"]).round(2)
    if "invoice_type" not in inv:
        inv["invoice_type"] = "SATIS"
    satis = yon == YON_SATIS
    if satis:  # Karşı taraf müşteridir; satışta maliyete eklenen vergi yoktur
        for kaynak, hedef in (("customer_name", "supplier_name"), ("customer_vkn", "supplier_vkn")):
            inv[hedef] = inv[kaynak] if kaynak in inv else ""
        inv["maliyet_vergisi_tl"] = 0.0
        inv["beklenen_tl"] = inv["net_tl"]

    jou = journal.copy()
    if jou.empty:
        jou = pd.DataFrame(columns=["entry_date", "document_no", "document_no_norm", "account_code", "account_norm",
                                    "amount"])
    jou["account_norm"] = jou["account_norm"].fillna("").astype(str)
    jou["selected"] = jou["account_norm"].map(lambda a: any(a.startswith(p) for p in accounts)).astype(bool)
    sel = jou[jou["selected"]]

    sel_grp = sel.groupby("document_no_norm").agg(
        Yevmiye_Belge_No=("document_no", "first"), Yevmiye_Tarihi=("entry_date", "min"),
        Yevmiye_Tutari=("amount", "sum"), Hesaplar=("account_code", lambda s: ", ".join(sorted(set(map(str, s))))),
    ).reset_index()
    sel_grp["Yevmiye_Tutari"] = sel_grp["Yevmiye_Tutari"].astype(float).round(2)
    sel_docs = sel_grp.set_index("document_no_norm", drop=False)
    other = jou[~jou["document_no_norm"].isin(sel_docs.index)]
    other_docs = {d: (str(g["document_no"].iloc[0]), set(map(str, g["account_code"])))
                  for d, g in other.groupby("document_no_norm")}

    # Aynı fatura numarası birden fazla tedarikçide → eşleşme belirsiz
    dup_mask = inv.duplicated("invoice_no_norm", keep=False)
    ambiguous = inv[dup_mask]
    clear = inv[~dup_mask]
    invoice_nos = set(inv["invoice_no_norm"])

    base_cols = {"invoice_no": "Fatura_No", "issue_date": "Fatura_Tarihi", "supplier_name": "Tedarikci",
                 "supplier_vkn": "Tedarikci_VKN", "net_tl": "KDV_Haric_Tutar_TL"}

    # 1) Tam eşleşme
    eslesme = {}  # fatura index → (document_no_norm, yöntem)
    for idx, no in clear["invoice_no_norm"].items():
        if no in sel_docs.index:
            eslesme[idx] = (no, YONTEM_TAM)

    # 2) Seri + sıra eşleşmesi (fatura no'su ile birebir aynı olmayan seçili hesap belgeleri arasında)
    acik_belgeler = [d for d in sel_docs.index if d not in invoice_nos]
    kalan = {idx: clear.at[idx, "invoice_no"] for idx in clear.index if idx not in eslesme}
    tekil, belirsiz = seri_sira_eslestir(kalan, {d: sel_docs.at[d, "Yevmiye_Belge_No"] for d in acik_belgeler})
    for idx, d in tekil.items():
        eslesme[idx] = (d, YONTEM_SERI_SIRA)
    belirsiz_belgeler = set().union(*belirsiz.values()) if belirsiz else set()

    # Seçili hesap dışı: belge no (tam ya da seri+sıra) yalnızca seçili olmayan hesaplarda geçiyor
    diger = {}  # fatura index → ([document_no_norm], yöntem)
    kalan = [idx for idx in clear.index if idx not in eslesme and idx not in belirsiz]
    for idx in kalan:
        no = clear.at[idx, "invoice_no_norm"]
        if no in other_docs:
            diger[idx] = ([no], YONTEM_TAM)
    ss = seri_sira_adaylari({idx: clear.at[idx, "invoice_no"] for idx in kalan if idx not in diger},
                            {d: v[0] for d, v in other_docs.items() if d not in invoice_nos})
    for idx, docs in ss.items():
        diger[idx] = (sorted(docs), YONTEM_SERI_SIRA)

    # 3) Tutar + tarih yedek eşleşmesi (belge no hiçbir yerde bulunamayan faturalar)
    eslesen_belgeler = {d for d, _ in eslesme.values()}
    kalan = {idx: (clear.at[idx, "issue_date"], clear.at[idx, "beklenen_tl"]) for idx in clear.index
             if idx not in eslesme and idx not in belirsiz and idx not in diger}
    acik = {d: (sel_docs.at[d, "Yevmiye_Tarihi"], sel_docs.at[d, "Yevmiye_Tutari"]) for d in acik_belgeler
            if d not in eslesen_belgeler and d not in belirsiz_belgeler}
    tekil, _ = tutar_tarih_eslestir(kalan, acik, tolerance)
    for idx, d in tekil.items():
        eslesme[idx] = (d, YONTEM_TUTAR_TARIH)
    eslesen_belgeler = {d for d, _ in eslesme.values()}

    not_booked = clear.loc[[idx for idx in clear.index
                            if idx not in eslesme and idx not in belirsiz and idx not in diger]]
    res["Muhasebeleşmemiş Faturalar"] = not_booked.rename(columns=base_cols)[list(base_cols.values())] \
        .reset_index(drop=True)

    other_acc = clear.loc[[idx for idx in clear.index if idx in diger]].copy()
    other_acc["Yevmiye_Belge_No"] = [", ".join(other_docs[d][0] for d in diger[i][0]) for i in other_acc.index]
    other_acc["Kullanilan_Hesaplar"] = [", ".join(sorted(set().union(*(other_docs[d][1] for d in diger[i][0]))))
                                        for i in other_acc.index]
    other_acc["Eslesme_Yontemi"] = [diger[i][1] for i in other_acc.index]
    res["Seçili Hesap Dışına Kaydedilmiş Faturalar"] = other_acc.rename(columns=base_cols)[
        list(base_cols.values()) + ["Kullanilan_Hesaplar", "Yevmiye_Belge_No", "Eslesme_Yontemi"]] \
        .reset_index(drop=True)

    matched = clear.loc[[idx for idx in clear.index if idx in eslesme]].copy()
    matched["document_no_norm"] = pd.Series([eslesme[i][0] for i in matched.index], index=matched.index, dtype=object)
    matched["Eslesme_Yontemi"] = pd.Series([eslesme[i][1] for i in matched.index], index=matched.index, dtype=object)
    matched = matched.merge(sel_grp, on="document_no_norm", how="left")
    matched["Fark"] = (matched["Yevmiye_Tutari"].abs() - matched["beklenen_tl"].abs()).round(2)
    matched["Dovizli"] = (matched["Para_Birimi"] != "TRY").map({True: "Evet", False: "Hayır"})
    diff, kur, res.kur_ozeti = kur_farki_ayir(matched, tolerance, kur_toleransi)
    diff = diff.assign(Olasi_Neden=[tutar_farki_nedeni(r, tolerance, kur_toleransi) for _, r in diff.iterrows()])
    tl_adlari = {"Yevmiye_Tutari": "Yevmiye_Tutari_TL", "Fark": "Fark_TL", "maliyet_vergisi_tl":
                 "Maliyete_Eklenen_Vergi_TL", "beklenen_tl": "Beklenen_Tutar_TL"}
    res["Tutar Farkları"] = diff.rename(columns=base_cols).rename(columns=tl_adlari)[TUTAR_FARKI_COLS] \
        .reset_index(drop=True)
    kur = kur.rename(columns=base_cols).rename(columns=tl_adlari)
    kur["Fark_Yuzdesi"] = (kur["Fark_TL"] / kur["Beklenen_Tutar_TL"].where(kur["Beklenen_Tutar_TL"] != 0) * 100) \
        .astype(float).round(2)
    res.kur_farki = kur[KUR_FARKI_COLS].reset_index(drop=True)

    if not matched.empty:
        matched["Fatura_Donemi"] = matched["issue_date"].map(lambda d: period_of(d, period_type))
        matched["Yevmiye_Donemi"] = matched["Yevmiye_Tarihi"].map(lambda d: period_of(d, period_type))
        period_diff = matched[matched["Fatura_Donemi"] != matched["Yevmiye_Donemi"]]
    else:
        period_diff = matched.assign(Fatura_Donemi=None, Yevmiye_Donemi=None)
    res["Dönem Farkları"] = period_diff.rename(columns=base_cols)[
        list(base_cols.values()) + ["Yevmiye_Tarihi", "Fatura_Donemi", "Yevmiye_Donemi", "Eslesme_Yontemi"]] \
        .reset_index(drop=True)

    # Belge no hatası da bir bulgudur: tutar + tarih ile eşleşenler ayrıca listelenir
    weak = matched[matched["Eslesme_Yontemi"] == YONTEM_TUTAR_TARIH]
    res["Belge No Uyuşmayan Eşleşmeler (Kontrol Edin)"] = weak.rename(columns=base_cols)[
        list(base_cols.values()) + ["Yevmiye_Belge_No", "Yevmiye_Tarihi", "Yevmiye_Tutari", "Hesaplar",
                                    "Eslesme_Yontemi"]] \
        .rename(columns={"Yevmiye_Tutari": "Yevmiye_Tutari_TL"}).reset_index(drop=True)

    orphan = sel_grp[~sel_grp["document_no_norm"].isin(invoice_nos | eslesen_belgeler | belirsiz_belgeler)]
    # Yalnızca borç yönlü ve hariç önekle başlamayan belgeler faturasız gider/alış adayı sayılır
    liste, haric, ozet = faturasiz_kayitlari_ayir(
        orphan[["Yevmiye_Belge_No", "Yevmiye_Tarihi", "Yevmiye_Tutari", "Hesaplar"]], haric_onekler, tolerance,
        yevmiye_isaretli(jou), "alacak" if satis else "borc")
    res["Faturası Bulunmayan Yevmiye Kayıtları"] = liste
    res.faturasiz_haric, res.faturasiz_ozeti = haric, ozet

    res["Belirsiz Eşleşme (Aynı No Farklı Tedarikçi)"] = ambiguous.rename(columns=base_cols)[
        list(base_cols.values())].sort_values("Fatura_No").reset_index(drop=True)

    multi = clear.loc[[idx for idx in clear.index if idx in belirsiz]].copy()
    multi["Aday_Belgeler"] = [", ".join(str(sel_docs.at[d, "Yevmiye_Belge_No"]) for d in belirsiz[i])
                              for i in multi.index]
    res["Belirsiz Eşleşme (Birden Fazla Seri+Sıra Adayı)"] = multi.rename(columns=base_cols)[
        list(base_cols.values()) + ["Aday_Belgeler"]].reset_index(drop=True)

    if satis and kdv_hesaplari:
        belgeler = {i: [d] for i, (d, _) in eslesme.items()}
        belgeler.update({i: list(d) for i, (d, _) in diger.items()})
        yontem = {i: y for i, (_, y) in eslesme.items()}
        yontem.update({i: y for i, (_, y) in diger.items()})
        res["KDV Farkları"], res.vergi_ozeti = satis_kdv_bulgulari(
            clear, belgeler, yontem, jou, kdv_hesaplari, tolerance, kur_toleransi, yevmiye_isaretli(jou))
    elif not satis and (kdv_hesaplari or tevkifat_hesaplari):
        belgeler = {i: [d] for i, (d, _) in eslesme.items()}
        belgeler.update({i: list(d) for i, (d, _) in diger.items()})
        yontem = {i: y for i, (_, y) in eslesme.items()}
        yontem.update({i: y for i, (_, y) in diger.items()})
        mf = {i: float(abs(sel_docs.at[d, "Yevmiye_Tutari"]) - abs(clear.at[i, "beklenen_tl"]))
              for i, (d, _) in eslesme.items()}
        kdv_fark, kdv_yok, tev, res.vergi_ozeti = alis_vergi_bulgulari(
            clear, belgeler, yontem, jou, kdv_hesaplari, tevkifat_hesaplari, tolerance, kur_toleransi,
            yevmiye_isaretli(jou), mf)
        if kdv_hesaplari:
            res["KDV Farkları"] = kdv_fark
            res["KDV'si Kaydedilmemiş Faturalar"] = kdv_yok
        if tevkifat_hesaplari:
            res["Tevkifat Kaydı Eksik/Farklı"] = tev

    yontemler = Counter(y for _, y in eslesme.values())
    res.eslesme_ozeti = OrderedDict([
        (YONTEM_TAM, yontemler[YONTEM_TAM]), (YONTEM_SERI_SIRA, yontemler[YONTEM_SERI_SIRA]),
        (YONTEM_TUTAR_TARIH, yontemler[YONTEM_TUTAR_TARIH]), ("Seçili hesap dışı", len(diger)),
        ("Belirsiz", len(belirsiz) + len(ambiguous)), ("Eşleşmeyen", len(not_booked)),
    ])
    if satis:
        _satis_bicimi(res)
    return res


SATIS_BASLIKLARI = OrderedDict([
    ("Muhasebeleşmemiş Faturalar", "Muhasebeleşmemiş Satış Faturaları"),
    ("Seçili Hesap Dışına Kaydedilmiş Faturalar", "Gelir Hesabı Dışına Kaydedilmiş Satış Faturaları"),
    ("Tutar Farkları", "Satış Tutar Farkları"),
    ("Dönem Farkları", "Satış Dönem Farkları"),
    ("Belge No Uyuşmayan Eşleşmeler (Kontrol Edin)", "Satış Belge No Uyuşmayan Eşleşmeler (Kontrol Edin)"),
    ("Faturası Bulunmayan Yevmiye Kayıtları", "Faturası Bulunmayan Gelir Kayıtları"),
    ("Belirsiz Eşleşme (Aynı No Farklı Tedarikçi)", "Satış Belirsiz Eşleşme (Aynı No Farklı Müşteri)"),
    ("Belirsiz Eşleşme (Birden Fazla Seri+Sıra Adayı)", "Satış Belirsiz Eşleşme (Birden Fazla Seri+Sıra Adayı)"),
    ("KDV Farkları", "Satış KDV Farkları"),
])
_SATIS_SUTUNLARI = {"Tedarikci": "Musteri", "Tedarikci_VKN": "Musteri_VKN"}
_SATIS_NEDENLERI = {NEDEN_KDV_MALIYETTE: "KDV gelire eklenmiş"}


def _satis_bicimi(res):
    """Alış mutabakatı biçimindeki sonucu satış başlık ve sütunlarına çevirir (yerinde)."""
    bolumler = list(res.items())
    res.clear()
    for baslik, df in bolumler:
        df = df.rename(columns=_SATIS_SUTUNLARI)
        if baslik == "Tutar Farkları":
            df = df.drop(columns=["Maliyete_Eklenen_Vergi_TL", "Beklenen_Tutar_TL"])
            df["Olasi_Neden"] = df["Olasi_Neden"].map(lambda n: _SATIS_NEDENLERI.get(n, n))
        res[SATIS_BASLIKLARI.get(baslik, baslik)] = df
    res.kur_farki = res.kur_farki.rename(columns=_SATIS_SUTUNLARI)


def satis_mutabakati(invoices, journal, gelir_hesaplari, kdv_hesaplari=None, tolerance=TOLERANCE_DEFAULT,
                     period_type="Aylık", haric_onekler=None, kur_toleransi=KUR_TOLERANSI_VARSAYILAN):
    """Satış faturalarını gelir hesapları (ör. 600, 601, 602; ALACAK yönlü) ve hesaplanan KDV hesaplarıyla
    (ör. 391) karşılaştırır. Firmanın kestiği iade faturaları alış iadesi sayıldığından buraya gelmez.

    Eşleştirme, tutar farkı (KDV hariç; kur toleransı dahil), dönem farkı ve faturasız gelir kaydı (gelir
    hesabında net alacak yönlü, hariç önekle başlamayan, faturası olmayan belge) reconcile() ile aynıdır;
    KDV: bkz. satis_kdv_bulgulari. Dönüş: MutabakatSonucu (başlıklar SATIS_BASLIKLARI)
    """
    return reconcile(invoices, journal, gelir_hesaplari, tolerance, period_type, haric_onekler, kur_toleransi,
                     kdv_hesaplari, None, YON_SATIS)


# ---------------------------------------------------------------------- tutarlılık kontrolleri
ARDISIK_EVET = "Evet"
ARDISIK_HAYIR = "Hayır"


def ardisik_numaralar(fatura_nolari):
    """Faturalar aynı serinin (ve yazılmışsa aynı yılın) ardışık sıra numaraları mı? (ör. ABC2024000000017,
    ABC2024000000018). Distribütörün aynı gün ayrı faturalarla kestiği alımlarda görülür; bilgi amaçlıdır."""
    anahtarlar = []
    for no in fatura_nolari:
        adaylar = belge_anahtarlari(no)
        if not adaylar:
            return False
        anahtarlar.append(adaylar[0])
    if len(anahtarlar) < 2 or len({(a[0], a[1]) for a in anahtarlar}) != 1:
        return False
    siralar = sorted(a[2] for a in anahtarlar)
    return all(b - a == 1 for a, b in zip(siralar, siralar[1:]))


def duplicate_suspects(invoices):
    """Aynı tedarikçiden, aynı tarihli ve aynı tutarlı farklı numaralı faturalar (olası mükerrer).

    Ardisik_Numara: grubun faturaları aynı serinin ardışık numaralarıysa "Evet" (bilgi; bulgu gizlenmez, ardışık
    olmayanlar — daha şüpheli — önce sıralanır).
    """
    cols = ["Tedarikci", "Tedarikci_VKN", "Fatura_Tarihi", "KDV_Haric_Tutar", "Adet", "Ardisik_Numara", "Fatura_No"]
    if invoices.empty:
        return pd.DataFrame(columns=cols)
    df = invoices.copy()
    df["tutar"] = df["total_amount"].round(2)
    g = df.groupby(["supplier_vkn", "issue_date", "tutar"]).agg(
        Tedarikci=("supplier_name", "first"), Fatura_No=("invoice_no", lambda s: ", ".join(sorted(s))),
        Nolar=("invoice_no", list), Adet=("id", "count")).reset_index()
    g = g[g["Adet"] > 1].rename(columns={"supplier_vkn": "Tedarikci_VKN", "issue_date": "Fatura_Tarihi",
                                         "tutar": "KDV_Haric_Tutar"})
    g["Ardisik_Numara"] = [ARDISIK_EVET if ardisik_numaralar(n) else ARDISIK_HAYIR for n in g["Nolar"]]
    g = g.sort_values(["Ardisik_Numara", "Fatura_Tarihi", "Tedarikci_VKN"], ascending=[False, True, True],
                      kind="mergesort")  # "Hayır" > "Evet": ardışık olmayanlar önce
    return g[cols].reset_index(drop=True)


def calculation_errors(invoices, lines, tolerance=TOLERANCE_DEFAULT):
    """XML faturalarda satır toplamları ile belge toplamlarının tutarlılığı."""
    cols = ["Fatura_No", "Tedarikci", "Tedarikci_VKN", "Kontrol", "Belgedeki_Tutar", "Hesaplanan_Tutar", "Fark"]
    if invoices.empty:
        return pd.DataFrame(columns=cols)
    xml_inv = invoices[invoices["source"] == "XML"]
    if xml_inv.empty:
        return pd.DataFrame(columns=cols)
    agg = lines.groupby("invoice_id").agg(satir_net=("line_net", "sum"),
                                         satir_kdv=("vat_amount", "sum")).reset_index()
    df = xml_inv.merge(agg, left_on="id", right_on="invoice_id", how="left").fillna({"satir_net": 0, "satir_kdv": 0})
    rows = []
    for _, r in df.iterrows():
        checks = [
            ("Satır toplamı ≠ Mal/Hizmet toplamı", r["line_extension_amount"], r["satir_net"]),
            ("Mal/Hizmet - İskonto ≠ KDV Hariç Tutar",
             r["total_amount"], (r["line_extension_amount"] or 0) - (r["allowance_total"] or 0)),
        ]
        # Belge geneli iskonto varsa satır KDV'leri belge KDV'sine eşit olmayabilir
        if not r["allowance_total"]:
            checks.append(("Satır KDV toplamı ≠ Belge KDV toplamı", r["vat_amount"], r["satir_kdv"]))
        for name, doc_val, calc_val in checks:
            if doc_val is None or pd.isna(doc_val):
                continue
            if abs(float(doc_val) - float(calc_val)) > tolerance:
                rows.append({"Fatura_No": r["invoice_no"], "Tedarikci": r["supplier_name"],
                             "Tedarikci_VKN": r["supplier_vkn"], "Kontrol": name,
                             "Belgedeki_Tutar": round(float(doc_val), 2), "Hesaplanan_Tutar": round(float(calc_val), 2),
                             "Fark": round(float(doc_val) - float(calc_val), 2)})
    return pd.DataFrame(rows, columns=cols)


def customer_mismatch(invoices, company_vkn):
    """Alıcı VKN'si firma VKN'sinden farklı XML faturalar (başka firmaya kesilmiş fatura)."""
    cols = ["Fatura_No", "Fatura_Tarihi", "Tedarikci", "Tedarikci_VKN", "Alici_VKN", "Alici_Unvan"]
    company_vkn = normalize_vkn(company_vkn)
    if not company_vkn or invoices.empty:
        return pd.DataFrame(columns=cols)
    # Firmanın kendi kestiği faturalar (satış, alış iadesi) kontrol edilmez: alıcısı karşı taraftır
    df = invoices[(invoices["source"] == "XML") & (invoices["customer_vkn"].fillna("") != company_vkn)
                  & (invoices["supplier_vkn"].map(normalize_vkn) != company_vkn)]
    return df.rename(columns={"invoice_no": "Fatura_No", "issue_date": "Fatura_Tarihi", "supplier_name": "Tedarikci",
                              "supplier_vkn": "Tedarikci_VKN", "customer_vkn": "Alici_VKN", "customer_name": "Alici_Unvan"})[cols].reset_index(drop=True)


# ---------------------------------------------------------------------- genel rapor
def run_full_audit(db, period_type, threshold, accounts, tolerance, company_vkn, haric_onekler=None,
                   fiyat_kurallari=None, kur_toleransi=KUR_TOLERANSI_VARSAYILAN, vergi_ayarlari=None):
    """Tüm kontrolleri çalıştırır.

    haric_onekler: faturasız kayıt kontrolünde hariç tutulan belge no önekleri (None → varsayılan)
    fiyat_kurallari: FiyatKurallari (None → varsayılan kurallar)
    kur_toleransi: dövizli faturalarda yüzde kur farkı toleransı (bkz. reconcile)
    vergi_ayarlari: VergiAyarlari (None → varsayılan: 191 / 360, gelir 600-602, satış KDV 391)
    Faturalar yönüne göre ayrılır (bkz. fatura_yonleri): fiyat analizi, alış mutabakatı, mükerrer ve alıcı VKN
    kontrolü yalnızca alışlara; satış mutabakatı (satis_mutabakati) satışlara; hesaplama kontrolü tümüne uygulanır.
    Dönüş: (özet DataFrame, OrderedDict(başlık → DataFrame), notlar, sayılar)
    sayılar: {"fatura", "alis_fatura", "satis_fatura", "yevmiye", "eslesme", "faturasiz_ozeti", "faturasiz_haric",
    "kur_ozeti", "kur_farki", "fiyat_ozeti", "fiyat_haric", "maliyet_vergisi", "vergi_ozeti", "satis"}; alış
    mutabakatı yapılmadıysa eslesme / faturasiz_ozeti / faturasiz_haric / kur_ozeti / kur_farki / vergi_ozeti None;
    satış mutabakatı yapılmadıysa satis None, yapıldıysa aynı anahtarlarla dict.
    """
    vergi = vergi_ayarlari or VergiAyarlari()
    invoices = maliyet_vergisi_ekle(db.get_invoices_df(), db.get_invoice_taxes_df(), vergi.maliyet_kodlari)
    lines = db.get_lines_df()
    journal = db.get_journal_df()
    alis = alis_faturalari(invoices, company_vkn)
    satis = satis_faturalari(invoices, company_vkn)
    sections = OrderedDict()
    notes = []
    eslesme = faturasiz_ozeti = faturasiz_haric = kur_ozeti = kur_farki = vergi_ozeti = satis_sayilari = None

    if not invoices.empty and "tax_detail" in invoices:
        eski = int(((invoices["source"] == "XML") & invoices["tax_detail"].isna()).sum())
        if eski:
            notes.append(f"{eski} XML fatura önceki bir sürümle yüklenmiş: ÖTV ve tevkifat bilgisi yok (maliyete eklenen "
                         "vergi ve tevkifat kontrolleri bu faturalarda eksik kalır). Tam kontrol için verileri silip "
                         "XML'leri yeniden yükleyin.")
    fiyat = fiyat_analizi(alis_faturalari(lines, company_vkn), period_type, threshold, fiyat_kurallari)
    sections[f"Fiyat Anomalileri (±%{threshold:g})"] = fiyat.riskli

    if accounts and not journal.empty:
        recon = reconcile(alis, journal, accounts, tolerance, period_type, haric_onekler, kur_toleransi,
                          vergi.kdv_hesaplari, vergi.tevkifat_hesaplari)
        vergi_ozeti = recon.vergi_ozeti
        sections.update(recon)
        eslesme = recon.eslesme_ozeti
        faturasiz_ozeti, faturasiz_haric = recon.faturasiz_ozeti, recon.faturasiz_haric
        kur_ozeti, kur_farki = recon.kur_ozeti, recon.kur_farki
        if not faturasiz_ozeti["isaretli"]:
            notes.append("Yevmiye tutarları işaretsiz (negatif tutar yok); faturasız kayıt kontrolünde "
                         "borç/alacak yönü filtresi uygulanmadı.")
    elif not accounts:
        notes.append("Mutabakat hesap kodu girilmediği için alış mutabakatı kontrolleri atlandı.")
    else:
        notes.append("Yevmiye kaydı yüklenmediği için mutabakat kontrolleri atlandı.")

    if not satis.empty and vergi.gelir_hesaplari and not journal.empty:
        srec = satis_mutabakati(satis, journal, vergi.gelir_hesaplari, vergi.satis_kdv_hesaplari, tolerance,
                                period_type, haric_onekler, kur_toleransi)
        sections.update(srec)
        satis_sayilari = {"fatura": len(satis), "eslesme": srec.eslesme_ozeti, "faturasiz_ozeti": srec.faturasiz_ozeti,
                          "faturasiz_haric": srec.faturasiz_haric, "kur_ozeti": srec.kur_ozeti,
                          "kur_farki": srec.kur_farki, "vergi_ozeti": srec.vergi_ozeti}
    elif not satis.empty and not vergi.gelir_hesaplari:
        notes.append("Gelir hesap kodu girilmediği için satış mutabakatı atlandı.")
    elif satis.empty and not normalize_vkn(company_vkn):
        notes.append("Firma VKN'si girilmediği için satış faturaları ayırt edilemedi (tüm faturalar alış sayıldı).")

    sections["Olası Mükerrer Faturalar"] = duplicate_suspects(alis)
    sections["Fatura Hesaplama Tutarsızlıkları"] = calculation_errors(invoices, lines, tolerance)
    if normalize_vkn(company_vkn):
        sections["Alıcısı Firma Olmayan Faturalar"] = customer_mismatch(alis, company_vkn)
    else:
        notes.append("Firma VKN'si girilmediği için alıcı VKN kontrolü atlandı.")

    summary = pd.DataFrame(
        [{"Kontrol": name, "Bulgu_Sayisi": len(df)} for name, df in sections.items()])
    return summary, sections, notes, {"fatura": len(invoices), "alis_fatura": len(alis), "satis_fatura": len(satis),
                                      "yevmiye": len(journal), "eslesme": eslesme,
                                      "faturasiz_ozeti": faturasiz_ozeti, "faturasiz_haric": faturasiz_haric,
                                      "kur_ozeti": kur_ozeti, "kur_farki": kur_farki,
                                      "fiyat_ozeti": fiyat.ozet, "fiyat_haric": fiyat.haric,
                                      "maliyet_vergisi": maliyet_vergisi_ozeti(alis), "vergi_ozeti": vergi_ozeti,
                                      "satis": satis_sayilari}
