"""Denetim kontrolleri: fiyat anomalisi, mutabakat ve tutarlılık kontrolleri."""
import math
import re
from bisect import bisect_left, bisect_right
from collections import Counter, OrderedDict, defaultdict
from datetime import date

import pandas as pd

from .utils import normalize_uom, normalize_vkn, period_of, tr_ascii_upper

TOLERANCE_DEFAULT = 0.01
# Dövizli (TRY dışı) faturalarda muhasebe fatura kuru yerine başka günün kurunu kullanabilir; TL tutarın bu
# yüzdesine kadar olan fark tutar farkı sayılmaz (firma başına değiştirilebilir). TL faturalara uygulanmaz:
# rakam yer değiştirme hataları (12.345,67 → 12.354,67) yalnızca birkaç TL fark yaratır.
KUR_TOLERANSI_VARSAYILAN = 1.0


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


def faturasiz_kayitlari_ayir(adaylar, onekler=None, tolerance=TOLERANCE_DEFAULT, isaretli=True):
    """Hiçbir faturayla eşleşmeyen seçili hesap belgelerinden faturasız gider/alış adaylarını ayırır.

    adaylar: Yevmiye_Belge_No, Yevmiye_Tarihi, Yevmiye_Tutari (seçili hesaplarda Borç − Alacak), Hesaplar
    onekler: parse_onek_listesi() çıktısı; None → VARSAYILAN_HARIC_ONEKLER, [] → önek filtresi yok
    isaretli: False ise (işaretsiz Tutar) yön bilinmediğinden borç/alacak filtresi uygulanmaz.
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
    neden = pd.Series([None] * len(df), index=df.index, dtype=object)
    if isaretli:
        neden[tutar <= tolerance] = NEDEN_ALACAK
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
    metin = (f"Faturasız kayıt: {fmt(ozet['listelenen'])} listelendi (Yüksek öncelik: {fmt(ozet['yuksek'])}, "
             f"Düşük: {fmt(ozet['dusuk'])}) | Listeye alınmayan: ")
    if ozet["isaretli"]:
        metin += f"alacak yönlü {fmt(ozet['alacak_yonlu'])}, "
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
    yon = ("Seçili hesaplardaki net tutar (Borç − Alacak) borç yönünde değil (ör. stoktan çıkış, SMM)"
           if ozet["isaretli"] else "Uygulanmadı: yevmiye tutarları işaretsiz (negatif tutar yok), yön bilinmiyor")
    rows = [
        ("Listelenen (Yüksek öncelik)", ozet["yuksek"], "Belge no GİB fatura numarası biçiminde"),
        ("Listelenen (Düşük öncelik)", ozet["dusuk"], "Belge no fatura numarasına benzemiyor"),
        ("Hariç: alacak yönlü", ozet["alacak_yonlu"] if ozet["isaretli"] else None, yon),
        ("Hariç: önek", ozet["haric_onek"],
         "Hariç tutulan önekler: " + (", ".join(ozet["onekler"]) or "(yok)")),
    ]
    return pd.DataFrame(rows, columns=cols)


# ---------------------------------------------------------------------- kur farkı toleransı
TUTAR_FARKI_COLS = ["Fatura_No", "Fatura_Tarihi", "Tedarikci", "Tedarikci_VKN", "KDV_Haric_Tutar_TL",
                    "Yevmiye_Tutari_TL", "Fark_TL", "Dovizli", "Para_Birimi", "Kur", "Hesaplar", "Yevmiye_Belge_No",
                    "Eslesme_Yontemi"]
KUR_FARKI_COLS = ["Fatura_No", "Fatura_Tarihi", "Tedarikci", "Tedarikci_VKN", "Para_Birimi", "Kur",
                  "KDV_Haric_Tutar_TL", "Yevmiye_Tutari_TL", "Fark_TL", "Fark_Yuzdesi", "Yevmiye_Belge_No"]


def izin_verilen_fark(tutar_tl, dovizli, tolerance=TOLERANCE_DEFAULT, kur_toleransi=KUR_TOLERANSI_VARSAYILAN):
    """Tutar farkı sayılmayacak en büyük fark (TL).

    TL faturada sabit TL toleransı; dövizli faturada max(TL toleransı, kur_toleransi % × TL tutar).
    """
    if not dovizli or not kur_toleransi:
        return tolerance
    return max(tolerance, abs(float(tutar_tl or 0)) * float(kur_toleransi) / 100)


def kur_farki_ayir(eslesen, tolerance=TOLERANCE_DEFAULT, kur_toleransi=KUR_TOLERANSI_VARSAYILAN):
    """Eşleşen faturaları (Fark sütunu TL tolerans üstü olanlar) tutar farkı ve tolerans içi kur farkı olarak ayırır.

    eslesen: reconcile() içindeki eşleşen faturalar (net_tl, Fark, Para_Birimi sütunlarıyla)
    Dönüş: (tutar farkı satırları, kur farkı (tolerans içi) satırları, özet dict)
    """
    fark = eslesen[eslesen["Fark"].abs() > tolerance]
    dovizli = fark["Para_Birimi"] != "TRY"
    sinir = [izin_verilen_fark(t, d, tolerance, kur_toleransi) for t, d in zip(fark["net_tl"], dovizli)]
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


# ---------------------------------------------------------------------- mutabakat
def reconcile(invoices, journal, accounts, tolerance=TOLERANCE_DEFAULT, period_type="Aylık", haric_onekler=None,
              kur_toleransi=KUR_TOLERANSI_VARSAYILAN):
    """Faturaları (KDV hariç, TL) seçili hesap kodlarındaki yevmiye kayıtlarıyla karşılaştırır.

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
    kalan = {idx: (clear.at[idx, "issue_date"], clear.at[idx, "net_tl"]) for idx in clear.index
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
    matched["Fark"] = (matched["Yevmiye_Tutari"].abs() - matched["net_tl"].abs()).round(2)
    matched["Dovizli"] = (matched["Para_Birimi"] != "TRY").map({True: "Evet", False: "Hayır"})
    diff, kur, res.kur_ozeti = kur_farki_ayir(matched, tolerance, kur_toleransi)
    tl_adlari = {"Yevmiye_Tutari": "Yevmiye_Tutari_TL", "Fark": "Fark_TL"}
    res["Tutar Farkları"] = diff.rename(columns=base_cols).rename(columns=tl_adlari)[TUTAR_FARKI_COLS] \
        .reset_index(drop=True)
    kur = kur.rename(columns=base_cols).rename(columns=tl_adlari)
    kur["Fark_Yuzdesi"] = (kur["Fark_TL"] / kur["KDV_Haric_Tutar_TL"].where(kur["KDV_Haric_Tutar_TL"] != 0) * 100) \
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
        yevmiye_isaretli(jou))
    res["Faturası Bulunmayan Yevmiye Kayıtları"] = liste
    res.faturasiz_haric, res.faturasiz_ozeti = haric, ozet

    res["Belirsiz Eşleşme (Aynı No Farklı Tedarikçi)"] = ambiguous.rename(columns=base_cols)[
        list(base_cols.values())].sort_values("Fatura_No").reset_index(drop=True)

    multi = clear.loc[[idx for idx in clear.index if idx in belirsiz]].copy()
    multi["Aday_Belgeler"] = [", ".join(str(sel_docs.at[d, "Yevmiye_Belge_No"]) for d in belirsiz[i])
                              for i in multi.index]
    res["Belirsiz Eşleşme (Birden Fazla Seri+Sıra Adayı)"] = multi.rename(columns=base_cols)[
        list(base_cols.values()) + ["Aday_Belgeler"]].reset_index(drop=True)

    yontemler = Counter(y for _, y in eslesme.values())
    res.eslesme_ozeti = OrderedDict([
        (YONTEM_TAM, yontemler[YONTEM_TAM]), (YONTEM_SERI_SIRA, yontemler[YONTEM_SERI_SIRA]),
        (YONTEM_TUTAR_TARIH, yontemler[YONTEM_TUTAR_TARIH]), ("Seçili hesap dışı", len(diger)),
        ("Belirsiz", len(belirsiz) + len(ambiguous)), ("Eşleşmeyen", len(not_booked)),
    ])
    return res


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
    df = invoices[(invoices["source"] == "XML") & (invoices["customer_vkn"].fillna("") != company_vkn)]
    return df.rename(columns={"invoice_no": "Fatura_No", "issue_date": "Fatura_Tarihi", "supplier_name": "Tedarikci",
                              "supplier_vkn": "Tedarikci_VKN", "customer_vkn": "Alici_VKN", "customer_name": "Alici_Unvan"})[cols].reset_index(drop=True)


# ---------------------------------------------------------------------- genel rapor
def run_full_audit(db, period_type, threshold, accounts, tolerance, company_vkn, haric_onekler=None,
                   fiyat_kurallari=None, kur_toleransi=KUR_TOLERANSI_VARSAYILAN):
    """Tüm kontrolleri çalıştırır.

    haric_onekler: faturasız kayıt kontrolünde hariç tutulan belge no önekleri (None → varsayılan)
    fiyat_kurallari: FiyatKurallari (None → varsayılan kurallar)
    kur_toleransi: dövizli faturalarda yüzde kur farkı toleransı (bkz. reconcile)
    Dönüş: (özet DataFrame, OrderedDict(başlık → DataFrame), notlar, sayılar)
    sayılar: {"fatura", "yevmiye", "eslesme", "faturasiz_ozeti", "faturasiz_haric", "kur_ozeti", "kur_farki",
    "fiyat_ozeti", "fiyat_haric"}; mutabakat yapılmadıysa eslesme / faturasiz_ozeti / faturasiz_haric / kur_ozeti /
    kur_farki None.
    """
    invoices = db.get_invoices_df()
    lines = db.get_lines_df()
    journal = db.get_journal_df()
    sections = OrderedDict()
    notes = []
    eslesme = faturasiz_ozeti = faturasiz_haric = kur_ozeti = kur_farki = None

    fiyat = fiyat_analizi(lines, period_type, threshold, fiyat_kurallari)
    sections[f"Fiyat Anomalileri (±%{threshold:g})"] = fiyat.riskli

    if accounts and not journal.empty:
        recon = reconcile(invoices, journal, accounts, tolerance, period_type, haric_onekler, kur_toleransi)
        sections.update(recon)
        eslesme = recon.eslesme_ozeti
        faturasiz_ozeti, faturasiz_haric = recon.faturasiz_ozeti, recon.faturasiz_haric
        kur_ozeti, kur_farki = recon.kur_ozeti, recon.kur_farki
        if not faturasiz_ozeti["isaretli"]:
            notes.append("Yevmiye tutarları işaretsiz (negatif tutar yok); faturasız kayıt kontrolünde "
                         "borç/alacak yönü filtresi uygulanmadı.")
    elif not accounts:
        notes.append("Mutabakat hesap kodu girilmediği için mutabakat kontrolleri atlandı.")
    else:
        notes.append("Yevmiye kaydı yüklenmediği için mutabakat kontrolleri atlandı.")

    sections["Olası Mükerrer Faturalar"] = duplicate_suspects(invoices)
    sections["Fatura Hesaplama Tutarsızlıkları"] = calculation_errors(invoices, lines, tolerance)
    if normalize_vkn(company_vkn):
        sections["Alıcısı Firma Olmayan Faturalar"] = customer_mismatch(invoices, company_vkn)
    else:
        notes.append("Firma VKN'si girilmediği için alıcı VKN kontrolü atlandı.")

    summary = pd.DataFrame(
        [{"Kontrol": name, "Bulgu_Sayisi": len(df)} for name, df in sections.items()])
    return summary, sections, notes, {"fatura": len(invoices), "yevmiye": len(journal), "eslesme": eslesme,
                                      "faturasiz_ozeti": faturasiz_ozeti, "faturasiz_haric": faturasiz_haric,
                                      "kur_ozeti": kur_ozeti, "kur_farki": kur_farki,
                                      "fiyat_ozeti": fiyat.ozet, "fiyat_haric": fiyat.haric}
