"""Denetim kontrolleri: fiyat anomalisi, mutabakat ve tutarlılık kontrolleri."""
import math
import re
from bisect import bisect_left, bisect_right
from collections import Counter, OrderedDict, defaultdict
from datetime import date

import pandas as pd

from .utils import normalize_uom, normalize_vkn, period_of

TOLERANCE_DEFAULT = 0.01


# ---------------------------------------------------------------------- fiyat analizi
def price_anomalies(lines, period_type="Aylık", threshold=15.0):
    """Dönem + ürün + birim bazında ağırlıklı ortalama birim fiyattan (AOBF) sapmaları hesaplar.

    İade faturaları ve miktarı 0 olan satırlar analize alınmaz. Tutarlar TL'ye çevrilir.
    Dönüş: tüm satırları içeren DataFrame (Risk_Durumu sütunuyla).
    """
    cols = ["Donem", "Tarih", "Fatura_No", "Tedarikci", "Tedarikci_VKN", "Urun_Adi", "Birim", "Miktar",
            "Birim_Fiyat_TL", "AOBF_TL", "Fark_Yuzdesi", "Donemdeki_Alim_Sayisi", "Risk_Durumu"]
    if lines.empty:
        return pd.DataFrame(columns=cols)
    df = lines[(lines["quantity"] > 0) & (lines["invoice_type"].fillna("") != "IADE")].copy()
    if df.empty:
        return pd.DataFrame(columns=cols)
    df["rate"] = df["exchange_rate"].fillna(1.0)
    df["net_tl"] = df["line_net"] * df["rate"]
    df["Birim_Fiyat_TL"] = df["net_tl"] / df["quantity"]
    df["Donem"] = df["issue_date"].map(lambda d: period_of(d, period_type))
    df["uom_key"] = df["uom"].map(normalize_uom)

    keys = ["Donem", "item_norm", "uom_key"]
    grp = df.groupby(keys).agg(top_tutar=("net_tl", "sum"), top_miktar=("quantity", "sum"),
                               Donemdeki_Alim_Sayisi=("id", "count")).reset_index()
    grp["AOBF_TL"] = grp["top_tutar"] / grp["top_miktar"]
    df = df.merge(grp[keys + ["AOBF_TL", "Donemdeki_Alim_Sayisi"]], on=keys)
    df["Fark_Yuzdesi"] = (df["Birim_Fiyat_TL"] - df["AOBF_TL"]) / df["AOBF_TL"] * 100
    df.loc[df["AOBF_TL"] == 0, "Fark_Yuzdesi"] = 0.0
    df["Risk_Durumu"] = df["Fark_Yuzdesi"].abs().gt(threshold).map({True: "YÜKSEK RİSK", False: "Normal"})

    out = df.rename(columns={"issue_date": "Tarih", "invoice_no": "Fatura_No", "supplier_name": "Tedarikci",
                             "supplier_vkn": "Tedarikci_VKN", "item_name": "Urun_Adi", "uom_key": "Birim",
                             "quantity": "Miktar"})[cols]
    for c in ("Birim_Fiyat_TL", "AOBF_TL", "Fark_Yuzdesi"):
        out[c] = out[c].round(2)
    return out.sort_values(["Donem", "Urun_Adi", "Tarih"]).reset_index(drop=True)


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
    """reconcile() sonucu: OrderedDict(başlık → DataFrame) + eslesme_ozeti (kategori → fatura sayısı)."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.eslesme_ozeti = OrderedDict()


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


# ---------------------------------------------------------------------- mutabakat
def reconcile(invoices, journal, accounts, tolerance=TOLERANCE_DEFAULT, period_type="Aylık"):
    """Faturaları (KDV hariç, TL) seçili hesap kodlarındaki yevmiye kayıtlarıyla karşılaştırır.

    accounts: normalize edilmiş hesap kodu önekleri listesi (ör. ['153', '770']).
    Fatura no ↔ yevmiye belge no kademeli eşleştirilir:
      1. Tam: normalize edilmiş Fatura_No = Belge_No
      2. Seri+Sıra: GİB numarasının seri + sıra (+ varsa yıl) anahtarı aynı, tek aday
      3. Tutar+Tarih: belge no uyuşmayan, aynı tutar ve ±TARIH_PENCERESI_GUN gün içinde tek aday (düşük güven)
    Seçili hesap dışı kontrolü 3. kademeden önce (tam + seri/sıra ile) yapılır.
    Dönüş: MutabakatSonucu (OrderedDict başlık → DataFrame, .eslesme_ozeti)
    """
    res = MutabakatSonucu()
    if not accounts:
        raise ValueError("En az bir hesap kodu girilmelidir")

    inv = invoices.copy()
    if inv.empty:
        inv = pd.DataFrame(columns=["invoice_no", "invoice_no_norm", "issue_date", "supplier_vkn", "supplier_name",
                                    "total_amount", "exchange_rate", "invoice_type"])
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
    diff = matched[matched["Fark"].abs() > tolerance]
    res["Tutar Farkları"] = diff.rename(columns=base_cols)[
        list(base_cols.values()) + ["Yevmiye_Tutari", "Fark", "Hesaplar", "Yevmiye_Belge_No", "Eslesme_Yontemi"]] \
        .rename(columns={"Yevmiye_Tutari": "Yevmiye_Tutari_TL", "Fark": "Fark_TL"}).reset_index(drop=True)

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
    res["Faturası Bulunmayan Yevmiye Kayıtları"] = orphan[
        ["Yevmiye_Belge_No", "Yevmiye_Tarihi", "Yevmiye_Tutari", "Hesaplar"]].reset_index(drop=True)

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
def duplicate_suspects(invoices):
    """Aynı tedarikçiden, aynı tarihli ve aynı tutarlı farklı numaralı faturalar (olası mükerrer)."""
    cols = ["Tedarikci", "Tedarikci_VKN", "Fatura_Tarihi", "KDV_Haric_Tutar", "Fatura_No", "Adet"]
    if invoices.empty:
        return pd.DataFrame(columns=cols)
    df = invoices.copy()
    df["tutar"] = df["total_amount"].round(2)
    g = df.groupby(["supplier_vkn", "issue_date", "tutar"]).agg(
        Tedarikci=("supplier_name", "first"), Fatura_No=("invoice_no", lambda s: ", ".join(sorted(s))),
        Adet=("id", "count")).reset_index()
    g = g[g["Adet"] > 1].rename(columns={"supplier_vkn": "Tedarikci_VKN", "issue_date": "Fatura_Tarihi",
                                         "tutar": "KDV_Haric_Tutar"})
    return g[cols].reset_index(drop=True)


def calculation_errors(invoices, lines, tolerance=TOLERANCE_DEFAULT):
    """XML faturalarda satır toplamları ile belge toplamlarının tutarlılığı."""
    cols = ["Fatura_No", "Tedarikci", "Kontrol", "Belgedeki_Tutar", "Hesaplanan_Tutar", "Fark"]
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
                rows.append({"Fatura_No": r["invoice_no"], "Tedarikci": r["supplier_name"], "Kontrol": name,
                             "Belgedeki_Tutar": round(float(doc_val), 2), "Hesaplanan_Tutar": round(float(calc_val), 2),
                             "Fark": round(float(doc_val) - float(calc_val), 2)})
    return pd.DataFrame(rows, columns=cols)


def customer_mismatch(invoices, company_vkn):
    """Alıcı VKN'si firma VKN'sinden farklı XML faturalar (başka firmaya kesilmiş fatura)."""
    cols = ["Fatura_No", "Fatura_Tarihi", "Tedarikci", "Alici_VKN", "Alici_Unvan"]
    company_vkn = normalize_vkn(company_vkn)
    if not company_vkn or invoices.empty:
        return pd.DataFrame(columns=cols)
    df = invoices[(invoices["source"] == "XML") & (invoices["customer_vkn"].fillna("") != company_vkn)]
    return df.rename(columns={"invoice_no": "Fatura_No", "issue_date": "Fatura_Tarihi", "supplier_name": "Tedarikci",
                              "customer_vkn": "Alici_VKN", "customer_name": "Alici_Unvan"})[cols].reset_index(drop=True)


# ---------------------------------------------------------------------- genel rapor
def run_full_audit(db, period_type, threshold, accounts, tolerance, company_vkn):
    """Tüm kontrolleri çalıştırır.

    Dönüş: (özet DataFrame, OrderedDict(başlık → DataFrame), notlar, sayılar)
    sayılar: {"fatura", "yevmiye", "eslesme"}; eslesme, mutabakat yapıldıysa yöntem → fatura sayısı,
    yapılmadıysa None.
    """
    invoices = db.get_invoices_df()
    lines = db.get_lines_df()
    journal = db.get_journal_df()
    sections = OrderedDict()
    notes = []
    eslesme = None

    prices = price_anomalies(lines, period_type, threshold)
    sections[f"Fiyat Anomalileri (±%{threshold:g})"] = prices[prices["Risk_Durumu"] == "YÜKSEK RİSK"] \
        .reset_index(drop=True)

    if accounts and not journal.empty:
        recon = reconcile(invoices, journal, accounts, tolerance, period_type)
        sections.update(recon)
        eslesme = recon.eslesme_ozeti
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
    return summary, sections, notes, {"fatura": len(invoices), "yevmiye": len(journal), "eslesme": eslesme}
