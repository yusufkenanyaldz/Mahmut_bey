"""Denetim kontrolleri: fiyat anomalisi, mutabakat ve tutarlılık kontrolleri."""
from collections import OrderedDict

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


# ---------------------------------------------------------------------- mutabakat
def reconcile(invoices, journal, accounts, tolerance=TOLERANCE_DEFAULT, period_type="Aylık"):
    """Faturaları (KDV hariç, TL) seçili hesap kodlarındaki yevmiye kayıtlarıyla karşılaştırır.

    accounts: normalize edilmiş hesap kodu önekleri listesi (ör. ['153', '770']).
    Belge numarası eşleşmesi normalize edilmiş Fatura_No = Belge_No üzerinden yapılır.
    Dönüş: OrderedDict(başlık → DataFrame)
    """
    res = OrderedDict()
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
    jou["selected"] = jou["account_norm"].map(lambda a: any(a.startswith(p) for p in accounts))
    sel = jou[jou["selected"]]

    sel_grp = sel.groupby("document_no_norm").agg(
        Yevmiye_Belge_No=("document_no", "first"), Yevmiye_Tarihi=("entry_date", "min"),
        Yevmiye_Tutari=("amount", "sum"), Hesaplar=("account_code", lambda s: ", ".join(sorted(set(map(str, s))))),
    ).reset_index()
    all_docs = jou.groupby("document_no_norm").agg(
        Kullanilan_Hesaplar=("account_code", lambda s: ", ".join(sorted(set(map(str, s)))))).reset_index()

    # Aynı fatura numarası birden fazla tedarikçide → eşleşme belirsiz
    dup_mask = inv.duplicated("invoice_no_norm", keep=False)
    ambiguous = inv[dup_mask]
    clear = inv[~dup_mask]

    base_cols = {"invoice_no": "Fatura_No", "issue_date": "Fatura_Tarihi", "supplier_name": "Tedarikci",
                 "supplier_vkn": "Tedarikci_VKN", "net_tl": "KDV_Haric_Tutar_TL"}

    merged = clear.merge(sel_grp, left_on="invoice_no_norm", right_on="document_no_norm", how="left")
    unmatched = merged[merged["document_no_norm"].isna()]
    unmatched = unmatched.drop(columns=["document_no_norm"]).merge(
        all_docs, left_on="invoice_no_norm", right_on="document_no_norm", how="left")

    not_booked = unmatched[unmatched["Kullanilan_Hesaplar"].isna()]
    res["Muhasebeleşmemiş Faturalar"] = not_booked.rename(columns=base_cols)[list(base_cols.values())] \
        .reset_index(drop=True)

    other_acc = unmatched[unmatched["Kullanilan_Hesaplar"].notna()]
    res["Seçili Hesap Dışına Kaydedilmiş Faturalar"] = other_acc.rename(columns=base_cols)[
        list(base_cols.values()) + ["Kullanilan_Hesaplar"]].reset_index(drop=True)

    matched = merged[merged["document_no_norm"].notna()].copy()
    matched["Yevmiye_Tutari"] = matched["Yevmiye_Tutari"].astype(float).round(2)
    matched["Fark"] = (matched["Yevmiye_Tutari"].abs() - matched["net_tl"].abs()).round(2)
    diff = matched[matched["Fark"].abs() > tolerance]
    res["Tutar Farkları"] = diff.rename(columns=base_cols)[
        list(base_cols.values()) + ["Yevmiye_Tutari", "Fark", "Hesaplar"]] \
        .rename(columns={"Yevmiye_Tutari": "Yevmiye_Tutari_TL", "Fark": "Fark_TL"}).reset_index(drop=True)

    if not matched.empty:
        matched["Fatura_Donemi"] = matched["issue_date"].map(lambda d: period_of(d, period_type))
        matched["Yevmiye_Donemi"] = matched["Yevmiye_Tarihi"].map(lambda d: period_of(d, period_type))
        period_diff = matched[matched["Fatura_Donemi"] != matched["Yevmiye_Donemi"]]
    else:
        period_diff = matched.assign(Fatura_Donemi=None, Yevmiye_Donemi=None)
    res["Dönem Farkları"] = period_diff.rename(columns=base_cols)[
        list(base_cols.values()) + ["Yevmiye_Tarihi", "Fatura_Donemi", "Yevmiye_Donemi"]].reset_index(drop=True)

    invoice_nos = set(inv["invoice_no_norm"])
    orphan = sel_grp[~sel_grp["document_no_norm"].isin(invoice_nos)]
    res["Faturası Bulunmayan Yevmiye Kayıtları"] = orphan[
        ["Yevmiye_Belge_No", "Yevmiye_Tarihi", "Yevmiye_Tutari", "Hesaplar"]] \
        .assign(Yevmiye_Tutari=lambda d: d["Yevmiye_Tutari"].astype(float).round(2)).reset_index(drop=True)

    res["Belirsiz Eşleşme (Aynı No Farklı Tedarikçi)"] = ambiguous.rename(columns=base_cols)[
        list(base_cols.values())].sort_values("Fatura_No").reset_index(drop=True)
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
    """Tüm kontrolleri çalıştırır. Dönüş: (özet DataFrame, OrderedDict(başlık → DataFrame), notlar)"""
    invoices = db.get_invoices_df()
    lines = db.get_lines_df()
    journal = db.get_journal_df()
    sections = OrderedDict()
    notes = []

    prices = price_anomalies(lines, period_type, threshold)
    sections[f"Fiyat Anomalileri (±%{threshold:g})"] = prices[prices["Risk_Durumu"] == "YÜKSEK RİSK"] \
        .reset_index(drop=True)

    if accounts and not journal.empty:
        sections.update(reconcile(invoices, journal, accounts, tolerance, period_type))
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
    return summary, sections, notes, {"fatura": len(invoices), "yevmiye": len(journal)}
