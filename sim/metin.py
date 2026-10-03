"""Denetçi Metin'in 1 aylık kullanımı: 40 firmayı programın kendi koduyla denetler ve sonuçları ölçer."""
import io
import json
import os
import sys
import time
from collections import Counter, defaultdict

import pandas as pd

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
from denetim import checks, importers  # noqa: E402
from denetim.export import export_sections  # noqa: E402
from denetim.firms import FirmRegistry  # noqa: E402
from denetim.utils import normalize_doc_no, parse_account_list  # noqa: E402

BASE = os.path.dirname(os.path.abspath(__file__))
FIRMS = os.path.join(BASE, "firmalar")
WORK = os.path.join(BASE, "calisma")
os.makedirs(WORK, exist_ok=True)

SECTION_TRUTH = {
    "Muhasebeleşmemiş Faturalar": ["muhasebelesmemis"],
    "Seçili Hesap Dışına Kaydedilmiş Faturalar": ["yanlis_hesap"],
    "Tutar Farkları": ["tutar_farki", "kdv_dahil_kayit", "cift_kayit"],
    "Dönem Farkları": ["donem_kaymasi"],
    "Faturası Bulunmayan Yevmiye Kayıtları": ["faturasiz_gider"],
    "Olası Mükerrer Faturalar": ["mukerrer_fatura"],
    "Fiyat Anomalileri": ["fiyat_sisirme"],
    "Alıcısı Firma Olmayan Faturalar": ["baska_firma_faturasi"],
    "Fatura Hesaplama Tutarsızlıkları": ["xml_hesap_hatasi"],
}
# Elle müdahale süresi tahmini (dakika)
FRICTION_MIN = {"baslik_satiri": 5, "sutun_adi": 5, "belge_no_aciklamada": 20, "firma_degisimi": 3}


def import_firm(meta, db, log):
    d = os.path.join(FIRMS, meta["code"])
    t0 = time.time()
    added = dup = 0
    errors, warnings = [], []
    xml_paths = []
    if os.path.exists(os.path.join(d, "efatura_2024Q3.zip")):
        xml_paths = [os.path.join(d, "efatura_2024Q3.zip")]
    elif os.path.isdir(os.path.join(d, "xml")):
        xml_paths = sorted(os.path.join(d, "xml", f) for f in os.listdir(os.path.join(d, "xml")))
    for path in xml_paths:  # GUI'deki import_xml_paths ile aynı akış
        for name, data in importers.iter_xml_payloads(path):
            digest = importers.file_hash(data)
            if db.find_import("XML", digest):
                continue
            res = importers.read_xml(data, name)
            errors += res.errors
            warnings += res.warnings
            if res.items:
                n, dups = db.save_invoices(res.items, "XML", name, digest)
                added += n
                dup += len(dups)
    t_xml = time.time() - t0

    t1 = time.time()
    xl = os.path.join(d, "faturalar.xlsx")
    if os.path.exists(xl):
        data = open(xl, "rb").read()
        res = importers.read_invoice_excel(data, "faturalar.xlsx")
        errors += res.errors
        warnings += res.warnings
        n, dups = db.save_invoices(res.items, "FATURA_EXCEL", "faturalar.xlsx", importers.file_hash(data))
        added += n
        dup += len(dups)
    t_xl = time.time() - t1

    # Yevmiye: önce dosyayı olduğu gibi dene, olmazsa Metin'in elle yaptığı düzeltmeler
    t2 = time.time()
    friction = []
    jpath = os.path.join(d, "yevmiye.xlsx")
    data = open(jpath, "rb").read()
    res = importers.read_journal_excel(data)
    attempts = [("ilk deneme", res.errors[:1])]
    if not res.items and res.errors:
        df = pd.read_excel(jpath, dtype=object)
        if df.columns[0] != "Tarih" and str(df.columns[0]).startswith(meta["name"][:5]):
            df = pd.read_excel(jpath, dtype=object, header=3)
            friction.append("baslik_satiri")
        if "Borç Tutarı" in df.columns:
            df = df.rename(columns={"Borç Tutarı": "Borç", "Alacak Tutarı": "Alacak"})
            friction.append("sutun_adi")
        if not any(c in df.columns for c in ("Belge No", "Evrak No", "Belge_No")) and "Açıklama" in df.columns:
            df["Belge_No"] = df["Açıklama"].astype(str).str.split(" - ").str[0]
            friction.append("belge_no_aciklamada")
        buf = io.BytesIO()
        df.to_excel(buf, index=False)
        data = buf.getvalue()
        res = importers.read_journal_excel(data)
        attempts.append(("düzeltilmiş dosya", res.errors[:1]))
    saved = db.save_journal(res.items, "yevmiye.xlsx", importers.file_hash(data)) if res.items else 0
    t_j = time.time() - t2
    log.update(added=added, dup=dup, import_errors=len(errors), import_warnings=len(warnings),
               journal_rows=saved, journal_errors=len(res.errors), friction=friction, attempts=attempts,
               t_xml=round(t_xml, 1), t_excel=round(t_xl, 1), t_journal=round(t_j, 1),
               sample_errors=errors[:3], sample_warnings=warnings[:3])


def classify_fp(section, row, meta):
    """Yanlış alarmın nedenini sınıflandırır."""
    foreign = set(meta["foreign_invoices"])
    if section == "Faturası Bulunmayan Yevmiye Kayıtları":
        doc = str(row["Yevmiye_Belge_No"])
        for p in ("BORDRO", "AMORT", "SMM", "GP-"):
            if doc.startswith(p):
                return f"faturasız olağan kayıt ({p})"
        return "belge no biçimi eşleşmedi"
    if section == "Muhasebeleşmemiş Faturalar":
        return "belge no biçimi eşleşmedi" if meta["belge_style"] == "kisa" else "diğer"
    if section == "Tutar Farkları":
        if row["Fatura_No"] in foreign:
            return "dövizli fatura kur farkı"
        return "diğer"
    if section == "Fiyat Anomalileri":
        return "ürün: " + str(row["Urun_Adi"])
    if section == "Olası Mükerrer Faturalar":
        return "aynı gün aynı fiyat araç alımı" if meta["sector"].startswith("Otomotiv") else "diğer"
    return "diğer"


def score(meta, sections):
    truth = meta["truth"]
    res = {}
    fp_causes = defaultdict(Counter)
    for title, df in sections.items():
        key = next((k for k in SECTION_TRUTH if title.startswith(k)), None)
        if key is None:
            continue
        exp_types = SECTION_TRUTH[key]
        if key == "Olası Mükerrer Faturalar":
            pairs = [tuple(p) for p in truth["mukerrer_fatura"]]
            found_sets = [set(str(s).split(", ")) for s in df["Fatura_No"]] if not df.empty else []
            tp = sum(1 for p in pairs if any(set(p) <= s for s in found_sets))
            fp_rows = [r for s, (_, r) in zip(found_sets, df.iterrows()) if not any(set(p) <= s for p in pairs)]
            res[key] = {"beklenen": len(pairs), "bulunan": tp, "yanlis_alarm": len(fp_rows)}
            for r in fp_rows:
                fp_causes[key][classify_fp(key, r, meta)] += 1
            continue
        expected = {normalize_doc_no(x) for t in exp_types for x in truth[t]}
        col = "Yevmiye_Belge_No" if key == "Faturası Bulunmayan Yevmiye Kayıtları" else "Fatura_No"
        found = {normalize_doc_no(x) for x in df[col]} if not df.empty else set()
        tp = expected & found
        res[key] = {"beklenen": len(expected), "bulunan": len(tp), "yanlis_alarm": len(found - expected)}
        if not df.empty:
            for _, r in df.iterrows():
                if normalize_doc_no(r[col]) not in expected:
                    fp_causes[key][classify_fp(key, r, meta)] += 1
        if key == "Fiyat Anomalileri":
            res[key]["yanlis_alarm_satir"] = int(sum(1 for x in df[col] if normalize_doc_no(x) not in expected))
    return res, fp_causes


def main():
    metas = json.load(open(os.path.join(FIRMS, "firmalar.json"), encoding="utf-8"))
    all_results = []
    # Metin her firmayı programın firma seçicisiyle açar: firma başına ayrı veritabanı (sim/calisma/veri/)
    registry = FirmRegistry(os.path.join(WORK, "veri"))
    for meta in metas:
        if registry.get(meta["code"]):
            registry.delete(meta["code"])  # Her çalıştırma temiz başlar
        firm = registry.create(meta["code"], meta["name"], meta["vkn"], meta["sector"])
        registry.set_last_firm(firm.code)
        db = registry.open_db(firm.code)
        log = {"code": meta["code"], "sector": meta["sector"], "name": meta["name"], "fatura": meta["n_invoices"],
               "format": meta["journal_format"], "belge": meta["belge_style"], "kaynak": meta["source"]}
        import_firm(meta, db, log)
        accounts = parse_account_list(",".join(meta["accounts"]))
        t = time.time()
        summary, sections, notes, counts = checks.run_full_audit(db, "Aylık", 15.0, accounts, 0.01, firm.vkn)
        log["t_audit"] = round(time.time() - t, 1)
        export_sections(os.path.join(WORK, f"{meta['code']}_Denetim_Raporu.xlsx"),
                        dict([("Özet", summary)] + list(sections.items())))
        sc, fp = score(meta, sections)
        # Metin'in ikinci denemesi: dövizli firmalarda tolerans 50 TL
        if meta["foreign_invoices"]:
            _, sec2, _, _ = checks.run_full_audit(db, "Aylık", 15.0, accounts, 50.0, firm.vkn)
            sc2, _ = score(meta, sec2)
            log["tolerans50_tutar"] = sc2["Tutar Farkları"]
        # Eşik duyarlılığı
        log["esik"] = {}
        for th in (10, 15, 25, 35):
            prices = checks.price_anomalies(db.get_lines_df(), "Aylık", th)
            risky = prices[prices["Risk_Durumu"] == "YÜKSEK RİSK"]
            exp = {normalize_doc_no(x) for x in meta["truth"]["fiyat_sisirme"]}
            got = {normalize_doc_no(x) for x in risky["Fatura_No"]}
            log["esik"][th] = {"bulunan": len(exp & got), "beklenen": len(exp), "yanlis_alarm": len(got - exp)}
        log["score"] = sc
        log["fp_causes"] = {k: dict(v) for k, v in fp.items()}
        all_results.append(log)
        tp = sum(v["bulunan"] for v in sc.values())
        ex = sum(v["beklenen"] for v in sc.values())
        fpn = sum(v["yanlis_alarm"] for v in sc.values())
        print(f"{meta['code']} {meta['sector'][:12]:12s} fatura={log['added']:5d} yevmiye={log['journal_rows']:5d} "
              f"elle={','.join(log['friction']) or '-':32s} yakalanan={tp:3d}/{ex:3d} yanlış_alarm={fpn:4d} "
              f"süre(xml/j/rapor)={log['t_xml']}/{log['t_journal']}/{log['t_audit']}s")
    json.dump(all_results, open(os.path.join(BASE, "sonuclar.json"), "w", encoding="utf-8"), ensure_ascii=False,
              indent=1, default=str)


if __name__ == "__main__":
    main()
