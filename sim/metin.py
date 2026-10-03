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
FATURASIZ = "Faturası Bulunmayan Yevmiye Kayıtları"
ESIKLER = (10, 15, 25, 35)
MIN_ALIMLAR = (1, 2, 3, 4, 5)
# Bilinen bir hatayı temsil etmeyen, bilgi amaçlı bölümler: yanlış alarm sayılmaz, ayrıca raporlanır.
# "Belge No Uyuşmayan Eşleşmeler" belge_style="kisa" vb. firmalarda dolabilir (belge no yazım hatası bulgusu).
INFO_SECTIONS = {
    "Belge No Uyuşmayan Eşleşmeler": "belge_no_uyusmayan",
    "Belirsiz Eşleşme (Birden Fazla Seri+Sıra Adayı)": "belirsiz_seri_sira",
}
# Elle müdahale süresi tahmini (dakika). Başlık satırı ve "Borç Tutarı" gibi sütun adlarını program kendisi
# çözdüğü için (madde 3) yalnızca belge no'su ayrı sütunda olmayan dökümler için muhasebeciden yeni döküm istenir.
FRICTION_MIN = {"belge_no_ayri_sutun_istendi": 20, "firma_degisimi": 3}


def ayri_belge_no_dokumu(path):
    """Muhasebecinin Belge No sütunu eklenmiş yeni dökümünü temsil eder (simülasyonda açıklamadan türetilir;
    program bunu yapmaz, belge numarasını yalnızca ayrı bir sütundan okur)."""
    df = pd.read_excel(path, dtype=object)
    df.insert(df.columns.get_loc("Tarih") + 1, "Belge No", df["Açıklama"].astype(str).str.split(" - ").str[0])
    df["Açıklama"] = df["Açıklama"].astype(str).str.split(" - ", n=1).str[1]
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue()


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

    # Yevmiye: dosya olduğu gibi yüklenir (başlık satırı ve sütun adlarını program kendisi bulur). Belge no ayrı
    # sütunda değilse program açık bir hata verir; Metin muhasebeciden Belge No sütunlu yeni döküm ister.
    t2 = time.time()
    friction = []
    jpath = os.path.join(d, "yevmiye.xlsx")
    data = open(jpath, "rb").read()
    res = importers.read_journal_excel(data)
    attempts = [("ilk deneme", res.errors[:1])]
    if not res.items and "Belge_No" in res.missing:
        friction.append("belge_no_ayri_sutun_istendi")
        data = ayri_belge_no_dokumu(jpath)
        res = importers.read_journal_excel(data)
        attempts.append(("muhasebeciden yeni döküm", res.errors[:1]))
    elif not res.items:
        friction.append("yuklenemedi")
    saved = db.save_journal(res.items, "yevmiye.xlsx", importers.file_hash(data)) if res.items else 0
    t_j = time.time() - t2
    log.update(added=added, dup=dup, import_errors=len(errors), import_warnings=len(warnings),
               journal_rows=saved, journal_errors=len(res.errors), friction=friction, attempts=attempts,
               journal_infos=res.infos,
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


def fiyat_puani(meta, lines, threshold, kurallar):
    """Fiyat şişirme yakalama / yanlış alarm (fatura bazında) ve eşik üstü olup yetersiz veri sayılan satırlar."""
    res = checks.fiyat_analizi(lines, "Aylık", threshold, kurallar)
    exp = {normalize_doc_no(x) for x in meta["truth"]["fiyat_sisirme"]}
    got = {normalize_doc_no(x) for x in res.riskli["Fatura_No"]}
    return {"bulunan": len(exp & got), "beklenen": len(exp), "yanlis_alarm": len(got - exp),
            "yetersiz_veri": res.ozet["yetersiz_veri"]}


def info(sections):
    """Bilgi amaçlı bölüm sayıları ve tutar+tarih eşleşmelerinin belge no anahtarına göre doğruluğu."""
    out = {}
    for title, df in sections.items():
        key = next((v for k, v in INFO_SECTIONS.items() if title.startswith(k)), None)
        if key is None:
            continue
        out[key] = len(df)
        if key == "belge_no_uyusmayan" and not df.empty:
            # Simülasyonda belge no fatura no'dan türetilir: anahtarlar uyuşuyorsa eşleşme doğrudur
            ok = sum(1 for _, r in df.iterrows()
                     if any(checks.anahtar_uyumlu(a, b) for a in checks.belge_anahtarlari(r["Fatura_No"])
                            for b in checks.belge_anahtarlari(r["Yevmiye_Belge_No"])))
            out["belge_no_uyusmayan_anahtar_uyumsuz"] = len(df) - ok
    return out


def print_totals(all_results):
    """Kontrol bazında toplam yakalama / yanlış alarm ve belge no biçimine göre kırılım."""
    tot = defaultdict(lambda: [0, 0, 0])
    style = defaultdict(lambda: [0, 0, 0, 0])
    bilgi = defaultdict(Counter)
    yontem = Counter()
    for r in all_results:
        for k, v in r["score"].items():
            t = tot[k]
            t[0] += v["bulunan"]; t[1] += v["beklenen"]; t[2] += v["yanlis_alarm"]
            st = style[r["belge"]]
            st[0] += v["bulunan"]; st[1] += v["beklenen"]; st[2] += v["yanlis_alarm"]
        style[r["belge"]][3] += 1
        bilgi[r["belge"]].update(r.get("bilgi", {}))
        yontem.update(r.get("eslesme") or {})
    print("\nKontrol bazında toplam")
    for k, (a, b, c) in tot.items():
        print(f"  {k:45s} {a:3d}/{b:3d}  yanlış alarm={c:5d}")
    print(f"  {'TOPLAM':45s} {sum(v[0] for v in tot.values()):3d}/{sum(v[1] for v in tot.values()):3d}  "
          f"yanlış alarm={sum(v[2] for v in tot.values()):5d}")
    print("Belge no biçimine göre")
    for k, (a, b, c, n) in style.items():
        print(f"  {k:9s} firma={n:2d} yakalanan={a:3d}/{b:3d} yanlış alarm={c:5d} bilgi={dict(bilgi[k])}")
    print("Eşleştirme yöntemleri:", dict(yontem))
    fz = Counter()
    for r in all_results:
        fz.update({k: v for k, v in r["faturasiz_ozeti"].items() if k != "isaretli"})
    bos = [sum(r["onek_bos_faturasiz"][k] for r in all_results) for k in ("bulunan", "beklenen", "yanlis_alarm")]
    print(f"Faturasız kayıt filtresi: listelenen={fz['listelenen']} (Yüksek={fz['yuksek']}, Düşük={fz['dusuk']}) "
          f"hariç: alacak yönlü={fz['alacak_yonlu']}, önek={fz['haric_onek']}  "
          f"işaretsiz yevmiye={sum(1 for r in all_results if not r['faturasiz_ozeti']['isaretli'])}")
    print(f"Faturasız kayıt, boş önek listesiyle (yalnızca borç yönü): {bos[0]}/{bos[1]} yanlış alarm={bos[2]} "
          f"(Yüksek öncelikli: {sum(r['onek_bos_faturasiz']['yuksek_oncelik'] for r in all_results)})")
    fo = Counter()
    for r in all_results:
        fo.update({k: v for k, v in r["fiyat_ozeti"].items() if isinstance(v, int) and not isinstance(v, bool)})
    print(f"Fiyat analizi (varsayılan kurallar): incelenen={fo['incelenen']} riskli={fo['riskli']} analiz dışı: "
          f"iade={fo['iade']}, miktar={fo['miktar']}, tevkifat={fo['tevkifat']}, anahtar kelime={fo['kelime']}  "
          f"yetersiz veri={fo['yetersiz_veri']}")

    def topla(alan, anahtar):
        return [sum(r[alan][anahtar][k] for r in all_results)
                for k in ("bulunan", "beklenen", "yanlis_alarm", "yetersiz_veri")]
    print("Fiyat şişirme eşik duyarlılığı (yakalanan / yanlış alarm [yetersiz veri satırı])")
    for th in ESIKLER:
        a, b = topla("esik", th), topla("esik_kelimesiz", th)
        print(f"  %{th:<3d} varsayılan kelimeler: {a[0]}/{a[1]} yanlış alarm={a[2]:4d} [{a[3]}]   "
              f"boş kelime listesi: {b[0]}/{b[1]} yanlış alarm={b[2]:4d} [{b[3]}]")
    print("Fiyat şişirme, en az alım sayısına göre (%15, varsayılan kelimeler)")
    for n in MIN_ALIMLAR:
        a = topla("min_alim", n)
        print(f"  en az {n}: {a[0]}/{a[1]} yanlış alarm={a[2]:4d} yetersiz veri satırı={a[3]}")
    elle = Counter(f for r in all_results for f in r["friction"])
    n_elle = sum(1 for r in all_results if r["friction"])
    print(f"Elle müdahale gereken yevmiye dosyası: {n_elle}/{len(all_results)} {dict(elle)}  "
          f"(tahmini {sum(FRICTION_MIN.get(k, 0) * v for k, v in elle.items())} dk)")


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
                        dict([("Özet", summary), ("Eşleşme Özeti", checks.eslesme_ozeti_df(counts["eslesme"])),
                              ("Faturasız Kayıt Özeti", checks.faturasiz_ozeti_df(counts["faturasiz_ozeti"])),
                              ("Fiyat Analizi Özeti", checks.fiyat_ozeti_df(counts["fiyat_ozeti"]))]
                             + list(sections.items())
                             + [("Faturasız Listeden Hariç Tutulanlar", counts["faturasiz_haric"]),
                                ("Fiyat Analizi Dışı Satırlar", counts["fiyat_haric"])]))
        sc, fp = score(meta, sections)
        log["faturasiz_ozeti"] = {k: v for k, v in counts["faturasiz_ozeti"].items() if k != "onekler"}
        # Önek listesinin katkısı: boş önek listesiyle (yalnızca borç yönü filtresi) aynı kontrol
        bos = checks.reconcile(db.get_invoices_df(), db.get_journal_df(), accounts, 0.01, "Aylık", haric_onekler=[])
        log["onek_bos_faturasiz"] = dict(score(meta, {FATURASIZ: bos[FATURASIZ]})[0][FATURASIZ],
                                         yuksek_oncelik=bos.faturasiz_ozeti["yuksek"])
        log["eslesme"] = dict(counts["eslesme"] or {})
        log["bilgi"] = info(sections)
        # Metin'in ikinci denemesi: dövizli firmalarda tolerans 50 TL
        if meta["foreign_invoices"]:
            _, sec2, _, _ = checks.run_full_audit(db, "Aylık", 15.0, accounts, 50.0, firm.vkn)
            sc2, _ = score(meta, sec2)
            log["tolerans50_tutar"] = sc2["Tutar Farkları"]
        # Fiyat analizi: eşik duyarlılığı varsayılan kurallarla ve boş anahtar kelime listesiyle (yalnızca
        # tevkifat + para birimi + yetersiz veri kuralları); en az alım sayısı duyarlılığı varsayılan kurallarla
        lines = db.get_lines_df()
        log["fiyat_ozeti"] = {k: v for k, v in counts["fiyat_ozeti"].items() if k != "kelimeler"}
        log["esik"], log["esik_kelimesiz"], log["min_alim"] = {}, {}, {}
        for th in ESIKLER:
            log["esik"][th] = fiyat_puani(meta, lines, th, checks.FiyatKurallari())
            log["esik_kelimesiz"][th] = fiyat_puani(meta, lines, th, checks.FiyatKurallari(kelimeler=[]))
        for n in MIN_ALIMLAR:
            log["min_alim"][n] = fiyat_puani(meta, lines, 15, checks.FiyatKurallari(min_alim=n))
        log["score"] = sc
        log["fp_causes"] = {k: dict(v) for k, v in fp.items()}
        all_results.append(log)
        tp = sum(v["bulunan"] for v in sc.values())
        ex = sum(v["beklenen"] for v in sc.values())
        fpn = sum(v["yanlis_alarm"] for v in sc.values())
        print(f"{meta['code']} {meta['sector'][:12]:12s} fatura={log['added']:5d} yevmiye={log['journal_rows']:5d} "
              f"elle={','.join(log['friction']) or '-':32s} yakalanan={tp:3d}/{ex:3d} yanlış_alarm={fpn:4d} "
              f"belge_no_uyuşmayan={log['bilgi'].get('belge_no_uyusmayan', 0):3d} "
              f"süre(xml/j/rapor)={log['t_xml']}/{log['t_journal']}/{log['t_audit']}s")
    print_totals(all_results)
    json.dump(all_results, open(os.path.join(BASE, "sonuclar.json"), "w", encoding="utf-8"), ensure_ascii=False,
              indent=1, default=str)


if __name__ == "__main__":
    main()
