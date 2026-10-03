"""Denetçi Metin'in 1 aylık kullanımı: 40 firmayı programın kendi koduyla denetler ve sonuçları ölçer."""
import io
import json
import os
import sys
import time
from collections import Counter, defaultdict

import pandas as pd

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
from denetim import checks, importers, inceleme  # noqa: E402
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
    "Tutar Farkları": ["tutar_farki", "kdv_dahil_kayit", "cift_kayit", "otv_maliyete_eklenmemis"],
    "Dönem Farkları": ["donem_kaymasi"],
    "Faturası Bulunmayan Yevmiye Kayıtları": ["faturasiz_gider"],
    "Olası Mükerrer Faturalar": ["mukerrer_fatura"],
    "Fiyat Anomalileri": ["fiyat_sisirme"],
    "Alıcısı Firma Olmayan Faturalar": ["baska_firma_faturasi"],
    "Fatura Hesaplama Tutarsızlıkları": ["xml_hesap_hatasi"],
    # Madde 7: KDV (191), tevkifat (360) ve satış mutabakatı. Çift kayıtta 191 de iki kez yazılır; KDV'nin maliyete
    # eklendiği kayıtta (kdv_dahil_kayit) 191 kaydı yoktur.
    "KDV Farkları": ["kdv_farki", "cift_kayit"],
    "KDV'si Kaydedilmemiş Faturalar": ["kdv_kaydedilmemis", "kdv_dahil_kayit"],
    "Tevkifat Kaydı Eksik/Farklı": ["tevkifat_kaydi_eksik"],
    "Muhasebeleşmemiş Satış Faturaları": ["muhasebelesmemis_satis"],
    "Gelir Hesabı Dışına Kaydedilmiş Satış Faturaları": [],
    "Satış Tutar Farkları": ["satis_tutar_farki"],
    "Satış Dönem Farkları": ["satis_donem_kaymasi"],
    "Satış KDV Farkları": ["satis_kdv_farki"],
    "Faturası Bulunmayan Gelir Kayıtları": ["faturasiz_gelir"],
}
FATURASIZ = "Faturası Bulunmayan Yevmiye Kayıtları"
FATURASIZ_GELIR = "Faturası Bulunmayan Gelir Kayıtları"
BELGE_SUTUNLU = {FATURASIZ, FATURASIZ_GELIR}  # Bulgu fatura değil yevmiye belgesi
ESIKLER = (10, 15, 25, 35)
KUR_TOLERANSLARI = (0, 0.5, 1, 2)
DURUM_KISA = {inceleme.DURUM_ACIK: "acik", inceleme.DURUM_SORUN_YOK: "sorun_yok",
              inceleme.DURUM_DUZELTME: "duzeltme"}
MIN_ALIMLAR = (1, 2, 3, 4, 5)
# Bilinen bir hatayı temsil etmeyen, bilgi amaçlı bölümler: yanlış alarm sayılmaz, ayrıca raporlanır.
# "Belge No Uyuşmayan Eşleşmeler" belge_style="kisa" vb. firmalarda dolabilir (belge no yazım hatası bulgusu).
INFO_SECTIONS = {
    "Belge No Uyuşmayan Eşleşmeler": "belge_no_uyusmayan",
    "Belirsiz Eşleşme (Birden Fazla Seri+Sıra Adayı)": "belirsiz_seri_sira",
    "Satış Belge No Uyuşmayan Eşleşmeler": "satis_belge_no_uyusmayan",
    "Satış Belirsiz Eşleşme (Birden Fazla Seri+Sıra Adayı)": "satis_belirsiz_seri_sira",
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
        res = importers.read_invoice_excel(data, "faturalar.xlsx", company_vkn=meta["vkn"])
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
    if section in BELGE_SUTUNLU:
        doc = str(row["Yevmiye_Belge_No"])
        for p in ("BORDRO", "AMORT", "SMM", "GP-", "SAT"):
            if doc.startswith(p):
                return f"faturasız olağan kayıt ({p})"
        return "belge no biçimi eşleşmedi"
    if section == "Muhasebeleşmemiş Faturalar":
        return "belge no biçimi eşleşmedi" if meta["belge_style"] == "kisa" else "diğer"
    if section == "Tutar Farkları":
        if row["Fatura_No"] in foreign:
            return "dövizli fatura kur farkı"
        return f"diğer ({row['Olasi_Neden']})" if row.get("Olasi_Neden") else "diğer"
    if section in ("KDV Farkları", "KDV'si Kaydedilmemiş Faturalar"):
        return f"neden: {row.get('Olasi_Neden') or '-'}"
    if section in ("Tevkifat Kaydı Eksik/Farklı", "Satış KDV Farkları"):
        return f"durum: {row.get('Durum')}"
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
        expected = {normalize_doc_no(x) for t in exp_types for x in truth.get(t, [])}
        col = "Yevmiye_Belge_No" if key in BELGE_SUTUNLU else "Fatura_No"
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


def gercek_bulgu_maskesi(key, df, meta):
    """Bölümdeki her satır bilinen bir hatayı mı temsil ediyor? (score() ile aynı ölçüt, satır bazında)"""
    truth = meta["truth"]
    if df.empty:
        return []
    if key == "Olası Mükerrer Faturalar":
        pairs = [set(p) for p in truth["mukerrer_fatura"]]
        return [any(p <= set(str(s).split(", ")) for p in pairs) for s in df["Fatura_No"]]
    expected = {normalize_doc_no(x) for t in SECTION_TRUTH[key] for x in truth.get(t, [])}
    col = "Yevmiye_Belge_No" if key in BELGE_SUTUNLU else "Fatura_No"
    return [normalize_doc_no(x) in expected for x in df[col]]


def inceleme_durumu(meta, sections, kayitlar):
    """Bulgu bölümlerine inceleme durumlarını uygular; açık / sorun yok / düzeltme sayılarını gerçek hata ve yanlış
    alarm olarak ayırır. Dönüş: (sayılar, yanlış alarm anahtarları)"""
    say = Counter()
    yanlis = []
    for title, df in inceleme.bolumlere_uygula(sections, kayitlar).items():
        key = next((k for k in SECTION_TRUTH if title.startswith(k)), None)
        if key is None or df.empty:
            continue
        for gercek, durum, anahtar in zip(gercek_bulgu_maskesi(key, df, meta), df[inceleme.DURUM_COL],
                                          df[inceleme.ANAHTAR_COL]):
            say[("gercek_" if gercek else "yanlis_") + DURUM_KISA[durum]] += 1
            if not gercek:
                yanlis.append(anahtar)
        if key == "Olası Mükerrer Faturalar":
            for gercek, ardisik in zip(gercek_bulgu_maskesi(key, df, meta), df["Ardisik_Numara"]):
                say[f"mukerrer_{'gercek' if gercek else 'yanlis'}_ardisik_{ardisik}"] += 1
    return dict(say), yanlis


def alis_faturalari_df(db, vkn):
    """Alış mutabakatına giden faturalar (run_full_audit ile aynı: yalnızca alışlar, maliyete eklenen vergilerle)."""
    inv = checks.maliyet_vergisi_ekle(db.get_invoices_df(), db.get_invoice_taxes_df())
    return checks.alis_faturalari(inv, vkn)


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
    print("Tutar farkı, dövizli faturalarda yüzde kur toleransına göre (TL faturalarda yalnızca 0,01 TL tolerans)")
    for pct in KUR_TOLERANSLARI:
        a = [sum(r["kur"][pct][k] for r in all_results) for k in ("bulunan", "beklenen", "yanlis_alarm",
                                                                    "tolerans_ici")]
        print(f"  %{pct:<4g} {a[0]}/{a[1]} yanlış alarm={a[2]:4d} kur farkı (tolerans içi, bilgi)={a[3]}")
    t50 = [sum(r["tolerans50_tutar"][k] for r in all_results if "tolerans50_tutar" in r)
           for k in ("bulunan", "beklenen", "yanlis_alarm")]
    print(f"  (karşılaştırma: dövizli firmalarda tüm faturalara 50 TL sabit tolerans, %0 kur: {t50[0]}/{t50[1]} "
          f"yanlış alarm={t50[2]})")
    i1, i2 = Counter(), Counter()
    for r in all_results:
        i1.update(r["inceleme_1"])
        i2.update(r["inceleme_2"])
    print("Bulgu inceleme (Metin 1. çalıştırmada yanlış alarmları 'İncelendi – Sorun Yok' işaretler; veriler silinip "
          "yeniden yüklenir, rapor yeniden çalıştırılır)")
    print(f"  1. çalıştırma: açık yanlış alarm={i1['yanlis_acik']} açık gerçek hata={i1['gercek_acik']}  "
          f"→ işaretlenen={sum(r['isaretlenen'] for r in all_results)}")
    print(f"  2. çalıştırma: açık yanlış alarm={i2['yanlis_acik']} sorun yok (gizli)={i2['yanlis_sorun_yok']} "
          f"| gerçek hata açık={i2['gercek_acik']} sorun yok={i2['gercek_sorun_yok']} "
          f"düzeltme={i2['gercek_duzeltme']}  yakalanan={sum(r['score2_bulunan'] for r in all_results)}/"
          f"{sum(r['score2_beklenen'] for r in all_results)}")
    print(f"  Mükerrer, ardışık numara: gerçek {i1['mukerrer_gercek_ardisik_Evet']}/"
          f"{i1['mukerrer_gercek_ardisik_Evet'] + i1['mukerrer_gercek_ardisik_Hayır']} ardışık, yanlış alarm "
          f"{i1['mukerrer_yanlis_ardisik_Evet']}/"
          f"{i1['mukerrer_yanlis_ardisik_Evet'] + i1['mukerrer_yanlis_ardisik_Hayır']} ardışık")
    vo = Counter()
    for r in all_results:
        vo.update(r.get("vergi_ozeti") or {})
    mv = [sum((r.get("maliyet_vergisi") or {}).get(k, 0) for r in all_results) for k in ("fatura", "toplam_tl")]
    print(f"Vergi mutabakatı (alış): KDV karşılaştırılan={vo['kdv_karsilastirilan']} KDV farkı={vo['kdv_farki']} "
          f"KDV'si kaydedilmemiş={vo['kdv_yok']} | tevkifatlı={vo['tevkifatli']} tevkifat bulgusu={vo['tevkifat_bulgu']}"
          f" | maliyete eklenen vergili fatura={mv[0]} (toplam {mv[1]:,.0f} TL)")
    so, se, sf, sk = Counter(), Counter(), Counter(), Counter()
    n_satis_firma = 0
    for r in all_results:
        if r.get("satis_ozeti"):
            n_satis_firma += 1
            so["fatura"] += r["satis_ozeti"]["fatura"]
            so["kur_tolerans_ici"] += r["satis_ozeti"]["kur_tolerans_ici"]
            se.update(r["satis_ozeti"]["eslesme"])
            sf.update(r["satis_ozeti"]["faturasiz"])
            sk.update(r["satis_ozeti"]["kdv"])
    print(f"Satış mutabakatı: {n_satis_firma} firma, {so['fatura']} satış faturası, eşleştirme={dict(se)} | "
          f"faturasız gelir listelenen={sf['listelenen']} (borç yönlü elenen={sf['borc_yonlu']}, önek={sf['haric_onek']})"
          f" | KDV karşılaştırılan={sk['kdv_karsilastirilan']} (istisna/ihracat={sk['istisna']}) fark={sk['kdv_farki']}"
          f" | kur farkı (tolerans içi)={so['kur_tolerans_ici']}")
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
        sc, fp = score(meta, sections)
        log["faturasiz_ozeti"] = {k: v for k, v in counts["faturasiz_ozeti"].items() if k != "onekler"}
        # Önek listesinin katkısı: boş önek listesiyle (yalnızca borç yönü filtresi) aynı kontrol
        alis = alis_faturalari_df(db, firm.vkn)
        bos = checks.reconcile(alis, db.get_journal_df(), accounts, 0.01, "Aylık", haric_onekler=[])
        log["onek_bos_faturasiz"] = dict(score(meta, {FATURASIZ: bos[FATURASIZ]})[0][FATURASIZ],
                                         yuksek_oncelik=bos.faturasiz_ozeti["yuksek"])
        log["eslesme"] = dict(counts["eslesme"] or {})
        log["bilgi"] = info(sections)
        # Dövizli faturalarda yüzde kur toleransının etkisi (%0 = madde 6 öncesi davranış)
        log["kur_ozeti"] = counts["kur_ozeti"]
        log["kur"] = {}
        inv_df, jou_df = alis, db.get_journal_df()
        for pct in KUR_TOLERANSLARI:
            r = checks.reconcile(inv_df, jou_df, accounts, 0.01, "Aylık", kur_toleransi=pct)
            log["kur"][pct] = dict(score(meta, {"Tutar Farkları": r["Tutar Farkları"]})[0]["Tutar Farkları"],
                                   tolerans_ici=r.kur_ozeti["tolerans_ici"])
        # Metin'in ikinci denemesi (madde 6 öncesi): dövizli firmalarda tolerans 50 TL, yüzde kur toleransı yok
        if meta["foreign_invoices"]:
            _, sec2, _, _ = checks.run_full_audit(db, "Aylık", 15.0, accounts, 50.0, firm.vkn, kur_toleransi=0)
            sc2, _ = score(meta, sec2)
            log["tolerans50_tutar"] = sc2["Tutar Farkları"]
        # Fiyat analizi: eşik duyarlılığı varsayılan kurallarla ve boş anahtar kelime listesiyle (yalnızca
        # tevkifat + para birimi + yetersiz veri kuralları); en az alım sayısı duyarlılığı varsayılan kurallarla
        lines = checks.alis_faturalari(db.get_lines_df(), firm.vkn)
        log["fiyat_ozeti"] = {k: v for k, v in counts["fiyat_ozeti"].items() if k != "kelimeler"}
        log["esik"], log["esik_kelimesiz"], log["min_alim"] = {}, {}, {}
        for th in ESIKLER:
            log["esik"][th] = fiyat_puani(meta, lines, th, checks.FiyatKurallari())
            log["esik_kelimesiz"][th] = fiyat_puani(meta, lines, th, checks.FiyatKurallari(kelimeler=[]))
        for n in MIN_ALIMLAR:
            log["min_alim"][n] = fiyat_puani(meta, lines, 15, checks.FiyatKurallari(min_alim=n))
        # Bulgu inceleme: Metin yanlış alarmları "İncelendi – Sorun Yok" işaretler; ay sonunda muhasebeciden gelen
        # dosyalar yeniden yüklenir (veriler silinir, aynı dosyalar tekrar içe aktarılır) ve rapor yeniden çalışır.
        log["inceleme_1"], yanlis = inceleme_durumu(meta, sections, inceleme.incelemeleri_oku(db))
        log["isaretlenen"] = inceleme.isaretle(db, yanlis, inceleme.DURUM_SORUN_YOK,
                                               "Metin: incelendi, meşru (kur farkı / aynı gün araç alımı / fiyat dalgalanması)")
        db.clear_data()
        import_firm(meta, db, {})
        summary, sections, notes, counts = checks.run_full_audit(db, "Aylık", 15.0, accounts, 0.01, firm.vkn)
        kayitlar = inceleme.incelemeleri_oku(db)
        log["inceleme_2"], _ = inceleme_durumu(meta, sections, kayitlar)
        sc2, _ = score(meta, sections)
        log["score2_bulunan"] = sum(v["bulunan"] for v in sc2.values())
        log["score2_beklenen"] = sum(v["beklenen"] for v in sc2.values())
        incelenen = inceleme.bolumlere_uygula(sections, kayitlar)
        st = counts["satis"]
        export_sections(os.path.join(WORK, f"{meta['code']}_Denetim_Raporu.xlsx"),
                        dict([("Özet", summary), ("İnceleme Özeti", inceleme.inceleme_ozeti(incelenen)),
                              ("Eşleşme Özeti", checks.eslesme_ozeti_df(counts["eslesme"])),
                              ("Faturasız Kayıt Özeti", checks.faturasiz_ozeti_df(counts["faturasiz_ozeti"])),
                              ("Fiyat Analizi Özeti", checks.fiyat_ozeti_df(counts["fiyat_ozeti"]))]
                             + ([("Satış Eşleşme Özeti", checks.eslesme_ozeti_df(st["eslesme"])),
                                 ("Faturasız Gelir Özeti", checks.faturasiz_ozeti_df(st["faturasiz_ozeti"]))]
                                if st else [])
                             + list(incelenen.items())
                             + [("Kur Farkı (Tolerans İçi)", counts["kur_farki"]),
                                ("Faturasız Listeden Hariç Tutulanlar", counts["faturasiz_haric"]),
                                ("Fiyat Analizi Dışı Satırlar", counts["fiyat_haric"])]
                             + ([("Satış Kur Farkı (Tolerans İçi)", st["kur_farki"]),
                                 ("Faturasız Gelir Listesinden Hariç Tutulanlar", st["faturasiz_haric"])] if st else [])))
        log["vergi_ozeti"] = {k: v for k, v in (counts["vergi_ozeti"] or {}).items() if not isinstance(v, list)}
        log["satis_ozeti"] = None if st is None else {
            "fatura": st["fatura"], "eslesme": dict(st["eslesme"]),
            "faturasiz": {k: v for k, v in st["faturasiz_ozeti"].items() if k not in ("onekler", "isaretli", "yon")},
            "kdv": {k: v for k, v in st["vergi_ozeti"].items() if not isinstance(v, list)},
            "kur_tolerans_ici": st["kur_ozeti"]["tolerans_ici"]}
        log["maliyet_vergisi"] = counts["maliyet_vergisi"]
        log["score"] = sc
        log["fp_causes"] = {k: dict(v) for k, v in fp.items()}
        all_results.append(log)
        tp = sum(v["bulunan"] for v in sc.values())
        ex = sum(v["beklenen"] for v in sc.values())
        fpn = sum(v["yanlis_alarm"] for v in sc.values())
        print(f"{meta['code']} {meta['sector'][:12]:12s} fatura={log['added']:5d} yevmiye={log['journal_rows']:5d} "
              f"elle={','.join(log['friction']) or '-':32s} yakalanan={tp:3d}/{ex:3d} yanlış_alarm={fpn:4d} "
              f"belge_no_uyuşmayan={log['bilgi'].get('belge_no_uyusmayan', 0):3d} "
              f"2.çalıştırma_açık_yanlış_alarm={log['inceleme_2'].get('yanlis_acik', 0):3d} "
              f"süre(xml/j/rapor)={log['t_xml']}/{log['t_journal']}/{log['t_audit']}s")
    print_totals(all_results)
    json.dump(all_results, open(os.path.join(BASE, "sonuclar.json"), "w", encoding="utf-8"), ensure_ascii=False,
              indent=1, default=str)


if __name__ == "__main__":
    main()
