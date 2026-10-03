"""40 firma için 2024 Q3 sentetik denetim verisi üretir (XML/Excel fatura, yevmiye, hata listesi)."""
import json
import os
import random
import shutil
import sys
import zipfile
from datetime import date, timedelta

import pandas as pd

random.seed(2024)
# İsteğe bağlı ilk argüman: çıktı klasörü (varsayılan sim/firmalar)
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(__file__), "firmalar")
START, END = date(2024, 7, 1), date(2024, 9, 30)
MONTHLY_INFLATION = 0.03

# (ad, birim, baz fiyat TL, kdv, hesap, tür)  tür: mal / hizmet / tevkifat / arac
CATALOG = {
    "İnşaat": [
        ("Nervürlü İnşaat Demiri Ø12", "TNE", 22000, 20, "170", "mal"),
        ("Çimento CEM I 42.5", "TNE", 3200, 20, "170", "mal"),
        ("Hazır Beton C30", "MTQ", 2600, 20, "170", "mal"),
        ("Tuğla 13.5'luk", "C62", 9, 20, "170", "mal"),
        ("Kalıp Kerestesi", "MTQ", 14000, 20, "170", "mal"),
        ("Taşeron İşçilik Hakedişi", "C62", 450000, 20, "170", "tevkifat"),
        ("İş Makinesi Kiralama", "DAY", 12000, 20, "170", "hizmet"),
        ("Motorin", "LTR", 43, 20, "770", "mal"),
    ],
    "Halı Üretimi": [
        ("Polipropilen İplik", "KGM", 95, 10, "150", "mal"),
        ("Akrilik İplik", "KGM", 160, 10, "150", "mal"),
        ("Jüt Taban", "MTK", 28, 20, "150", "mal"),
        ("Lateks", "KGM", 55, 20, "150", "mal"),
        ("Boya Kimyasalı", "KGM", 210, 20, "150", "mal"),
        ("Ambalaj Rulo", "C62", 35, 20, "150", "mal"),
        ("Elektrik Enerjisi", "KWH", 3.2, 20, "730", "hizmet"),
        ("Fason Dokuma Hizmeti", "MTK", 85, 20, "730", "tevkifat"),
    ],
    "Uluslararası Taşımacılık": [
        ("Motorin", "LTR", 43, 20, "740", "mal"),
        ("AdBlue", "LTR", 18, 20, "740", "mal"),
        ("Lastik 315/80 R22.5", "C62", 14500, 20, "740", "mal"),
        ("Otoyol Köprü Geçiş", "C62", 650, 20, "740", "hizmet"),
        ("Feribot Taşıma Bedeli", "C62", 38000, 0, "740", "hizmet"),
        ("Tır Bakım Onarım", "C62", 25000, 20, "740", "hizmet"),
    ],
    "Otomotiv Satış ve Kiralama": [
        ("Binek Otomobil Model A 1.5", "C62", 1450000, 20, "153", "arac"),
        ("Binek Otomobil Model B Hybrid", "C62", 1980000, 20, "153", "arac"),
        ("Hafif Ticari Araç", "C62", 1250000, 20, "153", "arac"),
        ("Kiralık Filo Aracı Model A", "C62", 1300000, 20, "254", "arac"),
        ("Fren Balatası", "SET", 3200, 20, "153", "mal"),
        ("Motor Yağı 5W30", "LTR", 420, 20, "153", "mal"),
        ("Binek Lastik 205/55 R16", "C62", 4200, 20, "153", "mal"),
    ],
    "Muhtelif İmalat": [
        ("DKP Sac Levha", "KGM", 34, 20, "150", "mal"),
        ("Alüminyum Profil", "KGM", 165, 20, "150", "mal"),
        ("Plastik Granül", "KGM", 62, 20, "150", "mal"),
        ("Rulman 6204", "C62", 280, 20, "150", "mal"),
        ("Elektrostatik Toz Boya", "KGM", 240, 20, "150", "mal"),
        ("Ambalaj Koli", "C62", 18, 20, "150", "mal"),
        ("Elektrik Enerjisi", "KWH", 3.2, 20, "730", "hizmet"),
        ("Nakliye Hizmeti", "C62", 8500, 20, "760", "hizmet"),
    ],
}
SECTOR_COUNTS = {"İnşaat": 8, "Halı Üretimi": 7, "Uluslararası Taşımacılık": 8,
                 "Otomotiv Satış ve Kiralama": 8, "Muhtelif İmalat": 9}
NAME_PARTS = {
    "İnşaat": ["Yapı", "İnşaat", "Taahhüt", "Konut"], "Halı Üretimi": ["Halı", "Tekstil", "Dokuma", "Kilim"],
    "Uluslararası Taşımacılık": ["Lojistik", "Nakliyat", "Transport", "Taşımacılık"],
    "Otomotiv Satış ve Kiralama": ["Otomotiv", "Oto Kiralama", "Motorlu Araçlar", "Filo"],
    "Muhtelif İmalat": ["Makina", "Metal", "Plastik", "Sanayi"],
}
CITIES = ["Gaziantep", "Kocaeli", "Bursa", "Ankara", "İstanbul", "Kayseri", "Konya", "Mersin", "Uşak", "Denizli"]
SURNAMES = ["Yılmaz", "Kaya", "Demir", "Şahin", "Çelik", "Öztürk", "Aydın", "Arslan", "Doğan", "Kılıç", "Aslan",
            "Koç", "Kurt", "Özdemir", "Polat", "Erdoğan", "Güneş", "Bulut", "Aksoy", "Tekin"]
# Muhasebe programı yevmiye döküm biçimleri
JOURNAL_FORMATS = ["standart", "luca", "logo", "zirve", "baslikli"]
FORMAT_WEIGHTS = [0.25, 0.25, 0.2, 0.15, 0.15]
BELGE_STYLES = ["tam", "bosluklu", "kisa"]
BELGE_WEIGHTS = [0.7, 0.18, 0.12]
ERROR_TYPES = ["muhasebelesmemis", "yanlis_hesap", "tutar_farki", "kdv_dahil_kayit", "cift_kayit", "donem_kaymasi",
               "faturasiz_gider", "mukerrer_fatura", "fiyat_sisirme", "baska_firma_faturasi", "xml_hesap_hatasi",
               # Madde 7
               "otv_maliyete_eklenmemis", "kdv_farki", "kdv_kaydedilmemis", "tevkifat_kaydi_eksik",
               "muhasebelesmemis_satis", "satis_tutar_farki", "satis_kdv_farki", "satis_donem_kaymasi",
               "faturasiz_gelir"]
# Motorlu taşıt ÖTV oranları (9077, ÖTV II; matrah = KDV hariç liste fiyatı). Bayi alımında ÖTV maliyete eklenir.
OTV_ORANLARI = {"Binek Otomobil Model A 1.5": 80, "Binek Otomobil Model B Hybrid": 60, "Hafif Ticari Araç": 25,
                "Kiralık Filo Aracı Model A": 80}
# Satış kalemleri: (ad, birim, KDV, tür) tür: yurtici / ihracat (IHRACAT senaryosu, ISTISNA) / istisna (ISTISNA)
SATIS_KATALOG = {
    "İnşaat": [("Konut Satışı", "C62", 20, "yurtici"), ("Taahhüt Hakediş Bedeli", "C62", 20, "yurtici")],
    "Halı Üretimi": [("Makine Halısı", "MTK", 10, "yurtici"), ("Makine Halısı (İhracat)", "MTK", 0, "ihracat")],
    "Uluslararası Taşımacılık": [("Yurtiçi Taşıma Hizmeti", "C62", 20, "yurtici"),
                                 ("Uluslararası Taşıma Hizmeti", "C62", 0, "istisna")],
    "Otomotiv Satış ve Kiralama": [("Araç Satışı", "C62", 20, "yurtici"), ("Araç Kiralama Hizmeti", "C62", 20, "yurtici")],
    "Muhtelif İmalat": [("Metal Aksam", "KGM", 20, "yurtici"), ("Metal Aksam (İhracat)", "KGM", 0, "ihracat")],
}
# Satışların ihracat / istisna payı (sektör başına)
DIS_SATIS_ORANI = {"Halı Üretimi": 0.4, "Muhtelif İmalat": 0.3, "Uluslararası Taşımacılık": 0.6}


def otv_line(ln):
    """Araç satırına ÖTV (9077) ekler (deterministik; rastgelelik kullanmaz)."""
    oran = OTV_ORANLARI.get(ln["name"]) if ln["kind"] == "arac" else None
    ln["otv"] = round(ln["qty"] * ln["price"] * oran / 100, 2) if oran else 0.0
    ln["otv_rate"] = oran
    return ln


def vkn():
    return str(random.randint(1000000000, 9999999999))


def rand_date(start=START, end=END):
    return start + timedelta(days=random.randint(0, (end - start).days))


def rand_date_rng(rng, start=START, end=END):
    return start + timedelta(days=rng.randint(0, (end - start).days))


def month_end(d):
    nxt = date(d.year + (d.month // 12), d.month % 12 + 1, 1)
    return nxt - timedelta(days=1)


def price_at(base, d):
    months = (d.year - 2024) * 12 + d.month - 7 + d.day / 30
    return base * (1 + MONTHLY_INFLATION) ** months


class Firm:
    def __init__(self, idx, sector):
        self.idx, self.sector = idx, sector
        self.name = f"{random.choice(SURNAMES)} {random.choice(NAME_PARTS[sector])} " \
                    f"{random.choice(['A.Ş.', 'Ltd. Şti.'])}"
        self.code = f"F{idx:02d}"
        self.vkn = vkn()
        self.city = random.choice(CITIES)
        self.catalog = CATALOG[sector]
        big = idx == 27  # bir büyük filo firması: performans testi
        self.n_invoices = 2500 if big else random.randint(80, 400)
        self.source = "excel" if random.random() < 0.2 else "xml"  # bazı firmalar portaldan Excel verir
        self.zip_xml = random.random() < 0.5
        self.journal_format = random.choices(JOURNAL_FORMATS, FORMAT_WEIGHTS)[0]
        self.belge_style = random.choices(BELGE_STYLES, BELGE_WEIGHTS)[0]
        self.suppliers = []
        for i in range(random.randint(6, 14)):
            prods = random.sample(self.catalog, k=min(len(self.catalog), random.randint(1, 3)))
            self.suppliers.append({
                "name": f"{random.choice(SURNAMES)} {random.choice(['Ticaret', 'Sanayi', 'Petrol', 'Yapı Market', 'Tedarik', 'Hizmet'])} "
                        f"{random.choice(['A.Ş.', 'Ltd. Şti.'])}",
                "vkn": vkn(), "prefix": "".join(random.choices("ABCDEFGHJKLMNPRSTUVYZ", k=3)), "seq": 0,
                "products": prods, "price_factor": random.uniform(0.95, 1.05)})
        # her ürünün en az bir tedarikçisi olsun
        for p in self.catalog:
            if not any(p in s["products"] for s in self.suppliers):
                random.choice(self.suppliers)["products"].append(p)
        self.invoices, self.journal, self.truth = [], [], {e: [] for e in ERROR_TYPES}
        # Madde 7 eklemeleri ayrı rastgele üreteçle: mevcut hataların üretimi ve `random` sırası değişmez
        self.rng = random.Random(7000 + idx)
        self.sales, self.sales_drafts = [], []

    @property
    def accounts(self):
        return sorted({p[4] for p in self.catalog})

    def next_no(self, sup):
        sup["seq"] += 1
        return f"{sup['prefix']}2024{sup['seq']:09d}"

    def make_invoice(self, sup=None, d=None, lines=None, inv_type=None):
        sup = sup or random.choice(self.suppliers)
        d = d or rand_date()
        if lines is None:
            lines = []
            for prod in random.sample(sup["products"], k=random.randint(1, len(sup["products"]))):
                name, uom, base, vat, acc, kind = prod
                if kind == "arac":
                    qty = 1
                elif kind in ("tevkifat",) or uom == "C62" and base > 5000:
                    qty = random.randint(1, 3) if kind != "tevkifat" else 1
                else:
                    qty = round(random.uniform(5, 400) * (1 if base > 100 else 20), 0)
                price = price_at(base, d) * sup["price_factor"] * random.gauss(1, 0.03)
                if kind == "tevkifat":
                    price = base * random.uniform(0.4, 1.8)  # hakediş tutarları her ay farklı
                if kind == "arac":
                    price = round(base * sup["price_factor"], -3)  # liste fiyatı
                lines.append({"name": name, "uom": uom, "qty": qty, "price": round(price, 4), "vat": vat,
                              "acc": acc, "kind": kind})
                otv_line(lines[-1])
        kinds = {ln["kind"] for ln in lines}
        inv = {"no": self.next_no(sup), "date": d, "sup": sup, "lines": lines, "currency": "TRY", "rate": 1.0,
               "type": inv_type or ("TEVKIFAT" if "tevkifat" in kinds else "SATIS"),
               "customer_vkn": self.vkn, "customer_name": self.name, "source": "XML", "calc_error": 0.0}
        self.invoices.append(inv)
        return inv

    def generate(self):
        n = self.n_invoices
        for _ in range(n):
            self.make_invoice()
        # Otomotiv: distribütörden aynı gün, aynı model birden çok araç (ayrı faturalarla)
        if self.sector == "Otomotiv Satış ve Kiralama":
            for _ in range(random.randint(2, 5)):
                sup = random.choice(self.suppliers)
                cars = [p for p in sup["products"] if p[5] == "arac"] or [self.catalog[0]]
                car, d = random.choice(cars), rand_date()
                for _ in range(random.randint(2, 4)):
                    self.make_invoice(sup, d, [otv_line({"name": car[0], "uom": car[1], "qty": 1,
                                                         "price": round(car[2] * sup["price_factor"], -3), "vat": 20,
                                                         "acc": car[4], "kind": "arac"})])
        # Taşımacılık: yurtdışı yakıt fişleri (EUR, Excel ile girilir)
        if self.sector == "Uluslararası Taşımacılık":
            foreign = {"name": "Yurtdışı Yakıt İstasyonu GmbH", "vkn": "9999999999", "prefix": "EUR", "seq": 0,
                       "products": [], "price_factor": 1}
            for _ in range(random.randint(15, 40)):
                d = rand_date()
                inv = self.make_invoice(foreign, d, [{"name": "Motorin", "uom": "LTR", "qty": random.randint(300, 900),
                                                     "price": round(random.uniform(1.55, 1.85), 3), "vat": 0,
                                                     "acc": "740", "kind": "mal"}])
                inv.update(currency="EUR", rate=round(36.5 + (d - START).days * 0.02, 4), source="EXCEL")
        # Birkaç iade faturası
        for _ in range(random.randint(0, 3)):
            base = random.choice(self.invoices)
            ln = dict(base["lines"][0])
            ln["qty"] = max(1, round(ln["qty"] * 0.2))
            self.make_invoice(base["sup"], min(END, base["date"] + timedelta(days=5)), [ln], "IADE")
        if self.source == "excel":
            for inv in self.invoices:
                inv["source"] = "EXCEL"
        self.inject_errors()
        self.inject_tax_errors()
        for inv in self.invoices:
            if not inv.get("skip_journal"):
                self.book(inv)
        self.noise_entries()
        self.make_sales()
        self.inject_sales_errors()
        for inv in self.sales:
            if not inv.get("skip_journal"):
                self.book_sale(inv)

    # ------------------------------------------------------------------ hata enjeksiyonu
    def pick(self, cond=lambda i: True):
        cands = [i for i in self.invoices if not i.get("err") and i["type"] == "SATIS" and i["currency"] == "TRY"
                 and cond(i)]
        inv = random.choice(cands)
        return inv

    def inject_errors(self):
        k = lambda: random.choice([0, 0, 1, 1, 2, 3])
        for _ in range(k()):
            inv = self.pick(); inv["err"] = "muhasebelesmemis"; inv["skip_journal"] = True
            self.truth["muhasebelesmemis"].append(inv["no"])
        for _ in range(k()):
            inv = self.pick(); inv["err"] = "yanlis_hesap"; inv["book_acc"] = random.choice(["689", "770" if "770" not in self.accounts else "689"])
            self.truth["yanlis_hesap"].append(inv["no"])
        for _ in range(k()):
            inv = self.pick(); inv["err"] = "tutar_farki"
            net = self.net(inv)
            s = f"{net:.2f}"
            # rakam yer değiştirme (ör. 12.345 → 12.354) ya da yuvarlama/yanlış giriş
            digits = [i for i, c in enumerate(s) if c.isdigit()]
            if len(digits) > 3 and random.random() < 0.6:
                a = digits[-4]; b = digits[-5] if len(digits) > 4 else digits[-3]
                lst = list(s); lst[a], lst[b] = lst[b], lst[a]
                wrong = float("".join(lst))
                if abs(wrong - net) < 1:
                    wrong = net * 1.1
            else:
                wrong = net * random.choice([0.9, 1.1, 0.5])
            inv["book_net"] = round(wrong, 2)
            self.truth["tutar_farki"].append(inv["no"])
        for _ in range(k()):
            inv = self.pick(lambda i: self.vat(i) > 0); inv["err"] = "kdv_dahil_kayit"; inv["vat_in_cost"] = True
            self.truth["kdv_dahil_kayit"].append(inv["no"])
        for _ in range(random.choice([0, 0, 1])):
            inv = self.pick(); inv["err"] = "cift_kayit"; inv["double"] = True
            self.truth["cift_kayit"].append(inv["no"])
        for _ in range(k()):
            inv = self.pick(lambda i: i["date"].month < 9); inv["err"] = "donem_kaymasi"
            inv["book_date"] = month_end(inv["date"]) + timedelta(days=random.randint(1, 10))
            self.truth["donem_kaymasi"].append(inv["no"])
        for _ in range(k()):
            # Faturası olmayan gider kaydı (sahte/eksik belge)
            sup = random.choice(self.suppliers)
            no = f"{sup['prefix']}2024{random.randint(800000000, 899999999):09d}"
            prod = random.choice(sup["products"])
            amt = round(price_at(prod[2], START) * random.randint(5, 50), 2)
            d = rand_date()
            self.journal_entry(d, no, prod[4], amt, 0, f"{sup['name']} fatura")
            self.journal_entry(d, no, "191.01", round(amt * prod[3] / 100, 2), 0, "İndirilecek KDV")
            self.journal_entry(d, no, "320.01", 0, round(amt * (1 + prod[3] / 100), 2), sup["name"])
            self.truth["faturasiz_gider"].append(no)
        for _ in range(k()):
            base = self.pick(lambda i: i["lines"][0]["kind"] != "arac")
            dup = self.make_invoice(base["sup"], base["date"], [dict(l) for l in base["lines"]])
            dup["err"] = "mukerrer_fatura"; dup["source"] = base["source"]
            self.truth["mukerrer_fatura"].append(sorted([base["no"], dup["no"]]))
        for _ in range(k()):
            inv = self.pick(lambda i: i["lines"][0]["kind"] == "mal")
            inv["err"] = "fiyat_sisirme"
            inv["lines"][0]["price"] = round(inv["lines"][0]["price"] * random.uniform(1.35, 1.8), 4)
            self.truth["fiyat_sisirme"].append(inv["no"])
        if self.source == "xml":
            for _ in range(random.choice([0, 0, 1, 2])):
                inv = self.pick(lambda i: i["source"] == "XML"); inv["err"] = "baska_firma_faturasi"
                inv["customer_vkn"], inv["customer_name"] = vkn(), f"{random.choice(SURNAMES)} Grup Şirketi"
                self.truth["baska_firma_faturasi"].append(inv["no"])
            for _ in range(random.choice([0, 0, 0, 1])):
                inv = self.pick(lambda i: i["source"] == "XML"); inv["err"] = "xml_hesap_hatasi"
                inv["calc_error"] = round(random.uniform(5, 500), 2)
                self.truth["xml_hesap_hatasi"].append(inv["no"])

    def pick7(self, cond):
        cands = [i for i in self.invoices if not i.get("err") and i["currency"] == "TRY" and cond(i)]
        return self.rng.choice(cands) if cands else None

    def inject_tax_errors(self):
        """Madde 7 alış hataları: ÖTV maliyete eklenmemiş, KDV farkı, KDV kaydedilmemiş, tevkifat kaydı eksik."""
        k = lambda: self.rng.choice([0, 0, 1, 1, 2])  # noqa: E731
        if self.sector.startswith("Otomotiv"):
            for _ in range(self.rng.choice([1, 1, 2, 3])):
                inv = self.pick7(lambda i: i["type"] == "SATIS" and any(l.get("otv") for l in i["lines"]))
                if inv:
                    inv["err"] = "otv_maliyete_eklenmemis"; inv["otv_gider"] = True
                    self.truth["otv_maliyete_eklenmemis"].append(inv["no"])
        for _ in range(k()):
            inv = self.pick7(lambda i: i["type"] == "SATIS" and self.vat(i) > 100)
            if inv:  # 191'e yanlış tutar (ör. %20 yerine %10 ya da eksik/fazla giriş); 320 buna göre
                inv["err"] = "kdv_farki"; inv["kdv_carpan"] = self.rng.choice([0.5, 0.9, 1.1])
                self.truth["kdv_farki"].append(inv["no"])
        for _ in range(k()):
            inv = self.pick7(lambda i: i["type"] == "SATIS" and self.vat(i) > 0)
            if inv:  # KDV hiç kaydedilmemiş: 191 yok, 320 = KDV hariç tutar
                inv["err"] = "kdv_kaydedilmemis"; inv["kdv_yok"] = True
                self.truth["kdv_kaydedilmemis"].append(inv["no"])
        for _ in range(self.rng.choice([0, 1, 1, 2])):
            inv = self.pick7(lambda i: i["type"] == "TEVKIFAT" and i["source"] == "XML")
            if inv:  # 360 kaydı yok ya da yanlış oranla (5/10) yazılmış
                inv["err"] = "tevkifat_kaydi_eksik"; inv["tev_hata"] = self.rng.choice(["yok", "farkli"])
                self.truth["tevkifat_kaydi_eksik"].append(inv["no"])

    # ------------------------------------------------------------------ satış faturaları
    def make_sales(self):
        """Gürültü satışlarının (eski SAT... kayıtları) tarih ve tutarlarıyla giden faturalar üretir."""
        used = {s["prefix"] for s in self.suppliers} | {"EUR", "SAT"}
        while True:
            self.sales_prefix = "".join(self.rng.choices("ABCDEFGHJKLMNPRSTUVYZ", k=3))
            if self.sales_prefix not in used:
                break
        customers = [{"name": f"{self.rng.choice(SURNAMES)} {self.rng.choice(['Ticaret', 'İnşaat', 'Holding', 'Mağazacılık'])} "
                              f"{self.rng.choice(['A.Ş.', 'Ltd. Şti.'])}", "vkn": str(self.rng.randint(10**9, 10**10 - 1))}
                     for _ in range(self.rng.randint(5, 12))]
        foreign = [{"name": f"{n} GmbH", "vkn": "2222222222"} for n in ("Müller Teppich", "Schmidt Logistik",
                                                                         "Weber Industrie")]
        katalog = SATIS_KATALOG[self.sector]
        seq = 0
        for d, amt in self.sales_drafts:
            seq += 1
            dis = self.rng.random() < DIS_SATIS_ORANI.get(self.sector, 0)
            kalem = self.rng.choice([k for k in katalog if (k[3] != "yurtici") == dis])
            name, uom, vat, tur = kalem
            cur, rate = ("EUR", round(36.5 + (d - START).days * 0.02, 4)) if tur == "ihracat" else ("TRY", 1.0)
            qty = self.rng.randint(1, 50) if uom == "C62" else self.rng.randint(100, 5000)
            price = round(amt / rate / qty, 4)
            inv = {"no": f"{self.sales_prefix}2024{seq:09d}", "date": d, "lines": [
                       {"name": name, "uom": uom, "qty": qty, "price": price, "vat": vat, "acc": "601" if tur == "ihracat"
                        else "600", "kind": "satis", "otv": 0.0}],
                   "currency": cur, "rate": rate, "type": "SATIS" if tur == "yurtici" else "ISTISNA",
                   "profile": "IHRACAT" if tur == "ihracat" else self.rng.choice(["TICARIFATURA", "TEMELFATURA"]),
                   "sup": {"name": self.name, "vkn": self.vkn}, "customer": self.rng.choice(foreign if dis else customers),
                   "source": "XML" if self.source == "xml" else "EXCEL", "calc_error": 0.0, "tur": tur}
            inv["customer_vkn"], inv["customer_name"] = inv["customer"]["vkn"], inv["customer"]["name"]
            self.sales.append(inv)

    def inject_sales_errors(self):
        k = lambda: self.rng.choice([0, 0, 1, 1, 2])  # noqa: E731
        temiz = lambda cond=lambda i: True: [i for i in self.sales if not i.get("err") and cond(i)]  # noqa: E731
        for _ in range(k()):
            inv = self.rng.choice(temiz()); inv["err"] = "muhasebelesmemis_satis"; inv["skip_journal"] = True
            self.truth["muhasebelesmemis_satis"].append(inv["no"])
        for _ in range(k()):
            inv = self.rng.choice(temiz()); inv["err"] = "satis_tutar_farki"
            inv["book_net"] = round(self.net(inv) * self.rate_of(inv) * self.rng.choice([0.9, 1.1, 0.5]), 2)
            self.truth["satis_tutar_farki"].append(inv["no"])
        for _ in range(k()):
            inv = self.rng.choice(temiz()); inv["err"] = "satis_kdv_farki"
            if inv["tur"] == "yurtici":  # %20 yerine %10, ya da 391 hiç yazılmamış
                inv["kdv_carpan"] = self.rng.choice([0.5, 0.0])
            else:  # İhracat / istisna faturasına KDV hesaplanmış
                inv["kdv_istisnaya"] = True
            self.truth["satis_kdv_farki"].append(inv["no"])
        for _ in range(k()):
            cands = temiz(lambda i: i["date"].month < 9)
            if cands:
                inv = self.rng.choice(cands); inv["err"] = "satis_donem_kaymasi"
                inv["book_date"] = month_end(inv["date"]) + timedelta(days=self.rng.randint(1, 10))
                self.truth["satis_donem_kaymasi"].append(inv["no"])
        for _ in range(k()):  # Faturası olmayan gelir kaydı (fatura kesilmemiş / iptal edilmiş satış)
            no = f"{self.sales_prefix}2024{self.rng.randint(800000000, 899999999):09d}"
            amt = round(self.rng.uniform(10000, 500000), 2)
            d = rand_date_rng(self.rng)
            # Belge no faturasız gider kaydında olduğu gibi tam GİB biçiminde yazılır
            self.journal_entry(d, no, "120.01", amt * 1.2, 0, "Satış")
            self.journal_entry(d, no, "600.01", 0, amt, "Satış")
            self.journal_entry(d, no, "391.01", 0, round(amt * 0.2, 2), "Hesaplanan KDV")
            self.truth["faturasiz_gelir"].append(no)

    @staticmethod
    def rate_of(inv):
        return inv["rate"] if inv["currency"] != "TRY" else 1.0

    def book_sale(self, inv):
        d = inv.get("book_date") or inv["date"]
        rate = self.rate_of(inv)  # İhracatta fatura kuru kullanılır
        net_tl = inv.get("book_net", round(self.net(inv) * rate, 2))
        vat_tl = round(self.vat(inv) * rate, 2) * inv.get("kdv_carpan", 1)
        if inv.get("kdv_istisnaya"):
            vat_tl = round(net_tl * 0.2, 2)
        doc = self.belge(inv["no"])
        desc = f"{inv['customer_name']} {inv['no']} nolu satış faturası"
        self.journal_entry(d, doc, "120.01", net_tl + vat_tl, 0, desc)
        self.journal_entry(d, doc, inv["lines"][0]["acc"] + ".01", 0, net_tl, desc)
        if vat_tl:
            self.journal_entry(d, doc, "391.01", 0, vat_tl, "Hesaplanan KDV")

    # ------------------------------------------------------------------ tutarlar
    @staticmethod
    def line_net(ln):
        return round(ln["qty"] * ln["price"], 2)

    def net(self, inv):
        return round(sum(self.line_net(l) for l in inv["lines"]), 2)

    def vat(self, inv):  # ÖTV KDV matrahına dahildir
        return round(sum(round((self.line_net(l) + l.get("otv", 0)) * l["vat"] / 100, 2) for l in inv["lines"]), 2)

    @staticmethod
    def otv(inv):
        return round(sum(l.get("otv", 0) for l in inv["lines"]), 2)

    # ------------------------------------------------------------------ yevmiye
    def belge(self, no):
        if self.belge_style == "bosluklu":
            return f"{no[:3]} {no[3:7]} {no[7:]}"
        if self.belge_style == "kisa":
            return f"{no[:3]}{int(no[7:])}"  # ör. ABC123 (seri + sıra)
        return no

    def journal_entry(self, d, doc, acc, debit, credit, desc):
        self.journal.append({"date": d, "doc": doc, "acc": acc, "debit": round(debit, 2),
                             "credit": round(credit, 2), "desc": desc})

    def book(self, inv, again=False):
        d = inv.get("book_date") or min(month_end(inv["date"]), inv["date"] + timedelta(days=random.randint(0, 4)))
        rate = inv["rate"] if inv["currency"] == "TRY" else round(inv["rate"] * random.gauss(1, 0.004), 4)
        doc = self.belge(inv["no"])
        sign = -1 if inv["type"] == "IADE" else 1
        by_acc = {}
        for ln in inv["lines"]:
            acc = inv.get("book_acc") or ln["acc"]
            by_acc[acc] = by_acc.get(acc, 0) + self.line_net(ln) * rate
        net_tl = sum(by_acc.values())
        if "book_net" in inv:
            scale = inv["book_net"] / net_tl
            by_acc = {a: v * scale for a, v in by_acc.items()}
        # ÖTV maliyete eklenir (hatalı kayıtta 770 giderlere yazılır)
        otv_gider = 0.0
        for ln in inv["lines"]:
            if ln.get("otv"):
                if inv.get("otv_gider"):
                    otv_gider += ln["otv"] * rate
                else:
                    acc = inv.get("book_acc") or ln["acc"]
                    by_acc[acc] = by_acc.get(acc, 0) + ln["otv"] * rate
        if otv_gider:
            by_acc["770"] = by_acc.get("770", 0) + otv_gider
        vat_tl = self.vat(inv) * rate
        if inv.get("vat_in_cost"):
            first = next(iter(by_acc))
            by_acc[first] += vat_tl
            vat_tl_191 = 0
        else:
            vat_tl_191 = vat_tl
        if inv.get("kdv_yok"):
            vat_tl_191 = 0
        vat_191_booked = round(vat_tl_191 * inv.get("kdv_carpan", 1), 2)
        gross = sum(by_acc.values()) + vat_191_booked
        desc = f"{inv['sup']['name']} {inv['no']} nolu fatura"
        for acc, amt in by_acc.items():
            sub = f"{acc}.01" if acc not in ("689",) else "689.01"
            self.journal_entry(d, doc, sub, amt if sign > 0 else 0, 0 if sign > 0 else amt, desc)
        if vat_191_booked:
            self.journal_entry(d, doc, "191.01", vat_191_booked if sign > 0 else 0, 0 if sign > 0 else vat_191_booked,
                               desc)
        withheld = round(vat_tl_191 * 0.4, 2) if inv["type"] == "TEVKIFAT" else 0
        booked_360 = {"yok": 0, "farkli": round(vat_tl_191 * 0.5, 2)}.get(inv.get("tev_hata"), withheld)
        if booked_360:
            self.journal_entry(d, doc, "360.02", 0, booked_360, "KDV tevkifatı")
        pay = gross - booked_360
        self.journal_entry(d, doc, "320.01", 0 if sign > 0 else pay, pay if sign > 0 else 0, desc)
        if inv.get("double") and not again:
            self.book(inv, again=True)

    def noise_entries(self):
        """Faturaya dayanmayan olağan kayıtlar: bordro, amortisman, SMM, satışlar, gider pusulası."""
        for m in (7, 8, 9):
            last = month_end(date(2024, m, 1))
            payroll = random.randint(300000, 3000000)
            self.journal_entry(last, f"BORDRO-2024-{m:02d}", "770.01" if self.sector != "İnşaat" else "170.02",
                               payroll, 0, "Personel ücretleri")
            self.journal_entry(last, f"BORDRO-2024-{m:02d}", "335.01", 0, payroll, "Personel ücretleri")
            dep = random.randint(50000, 400000)
            self.journal_entry(last, f"AMORT-{m:02d}", "770.05" if "770" in self.accounts else "730.05", dep, 0,
                               "Amortisman")
            self.journal_entry(last, f"AMORT-{m:02d}", "257.01", 0, dep, "Birikmiş amortisman")
            if "153" in self.accounts or "150" in self.accounts:
                acc = "153.01" if "153" in self.accounts else "150.01"
                cost = random.randint(500000, 5000000)
                self.journal_entry(last, f"SMM-{m:02d}", "621.01", cost, 0, "Satılan malın maliyeti")
                self.journal_entry(last, f"SMM-{m:02d}", acc, 0, cost, "Satılan malın maliyeti")
            for k in range(random.randint(10, 40)):
                d = rand_date(date(2024, m, 1), last)
                amt = round(random.uniform(10000, 900000), 2)
                # Madde 7: gürültü satışları yerine bu tarih ve tutarlarla giden satış faturası kesilir (make_sales)
                self.sales_drafts.append((d, amt))
            for k in range(random.randint(0, 4)):
                d = rand_date(date(2024, m, 1), last)
                amt = round(random.uniform(2000, 40000), 2)
                acc = self.accounts[0] + ".03"
                self.journal_entry(d, f"GP-{m:02d}{k:03d}", acc, amt, 0, "Gider pusulası")
                self.journal_entry(d, f"GP-{m:02d}{k:03d}", "100.01", 0, amt, "Gider pusulası")


# ---------------------------------------------------------------------- dosya yazımı
NSDECL = ('xmlns="urn:oasis:names:specification:ubl:schema:xsd:Invoice-2" '
          'xmlns:cac="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2" '
          'xmlns:cbc="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"')


def esc(s):
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def build_xml(firm, inv):
    cur = inv["currency"]
    nets = [Firm.line_net(l) for l in inv["lines"]]
    otvs = [l.get("otv", 0.0) for l in inv["lines"]]
    vats = [round((n + o) * l["vat"] / 100, 2) for n, o, l in zip(nets, otvs, inv["lines"])]
    net, vat, otv = round(sum(nets), 2), round(sum(vats), 2), round(sum(otvs), 2)
    withheld = round(vat * 0.4, 2) if inv["type"] == "TEVKIFAT" else 0

    def otv_sub(n, o, oran):
        return (f"""<cac:TaxSubtotal><cbc:TaxableAmount currencyID="{cur}">{n:.2f}</cbc:TaxableAmount><cbc:TaxAmount currencyID="{cur}">{o:.2f}</cbc:TaxAmount><cbc:Percent>{oran}</cbc:Percent>
        <cac:TaxCategory><cac:TaxScheme><cbc:Name>ÖTV</cbc:Name><cbc:TaxTypeCode>9077</cbc:TaxTypeCode></cac:TaxScheme></cac:TaxCategory></cac:TaxSubtotal>""")

    def kdv_sub(matrah, v, oran, istisna=False):
        muaf = ("<cbc:TaxExemptionReasonCode>301</cbc:TaxExemptionReasonCode><cbc:TaxExemptionReason>11/1-a Mal "
                "ihracatı</cbc:TaxExemptionReason>") if istisna else ""
        return (f"""<cac:TaxSubtotal><cbc:TaxableAmount currencyID="{cur}">{matrah:.2f}</cbc:TaxableAmount><cbc:TaxAmount currencyID="{cur}">{v:.2f}</cbc:TaxAmount><cbc:Percent>{oran}</cbc:Percent>
        <cac:TaxCategory>{muaf}<cac:TaxScheme><cbc:Name>KDV</cbc:Name><cbc:TaxTypeCode>0015</cbc:TaxTypeCode></cac:TaxScheme></cac:TaxCategory></cac:TaxSubtotal>""")

    istisna = inv["type"] == "ISTISNA"
    lines_xml = []
    for i, (l, n, o, v) in enumerate(zip(inv["lines"], nets, otvs, vats), 1):
        lines_xml.append(f"""
  <cac:InvoiceLine><cbc:ID>{i}</cbc:ID>
    <cbc:InvoicedQuantity unitCode="{l['uom']}">{l['qty']}</cbc:InvoicedQuantity>
    <cbc:LineExtensionAmount currencyID="{cur}">{n:.2f}</cbc:LineExtensionAmount>
    <cac:TaxTotal><cbc:TaxAmount currencyID="{cur}">{v + o:.2f}</cbc:TaxAmount>
      {otv_sub(n, o, l.get('otv_rate')) if o else ''}{kdv_sub(n + o, v, l['vat'], istisna)}
    </cac:TaxTotal>
    <cac:Item><cbc:Name>{esc(l['name'])}</cbc:Name></cac:Item>
    <cac:Price><cbc:PriceAmount currencyID="{cur}">{l['price']:.4f}</cbc:PriceAmount></cac:Price>
  </cac:InvoiceLine>""")
    wh = ""
    if withheld:
        wh = f"""<cac:WithholdingTaxTotal><cbc:TaxAmount currencyID="{cur}">{withheld:.2f}</cbc:TaxAmount>
    <cac:TaxSubtotal><cbc:TaxAmount currencyID="{cur}">{withheld:.2f}</cbc:TaxAmount><cbc:Percent>40</cbc:Percent>
    <cac:TaxCategory><cac:TaxScheme><cbc:TaxTypeCode>624</cbc:TaxTypeCode></cac:TaxScheme></cac:TaxCategory></cac:TaxSubtotal></cac:WithholdingTaxTotal>"""
    header_net = net + inv["calc_error"]
    belge_otv = "".join(otv_sub(n, o, l.get("otv_rate")) for n, o, l in zip(nets, otvs, inv["lines"]) if o)
    # Satış faturalarında senaryo faturanın kendisinde tutulur (eski alış faturaları için `random` sırası korunur)
    profile = inv.get("profile") or random.choice(["TICARIFATURA", "TEMELFATURA"])
    kur = f"""<cac:PricingExchangeRate><cbc:SourceCurrencyCode>{cur}</cbc:SourceCurrencyCode><cbc:TargetCurrencyCode>TRY</cbc:TargetCurrencyCode><cbc:CalculationRate>{inv['rate']}</cbc:CalculationRate></cac:PricingExchangeRate>""" \
        if cur != "TRY" else ""
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<Invoice {NSDECL}>
  <cbc:UBLVersionID>2.1</cbc:UBLVersionID><cbc:CustomizationID>TR1.2</cbc:CustomizationID>
  <cbc:ProfileID>{profile}</cbc:ProfileID>
  <cbc:ID>{inv['no']}</cbc:ID><cbc:IssueDate>{inv['date'].isoformat()}</cbc:IssueDate>
  <cbc:InvoiceTypeCode>{inv['type']}</cbc:InvoiceTypeCode><cbc:DocumentCurrencyCode>{cur}</cbc:DocumentCurrencyCode>{kur}
  <cac:AccountingSupplierParty><cac:Party>
    <cac:PartyIdentification><cbc:ID schemeID="VKN">{inv['sup']['vkn']}</cbc:ID></cac:PartyIdentification>
    <cac:PartyName><cbc:Name>{esc(inv['sup']['name'])}</cbc:Name></cac:PartyName></cac:Party></cac:AccountingSupplierParty>
  <cac:AccountingCustomerParty><cac:Party>
    <cac:PartyIdentification><cbc:ID schemeID="VKN">{inv['customer_vkn']}</cbc:ID></cac:PartyIdentification>
    <cac:PartyName><cbc:Name>{esc(inv['customer_name'])}</cbc:Name></cac:PartyName></cac:Party></cac:AccountingCustomerParty>
  <cac:TaxTotal><cbc:TaxAmount currencyID="{cur}">{vat + otv:.2f}</cbc:TaxAmount>
    {belge_otv}{kdv_sub(net + otv, vat, inv['lines'][0]['vat'] if len({l['vat'] for l in inv['lines']}) == 1 else 20, istisna)}</cac:TaxTotal>
  {wh}
  <cac:LegalMonetaryTotal>
    <cbc:LineExtensionAmount currencyID="{cur}">{header_net:.2f}</cbc:LineExtensionAmount>
    <cbc:TaxExclusiveAmount currencyID="{cur}">{header_net:.2f}</cbc:TaxExclusiveAmount>
    <cbc:TaxInclusiveAmount currencyID="{cur}">{header_net + otv + vat:.2f}</cbc:TaxInclusiveAmount>
    <cbc:AllowanceTotalAmount currencyID="{cur}">0.00</cbc:AllowanceTotalAmount>
    <cbc:PayableAmount currencyID="{cur}">{header_net + otv + vat - withheld:.2f}</cbc:PayableAmount>
  </cac:LegalMonetaryTotal>{''.join(lines_xml)}
</Invoice>"""


def tr_num(x):
    return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def write_firm(firm):
    d = os.path.join(OUT, firm.code)
    os.makedirs(d, exist_ok=True)
    xml_invs = [i for i in firm.invoices if i["source"] == "XML"]
    xml_sales = [i for i in firm.sales if i["source"] == "XML"]
    if xml_invs:
        if firm.zip_xml:
            with zipfile.ZipFile(os.path.join(d, "efatura_2024Q3.zip"), "w", zipfile.ZIP_DEFLATED) as zf:
                for inv in xml_invs:
                    zf.writestr(f"gelen/{inv['no']}.xml", build_xml(firm, inv))
                for inv in xml_sales:  # Giden e-fatura / e-arşiv
                    zf.writestr(f"giden/{inv['no']}.xml", build_xml(firm, inv))
        else:
            xd = os.path.join(d, "xml")
            os.makedirs(xd, exist_ok=True)
            for inv in xml_invs + xml_sales:
                with open(os.path.join(xd, f"{inv['no']}.xml"), "w", encoding="utf-8") as fh:
                    fh.write(build_xml(firm, inv))
    xl_invs = [i for i in firm.invoices if i["source"] == "EXCEL"]
    xl_sales = [i for i in firm.sales if i["source"] == "EXCEL"]
    if xl_invs or xl_sales:
        rows = []
        for inv in xl_invs + xl_sales:
            satis = inv in xl_sales
            for l in inv["lines"]:
                # Satış satırlarında Tedarikci_* alanlarına karşı taraf (müşteri) yazılır
                rows.append({"Fatura_No": inv["no"], "Tarih": inv["date"].strftime("%d.%m.%Y"),
                             "Tedarikci_VKN": inv["customer_vkn"] if satis else inv["sup"]["vkn"],
                             "Tedarikci_Ad": inv["customer_name"] if satis else inv["sup"]["name"],
                             "Urun_Adi": l["name"], "Miktar": l["qty"],
                             "Birim": {"C62": "Adet", "KGM": "Kg", "LTR": "Lt", "MTK": "m2", "MTQ": "m3",
                                       "TNE": "Ton", "KWH": "kWh", "DAY": "Gün", "SET": "Set"}[l["uom"]],
                             "Fiyat": l["price"], "KDV_Orani": l["vat"], "Para_Birimi": inv["currency"],
                             "Kur": inv["rate"], "ÖTV Tutarı": l.get("otv", 0.0), "Yön": "Satış" if satis else "Alış",
                             "Fatura Tipi": ("IHRACAT" if inv.get("profile") == "IHRACAT" else inv["type"]) if satis
                             else None})
        pd.DataFrame(rows).to_excel(os.path.join(d, "faturalar.xlsx"), index=False)

    j = pd.DataFrame(firm.journal).sort_values(["date", "doc"]).reset_index(drop=True)
    fmt = firm.journal_format
    if fmt == "standart":
        out = pd.DataFrame({"Tarih": j["date"], "Belge_No": j["doc"], "Hesap_Kodu": j["acc"], "Borc": j["debit"],
                            "Alacak": j["credit"], "Aciklama": j["desc"]})
        out.to_excel(os.path.join(d, "yevmiye.xlsx"), index=False)
    elif fmt == "luca":
        out = pd.DataFrame({"Fiş Tarihi": j["date"].map(lambda x: x.strftime("%d.%m.%Y")), "Evrak No": j["doc"],
                            "Hesap Kodu": j["acc"], "Açıklama": j["desc"], "Borç": j["debit"].map(tr_num),
                            "Alacak": j["credit"].map(tr_num)})
        out.to_excel(os.path.join(d, "yevmiye.xlsx"), index=False)
    elif fmt == "logo":
        out = pd.DataFrame({"Tarih": j["date"], "Fiş No": range(1, len(j) + 1), "Belge No": j["doc"],
                            "Hesap Kodu": j["acc"], "Hesap Adı": "", "Borç Tutarı": j["debit"],
                            "Alacak Tutarı": j["credit"], "Açıklama": j["desc"]})
        out.to_excel(os.path.join(d, "yevmiye.xlsx"), index=False)
    elif fmt == "zirve":
        # Belge numarası ayrı sütunda değil, açıklamanın içinde
        out = pd.DataFrame({"Yevmiye No": range(1, len(j) + 1), "Tarih": j["date"], "Hesap Kodu": j["acc"],
                            "Açıklama": j["doc"] + " - " + j["desc"], "Borç": j["debit"], "Alacak": j["credit"]})
        out.to_excel(os.path.join(d, "yevmiye.xlsx"), index=False)
    else:  # baslikli: üstte firma adı ve rapor başlığı satırları
        out = pd.DataFrame({"Tarih": j["date"], "Belge No": j["doc"], "Hesap Kodu": j["acc"], "Borç": j["debit"],
                            "Alacak": j["credit"], "Açıklama": j["desc"]})
        with pd.ExcelWriter(os.path.join(d, "yevmiye.xlsx")) as w:
            pd.DataFrame([[firm.name], ["YEVMİYE DEFTERİ DÖKÜMÜ 01.07.2024 - 30.09.2024"], [""]]) \
                .to_excel(w, index=False, header=False, startrow=0)
            out.to_excel(w, index=False, startrow=3)

    meta = {"code": firm.code, "name": firm.name, "sector": firm.sector, "vkn": firm.vkn, "city": firm.city,
            "accounts": firm.accounts, "source": firm.source, "zip": firm.zip_xml,
            "journal_format": firm.journal_format, "belge_style": firm.belge_style,
            "n_invoices": len(firm.invoices), "n_journal": len(firm.journal), "truth": firm.truth,
            "foreign_invoices": [i["no"] for i in firm.invoices if i["currency"] != "TRY"],
            "tevkifat_invoices": sum(1 for i in firm.invoices if i["type"] == "TEVKIFAT"),
            "sales_invoices": len(firm.sales), "sales_prefix": firm.sales_prefix,
            "export_invoices": sum(1 for i in firm.sales if i["tur"] != "yurtici"),
            "otv_invoices": sum(1 for i in firm.invoices if Firm.otv(i))}
    with open(os.path.join(d, "meta.json"), "w", encoding="utf-8") as fh:
        json.dump(meta, fh, ensure_ascii=False, indent=1)
    return meta


def main():
    shutil.rmtree(OUT, ignore_errors=True)
    idx = 1
    metas = []
    for sector, count in SECTOR_COUNTS.items():
        for _ in range(count):
            f = Firm(idx, sector)
            f.generate()
            metas.append(write_firm(f))
            print(f"{f.code} {sector:28s} {f.name:40s} fatura={len(f.invoices):5d} satış={len(f.sales):4d} "
                  f"yevmiye={len(f.journal):6d} "
                  f"kaynak={f.source:5s} format={f.journal_format:9s} belge={f.belge_style}")
            idx += 1
    with open(os.path.join(OUT, "firmalar.json"), "w", encoding="utf-8") as fh:
        json.dump(metas, fh, ensure_ascii=False, indent=1)


if __name__ == "__main__":
    main()
