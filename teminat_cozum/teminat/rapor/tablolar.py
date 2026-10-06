"""Rapor tablolarını sıra numarası (index) yerine içerik ve başlıktan bulma.

Aralık-2025 raporunda araya 339 tablosu girince index'ler kaymıştı; bu yüzden her tablo, kendi içinde
geçen sabit metinlerle (ve gerekirse hemen önündeki başlık paragrafıyla) tanınır. Firma ayarındaki
`tablolar:` bölümüyle tanımlar değiştirilebilir.
"""
from dataclasses import dataclass, field

from docx.oxml.ns import qn
from docx.table import Table

from .docx_araclari import ara, uniq_cells


@dataclass
class TabloTanimi:
    ad: str                       # rapordaki bölüm adı (kontrol listesinde görünür)
    icerik: list = field(default_factory=list)   # tablonun metninde hepsi geçmeli (regex)
    haric: list = field(default_factory=list)    # biri bile geçerse bu tablo değildir
    baslik: str = ''              # tablodan önceki son 3 paragraftan birinde geçmeli (regex)
    zorunlu: bool = True


VARSAYILAN_TABLOLAR = {
    'kapak': TabloTanimi('Kapak', icerik=[r'Rapor Sayısı', r'İncelemenin Dönemi']),
    'isci': TabloTanimi('1-6 İşçi sayısı', icerik=[r'\b(OCAK|ŞUBAT|MART|NİSAN|MAYIS|HAZİRAN|TEMMUZ|AĞUSTOS|EYLÜL|EKİM|KASIM|ARALIK)-\d{4}\b'],
                        haric=[r'Rapor Sayısı', r'TUTANAKLARI', r'Teslim Bedeli'], baslik=r'işçi sayısı|^1-6'),
    'matrah': TabloTanimi('3-1 Matrah ve vergi bildirimi', icerik=[r'TEVKİFAT UYGULANMAYAN'], haric=[r'İndirimler Toplamı']),
    'indirim': TabloTanimi('3-2 İndirilecek KDV ve oran dağılımı', icerik=[r'Yurtiçi Alımlara', r'GÖRE DAĞILIMI'],
                           haric=[r'TEVKİFAT UYGULANMAYAN']),
    'is_hacmi': TabloTanimi('3-3-1 İş hacmi', icerik=[r'TESLİM VE HİZMET TUTARI', r'İADE EDİLMESİ GEREKEN'],
                            haric=[r'TEVKİFAT UYGULANMAYAN']),
    'karsit': TabloTanimi('3-4 Karşıt inceleme özeti', icerik=[r'Karşıt İnceleme Yapılan Tutar']),
    'safha': TabloTanimi('3-4-3 İndirilecek KDV safhaları', icerik=[r'İNDİRİLECEK KDV TUTANAKLARI']),
    'yuklenilen': TabloTanimi('3-4-4 Yüklenilen KDV safhaları', icerik=[r'YÜKLENİLEN KDV TUTANAKLARI'], zorunlu=False),
    'tevkifat': TabloTanimi('3.7 Tevkifat', icerik=[r'KISMİ TEVKİFAT KAPSAMINA GİREN', r'İadeye Konu KDV', r'Dönemi'],
                            haric=[r'TEVKİFAT UYGULANMAYAN'], zorunlu=False),
    'dokum': TabloTanimi('3.8 Beyanname dökümü', icerik=[r'TEVKİFAT UYGULANMAYAN', r'İndirimler Toplamı']),
}


def tanimlari_al(ayar=None):
    tanimlar = {k: TabloTanimi(**vars(v)) for k, v in VARSAYILAN_TABLOLAR.items()}
    for ad, deg in ((ayar.tablolar if ayar else None) or {}).items():
        t = tanimlar.get(ad) or TabloTanimi(ad)
        for k, v in deg.items():
            setattr(t, k, v)
        tanimlar[ad] = t
    return tanimlar


def tablo_metni(t):
    return '\n'.join(' | '.join(c.text.strip() for c in uniq_cells(r)) for r in t.rows)


def govde_ogeleri(doc):
    """Gövdedeki (paragraf metni | Table) öğeleri sırayla; tablo öncesi başlıkları bulmak için."""
    for el in doc.element.body.iterchildren():
        if el.tag == qn('w:p'):
            yield ''.join(x.text or '' for x in el.iter(qn('w:t'))).strip()
        elif el.tag == qn('w:tbl'):
            yield Table(el, doc._body)


def tablolari_bul(doc, ayar=None):
    """{ad: Table} ve sorunlar listesi [(ad, açıklama)] döndürür.

    Bir tanıma birden çok tablo uyarsa ilki alınır ve sorun olarak bildirilir; bir tablo yalnızca bir
    tanıma atanır.
    """
    tanimlar = tanimlari_al(ayar)
    adaylar = []   # (Table, metin, önceki başlıklar)
    onceki = []
    for o in govde_ogeleri(doc):
        if isinstance(o, str):
            if o:
                onceki.append(o)
        else:
            adaylar.append((o, tablo_metni(o), onceki[-3:]))
            onceki = []
    bulunan, sorunlar, kullanilan = {}, [], set()
    for ad, t in tanimlar.items():
        uyan = []
        for i, (tbl, metin, bas) in enumerate(adaylar):
            if i in kullanilan:
                continue
            if not all(ara(p, metin) for p in t.icerik):
                continue
            if any(ara(p, metin) for p in t.haric):
                continue
            if t.baslik and not any(ara(t.baslik, b) for b in bas):
                continue
            uyan.append(i)
        if not uyan:
            if t.zorunlu:
                sorunlar.append((ad, f'{t.ad} tablosu şablonda bulunamadı.'))
            continue
        if len(uyan) > 1:
            sorunlar.append((ad, f'{t.ad} tanımına {len(uyan)} tablo uydu; ilki kullanıldı.'))
        kullanilan.add(uyan[0])
        bulunan[ad] = adaylar[uyan[0]][0]
    return bulunan, sorunlar
