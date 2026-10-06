"""Bir ayın girdi klasöründeki tüm dosyaları okur."""
import collections
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

from ..donusum import DonusumHatasi, docx_hazirla
from ..ortak import Donem
from . import kdv1 as kdv1_mod
from . import listeler, teminat


class GirdiHatasi(Exception):
    pass


@dataclass
class Girdiler:
    klasor: Path
    k: dict                       # KDV 1 beyannamesi
    donem: Donem
    liste_yolu: Path
    indirilecek: list
    liste_donemleri: collections.Counter
    takip_yolu: Path
    takip: list
    teminat_yolu: Path = None
    tem: dict = field(default_factory=dict)
    yuklenilen_yolu: Path = None
    yuklenilen: list = field(default_factory=list)
    notlar: list = field(default_factory=list)   # [(durum, konu, açıklama)]


def _sec(adaylar, ne, notlar, zorunlu=True):
    if not adaylar:
        if zorunlu:
            raise GirdiHatasi(f'{ne} bulunamadı.')
        return None
    if len(adaylar) > 1:
        notlar.append(('UYARI', 'Girdi dosyası', f'{ne} için birden fazla dosya var; seçilen: {adaylar[0].name}. '
                       f'Diğerleri: {", ".join(p.name for p in adaylar[1:])}'))
    return adaylar[0]


def girdileri_oku(klasor):
    klasor = Path(klasor)
    if not klasor.is_dir():
        raise GirdiHatasi(f'Girdi klasörü yok: {klasor}')
    notlar = []
    kp = _sec(kdv1_mod.kdv1_bul(klasor), '1 No.lu KDV beyannamesi (KDV 1.pdf)', notlar)
    k = kdv1_mod.kdv1_metinden(kdv1_mod.pdf_metni(kp), kp)
    yd = kdv1_mod.kdv1_donem(k)
    if not yd:
        raise GirdiHatasi(f'{kp.name}: beyannamenin dönemi (Yıl / Ay) okunamadı.')
    for alan in ('yurtici_alim', 'indirim_toplam', 'toplam_kdv', 'aylik_bedel'):
        if alan not in k:
            notlar.append(('HATA', 'KDV 1 okuma', f'{kp.name}: "{alan}" satırı okunamadı; ilgili tablolar 0 yazılacak.'))
    k.setdefault('yurtici_alim', 0.0)
    k.setdefault('indirim_toplam', 0.0)
    k.setdefault('toplam_kdv', 0.0)
    k.setdefault('aylik_bedel', 0.0)
    donem = Donem(*yd)

    lp = _sec(listeler.indirilecek_bul(klasor), 'İndirilecek KDV listesi (.xls)', notlar)
    ind = listeler.indirilecek_oku(lp)
    lst_period = collections.Counter(r['donem'] for r in ind if r['donem'] not in (None, ''))

    tp = _sec(listeler.takip_bul(klasor), 'Karşıt inceleme takip listesi (01 FİRMA VE MUH. BİLGİLERİ *.xls)', notlar)
    takip = listeler.takip_oku(tp)
    if not takip:
        raise GirdiHatasi(f'{tp.name}: takip listesinde firma satırı okunamadı.')

    g = Girdiler(klasor, k, donem, lp, ind, lst_period, tp, takip, notlar=notlar)

    dp = _sec(teminat.dilekce_bul(klasor), 'Teminat mektubu kabul dilekçesi', notlar, zorunlu=False)
    if dp:
        try:
            if dp.suffix.lower() == '.doc':
                with tempfile.TemporaryDirectory() as tmp:
                    g.tem = teminat.dilekce_metinden(teminat.dilekce_metni(docx_hazirla(dp, tmp)))
            else:
                g.tem = teminat.dilekce_metinden(teminat.dilekce_metni(dp))
            g.teminat_yolu = dp
        except DonusumHatasi as e:
            notlar.append(('HATA', 'Teminat dilekçesi', str(e)))

    yp = _sec(listeler.yuklenilen_bul(klasor), 'Yüklenilen tutanak çalışması', notlar, zorunlu=False)
    if yp:
        g.yuklenilen_yolu = yp
        g.yuklenilen = listeler.yuklenilen_oku(yp)
    return g
