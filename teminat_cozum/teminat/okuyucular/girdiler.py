"""Bir ayın girdilerini okur. Hangi dosyanın ne olduğu dosya adına göre değil, içerikten tanınarak seçilir
(teminat/tanima.py, PROJE_TALIMATI.md §13)."""
import collections
from dataclasses import dataclass, field
from pathlib import Path

from ..ortak import Donem
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
    secim: object = None          # tanima.Secim (tanıma tablosu için)


def secimden_oku(secim):
    """tanima.secim_yap sonucundaki dosyaları okur."""
    from ..tanima import dilekce_metni
    s = secim
    notlar = list(s.notlar)
    k = dict(s.k)
    kp = Path(s.yollar['KDV1'])
    for alan in ('yurtici_alim', 'indirim_toplam', 'toplam_kdv', 'aylik_bedel'):
        if alan not in k:
            notlar.append(('HATA', 'KDV 1 okuma', f'{kp.name}: "{alan}" satırı okunamadı; ilgili tablolar 0 yazılacak.'))
            k[alan] = 0.0
    if 'LISTE_INDIRILECEK' not in s.yollar:
        raise GirdiHatasi('İndirilecek KDV listesi bulunamadı (klasörde içeriği bu türe uyan ve okunabilen Excel dosyası yok).')
    if 'TAKIP' not in s.yollar:
        raise GirdiHatasi('Karşıt inceleme takip listesi bulunamadı (başlığında "FİRMA … AÇIKLAMA SMMM YMM" olan Excel dosyası yok).')
    lp, tp = Path(s.yollar['LISTE_INDIRILECEK']), Path(s.yollar['TAKIP'])
    ind = listeler.indirilecek_oku(lp)
    takip = listeler.takip_oku(tp)
    if not takip:
        raise GirdiHatasi(f'{tp.name}: takip listesinde firma satırı okunamadı.')
    g = Girdiler(s.klasor, k, s.donem, lp, ind, collections.Counter(r['donem'] for r in ind if r['donem'] not in (None, '')),
                 tp, takip, notlar=notlar, secim=s)
    if 'TEMINAT_DILEKCE' in s.yollar:
        dp = Path(s.yollar['TEMINAT_DILEKCE'])
        try:
            g.tem = teminat.dilekce_metinden(dilekce_metni(dp))
            g.teminat_yolu = dp
        except Exception as e:  # noqa: BLE001
            notlar.append(('HATA', 'Teminat dilekçesi', f'{dp.name} okunamadı: {e}'))
    if 'LISTE_YUKLENILEN' in s.yollar:
        g.yuklenilen_yolu = Path(s.yollar['LISTE_YUKLENILEN'])
        g.yuklenilen = listeler.yuklenilen_oku(g.yuklenilen_yolu)
    return g


def girdileri_oku(klasor, istenen=None, zorla=None, duzeltmeler=None):
    """Klasördeki belgeleri içerikten tanıyıp okur."""
    from ..tanima import TanimaHatasi, secim_yap, tani
    klasor = Path(klasor)
    if not klasor.is_dir():
        raise GirdiHatasi(f'Girdi klasörü yok: {klasor}')
    try:
        return secimden_oku(secim_yap(klasor, tani(klasor, duzeltmeler), istenen, zorla))
    except TanimaHatasi as e:
        raise GirdiHatasi(str(e)) from e
