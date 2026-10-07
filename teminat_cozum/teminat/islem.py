"""Bir ayın taslağını hazırlama adımları (komut satırı ve pencere ortak kullanır):

1. Seçilen klasördeki bütün belgeler içerikten tanınır (tanıma tablosu).
2. Raporda kullanılacak belgeler içerikten seçilir; karar verilemezse kullanıcıya sorulur (BelirsizSecim).
3. Şablon (önceki ayın bitmiş raporu) kullanıcı vermediyse çevredeki Word dosyalarında içerikten aranır.
"""
from dataclasses import dataclass, field
from pathlib import Path

from .tanima import (ROLLER, Onbellek, TanimaHatasi, TurDuzeltmeleri, metindeki_vknler, rapor_bilgisi, sablon_bul,
                     secim_yap, tablo_metni, tani, vkn_metinde)


def yol_temizle(yol):
    """Pencereye yapıştırılan yolun başındaki/sonundaki boşluk ve tırnakları temizler, mutlak yola çevirir."""
    if yol is None:
        return None
    s = str(yol).strip().strip('"').strip("'").strip()
    return Path(s).expanduser().resolve() if s else None


def duzeltme_dosyasi(ayar=None):
    """Elle tür düzeltmelerinin saklandığı dosya: firma ayar dosyasının yanında (yoksa kullanıcı klasöründe)."""
    if ayar is not None and getattr(ayar, '_yol', None):
        return Path(ayar._yol).parent / 'belge_turleri.yaml'
    return Path.home() / '.teminat_cozum' / 'belge_turleri.yaml'


@dataclass
class Hazirlik:
    klasor: Path
    secim: object
    sablon: Path = None
    sablon_kaynagi: str = ''          # 'elle' | 'içerikten'
    notlar: list = field(default_factory=list)

    def tablo(self):
        neden = ''
        if not self.sablon:
            neden = (f'{self.secim.donem.onceki().tire} dönemine ait rapor çevredeki klasörlerde bulunamadı — '
                     '"Önceki ayın raporu" alanından (komut satırında --sablon) seçin')
        return tablo_metni(self.secim, self.sablon, neden)

    def tum_notlar(self):
        return list(self.secim.notlar) + list(self.notlar)


def hazirla(klasor, ayar=None, sablon=None, zorla=None, istenen=None, ilerleme=None, onbellek=None):
    klasor = yol_temizle(klasor)
    if klasor is None or not klasor.is_dir():
        raise TanimaHatasi(f'Klasör bulunamadı: {klasor}' if klasor else 'Bu ayın klasörünü seçin.')
    ob = onbellek if onbellek is not None else Onbellek()
    duz = TurDuzeltmeleri(duzeltme_dosyasi(ayar))
    if ilerleme:
        ilerleme('Belgeler içerikten tanınıyor…')
    belgeler = tani(klasor, duz, ob, ilerleme)
    secim = secim_yap(klasor, belgeler, istenen, zorla)
    notlar = []
    metin = secim.k.get('_text') or ''
    if ayar is not None and ayar.vkn and metindeki_vknler(metin) and not vkn_metinde(ayar.vkn, metin):
        notlar.append(('UYARI', 'Ayar dosyası', f'Seçilen ayar dosyasındaki VKN ({ayar.vkn}) bu beyannamede geçmiyor — '
                       'başka firmanın ayar dosyası seçilmiş olabilir.'))
    onceki = secim.donem.onceki()
    kaynak = ''
    if sablon:
        sablon = yol_temizle(sablon)
        if not sablon.is_file():
            raise TanimaHatasi(f'Önceki ayın raporu olarak verilen dosya bulunamadı: {sablon}')
        kaynak = 'elle'
        bilgi = rapor_bilgisi(sablon, ob)
        if bilgi['donem'] and bilgi['donem'] != onceki:
            notlar.append(('UYARI', 'Şablon', f'Seçilen şablon raporun dönemi {bilgi["donem"].tire}, beklenen {onceki.tire}.'))
        if bilgi['vkn'] and metindeki_vknler(metin) and not vkn_metinde(bilgi['vkn'], metin):
            notlar.append(('UYARI', 'Şablon', f'Seçilen şablon raporun mükellef VKN\'si ({bilgi["vkn"]}) bu beyannamede geçmiyor — '
                           'başka firmanın raporu olabilir.'))
    else:
        if ilerleme:
            ilerleme(f'Önceki ayın ({onceki.tire}) raporu aranıyor…')
        sablon, n = sablon_bul(klasor, onceki, metin, ob, ilerleme=ilerleme)
        notlar += n
        kaynak = 'içerikten' if sablon else ''
    return Hazirlik(klasor, secim, sablon, kaynak, notlar)


def rol_adi(rol):
    return 'Şablon (önceki ayın raporu)' if rol == 'SABLON' else ROLLER.get(rol, (rol, False))[0]
