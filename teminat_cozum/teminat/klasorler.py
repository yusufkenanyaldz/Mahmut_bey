"""Ay klasörlerini ve şablon raporu (önceki ayın bitmiş raporu) bulma.

Önerilen düzen:

    TEMİNAT ÇÖZÜMÜ/<FİRMA>/2026/02 ŞUBAT/KDV 1.pdf ...
    TEMİNAT ÇÖZÜMÜ/<FİRMA>/2026/02 ŞUBAT/RAPOR/ŞUBAT-2026 RAPOR.doc

Ama ofisteki klasörler farklı adlandırılmış olabileceği için esnek aranır:

- Ay klasörü: adında ay adı geçen ("02 ŞUBAT", "02 ŞUBAT (GİRDİ)", "ESKA 12- ARALIK", "Şubat 2026") ya da adı ay
  numarasıyla başlayan ("02", "3- GİRDİ") ya da "2026-02" / "02.2026" biçimindeki klasör.
- Yıl: klasörün kendi adında (20xx) ya da üstündeki klasörlerden birinin adında ("2026", "2026 YILI").
- Kullanıcı firma klasörünü, yıl klasörünü ya da doğrudan o ayın klasörünü (KDV 1.pdf'in olduğu klasör) seçebilir;
  ayın klasörü seçilirse dönem KDV 1 beyannamesinden okunur.
- Önceki ayın raporu: önceki ayın klasöründeki "RAPOR..." alt klasöründe ya da adında RAPOR geçen Word dosyası.
"""
import re
from dataclasses import dataclass, field
from pathlib import Path

from .ortak import AYLAR, Donem, katla

TARAMA_DERINLIGI = 4
# Bu klasörlerin içine girilmez (çıktılarımız, tutanaklar, fatura klasörleri ay klasörü sanılmasın).
ATLANAN = {'CLAUDE TASLAK', 'CLAUDE OUTPUTS', '_CEVRILEN', 'GERIYE DONUK TEST', 'TUTANAK CALISMASI'}
_AY_ADLARI = [(i + 1, katla(a).upper()) for i, a in enumerate(AYLAR)]
WORD = ('.doc', '.docx')


class KlasorHatasi(Exception):
    pass


def _buyuk(ad):
    return katla(ad).upper()


def yil_coz(ad):
    """Addaki tek yıl (20xx) ya da None."""
    y = set(re.findall(r'(?<!\d)(20\d\d)(?!\d)', katla(ad)))
    return int(y.pop()) if len(y) == 1 else None


def ay_coz(ad):
    """Klasör adından (ay, yıl|None); ay klasörü değilse None."""
    n = _buyuk(ad)
    m = re.search(r'(?<!\d)(20\d\d)\s*[-_./]\s*(\d{1,2})(?!\d)', n)
    if m and 1 <= int(m.group(2)) <= 12:
        return int(m.group(2)), int(m.group(1))
    m = re.search(r'(?<!\d)(\d{1,2})\s*[-_./]\s*(20\d\d)(?!\d)', n)
    if m and 1 <= int(m.group(1)) <= 12:
        return int(m.group(1)), int(m.group(2))
    adlar = {no for no, a in _AY_ADLARI if re.search(rf'(?<![A-Z]){a}(?![A-Z])', n)}
    if len(adlar) > 1:
        return None                    # "OCAK-ŞUBAT" gibi: belirsiz
    if adlar:
        return adlar.pop(), yil_coz(n)
    m = re.match(r'^\s*(\d{1,2})(?!\d)', n)
    if m and 1 <= int(m.group(1)) <= 12:
        return int(m.group(1)), yil_coz(n)
    return None


def _yil_klasoru_mu(ad):
    return yil_coz(ad) is not None and ay_coz(ad) is None


def _ust_yil(p, adim=3):
    for q in [p, *p.parents][:adim]:
        y = yil_coz(q.name)
        if y and ay_coz(q.name) is None:
            return y
    return None


def _atla(ad):
    n = _buyuk(ad)
    return ad.startswith(('.', '~', '$')) or n in ATLANAN or n.startswith('RAPOR')


@dataclass
class AyKlasoru:
    yol: Path
    ay: int
    yil: int = None        # klasör adından ya da üst klasörlerden anlaşılamadıysa None

    def etiket(self):
        return f'{AYLAR[self.ay - 1]}-{self.yil if self.yil else "????"}'


def ay_klasorlerini_tara(kok, derinlik=TARAMA_DERINLIGI):
    """Kök (kendisi bir ay klasörü olabilir) ve altındaki ay klasörleri. Ay klasörlerinin içine girilmez."""
    kok = Path(kok)
    r = ay_coz(kok.name)
    if r:
        return [AyKlasoru(kok, r[0], r[1] or _ust_yil(kok.parent))]
    out = []

    def gez(d, kalan):
        try:
            altlar = sorted((x for x in d.iterdir() if x.is_dir()), key=lambda x: _buyuk(x.name))
        except OSError:
            return
        for x in altlar:
            if _atla(x.name):
                continue
            r = ay_coz(x.name)
            if r:
                out.append(AyKlasoru(x, r[0], r[1] or _ust_yil(x.parent)))
            elif kalan > 1:
                gez(x, kalan - 1)
    gez(kok, derinlik)
    return out


def kdv1_klasorleri(ay_klasoru, derinlik=2):
    """Ay klasörü ya da alt klasörlerinden KDV 1.pdf içerenler (önce kendisi)."""
    from .okuyucular.kdv1 import kdv1_bul
    ay_klasoru = Path(ay_klasoru)
    if kdv1_bul(ay_klasoru):
        return [ay_klasoru]
    out = []

    def gez(d, kalan):
        try:
            altlar = sorted(x for x in d.iterdir() if x.is_dir() and not _atla(x.name))
        except OSError:
            return
        for x in altlar:
            if kdv1_bul(x):
                out.append(x)
            elif kalan > 1:
                gez(x, kalan - 1)
    gez(ay_klasoru, derinlik)
    return out


def kdv1_donemi(klasor):
    """Klasördeki KDV 1 beyannamesinin dönemi ya da None."""
    from .okuyucular.kdv1 import kdv1_bul, kdv1_donem, kdv1_metinden, pdf_metni
    c = kdv1_bul(klasor)
    if not c:
        return None
    try:
        yd = kdv1_donem(kdv1_metinden(pdf_metni(c[0])))
    except Exception:  # noqa: BLE001 - bozuk / okunamayan PDF: dönem bilinmiyor
        return None
    return Donem(*yd) if yd else None


def rapor_dosyalari(ay_klasoru):
    """Ay klasöründeki ofis raporu adayları: 'RAPOR...' alt klasöründeki Word dosyaları; o yoksa ay klasöründe adında
    RAPOR geçen Word dosyaları (taslaklar ve dilekçeler hariç)."""
    ay_klasoru = Path(ay_klasoru)
    adaylar = []
    try:
        icerik = sorted(ay_klasoru.iterdir())
    except OSError:
        return []
    for x in icerik:
        if x.is_dir() and _buyuk(x.name).startswith('RAPOR'):
            for f in sorted(x.rglob('*')):
                if f.is_file() and f.suffix.lower() in WORD and not f.name.startswith('~'):
                    adaylar.append(f)
    if not adaylar:
        for f in icerik:
            n = _buyuk(f.stem)
            if (f.is_file() and f.suffix.lower() in WORD and not f.name.startswith('~') and 'RAPOR' in n
                    and 'TASLAK' not in n and 'DILEKCE' not in n):
                adaylar.append(f)
    return adaylar


def rapor_sec(adaylar, yer=''):
    """Birden çok aday varsa 'ORJ' içermeyen, adında RAPOR geçen, aynı adın .docx hali tercih edilir."""
    c = list(adaylar)
    if len(c) > 1:
        c = [p for p in c if 'ORJ' not in _buyuk(p.name)] or c
    if len(c) > 1:
        c = [p for p in c if 'RAPOR' in _buyuk(p.stem)] or c
    if len(c) > 1 and len({p.with_suffix('') for p in c}) == 1:
        c = [p for p in c if p.suffix.lower() == '.docx'] or c
    if len(c) > 1:
        raise KlasorHatasi(f'{yer} birden fazla rapor var: {", ".join(p.name for p in c)} — şablon olarak hangisinin '
                           'kullanılacağını "Önceki ayın raporu" alanından (komut satırında --sablon) seçin.')
    return c[0] if c else None


def rapor_dosyasi(ay_klasoru):
    """Ay klasöründeki ofis raporu (yoksa None)."""
    return rapor_sec(rapor_dosyalari(ay_klasoru), f'{ay_klasoru} içinde')


def _ay_klasoru_mu(p):
    from .okuyucular.kdv1 import kdv1_bul
    return ay_coz(p.name) is not None or bool(kdv1_bul(p))


def arama_koku(secilen):
    """Önceki ayı da görebilmek için aramanın yapılacağı klasör: ay klasörü seçildiyse üstü; üstü (ya da seçilen)
    yıl klasörüyse onun da üstü (Ocak raporunun şablonu önceki yılın Aralık klasöründedir)."""
    p = Path(secilen)
    while _ay_klasoru_mu(p) and p.parent != p:      # "03 MART/GİRDİ" seçildiyse "03 MART"ın da üstüne çık
        p = p.parent
    if _yil_klasoru_mu(p.name) and p.parent != p:
        p = p.parent
    return p


def _liste(klasorler, kok, en_cok=24):
    if not klasorler:
        return 'hiç ay klasörü bulunamadı'
    s = []
    for a in sorted(klasorler, key=lambda a: (a.yil or 0, a.ay, str(a.yol)))[:en_cok]:
        try:
            yol = a.yol.relative_to(kok)
        except ValueError:
            yol = a.yol
        s.append(f'{a.etiket()} → {yol}')
    ek = f' … (+{len(klasorler) - en_cok})' if len(klasorler) > en_cok else ''
    return '; '.join(s) + ek


def _donemin_klasorleri(tum, d):
    """Döneme uyan ay klasörleri: yılı bilinip uyanlar; hiç yoksa yılı bilinmeyen aynı aylar."""
    kesin = [a for a in tum if a.ay == d.ay and a.yil == d.yil]
    return kesin or [a for a in tum if a.ay == d.ay and a.yil is None]


@dataclass
class Bulunan:
    girdi: Path
    donem: Donem
    sablon: Path = None
    notlar: list = field(default_factory=list)
    kok: Path = None

    def ozet(self):
        s = [f'Dönem        : {self.donem.tire}', f'Bu ayın klasörü: {self.girdi}',
             f'Şablon (önceki ayın raporu): {self.sablon or "BULUNAMADI"}']
        return '\n'.join(s + [f'Not: {n}' for n in self.notlar])


def girdi_ve_sablon_bul(secilen, donem=None):
    """Seçilen klasörden (firma / yıl / ay klasörü) bu ayın girdi klasörünü ve önceki ayın raporunu bulur."""
    secilen = Path(secilen)
    if not secilen.is_dir():
        raise KlasorHatasi(f'Klasör bulunamadı: {secilen}')
    notlar = []
    kok = arama_koku(secilen)
    tum = None
    if _ay_klasoru_mu(secilen):                       # doğrudan ayın klasörü seçilmiş
        girdiler = kdv1_klasorleri(secilen)
        if not girdiler:
            raise KlasorHatasi(f'Seçilen ay klasöründe KDV 1.pdf yok: {secilen}. Beyanname dosyasının adı "KDV 1.pdf" olmalı.')
        if len(girdiler) > 1:
            raise KlasorHatasi(f'Seçilen klasörün altında birden fazla KDV 1.pdf var: {"; ".join(map(str, girdiler))}. '
                               'Beyannamenin bulunduğu klasörü seçin.')
        girdi = girdiler[0]
        d_pdf = kdv1_donemi(girdi)
        if donem and d_pdf and donem != d_pdf:
            raise KlasorHatasi(f'Seçilen klasördeki KDV 1 beyannamesi {d_pdf.tire} dönemine ait, dönem olarak {donem.tire} '
                               'yazılmış. Dönemi düzeltin ya da doğru ayın klasörünü seçin.')
        donem = donem or d_pdf
        if not donem:
            raise KlasorHatasi(f'{girdi} içindeki KDV 1 beyannamesinden dönem okunamadı; dönemi yazın (ör. 2026-03).')
    else:
        if not donem:
            raise KlasorHatasi('Dönemi yazın (ör. 2026-03) ya da doğrudan o ayın klasörünü (KDV 1.pdf\'in olduğu klasör) seçin.')
        tum = ay_klasorlerini_tara(kok)
        adaylar = _donemin_klasorleri(tum, donem)
        girdiler = []
        for a in adaylar:
            girdiler += kdv1_klasorleri(a.yol)
        if not girdiler:
            if adaylar:
                raise KlasorHatasi(f'{donem.tire} klasörü bulundu ama içinde KDV 1.pdf yok: '
                                   f'{", ".join(str(a.yol) for a in adaylar)}. Beyanname dosyasının adı "KDV 1.pdf" olmalı.')
            raise KlasorHatasi(f'{donem.tire} için ay klasörü bulunamadı. Seçilen klasör: {secilen}. Programın bu klasörün '
                               f'altında gördüğü ay klasörleri: {_liste(tum, kok)}. Firma klasörünü (ör. ...\\TEMİNAT ÇÖZÜMÜ\\'
                               '<FİRMA>) ya da doğrudan o ayın klasörünü seçin.')
        if len(girdiler) > 1:
            raise KlasorHatasi(f'{donem.tire} için birden fazla klasörde KDV 1.pdf var: {"; ".join(map(str, girdiler))}. '
                               'Doğrudan o ayın klasörünü seçin.')
        girdi = girdiler[0]
        if all(a.yil is None for a in adaylar):
            notlar.append(f'{girdi.name} klasörünün yılı klasör adlarından anlaşılamadı; ay adına göre seçildi '
                          '(yıl, KDV 1 beyannamesinden kontrol edilecek).')
    if tum is None:
        tum = ay_klasorlerini_tara(kok)
    onceki = donem.onceki()
    oncekiler = [a for a in _donemin_klasorleri(tum, onceki) if a.yol.resolve() not in (girdi.resolve(), *girdi.resolve().parents)]
    raporlar = []
    for a in oncekiler:
        raporlar += rapor_dosyalari(a.yol)
    sablon = rapor_sec(raporlar, f'{onceki.tire} klasörlerinde') if raporlar else None
    if not sablon:
        if oncekiler:
            notlar.append(f'{onceki.tire} klasöründe ({", ".join(str(a.yol) for a in oncekiler)}) ofis raporu bulunamadı '
                          '(RAPOR alt klasörü ya da adında RAPOR geçen Word dosyası yok).')
        else:
            notlar.append(f'{onceki.tire} klasörü bulunamadı (aranan yer: {kok}).')
    return Bulunan(girdi, donem, sablon, notlar, kok)


def donem_haritasi(kok):
    """Geriye dönük test için: {Donem: {'girdi': [klasör], 'rapor': [dosya]}}. Yılı adlardan anlaşılamayan ay
    klasörlerinin dönemi KDV 1 beyannamesinden okunur."""
    out = {}
    atlanan = []
    for a in ay_klasorlerini_tara(kok):
        girdiler = kdv1_klasorleri(a.yol)
        d = Donem(a.yil, a.ay) if a.yil else None
        if d is None and girdiler:
            d = kdv1_donemi(girdiler[0])
        if d is None:
            atlanan.append(a.yol)
            continue
        k = out.setdefault(d, {'girdi': [], 'rapor': []})
        k['girdi'] += girdiler
        k['rapor'] += rapor_dosyalari(a.yol)
    return out, atlanan
