"""Komut satırı. Programa yalnızca "bu firmanın bu ayki klasörü" verilir; içindeki belgeler dosya adına değil İÇERİĞE
göre tanınır (PROJE_TALIMATI.md §13). Her çalıştırmada önce tanıma tablosu gösterilir.

    python -m teminat tani "<firmanın o ayki klasörü>"                       (yalnızca tanıma tablosu)
    python -m teminat taslak "<firmanın o ayki klasörü>" --ayar firmalar/<firma>/ayar.yaml
    python -m teminat taslak "<klasör>" --sablon "<önceki ayın raporu.doc>"   (şablonu elle vermek)
    python -m teminat taslak "<klasör>" --sec LISTE_INDIRILECEK="<dosya>"     (belirsiz seçimde dosyayı elle seçmek)
    python -m teminat karsilastir "gerçek rapor.doc(x)" "taslak.docx"
    python -m teminat denetle "rapor.doc(x)"
    python -m teminat geriye-donuk "<firma klasörü>" --ayar ... --cikti "...\\GERİYE DÖNÜK TEST"
    python -m teminat arayuz          (pencereli arayüz; .exe çift tıklanınca da bu açılır)
"""
import argparse
import sys
import tempfile
from pathlib import Path

from .ayar import ayar_oku
from .donusum import DonusumHatasi, docx_hazirla
from .islem import hazirla, rol_adi, yol_temizle
from .okuyucular.girdiler import GirdiHatasi
from .ortak import Donem
from .rapor.olustur import SablonHatasi
from .tanima import ROLLER, BelirsizSecim, TanimaHatasi

SECILEBILIR = list(ROLLER) + ['SABLON']


def _klasor(a):
    k = a.klasor or a.girdi or a.kok
    if not k:
        raise SystemExit('Firmanın o ayki klasörünü verin (ör. python -m teminat taslak "C:\\...\\03 MART").')
    return k


def _zorla(a):
    out = {}
    for x in a.sec or []:
        rol, _, yol = x.partition('=')
        rol = rol.strip().upper()
        if rol not in SECILEBILIR or not yol.strip():
            raise SystemExit(f'--sec biçimi: ROL="dosya yolu"; roller: {", ".join(SECILEBILIR)}')
        out[rol] = yol_temizle(yol)
    return out


def _hazirla(a, ayar):
    zorla = _zorla(a)
    sablon = a.sablon or zorla.pop('SABLON', None)
    hz = hazirla(_klasor(a), ayar, sablon, zorla, Donem.coz(a.donem) if a.donem else None, ilerleme=lambda m: print(m))
    print('\n=== TANIMA TABLOSU ===')
    print(hz.tablo())
    notlar = [n for n in hz.tum_notlar() if n[0] in ('HATA', 'UYARI')]
    if notlar:
        print('\nNotlar:')
        for d, konu, ac in notlar:
            print(f'  [{d}] {konu}: {ac}')
    print()
    return hz


def _tani(a):
    hz = _hazirla(a, ayar_oku(a.ayar))
    return 0 if hz.sablon and not hz.secim.eksik_zorunlu() else 1


def _taslak(a):
    from .kontroller.cikti import konsol_ozeti
    from .taslak import taslak_uret
    ayar = ayar_oku(a.ayar)
    hz = _hazirla(a, ayar)
    tc = taslak_uret(ayar, None, hz.klasor, a.cikti, kati=a.kati, uzerine_yaz=a.uzerine_yaz, hz=hz)
    print(f'Taslak : {tc.docx}\nKontrol: {tc.xlsx}\n         {tc.html}')
    if tc.farklar:
        print(f'Farklar: {tc.farklar} (tutar farkı: {tc.karsilastirma.tutar_farki})')
    print()
    print(konsol_ozeti(tc.sonuc.kontroller))
    return 1 if tc.sonuc.kontroller.say('HATA') else 0


def _karsilastir(a):
    from .karsilastir import karsilastir
    with tempfile.TemporaryDirectory() as tmp:
        k = karsilastir(docx_hazirla(yol_temizle(a.gercek), tmp), docx_hazirla(yol_temizle(a.taslak), tmp))
    m = k.metin()
    if a.cikti:
        Path(a.cikti).write_text(m, encoding='utf-8')
        print(f'Farklar: {a.cikti}')
    print(m if not a.cikti else m.splitlines()[-1])
    return 0


def _denetle(a):
    from docx import Document

    from .kontroller.belge import safha_tablosu_denetle
    from .kontroller.cikti import konsol_ozeti
    from .kontroller.liste import KontrolListesi
    from .ortak import PARA_RE, num
    from .rapor.docx_araclari import find_row, uniq_cells
    from .rapor.tablolar import tablolari_bul
    kl = KontrolListesi()
    with tempfile.TemporaryDirectory() as tmp:
        d = Document(str(docx_hazirla(yol_temizle(a.rapor), tmp)))
    T, sorunlar = tablolari_bul(d)
    for _, ac in sorunlar:
        kl.uyari('Tablo bulma', ac)
    base = None
    if 'karsit' in T:
        r = find_row(T['karsit'], r'İndirilecek KDV Tutarı')
        if r is not None:
            p = [c.text.strip() for c in uniq_cells(r) if PARA_RE.match(c.text.strip())]
            base = num(p[0]) if p else None
    if 'safha' in T:
        safha_tablosu_denetle(kl, T['safha'], base)
    print(konsol_ozeti(kl))
    for k in kl.sirali():
        if k.durum == 'TAMAM':
            print(f'  [TAMAM] {k.konu}: {k.aciklama}')
    return 1 if kl.say('HATA') else 0


def _geriye_donuk(a):
    from .geriye_donuk import geriye_donuk, metin_ozeti
    ayar = ayar_oku(a.ayar)
    kok = yol_temizle(_klasor(a))
    cikti = Path(a.cikti) if a.cikti else kok / 'CLAUDE TASLAK' / 'GERİYE DÖNÜK TEST'
    s = geriye_donuk(ayar, kok, cikti, Donem.coz(a.baslangic) if a.baslangic else None,
                     Donem.coz(a.bitis) if a.bitis else None, kati=a.kati, ilerleme=lambda m: print(m))
    print(metin_ozeti(s))
    print(f'\nAyrıntı: {cikti}')
    return 0


def _arayuz(a):
    from .arayuz import main as arayuz_main
    arayuz_main(getattr(a, 'klasor', None))
    return 0


def _klasor_argumanlari(p, eski=True):
    p.add_argument('klasor', nargs='?', help='firmanın o ayki klasörü (içindeki belgeler içerikten tanınır)')
    if eski:
        p.add_argument('--girdi', help=argparse.SUPPRESS)
        p.add_argument('--kok', help=argparse.SUPPRESS)


def main(argv=None):
    for akis in (sys.stdout, sys.stderr):
        try:
            akis.reconfigure(errors='replace')
        except AttributeError:
            pass
    p = argparse.ArgumentParser(prog='python -m teminat', description='KDV iadesi teminat çözüm raporu taslak programı')
    alt = p.add_subparsers(dest='komut', required=True)

    for ad, yardim, fn in (('tani', 'klasördeki belgeleri içerikten tanır ve tanıma tablosunu gösterir', _tani),
                           ('taslak', 'önce tanıma tablosunu gösterir, sonra taslağı ve kontrol listesini üretir', _taslak)):
        t = alt.add_parser(ad, help=yardim)
        _klasor_argumanlari(t)
        t.add_argument('--ayar', help='firma ayar dosyası (firmalar/<firma>/ayar.yaml)')
        t.add_argument('--sablon', help='önceki ayın bitmiş raporu (.doc/.docx); verilmezse içerikten aranır')
        t.add_argument('--sec', action='append', metavar='ROL=DOSYA',
                       help=f'içerikten karar verilemeyen belgeyi elle seç; roller: {", ".join(SECILEBILIR)}')
        t.add_argument('--donem', help='klasörde birden fazla dönemin beyannamesi varsa hangisi (ör. 2026-03)')
        if ad == 'taslak':
            t.add_argument('--cikti', help='çıktı klasörü (varsayılan: <klasör>\\CLAUDE TASLAK)')
            t.add_argument('--uzerine-yaz', action='store_true', help='aynı adlı eski taslağın üzerine yaz')
            t.add_argument('--kati', action='store_true', help='ilk hatada dur (geliştirme için)')
        t.set_defaults(fn=fn)

    k = alt.add_parser('karsilastir', help='taslak ile ofisin bitmiş raporunu karşılaştırır')
    k.add_argument('gercek')
    k.add_argument('taslak')
    k.add_argument('--cikti', help='farkları bu dosyaya yaz')
    k.set_defaults(fn=_karsilastir)

    dn = alt.add_parser('denetle', help='bir raporun safha tablosunda toplam / satır kayması kontrolü')
    dn.add_argument('rapor')
    dn.set_defaults(fn=_denetle)

    g = alt.add_parser('geriye-donuk', help='firma klasöründeki her ayın taslağını önceki aydan üretip gerçek raporla karşılaştırır')
    _klasor_argumanlari(g)
    g.add_argument('--ayar')
    g.add_argument('--cikti')
    g.add_argument('--baslangic', help='ör. 2025-01')
    g.add_argument('--bitis', help='ör. 2026-02')
    g.add_argument('--kati', action='store_true')
    g.set_defaults(fn=_geriye_donuk)

    ar = alt.add_parser('arayuz', help='pencereli arayüzü açar')
    ar.add_argument('klasor', nargs='?', help='açılışta seçili olacak klasör')
    ar.set_defaults(fn=_arayuz)

    a = p.parse_args(argv)
    try:
        return a.fn(a)
    except BelirsizSecim as e:
        print(f'SEÇİM GEREKLİ ({rol_adi(e.rol)}): {e.aciklama}', file=sys.stderr)
        for i, y in enumerate(e.adaylar, 1):
            print(f'  {i}) {y}', file=sys.stderr)
        secenek = '--sablon "<dosya>"' if e.rol == 'SABLON' else f'--sec {e.rol}="<dosya>"'
        print(f'Hangisinin kullanılacağını {secenek} ile belirtip yeniden çalıştırın.', file=sys.stderr)
        return 3
    except (GirdiHatasi, SablonHatasi, TanimaHatasi, DonusumHatasi, ValueError) as e:
        print(f'HATA: {e}', file=sys.stderr)
        return 2
