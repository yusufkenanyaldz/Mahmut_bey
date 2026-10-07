"""Komut satırı.

    python -m teminat taslak --ayar firmalar/<firma>/ayar.yaml --kok "C:\\...\\TEMİNAT ÇÖZÜMÜ\\<FİRMA>" --donem 2026-03
    python -m teminat taslak --ayar firmalar/<firma>/ayar.yaml --girdi "...\\2026\\03 MART"   (dönem KDV 1'den, şablon önceki aydan)
    python -m teminat taslak --ayar firmalar/<firma>/ayar.yaml --sablon "ŞUBAT-2026 RAPOR.doc" --girdi "...\\2026\\03 MART"
    python -m teminat klasor --kok "...\\<FİRMA>" --donem 2026-03                        (yalnızca klasörleri bul)
    python -m teminat karsilastir "gerçek rapor.doc(x)" "taslak.docx"
    python -m teminat denetle "rapor.doc(x)"
    python -m teminat geriye-donuk --ayar firmalar/<firma>/ayar.yaml --kok "...\\<FİRMA>" --cikti "...\\GERİYE DÖNÜK TEST"
    python -m teminat arayuz          (pencereli arayüz; .exe çift tıklanınca da bu açılır)
"""
import argparse
import sys
import tempfile
from pathlib import Path

from .ayar import ayar_oku
from .donusum import DonusumHatasi, docx_hazirla
from .klasorler import KlasorHatasi, girdi_ve_sablon_bul
from .okuyucular.girdiler import GirdiHatasi
from .ortak import Donem
from .rapor.olustur import SablonHatasi


def _bul(a):
    """Seçilen klasörden bu ayın klasörünü ve önceki ayın raporunu bulur; bulamazsa nedenini anlatan hata verir."""
    secilen = a.girdi or a.kok
    if not secilen:
        raise SystemExit('--kok (firma klasörü + --donem) ya da --girdi (o ayın klasörü) verilmeli.')
    return girdi_ve_sablon_bul(secilen, Donem.coz(a.donem) if a.donem else None)


def _klasor(a):
    b = _bul(a)
    print(b.ozet())
    return 0 if b.sablon else 1


def _taslak(a):
    from .kontroller.cikti import konsol_ozeti
    from .taslak import taslak_uret
    ayar = ayar_oku(a.ayar)
    if a.sablon and a.girdi:
        sablon, girdi = Path(a.sablon), Path(a.girdi)
    else:
        b = _bul(a)
        for n in b.notlar:
            print(f'Not: {n}')
        girdi = b.girdi
        sablon = Path(a.sablon) if a.sablon else b.sablon
        if not sablon:
            raise KlasorHatasi(f'Önceki ayın ({b.donem.onceki().tire}) bitmiş raporu bulunamadı; şablon olarak kullanılacak '
                               'raporu "Önceki ayın raporu" alanından (komut satırında --sablon) seçin.')
    print(f'Şablon : {sablon}\nGirdi  : {girdi}')
    tc = taslak_uret(ayar, sablon, girdi, a.cikti, kati=a.kati, uzerine_yaz=a.uzerine_yaz)
    print(f'Taslak : {tc.docx}\nKontrol: {tc.xlsx}\n         {tc.html}\n')
    print(konsol_ozeti(tc.sonuc.kontroller))
    return 1 if tc.sonuc.kontroller.say('HATA') else 0


def _karsilastir(a):
    from .karsilastir import karsilastir
    with tempfile.TemporaryDirectory() as tmp:
        k = karsilastir(docx_hazirla(a.gercek, tmp), docx_hazirla(a.taslak, tmp))
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
        d = Document(str(docx_hazirla(a.rapor, tmp)))
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
    cikti = Path(a.cikti) if a.cikti else Path(a.kok) / 'CLAUDE TASLAK' / 'GERİYE DÖNÜK TEST'
    s = geriye_donuk(ayar, a.kok, cikti, Donem.coz(a.baslangic) if a.baslangic else None,
                     Donem.coz(a.bitis) if a.bitis else None, kati=a.kati)
    print(metin_ozeti(s))
    print(f'\nAyrıntı: {cikti}')
    return 0


def _arayuz(a):
    from .arayuz import main as arayuz_main
    arayuz_main()
    return 0


def main(argv=None):
    for akis in (sys.stdout, sys.stderr):
        try:
            akis.reconfigure(errors='replace')
        except AttributeError:
            pass
    p = argparse.ArgumentParser(prog='python -m teminat', description='KDV iadesi teminat çözüm raporu taslak programı')
    alt = p.add_subparsers(dest='komut', required=True)

    t = alt.add_parser('taslak', help='önceki ayın raporundan bu ayın taslağını ve kontrol listesini üretir')
    t.add_argument('--ayar', help='firma ayar dosyası (firmalar/<firma>/ayar.yaml)')
    t.add_argument('--kok', help='firma klasörü (ör. ...\\TEMİNAT ÇÖZÜMÜ\\<FİRMA>); --donem ile birlikte')
    t.add_argument('--donem', help='rapor dönemi: 2026-03 ya da MART-2026 (--girdi verilirse KDV 1\'den okunur)')
    t.add_argument('--sablon', help='önceki ayın bitmiş raporu (.doc/.docx); verilmezse önceki ayın klasöründe aranır')
    t.add_argument('--girdi', help='bu ayın girdi klasörü (KDV 1.pdf\'in olduğu klasör)')
    t.add_argument('--cikti', help='çıktı klasörü (varsayılan: <girdi>\\CLAUDE TASLAK)')
    t.add_argument('--uzerine-yaz', action='store_true', help='aynı adlı eski taslağın üzerine yaz')
    t.add_argument('--kati', action='store_true', help='ilk hatada dur (geliştirme için)')
    t.set_defaults(fn=_taslak)

    b = alt.add_parser('klasor', help='seçilen klasörde bu ayın klasörünü ve önceki ayın raporunu bulur (deneme)')
    b.add_argument('--kok', help='firma klasörü')
    b.add_argument('--girdi', help='o ayın klasörü')
    b.add_argument('--donem', help='2026-03')
    b.set_defaults(fn=_klasor)

    k = alt.add_parser('karsilastir', help='taslak ile ofisin bitmiş raporunu karşılaştırır')
    k.add_argument('gercek')
    k.add_argument('taslak')
    k.add_argument('--cikti', help='farkları bu dosyaya yaz')
    k.set_defaults(fn=_karsilastir)

    dn = alt.add_parser('denetle', help='bir raporun safha tablosunda toplam / satır kayması kontrolü')
    dn.add_argument('rapor')
    dn.set_defaults(fn=_denetle)

    g = alt.add_parser('geriye-donuk', help='her ayın taslağını önceki aydan üretip gerçek raporla karşılaştırır')
    g.add_argument('--ayar')
    g.add_argument('--kok', required=True)
    g.add_argument('--cikti')
    g.add_argument('--baslangic', help='ör. 2025-01')
    g.add_argument('--bitis', help='ör. 2026-02')
    g.add_argument('--kati', action='store_true')
    g.set_defaults(fn=_geriye_donuk)

    ar = alt.add_parser('arayuz', help='pencereli arayüzü açar')
    ar.set_defaults(fn=_arayuz)

    a = p.parse_args(argv)
    try:
        return a.fn(a)
    except (GirdiHatasi, SablonHatasi, KlasorHatasi, DonusumHatasi, ValueError) as e:
        print(f'HATA: {e}', file=sys.stderr)
        return 2
