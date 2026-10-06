"""Komut satırı.

    python -m teminat taslak --ayar firmalar/<firma>/ayar.yaml --kok "C:\\...\\TEMİNAT ÇÖZÜMÜ\\<FİRMA>" --donem 2026-03
    python -m teminat taslak --ayar firmalar/<firma>/ayar.yaml --sablon "ŞUBAT-2026 RAPOR.doc" --girdi "...\\2026\\03 MART"
    python -m teminat karsilastir "gerçek rapor.doc(x)" "taslak.docx"
    python -m teminat denetle "rapor.doc(x)"
    python -m teminat geriye-donuk --ayar firmalar/<firma>/ayar.yaml --kok "...\\<FİRMA>" --cikti "...\\GERİYE DÖNÜK TEST"
"""
import argparse
import sys
import tempfile
from pathlib import Path

from .ayar import ayar_oku
from .donusum import DonusumHatasi, docx_hazirla
from .klasorler import KlasorHatasi, ay_klasoru, rapor_dosyasi
from .okuyucular.girdiler import GirdiHatasi
from .ortak import Donem
from .rapor.olustur import SablonHatasi


def _taslak(a):
    from .kontroller.cikti import konsol_ozeti
    from .taslak import taslak_uret
    ayar = ayar_oku(a.ayar)
    if a.kok:
        if not a.donem:
            raise SystemExit('--kok ile birlikte --donem verilmeli (ör. 2026-03).')
        d = Donem.coz(a.donem)
        girdi = Path(a.girdi) if a.girdi else ay_klasoru(a.kok, d)
        sablon = Path(a.sablon) if a.sablon else rapor_dosyasi(ay_klasoru(a.kok, d.onceki()))
        if not sablon:
            raise SystemExit(f'{d.onceki().tire} RAPOR klasöründe rapor bulunamadı; --sablon ile verin.')
    else:
        if not (a.sablon and a.girdi):
            raise SystemExit('--sablon ve --girdi (ya da --kok ve --donem) verilmeli.')
        sablon, girdi = Path(a.sablon), Path(a.girdi)
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
    t.add_argument('--donem', help='rapor dönemi: 2026-03 ya da MART-2026')
    t.add_argument('--sablon', help='önceki ayın bitmiş raporu (.doc/.docx)')
    t.add_argument('--girdi', help='bu ayın girdi klasörü')
    t.add_argument('--cikti', help='çıktı klasörü (varsayılan: <girdi>\\CLAUDE TASLAK)')
    t.add_argument('--uzerine-yaz', action='store_true', help='aynı adlı eski taslağın üzerine yaz')
    t.add_argument('--kati', action='store_true', help='ilk hatada dur (geliştirme için)')
    t.set_defaults(fn=_taslak)

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

    a = p.parse_args(argv)
    try:
        return a.fn(a)
    except (GirdiHatasi, SablonHatasi, KlasorHatasi, DonusumHatasi, ValueError) as e:
        print(f'HATA: {e}', file=sys.stderr)
        return 2
