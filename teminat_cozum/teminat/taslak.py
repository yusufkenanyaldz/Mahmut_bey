"""Bir ay için taslak + kontrol listesi üretimi (komut satırı ve geriye dönük test bunu kullanır)."""
from dataclasses import dataclass
from pathlib import Path

from .donusum import docx_hazirla
from .kontroller import cikti
from .okuyucular.girdiler import girdileri_oku
from .okuyucular.listeler import dosya_adi
from .rapor.olustur import taslak_olustur

CIKTI_KLASORU = 'CLAUDE TASLAK'


@dataclass
class TaslakCiktisi:
    docx: Path
    xlsx: Path
    html: Path
    sonuc: object


def _bos_ad(klasor, adlar, uzerine_yaz):
    """Var olan dosyaların üzerine yazmamak için ' (2)', ' (3)'... ekler (üç dosya aynı eki alır)."""
    if uzerine_yaz:
        return [klasor / a for a in adlar]
    n = 1
    while True:
        ek = '' if n == 1 else f' ({n})'
        yollar = [klasor / (Path(a).stem + ek + Path(a).suffix) for a in adlar]
        if not any(p.exists() for p in yollar):
            return yollar
        n += 1


def taslak_uret(ayar, sablon, girdi, cikti_klasoru=None, kati=False, uzerine_yaz=False):
    girdi = Path(girdi)
    cikti_klasoru = Path(cikti_klasoru) if cikti_klasoru else girdi / CIKTI_KLASORU
    cikti_klasoru.mkdir(parents=True, exist_ok=True)
    g = girdileri_oku(girdi)
    sablon_docx = docx_hazirla(sablon, cikti_klasoru / '_cevrilen')
    d = g.donem
    docx, xlsx, html = _bos_ad(cikti_klasoru, [f'{d.tire} RAPOR TASLAK.docx', f'{d.tire} KONTROL LİSTESİ.xlsx',
                                               f'{d.tire} KONTROL LİSTESİ.html'], uzerine_yaz)
    girdi_dosyalari = {Path(p).resolve() for p in (sablon, g.liste_yolu, g.takip_yolu, g.teminat_yolu, g.yuklenilen_yolu) if p}
    if any(p.resolve() in girdi_dosyalari for p in (docx, xlsx, html)):
        raise ValueError('Çıktı bir girdi dosyasının üzerine yazılamaz.')
    sonuc = taslak_olustur(sablon_docx, g, ayar, docx, kati=kati)
    sonuc.ozet['dosyalar'] = {
        'KDV 1': dosya_adi(g.k.get('_file')), 'İndirilecek liste': dosya_adi(g.liste_yolu), 'Takip listesi': dosya_adi(g.takip_yolu),
        'Teminat dilekçesi': dosya_adi(g.teminat_yolu), 'Yüklenilen tutanak': dosya_adi(g.yuklenilen_yolu),
    }
    baslik = f'{ayar.kisa_ad or ayar.firma} {d.tire} — Kontrol Listesi'.strip(' —')
    cikti.xlsx_yaz(sonuc.kontroller, sonuc.ozet, xlsx)
    cikti.html_yaz(sonuc.kontroller, sonuc.ozet, html, baslik)
    return TaslakCiktisi(docx, xlsx, html, sonuc)
