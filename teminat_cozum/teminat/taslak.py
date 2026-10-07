"""Bir ay için taslak + kontrol listesi üretimi (komut satırı, pencere ve geriye dönük test bunu kullanır)."""
from dataclasses import dataclass
from pathlib import Path

from .donusum import docx_hazirla
from .islem import hazirla
from .kontroller import cikti
from .okuyucular.girdiler import secimden_oku
from .rapor.olustur import SablonHatasi, taslak_olustur
from .tanima import EK_TURLER

CIKTI_KLASORU = 'CLAUDE TASLAK'


@dataclass
class TaslakCiktisi:
    docx: Path
    xlsx: Path
    html: Path
    sonuc: object
    hazirlik: object = None
    farklar: Path = None
    karsilastirma: object = None


def _bos_ad(klasor, adlar, uzerine_yaz):
    """Var olan dosyaların üzerine yazmamak için ' (2)', ' (3)'... ekler (bütün dosyalar aynı eki alır)."""
    if uzerine_yaz:
        return [klasor / a for a in adlar]
    n = 1
    while True:
        ek = '' if n == 1 else f' ({n})'
        yollar = [klasor / (Path(a).stem + ek + Path(a).suffix) for a in adlar]
        if not any(p.exists() for p in yollar):
            return yollar
        n += 1


def belge_satirlari(hz):
    """Kontrol listesindeki "Belgeler" sayfası için tanıma tablosu satırları."""
    out = []
    for b in sorted(hz.secim.belgeler, key=lambda b: (b.rol == '', b.tur in EK_TURLER, b.tur, b.goreli)):
        kullanim = {'BU_AYIN_RAPORU': 'karşılaştırma'}.get(b.rol, 'KULLANILDI' if b.rol else '')
        out.append((b.goreli, b.tur + (' (elle)' if b.elle else ''), b.aciklama, b.donem.tire if b.donem else '', kullanim, b.notu))
    if hz.sablon:
        out.insert(0, (str(hz.sablon), 'RAPOR', f'Şablon: önceki ayın raporu ({hz.sablon_kaynagi})',
                       hz.secim.donem.onceki().tire, 'ŞABLON', ''))
    return out


def taslak_uret(ayar, sablon, girdi, cikti_klasoru=None, kati=False, uzerine_yaz=False, zorla=None, hz=None, ilerleme=None):
    hz = hz or hazirla(girdi, ayar, sablon, zorla, ilerleme=ilerleme)
    if not hz.sablon:
        raise SablonHatasi(f'Önceki ayın ({hz.secim.donem.onceki().tire}) bitmiş raporu bulunamadı; şablon olarak kullanılacak '
                           'raporu "Önceki ayın raporu" alanından (komut satırında --sablon) seçin.')
    girdi = hz.klasor
    cikti_klasoru = Path(cikti_klasoru) if cikti_klasoru else girdi / CIKTI_KLASORU
    cikti_klasoru.mkdir(parents=True, exist_ok=True)
    g = secimden_oku(hz.secim)
    g.notlar += hz.notlar
    sablon_docx = docx_hazirla(hz.sablon, cikti_klasoru / '_cevrilen')
    d = g.donem
    docx, xlsx, html, fark = _bos_ad(cikti_klasoru, [f'{d.tire} RAPOR TASLAK.docx', f'{d.tire} KONTROL LİSTESİ.xlsx',
                                                     f'{d.tire} KONTROL LİSTESİ.html', f'{d.tire} FARKLAR.txt'], uzerine_yaz)
    girdi_dosyalari = {Path(p).resolve() for p in [hz.sablon, *hz.secim.yollar.values()] if p}
    if any(p.resolve() in girdi_dosyalari for p in (docx, xlsx, html, fark)):
        raise ValueError('Çıktı bir girdi dosyasının üzerine yazılamaz.')
    if ilerleme:
        ilerleme('Taslak oluşturuluyor…')
    sonuc = taslak_olustur(sablon_docx, g, ayar, docx, kati=kati)
    sonuc.ozet['dosyalar'] = {**{rol: Path(y).relative_to(girdi).as_posix() if Path(y).is_relative_to(girdi) else str(y)
                                 for rol, y in hz.secim.yollar.items()},
                              'Şablon': str(hz.sablon)}
    sonuc.ozet['belgeler'] = belge_satirlari(hz)
    tc = TaslakCiktisi(docx, xlsx, html, sonuc, hz)
    if hz.secim.bu_ayin_raporu:                   # bu ayın bitmiş raporu da klasördeyse: öğrenme döngüsü
        from .karsilastir import karsilastir
        try:
            k = karsilastir(docx_hazirla(hz.secim.bu_ayin_raporu, cikti_klasoru / '_cevrilen'), docx)
            fark.write_text(k.metin(), encoding='utf-8')
            tc.farklar, tc.karsilastirma = fark, k
            sonuc.kontroller.ekle('BİLGİ' if k.tutar_farki == 0 else 'UYARI', 'Gerçek raporla karşılaştırma',
                                  f'Klasördeki bu ayın bitmiş raporuyla karşılaştırıldı: tutar farkı {k.tutar_farki}, '
                                  f'farklı satır {k.farkli_satir}. Ayrıntı: {fark.name}')
        except Exception as e:  # noqa: BLE001 - karşılaştırma yapılamasa da taslak üretilmiş olsun
            sonuc.kontroller.bilgi('Gerçek raporla karşılaştırma', f'Karşılaştırma yapılamadı: {e}')
    baslik = f'{ayar.kisa_ad or ayar.firma} {d.tire} — Kontrol Listesi'.strip(' —')
    cikti.xlsx_yaz(sonuc.kontroller, sonuc.ozet, xlsx)
    cikti.html_yaz(sonuc.kontroller, sonuc.ozet, html, baslik)
    return tc
