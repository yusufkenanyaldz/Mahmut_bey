"""Geriye dönük test: her ayın taslağını bir önceki ayın gerçek raporundan üretir ve o ayın gerçek raporuyla
karşılaştırır. Sonuç: dönem başına tutar farkı, farklı satır sayısı, kontrol listesi özeti.
"""
import traceback
from dataclasses import dataclass, field
from pathlib import Path

from .donusum import docx_hazirla
from .karsilastir import karsilastir
from .klasorler import ay_klasorleri, rapor_dosyasi
from .taslak import taslak_uret


@dataclass
class AySonucu:
    donem: object
    durum: str                 # 'tamam' | 'atlandı' | 'hata'
    aciklama: str = ''
    tutar_farki: int = None
    farkli_satir: int = None
    kontrol: dict = field(default_factory=dict)
    fark_dosyasi: Path = None
    taslak: Path = None


def geriye_donuk(ayar, kok, cikti_klasoru, baslangic=None, bitis=None, kati=False):
    cikti_klasoru = Path(cikti_klasoru)
    cikti_klasoru.mkdir(parents=True, exist_ok=True)
    cev = cikti_klasoru / '_cevrilen'
    aylar = ay_klasorleri(kok)
    sonuc = []
    for d in sorted(aylar):
        if (baslangic and d < baslangic) or (bitis and d > bitis):
            continue
        onceki = d.onceki()
        if onceki not in aylar:
            sonuc.append(AySonucu(d, 'atlandı', f'{onceki.tire} klasörü yok (şablon için önceki ayın raporu gerekli).'))
            continue
        try:
            sablon = rapor_dosyasi(aylar[onceki])
            gercek = rapor_dosyasi(aylar[d])
            if not sablon:
                sonuc.append(AySonucu(d, 'atlandı', f'{onceki.tire}/RAPOR içinde rapor yok.'))
                continue
            if not gercek:
                sonuc.append(AySonucu(d, 'atlandı', f'{d.tire}/RAPOR içinde karşılaştırılacak rapor yok.'))
                continue
            ay_cikti = cikti_klasoru / d.tire
            tc = taslak_uret(ayar, sablon, aylar[d], ay_cikti, kati=kati, uzerine_yaz=True)
            k = karsilastir(docx_hazirla(gercek, cev), tc.docx)
            fark = ay_cikti / f'{d.tire} FARKLAR.txt'
            fark.write_text(k.metin(), encoding='utf-8')
            sonuc.append(AySonucu(d, 'tamam', '', k.tutar_farki, k.farkli_satir, tc.sonuc.kontroller.ozet(), fark, tc.docx))
        except Exception as e:  # noqa: BLE001 - bir ayın hatası diğer ayları durdurmasın
            if kati:
                raise
            sonuc.append(AySonucu(d, 'hata', f'{type(e).__name__}: {e}\n{traceback.format_exc(limit=3)}'))
    ozet_yaz(sonuc, cikti_klasoru)
    return sonuc


def ozet_yaz(sonuc, klasor):
    from openpyxl import Workbook
    from openpyxl.styles import Font

    wb = Workbook()
    ws = wb.active
    ws.title = 'Geriye dönük test'
    ws.append(['Dönem', 'Durum', 'Tutar farkı', 'Farklı satır', 'HATA', 'UYARI', 'ELLE', 'Açıklama', 'Fark dosyası'])
    for c in ws[1]:
        c.font = Font(bold=True)
    for s in sonuc:
        ws.append([s.donem.tire, s.durum, s.tutar_farki, s.farkli_satir, s.kontrol.get('HATA'), s.kontrol.get('UYARI'),
                   s.kontrol.get('ELLE'), s.aciklama.split('\n')[0], str(s.fark_dosyasi or '')])
    for col, w in zip('ABCDEFGHI', (14, 10, 12, 12, 8, 8, 8, 70, 60)):
        ws.column_dimensions[col].width = w
    wb.save(str(Path(klasor) / 'GERİYE DÖNÜK TEST.xlsx'))


def metin_ozeti(sonuc):
    out = ['Dönem        Durum     Tutar farkı  Farklı satır  HATA/UYARI']
    for s in sonuc:
        if s.durum == 'tamam':
            out.append(f'{s.donem.tire:<12} {s.durum:<9} {s.tutar_farki:>11}  {s.farkli_satir:>12}  '
                       f'{s.kontrol.get("HATA", 0)}/{s.kontrol.get("UYARI", 0)}')
        else:
            out.append(f'{s.donem.tire:<12} {s.durum:<9} {s.aciklama.splitlines()[0]}')
    return '\n'.join(out)
