"""Geriye dönük test: her ayın taslağını bir önceki ayın gerçek raporundan üretir ve o ayın gerçek raporuyla
karşılaştırır. Sonuç: dönem başına tutar farkı, farklı satır sayısı, kontrol listesi özeti.
"""
import os
import traceback
from dataclasses import dataclass, field
from pathlib import Path

from .islem import hazirla
from .ortak import Donem
from .tanima import (BelirsizSecim, Onbellek, TanimaHatasi, _pdf_metni, atlanan_klasor, atlanir, cikti_klasoru_hazirla,
                     cikti_klasoru_mu, nf, siniflandir)
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


def _kdv1_donemi(p, onbellek):
    c = onbellek.al(p, 'kdv1')
    if c is not None:
        return Donem.coz(c['donem']) if c.get('donem') else None
    from .okuyucular.kdv1 import kdv1_donem, kdv1_metinden, pdf_metni
    try:
        yd = kdv1_donem(kdv1_metinden(pdf_metni(p)))
    except Exception:  # noqa: BLE001
        yd = None
    d = Donem(*yd) if yd else None
    onbellek.koy(p, {'donem': d.tire if d else None}, 'kdv1')
    return d


def ay_klasorlerini_bul(kok, onbellek=None, ilerleme=None):
    """Kök altındaki ay klasörlerini İÇERİKTEN bulur: 1 No.lu beyanname içeren en üst klasör (altında başka dönemin
    beyannamesi olmayan). Dönüş: {Donem: [klasör]}. Klasör ve dosya adlarına bakılmaz."""
    kok = Path(kok).resolve()
    onbellek = onbellek if onbellek is not None else Onbellek()
    yerler = []
    n = 0
    for d, dirs, files in os.walk(kok):
        if cikti_klasoru_mu(d):
            dirs[:] = []
            continue
        dirs[:] = sorted(x for x in dirs if not atlanan_klasor(x))
        for f in sorted(files):
            if atlanir(f) or not f.lower().endswith('.pdf'):
                continue
            p = Path(d) / f
            n += 1
            if ilerleme and n % 200 == 0:
                ilerleme(f'{n} PDF tarandı…')
            c = onbellek.al(p)
            if c is None:
                try:
                    c = {'tur': siniflandir(nf(_pdf_metni(p)))}
                except Exception:  # noqa: BLE001
                    continue
                onbellek.koy(p, c)
            if c['tur'] == 'KDV1':
                donem = _kdv1_donemi(p, onbellek)
                if donem:
                    yerler.append((Path(d), donem))
    onbellek.kaydet()
    alt = {}
    for d, donem in yerler:
        for q in [d, *d.parents]:
            alt.setdefault(q, set()).add(donem)
            if q == kok:
                break
    aylar = {}
    for d, donem in yerler:
        m = d
        while m != kok and m.parent != kok and alt.get(m.parent) == {donem}:
            m = m.parent
        aylar.setdefault(donem, set()).add(m)
    return {d: sorted(v) for d, v in aylar.items()}


def geriye_donuk(ayar, kok, cikti_klasoru, baslangic=None, bitis=None, kati=False, ilerleme=None):
    """Her ay klasörü için: belgeleri içerikten tanı → önceki ayın raporunu şablon olarak bul → taslak üret → o ayın
    klasöründeki bitmiş raporla karşılaştır."""
    cikti_klasoru = cikti_klasoru_hazirla(cikti_klasoru)
    ob = Onbellek()
    aylar = ay_klasorlerini_bul(kok, ob, ilerleme)
    if not aylar:
        raise TanimaHatasi(f'{kok} altında 1 No.lu KDV beyannamesi içeren klasör bulunamadı; firmanın klasörünü seçin.')
    sonuc = []
    for d in sorted(aylar):
        if (baslangic and d < baslangic) or (bitis and d > bitis):
            continue
        if ilerleme:
            ilerleme(f'{d.tire} hazırlanıyor…')
        try:
            if len(aylar[d]) > 1:
                sonuc.append(AySonucu(d, 'atlandı', f'{d.tire} beyannamesi birden fazla klasörde var: '
                                      + '; '.join(map(str, aylar[d]))))
                continue
            klasor = aylar[d][0]
            hz = hazirla(klasor, ayar, istenen=d, onbellek=ob)
            if not hz.sablon:
                sonuc.append(AySonucu(d, 'atlandı', f'Şablon için {d.onceki().tire} raporu bulunamadı.'))
                continue
            if not hz.secim.bu_ayin_raporu:
                neden = [ac for _, konu, ac in hz.secim.notlar if konu == 'Gerçek raporla karşılaştırma']
                sonuc.append(AySonucu(d, 'atlandı', neden[0] if neden else f'{klasor} içinde karşılaştırılacak {d.tire} raporu yok.'))
                continue
            tc = taslak_uret(ayar, None, klasor, cikti_klasoru / d.tire, kati=kati, uzerine_yaz=True, hz=hz)
            k = tc.karsilastirma
            if k is None:
                sonuc.append(AySonucu(d, 'hata', 'Taslak üretildi ama gerçek raporla karşılaştırılamadı.', taslak=tc.docx))
                continue
            sonuc.append(AySonucu(d, 'tamam', '', k.tutar_farki, k.farkli_satir, tc.sonuc.kontroller.ozet(), tc.farklar, tc.docx))
        except BelirsizSecim as e:
            sonuc.append(AySonucu(d, 'atlandı', f'Belge seçimi belirsiz ({e.rol}): {e.aciklama} '
                                  + '; '.join(map(str, e.adaylar))))
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
