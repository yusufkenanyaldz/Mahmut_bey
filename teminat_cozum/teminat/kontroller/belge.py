"""Word belgesi üzerinde yapılan kontroller: safha tablosu toplamları, eski kalmış metin, güncellenmemiş tutarlar.

Bunlar hem programın ürettiği taslağa hem de ofisin bitmiş raporuna (`python -m teminat denetle`) uygulanabilir.
"""
import collections
import re

from ..ortak import PARA_RE, num, tr
from ..rapor.docx_araclari import YAZILAN, tum_paragraflar, uniq_cells

GENEL = 'GENEL'


def _satir_tutari(cs):
    """Satırdaki son tutar hücresi (KDV TUTARI sütunu)."""
    for c in reversed(cs):
        t = c.text.strip()
        if PARA_RE.match(t):
            return num(t)
    return None


def safha_tablosu_denetle(kl, tablo, base=None, yer='Tablo 3-4-3', tol=0.01):
    """5. Satır toplamı ↔ TOPLAM hücreleri; satır kayması (aynı tutar iki firmada); sıra numaraları."""
    ekli, muh, toplamlar = [], [], []
    for r in list(tablo.rows)[1:]:
        cs = uniq_cells(r)
        metin = ' | '.join(c.text.strip() for c in cs)
        v = _satir_tutari(cs)
        if 'TOPLAM' in metin:
            toplamlar.append((metin, v))
            continue
        if v is None:
            continue
        durum = next((c.text.strip() for c in cs if re.match(r'^\d+\s*-\s*(EKLİ|İTHALAT)', c.text.strip())), None)
        firma = cs[1].text.strip() if len(cs) > 1 else ''
        if durum:
            ekli.append((firma, durum, v))
        elif 'Muhafaza' in metin:
            muh.append((firma, v))
    sorun = 0
    if len(toplamlar) < 3:
        kl.uyari('Safha tablosu', f'{yer}: EKLİ / muhafaza / genel toplam satırları beklenen yapıda değil ({len(toplamlar)} toplam satırı).', no='5', yer=yer)
        return
    (_, t_e), (_, t_m), (_, t_g) = toplamlar[0], toplamlar[-2], toplamlar[-1]
    s_e, s_m = sum(v for *_, v in ekli), sum(v for _, v in muh)
    for ad, hesap, hucre in (('EKLİ toplamı', s_e, t_e), ('Muhafaza toplamı', s_m, t_m), ('Genel toplam', s_e + s_m, t_g)):
        if hucre is None or abs(hesap - hucre) > tol:
            sorun += 1
            kl.hata('Safha toplamı', f'{yer}: {ad} satırların toplamı {tr(hesap)}, tablodaki TOPLAM hücresi '
                    f'{tr(hucre) if hucre is not None else "boş"}.', no='5', yer=yer, beklenen=tr(hesap),
                    bulunan=tr(hucre) if hucre is not None else '')
    numara = [int(re.match(r'^(\d+)', d).group(1)) for _, d, _ in ekli]
    if numara != list(range(1, len(numara) + 1)):
        sorun += 1
        kl.uyari('Safha sıra no', f'{yer}: EKLİ/İTHALAT sıra numaraları ardışık değil: {numara}', no='5', yer=yer)
    sayac = collections.Counter(round(v, 2) for *_, v in ekli + [(f, v) for f, v in muh])
    for v, n in sayac.items():
        if n > 1:
            sorun += 1
            adlar = [f for f, *_, x in ekli + muh if round(x, 2) == v]
            kl.uyari('Satır kayması', f'{yer}: {tr(v)} tutarı {n} satırda: {", ".join(adlar)} — satır kayması olabilir.', no='5', yer=yer)
    if base:
        if t_g is not None and t_g < base * 0.8 - tol:
            sorun += 1
            kl.hata('Karşıt inceleme oranı', f'{yer}: genel toplam {tr(t_g)} indirilecek KDV\'nin %80\'i ({tr(base * 0.8)}) altında.', no='3', yer=yer)
    if not sorun:
        kl.tamam('Safha tablosu', f'{yer}: {len(ekli)} EKLİ/İTHALAT + {len(muh)} muhafaza satırı; toplamlar tutuyor, tekrar eden tutar yok.', no='5', yer=yer)


def eski_metin_tara(kl, doc, desenler, haric_tablolar=(), atla=None):
    """11. Önceki dönem adı gibi eski kalmış metinleri bulur. desenler: [(regex, açıklama)]."""
    bulgular = []
    for yer, p in tum_paragraflar(doc):
        if yer in haric_tablolar:
            continue
        t = p.text
        if atla and atla(t):
            continue
        for pat, ac in desenler:
            m = re.search(pat, t)
            if m:
                s = max(0, m.start() - 50)
                bulgular.append((yer, ac, '…' + ' '.join(t[s:m.end() + 50].split()) + '…'))
    for yer, ac, ozet in bulgular:
        kl.uyari('Eski metin', f'{ac}: {ozet}', no='11', yer=yer)
    if not bulgular:
        kl.tamam('Eski metin', 'Önceki dönem adı taslakta kalmadı.', no='11')
    return bulgular


def guncellenmemis_tutarlar(kl, tablolar):
    """Doldurulan tablolarda program tarafından yazılmamış tutar hücreleri (önceki aydan kalmış olabilir)."""
    sifir, dolu = [], []
    for ad, (baslik, t) in tablolar.items():
        for r in t.rows:
            cs = uniq_cells(r)
            for c in cs:
                tx = c.text.strip()
                if not PARA_RE.match(tx) or c._tc in YAZILAN:
                    continue
                etiket = cs[0].text.strip() or (cs[1].text.strip() if len(cs) > 1 else '')
                (sifir if num(tx) == 0 else dolu).append(f'{baslik}: "{etiket[:60]}" = {tx}')
    for x in dolu:
        kl.uyari('Güncellenmeyen tutar', f'{x} — bu hücre beyannameden doldurulmadı, önceki aydan kalmış olabilir.', no='11')
    if sifir:
        kl.bilgi('Güncellenmeyen tutar', f'Program tarafından doldurulmayan {len(sifir)} sıfır tutarlı hücre: ' + '; '.join(sifir), no='11')
