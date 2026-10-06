"""Girdi klasörü standardı: <kök>/<yıl>/<NN AY>/ ve içinde RAPOR/ (ofisin bitmiş raporu).

    TEMİNAT ÇÖZÜMÜ/<FİRMA>/2026/02 ŞUBAT/KDV 1.pdf ...
    TEMİNAT ÇÖZÜMÜ/<FİRMA>/2026/02 ŞUBAT/RAPOR/ŞUBAT-2026 RAPOR.doc
"""
import re
from pathlib import Path

from .ortak import AYLAR, Donem, katla


class KlasorHatasi(Exception):
    pass


def ay_klasorleri(kok):
    """{Donem: klasör} — <kök>/<yyyy>/<NN ...> ya da ay adıyla başlayan klasörler."""
    kok = Path(kok)
    out = {}
    if not kok.is_dir():
        raise KlasorHatasi(f'Klasör yok: {kok}')
    for yd in sorted(kok.iterdir()):
        if not (yd.is_dir() and re.fullmatch(r'\d{4}', yd.name)):
            continue
        for ad in sorted(yd.iterdir()):
            if not ad.is_dir():
                continue
            n = katla(ad.name).upper()
            m = re.match(r'^(\d{1,2})\b', n)
            ay = int(m.group(1)) if m else None
            if ay is None:
                ay = next((i + 1 for i, a in enumerate(AYLAR) if n.startswith(katla(a).upper())), None)
            if not ay or not 1 <= ay <= 12:
                continue
            d = Donem(int(yd.name), ay)
            if d in out:
                raise KlasorHatasi(f'{d.tire} için birden fazla klasör var: {out[d].name}, {ad.name} — --girdi ile belirtin.')
            out[d] = ad
    return out


def ay_klasoru(kok, donem):
    k = ay_klasorleri(kok)
    if donem not in k:
        raise KlasorHatasi(f'{donem.tire} klasörü bulunamadı ({Path(kok) / str(donem.yil)} altında "{donem.ay:02d} {donem.ad}" bekleniyor).')
    return k[donem]


def rapor_dosyasi(ay_klasoru):
    """Ayın RAPOR/ klasöründeki ofis raporu (.docx/.doc). Birden çok aday varsa 'ORJ' içermeyen ve
    adında RAPOR geçen tercih edilir; yine de belirsizse hata verir."""
    rk = next((p for p in Path(ay_klasoru).iterdir() if p.is_dir() and katla(p.name).upper() == 'RAPOR'), None)
    if rk is None:
        return None
    c = [p for p in sorted(rk.iterdir()) if p.is_file() and p.suffix.lower() in ('.doc', '.docx') and not p.name.startswith('~')]
    if len(c) > 1:
        c2 = [p for p in c if 'ORJ' not in katla(p.name).upper()]
        c = c2 or c
    if len(c) > 1:
        c2 = [p for p in c if 'RAPOR' in katla(p.stem).upper()]
        c = c2 or c
    if len(c) > 1:
        # aynı adın .doc ve .docx hali → .docx
        stems = {p.stem for p in c}
        if len(stems) == 1:
            c = [p for p in c if p.suffix.lower() == '.docx']
    if len(c) > 1:
        raise KlasorHatasi(f'{rk} içinde birden fazla rapor var: {", ".join(p.name for p in c)} — --sablon ile belirtin.')
    return c[0] if c else None
