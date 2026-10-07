"""Klasör adından ay/yıl okuma — YALNIZCA İPUCU olarak kullanılır (PROJE_TALIMATI.md §13).

Hangi dosyanın ne olduğu ve hangi dönemin raporu olduğu her zaman İÇERİKTEN anlaşılır (teminat/tanima.py). Klasör
adındaki ay yalnızca (a) belgedeki dönemle karşılaştırılıp uyuşmazlık uyarısı vermek ve (b) önceki ayın raporunu
ararken hangi klasörlere önce bakılacağını sıralamak için kullanılır; hiçbir dosya klasör adına göre seçilmez.
"""
import re

from .ortak import AYLAR, katla

_AY_ADLARI = [(i + 1, katla(a).upper()) for i, a in enumerate(AYLAR)]


def yil_coz(ad):
    """Addaki tek yıl (20xx) ya da None."""
    y = set(re.findall(r'(?<!\d)(20\d\d)(?!\d)', katla(ad)))
    return int(y.pop()) if len(y) == 1 else None


def ay_coz(ad):
    """Klasör adından (ay, yıl|None) ipucu; ay görünmüyorsa None.

    Önce ay adı ("02 ŞUBAT", "ABC 12- ARALIK", "Şubat 2026 (20.03.2026 verildi)"), sonra "2026-02" / "02.2026" /
    "2026 02" biçimi, sonra baştaki ay numarası ("02", "3- GİRDİ") denenir.
    """
    n = katla(ad).upper()
    adlar = {no for no, a in _AY_ADLARI if re.search(rf'(?<![A-Z]){a}(?![A-Z])', n)}
    if len(adlar) > 1:
        return None                    # "OCAK-ŞUBAT" gibi: belirsiz
    if adlar:
        return adlar.pop(), yil_coz(n)
    m = re.search(r'(?<!\d)(20\d\d)\s*[-_./ ]\s*(\d{1,2})(?!\d)', n)
    if m and 1 <= int(m.group(2)) <= 12:
        return int(m.group(2)), int(m.group(1))
    m = re.search(r'(?<!\d)(\d{1,2})\s*[-_./ ]\s*(20\d\d)(?!\d)', n)
    if m and 1 <= int(m.group(1)) <= 12:
        return int(m.group(1)), int(m.group(2))
    m = re.match(r'^\s*(\d{1,2})(?!\d)', n)
    if m and 1 <= int(m.group(1)) <= 12:
        return int(m.group(1)), yil_coz(n)
    return None
