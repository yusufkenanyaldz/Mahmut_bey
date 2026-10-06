"""Duman testi verisi: gerçek PDF'li sentetik bir firma klasörü + ayar dosyası üretir.

    python tests/duman_verisi.py <hedef klasör>

<hedef>/ORNEK/2026/02 ŞUBAT/RAPOR/ŞUBAT-2026 RAPOR.docx (şablon), <hedef>/ORNEK/2026/03 MART/... (girdiler),
<hedef>/ayar.yaml. Derlenen .exe'nin uçtan uca çalıştığını (pdfplumber dahil) denemek için kullanılır.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import sentetik as S  # noqa: E402
from teminat.ortak import Donem  # noqa: E402


def uret(hedef):
    hedef = Path(hedef)
    kok = hedef / 'ORNEK'
    eski = S.Senaryo(Donem(2026, 2))
    yeni = S.Senaryo(Donem(2026, 3), devreden_onceki=eski.sonraki_devreden)
    S.girdi_klasoru(kok, eski)
    (kok / '2026' / '02 ŞUBAT' / 'RAPOR').mkdir(parents=True, exist_ok=True)
    S.sablon_docx(kok / '2026' / '02 ŞUBAT' / 'RAPOR' / 'ŞUBAT-2026 RAPOR.docx', eski)
    g = S.girdi_klasoru(kok, yeni)
    S.kdv1_pdf(g / 'KDV 1.pdf', yeni)          # düz metin yerine gerçek PDF
    (hedef / 'ayar.yaml').write_text('kisa_ad: ÖRNEK\nymm_no: "00000000"\nrapor_referanslari:\n'
                                     '  OCAK/2026: {tarih: "15.04.2026", sayi: "2026-50"}\n'
                                     '  ŞUBAT/2026: {tarih: "20.05.2026", sayi: "2026-60"}\n', encoding='utf-8')
    return kok


if __name__ == '__main__':
    print(uret(sys.argv[1]))
