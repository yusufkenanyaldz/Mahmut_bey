import sys
from pathlib import Path

import pytest

KOK = Path(__file__).resolve().parents[1]
for p in (KOK, KOK / 'tests'):
    if str(p) not in sys.path:
        sys.path.insert(0, str(p))

from teminat.ayar import FirmaAyari  # noqa: E402
from teminat.okuyucular import kdv1  # noqa: E402
from teminat.ortak import Donem  # noqa: E402

import sentetik as S  # noqa: E402


GERCEK_PDF_METNI = kdv1.pdf_metni


GERCEK_PDF_ILK_SAYFA = kdv1.pdf_ilk_sayfa


def _metin_ya_da_pdf(p, gercek):
    """Sentetik '.pdf' dosyaları çoğunlukla pdfplumber metnini düz metin olarak içerir; gerçek PDF ise gerçekten okunur."""
    with open(p, 'rb') as f:
        if f.read(5) == b'%PDF-':
            return gercek(p)
    return Path(p).read_text(encoding='utf-8')


@pytest.fixture(autouse=True)
def sahte_pdf(monkeypatch, tmp_path_factory):
    monkeypatch.setattr(kdv1, 'pdf_metni', lambda p: _metin_ya_da_pdf(p, GERCEK_PDF_METNI))
    monkeypatch.setattr(kdv1, 'pdf_ilk_sayfa', lambda p: _metin_ya_da_pdf(p, GERCEK_PDF_ILK_SAYFA))
    # tanıma önbelleği ve hatırlanan ayarlar kullanıcının ev klasörüne değil geçici klasöre yazılsın
    ev = tmp_path_factory.mktemp('ev')
    monkeypatch.setenv('HOME', str(ev))
    monkeypatch.setenv('USERPROFILE', str(ev))


@pytest.fixture
def ayar():
    return FirmaAyari(ymm_no='00000000', kisa_ad='ÖRNEK',
                      rapor_referanslari={'OCAK/2026': ('15.04.2026', '2026-50'), 'ŞUBAT/2026': ('20.05.2026', '2026-60')})


@pytest.fixture
def ay(tmp_path):
    """ŞUBAT-2026 şablonu + MART-2026 girdileri."""
    eski = S.Senaryo(Donem(2026, 2))
    yeni = S.Senaryo(Donem(2026, 3), devreden_onceki=eski.sonraki_devreden)
    sablon = S.sablon_docx(tmp_path / 'sablon.docx', eski)
    girdi = S.girdi_klasoru(tmp_path / 'ORNEK', yeni)
    return dict(eski=eski, yeni=yeni, sablon=sablon, girdi=girdi, kok=tmp_path)
