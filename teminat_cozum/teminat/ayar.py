"""Firma ayar dosyası (firmalar/<firma>/ayar.yaml) okuyucusu.

Ayar dosyasında olmayan her alan için buradaki varsayılanlar kullanılır. Mükellefe ait bilgiler
(VKN, rapor referansları vb.) yalnızca ofis bilgisayarındaki ayar dosyasında tutulur; depoya konmaz.
"""
import json
from dataclasses import dataclass, field
from pathlib import Path

import yaml

from .ortak import DOLD

# Safha tablosunda özel inceleme şekilleri: sırayla denenir, ilk uyan kullanılır.
# 'kosul' içindeki tüm parçalar firma adında (büyük harf, Türkçe karakter duyarsız) geçmelidir.
VARSAYILAN_OZEL_SEKILLER = [
    {'kosul': ['STANDARD'], 'sekil': 'TSE'},
    {'kosul': ['ORGAN', 'SANAYI', 'IV. KISIM'], 'sekil': 'OSB-ELEKTRİK'},
    {'kosul': ['ORGAN', 'SANAYI', 'IKINCI'], 'sekil': 'OSB-DOĞALGAZ'},
    {'kosul': ['ORGAN', 'SANAYI', 'YALOVA'], 'sekil': 'OSB-' + DOLD},
    {'kosul': ['ORGAN', 'SANAYI'], 'sekil': 'OSB-SU'},
]


@dataclass
class EkliKurali:
    yurtici_adet: int = 10        # tutarı en büyük kaç yurtiçi firma "EKLİ" olur
    ithalat_ayri: bool = True     # ithalat firmaları EKLİ'lerden sonra "n- İTHALAT" olarak
    osb_listeden: bool = False    # OSB/TSE satırları takipte yoksa indirilecek listeden eklensin mi
    asgari_oran: float = 0.80     # EKLİ toplamı bu oranı geçmiyorsa EKLİ sayısı artırılır


@dataclass
class FirmaAyari:
    firma: str = ''
    kisa_ad: str = ''
    vkn: str = ''
    vergi_dairesi: str = ''
    ymm_no: str = ''                       # 'YMM 12345678/2026-50' içindeki 12345678
    iade_turleri: list = field(default_factory=list)
    tablolar: dict = field(default_factory=dict)          # tablo tanımı üzerine yazmaları (rapor/tablolar.py)
    ocak_raporu: dict = field(default_factory=dict)       # {yıl: {tarih: '15/04/2026', sayi: '2026-50'}}
    rapor_referanslari: dict = field(default_factory=dict)  # {'OCAK/2026': ('15.04.2026', '2026-50')}
    ekli_kurali: EkliKurali = field(default_factory=EkliKurali)
    ozel_inceleme_sekilleri: list = field(default_factory=lambda: [dict(x) for x in VARSAYILAN_OZEL_SEKILLER])
    firma_adi_duzeltmeleri: dict = field(default_factory=dict)  # takip listesindeki ad → raporda yazılacak ad
    iade_turu_etiketleri: dict = field(default_factory=dict)    # iş hacmi tablosunda yeni eklenecek tür etiketi
    tolerans: float = 0.01
    _yol: Path = None

    @property
    def rapor_oneki(self):
        return f'YMM {self.ymm_no}' if self.ymm_no else f'YMM {DOLD}'

    def ocak_referansi(self, yil):
        """Ocak raporu 'tarih ve YMM no/sayı' metni ('15/04/2026 tarih ve YMM 12345678/2026-50') ya da None."""
        o = self.ocak_raporu.get(yil) or self.ocak_raporu.get(str(yil))
        if o:
            tarih, sayi = o['tarih'], o['sayi']
        elif f'OCAK/{yil}' in self.rapor_referanslari:
            tarih, sayi = self.rapor_referanslari[f'OCAK/{yil}']
            tarih = tarih.replace('.', '/')
        else:
            return None
        return f'{tarih} tarih ve {self.rapor_oneki}/{sayi}'


def _referanslari_coz(v, kok):
    if isinstance(v, str):
        p = Path(v)
        if not p.is_absolute() and kok is not None:
            p = kok / p
        with open(p, encoding='utf-8') as f:
            v = json.load(f) if p.suffix.lower() == '.json' else yaml.safe_load(f)
    out = {}
    for k, x in (v or {}).items():
        if k.startswith('_'):
            continue
        if isinstance(x, dict):
            out[k] = (str(x['tarih']), str(x['sayi']))
        else:
            out[k] = (str(x[0]), str(x[1]))
    return out


def ayar_oku(yol=None):
    """Ayar dosyasını okur; yol None ise tüm varsayılanlarla boş ayar döner."""
    if yol is None:
        return FirmaAyari()
    yol = Path(str(yol).strip().strip('"').strip("'"))
    try:
        with open(yol, encoding='utf-8') as f:
            d = yaml.safe_load(f) or {}
    except OSError as e:
        raise ValueError(f'Ayar dosyası açılamadı: {yol} ({e.strerror})') from e
    except yaml.YAMLError as e:
        yer = getattr(e, 'problem_mark', None)
        raise ValueError(f'Ayar dosyasında yazım hatası: {yol.name}'
                         + (f', satır {yer.line + 1}' if yer else '') + ' — girintileri ve iki nokta üst üsteleri kontrol edin.') from e
    if not isinstance(d, dict):
        raise ValueError(f'{yol.name}: ayar dosyası "alan: değer" satırlarından oluşmalı.')
    a = FirmaAyari(_yol=yol)
    bilinen = {k for k in FirmaAyari.__dataclass_fields__ if not k.startswith('_')}
    bilinmeyen = set(d) - bilinen
    if bilinmeyen:
        raise ValueError(f'{yol.name}: bilinmeyen ayar(lar): {", ".join(sorted(bilinmeyen))}')
    for k, v in d.items():
        try:
            if k == 'ekli_kurali':
                a.ekli_kurali = EkliKurali(**(v or {}))
            elif k == 'rapor_referanslari':
                a.rapor_referanslari = _referanslari_coz(v, yol.parent)
            elif k in ('iade_turleri',):
                a.iade_turleri = [str(x) for x in (v or [])]
            elif k == 'ymm_no':
                a.ymm_no = str(v)
            elif k == 'vkn':
                a.vkn = str(v)
            elif v is not None:
                setattr(a, k, v)
        except ValueError as e:
            raise ValueError(f'{yol.name}, "{k}" alanı: {e}') from e
        except (TypeError, KeyError, AttributeError, OSError, yaml.YAMLError) as e:
            raise ValueError(f'{yol.name}, "{k}" alanı anlaşılamadı ({type(e).__name__}: {e}). Örnek ayar dosyasındaki '
                             'biçime bakın: firmalar/ornek/ayar.yaml') from e
    return a
