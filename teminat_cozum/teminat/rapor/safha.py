"""3-4-3 İndirilecek KDV safha tablosunun satırlarını hesaplar (Word'e dokunmaz; birim testi yapılabilir).

Kurallar (PROJE_TALIMATI.md bölüm 5, "Safha (T12) kuralları"):
1. Satırlar takip listesinden gelir; tutar takip listesindeki tutardır. İndirilecek listeyle firma bazında
   karşılaştırılır.
2. İnceleme şekli: ithalat (VKN 1111111111) → İTHALAT; ayardaki özel şekiller (OSB/TSE);
   YMM sütunu dolu → BİLGİ İSTEME; diğerleri → K.İ.T.
3. Yurtiçi firmalardan tutarı en büyük N tanesi EKLİ, ardından ithalat firmaları; EKLİ toplamı asgari
   oranı geçmiyorsa EKLİ sayısı artırılır.
9. KDV'si olmayan firma safha tablosuna ve karşıt inceleme tutarına alınmaz.
"""
import collections
import re
from dataclasses import dataclass, field

from ..ortak import katla, tr
from ..okuyucular.listeler import ITHALAT_VKN

ITHALAT, BILGI_ISTEME, KIT = 'İTHALAT', 'BİLGİ İSTEME', 'K.İ.T.'


def firma_anahtari(s):
    """Firma adının ilk iki kelimesi (büyük harf, Türkçe karakter duyarsız): eşleştirme anahtarı."""
    w = re.findall(r'[A-Z0-9]+', katla(s).upper())
    return ' '.join(w[:2])


def ozel_sekil(firma, kurallar):
    u = katla(firma).upper()
    for k in kurallar or []:
        if all(katla(x).upper() in u for x in k['kosul']):
            return k['sekil']
    return None


@dataclass
class SafhaSatiri:
    firma: str
    sekil: str
    kdv: float
    liste_kdv: float = None     # indirilecek listedeki toplam (eşleşme varsa)
    liste_adlari: list = field(default_factory=list)
    kaynak: str = 'takip'       # 'takip' | 'liste'


@dataclass
class SafhaSonucu:
    satirlar: list              # tüm satırlar (takip sırası + listeden eklenen OSB/TSE)
    ekli: list                  # EKLİ + İTHALAT satırları, rapordaki sırasıyla
    muhafaza: list              # "Tarafımızdan Muhafaza Edilmektedir." satırları
    ekli_adet: int
    tot_e: float
    tot_m: float
    notlar: list = field(default_factory=list)   # [(durum, konu, açıklama)]

    @property
    def tot(self):
        return self.tot_e + self.tot_m


def liste_firma_toplamlari(indirilecek):
    by = collections.OrderedDict()
    vk = {}
    for r in indirilecek:
        n = r['satici'].strip()
        by[n] = by.get(n, 0.0) + r['kdv']
        vk[n] = r['vkn']
    return by, vk


def safha_hesapla(takip, indirilecek, base, ayar, tol=0.01):
    ek = ayar.ekli_kurali
    notlar = []
    by, vk = liste_firma_toplamlari(indirilecek)
    lk = collections.defaultdict(list)
    for n in by:
        lk[firma_anahtari(n)].append(n)
    satirlar = []
    for t in takip:
        if t['kdv'] <= 0:   # ofis kuralı: KDV'si olmayan fatura karşıt incelemeye girmez
            notlar.append(('BİLGİ', 'Safha', f"{t['firma']} KDV'siz ({tr(t['kdv'])}); safha tablosuna alınmadı."))
            continue
        ln = lk.get(firma_anahtari(t['firma']), [])
        if not ln:
            ilk = (firma_anahtari(t['firma']).split() or [''])[0]
            ln = [n for n in by if (firma_anahtari(n).split() or [''])[0] == ilk and abs(by[n] - t['kdv']) < tol]
        lv = sum(by[x] for x in ln) if ln else None
        ithal = any(vk[x].startswith(ITHALAT_VKN) for x in ln)
        if ithal:
            sek = ITHALAT
        else:
            sek = ozel_sekil(t['firma'], ayar.ozel_inceleme_sekilleri) or (BILGI_ISTEME if t['ymm'] else KIT)
        if lv is None:
            notlar.append(('UYARI', 'Takip ↔ liste', f"{t['firma']} indirilecek listede bulunamadı (takipte {tr(t['kdv'])})."))
        elif abs(lv - t['kdv']) > tol:
            notlar.append(('UYARI', 'Takip ↔ liste',
                           f"{t['firma']}: takip listesinde {tr(t['kdv'])}, indirilecek listede {tr(lv)} (fark {tr(t['kdv'] - lv)})."))
        ad = ayar.firma_adi_duzeltmeleri.get(t['firma'], t['firma'])
        satirlar.append(SafhaSatiri(ad, sek, t['kdv'], lv, ln))
    # OSB / TSE satırları takipte yoksa listeden (firma ayarına bağlı)
    if ek.osb_listeden:
        for n, v in by.items():
            sek = ozel_sekil(n, ayar.ozel_inceleme_sekilleri)
            if not sek or v <= 0:
                continue
            if any(firma_anahtari(n) == firma_anahtari(r.firma) for r in satirlar):
                continue
            satirlar.append(SafhaSatiri(n.strip(), sek, v, v, [n], kaynak='liste'))
            notlar.append(('BİLGİ', 'Safha', f'{n.strip()} ({sek}) takip listesinde yok; indirilecek listeden eklendi.'))
    # takip listesinde olmayan liste firmaları (bilgi)
    takip_anahtar = {firma_anahtari(r.firma) for r in satirlar} | {firma_anahtari(t['firma']) for t in takip}
    eslesen_liste = {x for r in satirlar for x in r.liste_adlari}
    disarida = [(n, v) for n, v in by.items() if n not in eslesen_liste and firma_anahtari(n) not in takip_anahtar and v > 0]
    if disarida:
        buyuk = sorted(disarida, key=lambda x: -x[1])[:10]
        notlar.append(('BİLGİ', 'Takip ↔ liste',
                       f'İndirilecek listede olup takip listesinde olmayan {len(disarida)} firma, toplam '
                       f'{tr(sum(v for _, v in disarida))} TL. En büyükleri: ' + '; '.join(f'{n} {tr(v)}' for n, v in buyuk)))

    # listeden eklenen OSB/TSE satırları EKLİ seçimine girmez, muhafaza satırlarının sonuna yazılır
    aday = [r for r in satirlar if r.kaynak == 'takip']
    ith = [r for r in aday if r.sekil == ITHALAT] if ek.ithalat_ayri else []
    ith_ids = {id(r) for r in ith}
    srt = sorted([r for r in aday if id(r) not in ith_ids], key=lambda r: -r.kdv)
    n_e = ek.yurtici_adet
    while n_e < len(srt) and sum(r.kdv for r in srt[:n_e] + ith) < base * ek.asgari_oran:
        n_e += 1
    ekli = srt[:n_e] + sorted(ith, key=lambda r: -r.kdv)
    if n_e > ek.yurtici_adet:
        notlar.append(('UYARI', 'EKLİ seçimi',
                       f'İlk {ek.yurtici_adet} firma %{ek.asgari_oran * 100:.0f}\'i karşılamadığı için EKLİ tutanak sayısı {n_e}\'e çıkarıldı.'))
    ekli_ids = {id(r) for r in ekli}
    muh = [r for r in aday if id(r) not in ekli_ids] + [r for r in satirlar if r.kaynak != 'takip']
    tot_e = sum(r.kdv for r in ekli)
    tot_m = sum(r.kdv for r in muh)
    sonuc = SafhaSonucu(satirlar, ekli, muh, n_e, tot_e, tot_m, notlar)
    # takdir gerektiren seçim → kontrol listesine
    yedek = srt[n_e] if n_e < len(srt) else None
    notlar.append(('UYARI', 'EKLİ seçimi',
                   'EKLİ firmalar tutara göre otomatik seçildi (ofis bazen sıradaki büyük firmayı atlayıp sonrakini ekliyor) — '
                   'tutanağı hazır olanlarla karşılaştırın: ' + ', '.join(r.firma for r in ekli)
                   + (f'. Sıradaki aday: {yedek.firma} ({tr(yedek.kdv)}).' if yedek else '.')))
    muh_ymm = [r.firma for r in muh if r.sekil == BILGI_ISTEME]
    if muh_ymm:
        notlar.append(('BİLGİ', 'Muhafaza satırları',
                       'Muhafaza satırlarında YMM sütunu dolu olanlar "BİLGİ İSTEME" yazıldı (ofis bazen K.İ.T. yazıyor): '
                       + ', '.join(muh_ymm)))
    tutarlar = collections.Counter(round(r.kdv, 2) for r in satirlar)
    for v, n in tutarlar.items():
        if n > 1:
            adlar = [r.firma for r in satirlar if round(r.kdv, 2) == v]
            notlar.append(('UYARI', 'Satır kayması', f'{tr(v)} tutarı {n} firmada görünüyor: {", ".join(adlar)} — takip listesinde satır kayması olabilir.'))
    return sonuc
