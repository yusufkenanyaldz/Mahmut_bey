"""Girdilere dayalı sayısal kontroller (PROJE_TALIMATI.md bölüm 6). Word belgesine dokunmaz."""
import collections

from ..okuyucular.listeler import donem_degeri
from ..ortak import AYLAR, tr


def liste_toplami(kl, liste_toplam, base, tol=1.0):
    """1. İndirilecek KDV listesi toplamı ↔ beyandaki yurtiçi + sorumlu + ithal."""
    fark = liste_toplam - base
    kl.sonuc(abs(fark) <= tol, 'İndirilecek liste ↔ beyan',
             f'Liste toplamı beyanla tutuyor ({tr(base)} TL).',
             f'İndirilecek KDV listesi toplamı {tr(liste_toplam)} TL, beyandaki yurtiçi+sorumlu+ithal {tr(base)} TL — '
             f'fark {tr(fark)} TL. Karşıt inceleme tabanı olarak beyandaki tutar kullanıldı.',
             durum='UYARI', no='1', beklenen=tr(base), bulunan=tr(liste_toplam))


def liste_donemi(kl, donemler: collections.Counter, donem):
    """2. Liste dosyasının dönemi (sütun 15) ↔ rapor dönemi."""
    if not donemler:
        kl.uyari('Liste dönemi', 'İndirilecek KDV listesinde dönem sütunu (15) boş; listenin bu aya ait olduğunu elle kontrol edin.', no='2')
        return
    okunan = collections.Counter()
    okunamayan = collections.Counter()
    for v, n in donemler.items():
        d = donem_degeri(v)
        (okunan if d else okunamayan)[d or v] += n
    yanlis = {d: n for d, n in okunan.items() if d != (donem.yil, donem.ay)}
    if okunamayan:
        kl.bilgi('Liste dönemi', f'Dönem sütununda anlaşılamayan değerler: {dict(okunamayan)}', no='2')
    if yanlis:
        ac = ', '.join(f'{AYLAR[a - 1]}/{y}: {n} satır' for (y, a), n in sorted(yanlis.items()))
        kl.hata('Liste dönemi', f'İndirilecek KDV listesinde rapor dönemi ({donem.egik}) dışında satırlar var — {ac}. '
                'Klasörde önceki ayın listesi kalmış olabilir.', no='2')
    elif okunan:
        kl.tamam('Liste dönemi', f'Listedeki tüm satırlar {donem.egik} dönemine ait.', no='2')


def iade_turleri(kl, kalemler, iade_gereken, tol=0.01):
    """6. İade türleri toplamı ↔ 'İade edilmesi gereken KDV'."""
    top = sum(v for _, v in kalemler.values() if v > 0)
    kl.sonuc(abs(top - iade_gereken) <= tol, 'İade türleri toplamı',
             f'İade türleri toplamı iade edilmesi gereken KDV ile tutuyor ({tr(iade_gereken)}).',
             f'İade türleri toplamı {tr(top)} ≠ beyandaki iade edilmesi gereken KDV {tr(iade_gereken)} '
             f'(fark {tr(top - iade_gereken)}).', no='6', beklenen=tr(iade_gereken), bulunan=tr(top))


def teminat(kl, tem, kalemler, iade_gereken, tol=0.01):
    """7. Teminat dilekçesinde türler toplamı ↔ toplam; teminat ≤ iade."""
    if not tem:
        kl.hata('Teminat dilekçesi', 'Teminat mektubu kabul dilekçesi bulunamadı veya okunamadı; teminat tarihi, numarası '
                've tutarları (Raporun amacı, 4-5, 4-6) önceki aydan kalmış olabilir — elle girilmeli.', no='7')
        return
    eksik = [a for a in ('tarih', 'no', 'toplam') if a not in tem]
    if eksik:
        kl.hata('Teminat dilekçesi', f'Dilekçeden okunamayan bilgi: {", ".join(eksik)}.', no='7')
    turler = {k: v for k, v in tem.items() if k in ('301', '410', '339', '318', '448')}
    if 'toplam' in tem:
        top = sum(turler.values())
        kl.sonuc(abs(top - tem['toplam']) <= tol, 'Teminat türleri toplamı',
                 f'Dilekçede türler toplamı toplam teminatla tutuyor ({tr(tem["toplam"])}).',
                 f'Dilekçede türler toplamı {tr(top)} ≠ toplam {tr(tem["toplam"])}.', no='7',
                 beklenen=tr(tem['toplam']), bulunan=tr(top))
        kl.sonuc(tem['toplam'] <= iade_gereken + tol, 'Teminat ≤ iade',
                 f'Teminat toplamı ({tr(tem["toplam"])}) iade edilmesi gereken KDV\'yi ({tr(iade_gereken)}) aşmıyor.',
                 f'Teminat toplamı {tr(tem["toplam"])} > iade edilmesi gereken KDV {tr(iade_gereken)}.', no='7')
    for kod, v in turler.items():
        iade = kalemler.get(kod, (0, 0))[1]
        if v > iade + tol:
            kl.hata('Teminat ≤ iade', f'{kod} teminatı {tr(v)} > beyandaki {kod} iadesi {tr(iade)}.', no='7')


def tevkifat_410(kl, matrah, oran, tevkifat, iade, tol=1.0):
    """8. 410: matrah × %20 × 4/10 ↔ beyandaki iade."""
    pay, payda = (int(x) for x in str(tevkifat).split('/'))
    hesap = matrah * oran / 100 * pay / payda
    kl.sonuc(abs(hesap - iade) <= tol, '410 tevkifat hesabı',
             f'410: {tr(matrah)} × %{oran} × {tevkifat} = {tr(hesap)} beyanla tutuyor.',
             f'410: {tr(matrah)} × %{oran} × {tevkifat} = {tr(hesap)}, beyandaki iade {tr(iade)} (fark {tr(hesap - iade)}).',
             durum='UYARI', no='8', beklenen=tr(iade), bulunan=tr(hesap))


def yuklenilen(kl, ytot, hedef, tol=0.01):
    """9. Yüklenilen tutanak toplamı ↔ 301 yüklenilen (+339)."""
    kl.sonuc(abs(ytot - hedef) <= tol, 'Yüklenilen tutanak toplamı',
             f'Yüklenilen tutanak toplamı 301 (+339) yüklenilen KDV ile tutuyor ({tr(hedef)}).',
             f'Yüklenilen tutanak toplamı {tr(ytot)} ≠ 301 (+339) yüklenilen {tr(hedef)} (fark {tr(ytot - hedef)}).',
             no='9', beklenen=tr(hedef), bulunan=tr(ytot))


def devreden(kl, onceki_sonraki, bu_onceki, tol=0.01):
    """10. Önceki ay 'sonraki döneme devreden' ↔ bu ay 'önceki dönemden devreden'."""
    if onceki_sonraki is None:
        kl.bilgi('Devreden KDV', 'Şablonda (önceki ay raporu) "Sonraki döneme devreden" tutarı okunamadı; kontrol yapılamadı.', no='10')
        return
    kl.sonuc(abs(onceki_sonraki - bu_onceki) <= tol, 'Devreden KDV',
             f'Önceki ay sonraki döneme devreden = bu ay önceki dönemden devreden ({tr(bu_onceki)}).',
             f'Önceki ay raporunda sonraki döneme devreden {tr(onceki_sonraki)}, bu ay beyanda önceki dönemden '
             f'devreden {tr(bu_onceki)} (fark {tr(bu_onceki - onceki_sonraki)}). Düzeltme beyannamesi olabilir.',
             no='10', beklenen=tr(onceki_sonraki), bulunan=tr(bu_onceki))


def beyan_tutarliligi(kl, k, tol=0.05):
    """KDV 1 okumasının iç tutarlılığı: okunan satırlar toplamı ↔ beyandaki toplamlar.

    PDF'ten bir satır okunamazsa (desen tutmazsa) tabloda eksik kalır; bu kontrol bunu yakalar.
    """
    satirlar = []
    if k.get('r701'):
        satirlar.append((k['r701'][0], k['r701'][2]))
    satirlar += [(a, c) for a, _, c in k.get('yurtici', [])]
    for x in ('k410', 'k448'):
        if k.get(x):
            satirlar.append((k[x][0], k[x][3]))
    for x in ('d_iade', 'd_kur', 'd_diger', 'd_amort'):
        if k.get(x):
            satirlar.append((k[x][0], k[x][1]))
    m = sum(a for a, _ in satirlar)
    v = sum(b for _, b in satirlar)
    for ad, hesap, beyan in (('Matrah toplamı', m, k.get('matrah_toplam')), ('Hesaplanan KDV', v, k.get('hesaplanan'))):
        if beyan is None:
            kl.uyari('KDV 1 okuma', f'{ad} beyannameden okunamadı; satırlar kontrol edilemedi.')
            continue
        kl.sonuc(abs(hesap - beyan) <= tol, 'KDV 1 okuma',
                 f'{ad}: okunan işlem satırları toplamı beyanla tutuyor ({tr(beyan)}).',
                 f'{ad}: okunan işlem satırları toplamı {tr(hesap)}, beyanda {tr(beyan)} (fark {tr(beyan - hesap)}) — '
                 'beyannamedeki bazı satırlar okunamamış olabilir; 3-1 ve 3.8 tablolarını beyannameyle karşılaştırın.',
                 beklenen=tr(beyan), bulunan=tr(hesap))
    if k.get('oranlar'):
        o = sum(x[2] for x in k['oranlar'])
        kl.sonuc(abs(o - k.get('yurtici_alim', 0)) <= tol, 'KDV 1 okuma',
                 f'Oran dağılımı toplamı yurtiçi alımlar KDV\'si ile tutuyor ({tr(o)}).',
                 f'Oran dağılımı KDV toplamı {tr(o)}, yurtiçi alımlar KDV\'si {tr(k.get("yurtici_alim", 0))} — oran satırlarını kontrol edin.',
                 durum='UYARI', beklenen=tr(k.get('yurtici_alim', 0)), bulunan=tr(o))
    ind = sum(k.get(x, 0) for x in ('devreden_onceki', 'yurtici_alim', 'sorumlu', 'ithal', 'satis_iade'))
    if k.get('indirim_toplam'):
        kl.sonuc(abs(ind - k['indirim_toplam']) <= tol, 'KDV 1 okuma',
                 f'İndirimler toplamı okunan kalemlerle tutuyor ({tr(ind)}).',
                 f'Devreden + yurtiçi + sorumlu + ithal + satıştan iade = {tr(ind)}, beyandaki indirimler toplamı '
                 f'{tr(k["indirim_toplam"])} — okunmayan bir indirim kalemi olabilir.',
                 durum='UYARI', beklenen=tr(k['indirim_toplam']), bulunan=tr(ind))
