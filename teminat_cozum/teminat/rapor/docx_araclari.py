"""python-docx yardımcıları: hücre/paragraf yazma, satır kopyalama/silme, satır grubu eşitleme.

Yazılan her hücre `YAZILAN` kümesine kaydedilir; böylece doldurma sonunda güncellenmemiş (önceki aydan
kalmış olabilecek) tutar hücreleri bulunabilir.
"""
import copy
import re

from docx.enum.text import WD_COLOR_INDEX
from docx.oxml.ns import qn
from docx.table import _Row

from ..ortak import katla

SARI = WD_COLOR_INDEX.YELLOW
YAZILAN = set()   # yazılan hücrelerin <w:tc> öğeleri (kayit_baslat ile sıfırlanır)


def kayit_baslat():
    YAZILAN.clear()
    return YAZILAN


def set_par(p, text, hl=None):
    runs = p.runs
    if not runs:
        r = p.add_run(text)
    else:
        runs[0].text = text
        for r in runs[1:]:
            r._r.getparent().remove(r._r)
        r = runs[0]
    if hl:
        r.font.highlight_color = hl
    return r


def paragraf_degistir(p, desen, yeni, count=0, flags=0):
    """Paragrafta regex değişimi yapar; biçimi korumak için mümkünse yalnızca ilgili run'ı değiştirir.

    Eşleşme birden çok run'a yayılıyorsa paragraf tek run'a indirilir (ilk run'ın biçimiyle).
    Değişiklik yapıldıysa True döner.
    """
    eski = p.text
    yeni_metin = re.sub(desen, yeni, eski, count=count, flags=flags)
    if yeni_metin == eski:
        return False
    runs = p.runs
    if runs and ''.join(r.text for r in runs) == eski:
        sinirlar, i = [], 0
        for r in runs:
            sinirlar.append((i, i + len(r.text)))
            i += len(r.text)
        eslesmeler = list(re.finditer(desen, eski, flags))
        if count:
            eslesmeler = eslesmeler[:count]
        hedefler = []
        for m in eslesmeler:
            ri = [j for j, (a, b) in enumerate(sinirlar) if a <= m.start() and m.end() <= b and (m.end() > m.start() or a < b)]
            if not ri:
                break
            hedefler.append((ri[0], m))
        else:
            for j, m in reversed(hedefler):
                a, _ = sinirlar[j]
                t = runs[j].text
                runs[j].text = t[:m.start() - a] + m.expand(yeni) + t[m.end() - a:]
            if p.text == yeni_metin:
                return True
    set_par(p, yeni_metin)
    return True


def metin_degistir(p, eski, yeni):
    """Düz metin değişimi (regex değil)."""
    return paragraf_degistir(p, re.escape(eski), yeni.replace('\\', r'\\'))


def isaretle(p, renk=SARI):
    for r in p.runs:
        r.font.highlight_color = renk


def set_cell(cell, text, hl=None):
    ps = cell.paragraphs
    set_par(ps[0], text, hl)
    for p in ps[1:]:
        p._p.getparent().remove(p._p)
    YAZILAN.add(cell._tc)


def hucre_isaretle(cell, renk=SARI):
    for p in cell.paragraphs:
        isaretle(p, renk)


def uniq_cells(row):
    out, seen = [], set()
    for c in row.cells:
        if id(c._tc) in seen:
            continue
        seen.add(id(c._tc))
        out.append(c)
    return out


def clone_row_after(row, ref_row=None):
    new = copy.deepcopy((ref_row or row)._tr)
    row._tr.addnext(new)
    return _Row(new, row._parent)


def del_row(row):
    row._tr.getparent().remove(row._tr)


def row_text(row):
    return ' | '.join(c.text.strip() for c in uniq_cells(row))


def ara(pat, metin):
    """Türkçe karakter ve büyük/küçük harf duyarsız regex araması."""
    return re.search(katla(pat), katla(metin), re.I)


def find_row(table, pat):
    for r in table.rows:
        if ara(pat, row_text(r)):
            return r


def find_rows(table, pat):
    return [r for r in table.rows if ara(pat, row_text(r))]


def set_row_vals(row, vals, start=None):
    """vals: satırın SON len(vals) benzersiz hücresine (ya da start'tan itibaren) yazılır."""
    cs = uniq_cells(row)
    if start is None:
        if len(cs) <= len(vals):
            raise SablonYok(f'satırda {len(cs)} hücre var, etiket + {len(vals)} değer bekleniyordu: "{row_text(row)[:80]}"')
        start = len(cs) - len(vals)
    for c, v in zip(cs[start:], vals):
        if v is not None:
            set_cell(c, v)


def ilk_hucre(row):
    return uniq_cells(row)[0]


class SablonYok(Exception):
    """Gruba satır eklenmesi gerekiyor ama kopyalanacak örnek satır yok."""


def grubu_esitle(mevcut, istenen, kalem_anahtari, satir_anahtari, *, sonra=None, sablon=None):
    """Bir tablodaki ardışık satır grubunu istenen kalemlere göre yeniden kurar.

    mevcut         : gruptaki mevcut satırlar (_Row), tablodaki sırasıyla
    istenen        : kalemler listesi (sıra bu listeye göre olur)
    kalem_anahtari : kalem → anahtar;  satir_anahtari : mevcut satır → anahtar (aynı uzayda)
    sonra          : grup boşsa yeni satırların ekleneceği yerden önceki satır (_Row)
    sablon         : yeni satır için kopyalanacak satır (_Row); verilmezse mevcut[0]

    Aynı anahtarlı mevcut satır yeniden kullanılır (ofisin satır biçimi ve etiketi korunur), olmayanlar
    şablondan kopyalanır, istenmeyenler silinir. Dönüş: ([(row, kalem, yeni_mi)] istenen sırasıyla,
    [silinen satırların ilk hücre metni]).
    """
    ornek = sablon if sablon is not None else (mevcut[0] if mevcut else None)
    sablon_tr = copy.deepcopy(ornek._tr) if ornek is not None else None
    if mevcut:
        onceki = mevcut[0]._tr.getprevious()
        tablo = mevcut[0]._parent
    elif sonra is not None:
        onceki = sonra._tr
        tablo = sonra._parent
    else:
        raise SablonYok('grubun yeri belirsiz')
    havuz = list(mevcut)
    secilen = []
    for k in istenen:
        a = kalem_anahtari(k)
        eslesen = next((r for r in havuz if satir_anahtari(r) == a), None)
        if eslesen is not None:
            havuz.remove(eslesen)
        secilen.append((eslesen, k))
    if sablon_tr is None and any(r is None for r, _ in secilen):
        raise SablonYok('şablonda kopyalanacak örnek satır yok')
    silinen = [r.cells[0].text.strip() for r in havuz]
    for r in havuz:
        del_row(r)
    out = []
    for r, k in secilen:
        yeni = r is None
        tr_el = copy.deepcopy(sablon_tr) if yeni else r._tr
        onceki.addnext(tr_el)
        onceki = tr_el
        out.append((_Row(tr_el, tablo), k, yeni))
    return out, silinen


def vmerge_devam_yap(tr_el, sutun):
    """Satırın `sutun`. hücresini (tc sırası) dikey birleşimin devamı yapar ve metnini siler."""
    tcs = tr_el.findall(qn('w:tc'))
    tc = tcs[sutun]
    pr = tc.get_or_add_tcPr()
    vm = pr.find(qn('w:vMerge'))
    if vm is None:
        vm = pr.makeelement(qn('w:vMerge'), {})
        pr.append(vm)
    if qn('w:val') in vm.attrib:
        del vm.attrib[qn('w:val')]
    for t in tc.iter(qn('w:t')):
        t.text = ''


def vmerge_baslangic_sutunu(tr_el):
    """Satırda dikey birleşimi başlatan (vMerge=restart) hücrenin tc sırası; yoksa None."""
    for i, tc in enumerate(tr_el.findall(qn('w:tc'))):
        pr = tc.tcPr
        vm = pr.find(qn('w:vMerge')) if pr is not None else None
        if vm is not None and vm.get(qn('w:val')) == 'restart':
            if i > 0:
                return i
    return None


def tum_paragraflar(doc):
    """Gövde, tablo hücreleri, üst/alt bilgi paragrafları: (yer, paragraf)."""
    for p in doc.paragraphs:
        yield 'metin', p
    for ti, t in enumerate(doc.tables):
        for c in _tablo_hucreleri(t):
            for p in c.paragraphs:
                yield f'tablo{ti}', p
    for s in doc.sections:
        for hf, ad in ((s.header, 'üst bilgi'), (s.footer, 'alt bilgi')):
            if hf.is_linked_to_previous:
                continue
            for p in hf.paragraphs:
                yield ad, p
            for t in hf.tables:
                for c in _tablo_hucreleri(t):
                    for p in c.paragraphs:
                        yield ad, p


def _tablo_hucreleri(t):
    seen = set()   # <w:tc> öğelerinin kendisi tutulur (id() yeniden kullanılabilir)
    for r in t.rows:
        for c in r.cells:
            if c._tc in seen:
                continue
            seen.add(c._tc)
            yield c
