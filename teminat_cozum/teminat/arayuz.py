"""Masaüstü penceresi (tkinter). Kullanıcı yalnızca "bu firmanın bu ayki klasörü"nü seçer; program içindeki belgeleri
İÇERİKTEN tanır, önce tanıma tablosunu gösterir, sonra taslağı üretir (PROJE_TALIMATI.md §13).

- Türü yanlış tanınan dosyaya tablo üzerinde çift tıklanarak türü düzeltilir; düzeltme firma ayar dosyasının yanındaki
  belge_turleri.yaml'a yazılır ve sonraki aylarda hatırlanır.
- İçerikten karar verilemeyen durumlarda (ör. aynı dönemin iki farklı beyannamesi) program adayları listeleyip sorar.
"""
import json
import os
import queue
import subprocess
import sys
import threading
import traceback
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from . import __version__, cli
from .donusum import soffice_yolu
from .islem import duzeltme_dosyasi, hazirla, rol_adi, yol_temizle
from .tanima import ACIKLAMA, TUR_LISTESI, BelirsizSecim, TurDuzeltmeleri

OTOMATIK = 'OTOMATIK'

HATIRLA = Path.home() / '.teminat_cozum.json'
KULLANIM = {'KDV1': '✔ beyanname', 'LISTE_INDIRILECEK': '✔ indirilecek liste', 'TAKIP': '✔ takip listesi',
            'TEMINAT_DILEKCE': '✔ teminat dilekçesi', 'LISTE_YUKLENILEN': '✔ yüklenilen', 'BU_AYIN_RAPORU': '✔ karşılaştırma'}


def program_klasoru():
    """.exe'nin (ya da kaynak kodun) bulunduğu klasör; firmalar/ burada aranır."""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parents[1]


class _KuyrukYazici:
    def __init__(self, q, kayit=None):
        self.q = q
        self.kayit = kayit if kayit is not None else []

    def write(self, s):
        self.q.put(s)
        self.kayit.append(s)
        return len(s)

    def flush(self):
        pass

    def reconfigure(self, **_):
        pass


def klasoru_ac(yol):
    yol = str(yol)
    if sys.platform.startswith('win'):
        os.startfile(yol)  # noqa: S606 - kullanıcının kendi klasörü
    elif sys.platform == 'darwin':
        subprocess.Popen(['open', yol])
    else:
        subprocess.Popen(['xdg-open', yol])


class Uygulama(tk.Tk):
    def __init__(self, klasor=None):
        super().__init__()
        self.title(f'Teminat Çözüm Raporu — Taslak Programı {__version__}')
        self.geometry('1100x720')
        self.minsize(820, 520)
        self.q = queue.Queue()
        self.calisiyor = False
        self.v_ayar, self.v_klasor, self.v_sablon = tk.StringVar(), tk.StringVar(), tk.StringVar()
        self.hz = None              # son tanıma (islem.Hazirlik)
        self.zorla = {}             # kullanıcının belirsiz seçimlerde seçtikleri {rol: yol}
        self.son_cikti = None
        self._sablon_klasoru = None  # "Önceki ayın raporu" hangi klasör için seçildi
        self._hatirla_yukle()
        if klasor:
            self.v_klasor.set(str(yol_temizle(klasor)))
        self._kur()
        self.v_klasor.trace_add('write', lambda *_: self._klasor_degisti())
        self.v_sablon.trace_add('write', lambda *_: self._sablon_degisti())
        self.v_ayar.trace_add('write', lambda *_: self._sifirla(zorla=False))
        self.after(100, self._kuyrugu_bosalt)
        self._yaz(
            'Nasıl çalışır:\n'
            '1) Ayar dosyasını seçin (firmalar\\<firma>\\ayar.yaml).\n'
            '2) Bu ayın klasörünü seçin — firmanın o ayki klasörü. Dosya adları ve klasör düzeni önemli değildir:\n'
            '   program her dosyayı açıp ne olduğunu içeriğinden anlar.\n'
            '3) "1) Belgeleri tanı": tanıma tablosu gösterilir (hangi dosya ne, eksik zorunlu belge var mı, önceki ayın raporu\n'
            '   bulundu mu). Türü yanlış olan satıra çift tıklayıp düzeltebilirsiniz.\n'
            '4) "2) Taslak oluştur": taslak ve kontrol listesi o klasördeki "CLAUDE TASLAK" klasörüne yazılır;\n'
            '   orijinal dosyalara dokunulmaz.\n')
        if not soffice_yolu():
            self._yaz('\nNot: LibreOffice bulunamadı. Eski Word (.doc) dosyaları okunamaz; LibreOffice kurun.\n')

    # ---------------------------------------------------------------- yerleşim
    def _kur(self):
        ust = ttk.Frame(self, padding=10)
        ust.pack(fill='x')
        ust.columnconfigure(1, weight=1)
        for i, (etiket, deg, komut) in enumerate([
                ('Ayar dosyası (ayar.yaml):', self.v_ayar, self._sec_ayar),
                ('Bu ayın klasörü:', self.v_klasor, self._sec_klasor),
                ('Önceki ayın raporu (isteğe bağlı):', self.v_sablon, self._sec_sablon)]):
            ttk.Label(ust, text=etiket).grid(row=i, column=0, sticky='w', pady=3)
            ttk.Entry(ust, textvariable=deg).grid(row=i, column=1, sticky='ew', padx=6)
            ttk.Button(ust, text='Seç…', command=komut).grid(row=i, column=2)

        dugmeler = ttk.Frame(self, padding=(10, 0))
        dugmeler.pack(fill='x')
        self.dugmeler = []          # iş sürerken kapatılanlar (alanlar dahil)
        for w in ust.winfo_children():
            if isinstance(w, (ttk.Entry, ttk.Button)):
                self.dugmeler.append(w)
        for metin, komut in [('1) Belgeleri tanı', self.tani), ('2) Taslak oluştur', self.taslak),
                             ('Çıktı klasörünü aç', self.cikti_ac), ('Gerçek raporla karşılaştır…', self.karsilastir),
                             ('Rapor denetle…', self.denetle), ('Geriye dönük test…', self.geriye_donuk)]:
            b = ttk.Button(dugmeler, text=metin, command=komut)
            b.pack(side='left', padx=(0, 6), pady=6)
            self.dugmeler.append(b)

        self.defter = ttk.Notebook(self, padding=(10, 0, 10, 4))
        self.defter.pack(fill='both', expand=True)
        sekme1 = ttk.Frame(self.defter)
        sekme2 = ttk.Frame(self.defter)
        self.defter.add(sekme1, text='Tanıma tablosu')
        self.defter.add(sekme2, text='Çıktı')
        ttk.Label(sekme1, text='Türü yanlış görünen satıra çift tıklayıp düzeltin (düzeltme sonraki aylarda da hatırlanır).',
                  foreground='#555').pack(anchor='w', pady=(4, 2))
        sutunlar = ('tur', 'aciklama', 'donem', 'kullanim', 'not')
        self.agac = ttk.Treeview(sekme1, columns=sutunlar, show='tree headings')
        self.agac.heading('#0', text='Dosya')
        for s, baslik, gen in zip(sutunlar, ('Tür', 'Açıklama', 'Dönem', 'Kullanım', 'Not'), (150, 230, 105, 150, 260)):
            self.agac.heading(s, text=baslik)
            self.agac.column(s, width=gen, stretch=s in ('aciklama', 'not'))
        self.agac.column('#0', width=330, stretch=True)
        kay = ttk.Scrollbar(sekme1, command=self.agac.yview)
        self.agac.configure(yscrollcommand=kay.set)
        kay.pack(side='right', fill='y')
        self.agac.pack(fill='both', expand=True)
        self.agac.bind('<Double-1>', self._tur_duzelt)
        self.agac.tag_configure('kullanilan', background='#e6f4ea')
        self.agac.tag_configure('sorun', foreground='#9a3412')
        self.agac.tag_configure('uyari', background='#fff3cd', foreground='#664d03')
        self.agac.tag_configure('hata', background='#f8d7da', foreground='#842029')
        self.metin = tk.Text(sekme2, wrap='word', font=('Consolas', 10) if sys.platform.startswith('win') else None)
        kay2 = ttk.Scrollbar(sekme2, command=self.metin.yview)
        self.metin.configure(yscrollcommand=kay2.set, state='disabled')
        kay2.pack(side='right', fill='y')
        self.metin.pack(side='left', fill='both', expand=True)
        self.durum = ttk.Label(self, text='Hazır.', anchor='w', padding=(10, 0, 10, 6))
        self.durum.pack(fill='x')
        self.satir_yolu = {}

    # ---------------------------------------------------------------- yardımcılar
    def _hatirla_yukle(self):
        try:
            d = json.loads(HATIRLA.read_text(encoding='utf-8'))
        except (OSError, ValueError):
            d = {}
        self.v_ayar.set(d.get('ayar', ''))
        self.v_klasor.set(d.get('klasor', ''))
        if not self.v_ayar.get():
            firmalar = program_klasoru() / 'firmalar'
            aday = sorted(p for p in firmalar.glob('*/ayar.yaml') if p.parent.name != 'ornek') if firmalar.is_dir() else []
            if len(aday) == 1:
                self.v_ayar.set(str(aday[0]))

    def _hatirla_kaydet(self):
        try:
            HATIRLA.write_text(json.dumps({'ayar': self.v_ayar.get(), 'klasor': self.v_klasor.get()}, ensure_ascii=False),
                               encoding='utf-8')
        except OSError:
            pass

    def _sifirla(self, zorla=True):
        """Alanlar değişince eski tanıma geçersizdir: tablo temizlenir, taslak yeniden tanıma yapar."""
        self.hz = None
        if zorla:
            self.zorla = {}
            self.son_cikti = None
        if hasattr(self, 'agac'):
            self.agac.delete(*self.agac.get_children())
            self.satir_yolu = {}
            if not self.calisiyor:
                self.durum.configure(text='Hazır.')

    def _klasor_degisti(self):
        # başka bir klasör için seçilmiş "önceki ayın raporu" yeni klasöre taşınmasın
        if self.v_sablon.get() and self._sablon_klasoru != self.v_klasor.get():
            self.v_sablon.set('')
        self._sifirla()

    def _sablon_degisti(self):
        self._sablon_klasoru = self.v_klasor.get()
        self._sifirla(zorla=False)

    def _yaz(self, s):
        self.metin.configure(state='normal')
        self.metin.insert('end', s)
        self.metin.see('end')
        self.metin.configure(state='disabled')

    def _kuyrugu_bosalt(self):
        try:
            while True:
                self._yaz(self.q.get_nowait())
        except queue.Empty:
            pass
        self.after(100, self._kuyrugu_bosalt)

    def _sec_ayar(self):
        bas = Path(self.v_ayar.get()).parent if self.v_ayar.get() else program_klasoru() / 'firmalar'
        p = filedialog.askopenfilename(title='Firma ayar dosyası', initialdir=str(bas) if bas.is_dir() else None,
                                       filetypes=[('Ayar dosyası', '*.yaml *.yml'), ('Tüm dosyalar', '*.*')])
        if p:
            self.v_ayar.set(p)

    def _sec_klasor(self):
        p = filedialog.askdirectory(title='Firmanın bu ayki klasörü', initialdir=self.v_klasor.get() or None)
        if p:
            self.v_klasor.set(p)

    def _sec_sablon(self):
        p = filedialog.askopenfilename(title='Önceki ayın bitmiş raporu (şablon)', filetypes=[('Word', '*.doc *.docx')],
                                       initialdir=self.v_klasor.get() or None)
        if p:
            self.v_sablon.set(p)

    def _ayar(self):
        from .ayar import ayar_oku
        yol = self.v_ayar.get().strip()
        return ayar_oku(yol or None)

    def _ayar_ya_da_hata(self):
        try:
            return True, self._ayar()
        except Exception as e:  # noqa: BLE001 - .exe'de konsol yok: her hata pencerede gösterilmeli
            messagebox.showerror('Ayar dosyası', str(e) or type(e).__name__)
            return False, None

    def _anlik(self):
        return (self.v_klasor.get(), self.v_ayar.get(), self.v_sablon.get(), tuple(sorted((k, str(v)) for k, v in self.zorla.items())))

    def _kontrol(self):
        if self.calisiyor:
            return False
        k = yol_temizle(self.v_klasor.get())
        if k is None:
            messagebox.showerror('Klasör', 'Firmanın bu ayki klasörünü seçin.')
            return False
        if not k.is_dir():
            messagebox.showerror('Klasör', f'Klasör bulunamadı:\n{k}')
            return False
        s = yol_temizle(self.v_sablon.get())
        if s is not None and not s.is_file():
            messagebox.showerror('Önceki ayın raporu', f'Dosya bulunamadı:\n{s}')
            return False
        return True

    def _mesgul(self, evet, metin=''):
        self.calisiyor = evet
        for b in self.dugmeler:
            b.state(['disabled'] if evet else ['!disabled'])
        if metin:
            self.durum.configure(text=metin)

    def _arka_planda(self, baslik, is_, bitince, tekrar=None):
        """is_() bir iş parçacığında çalışır (print'leri Çıktı sekmesine gider); sonucu ana iş parçacığında bitince(sonuc, hata)."""
        self._mesgul(True, f'{baslik}…')
        self._yaz(f'\n===== {baslik} =====\n')
        anlik = self._anlik()

        def calis():
            yazici = _KuyrukYazici(self.q)
            sonuc = hata = None
            try:
                with redirect_stdout(yazici), redirect_stderr(yazici):
                    sonuc = is_()
            except BaseException as e:  # noqa: BLE001 - hata pencerede gösterilir
                hata = e
                if not isinstance(e, (BelirsizSecim, ValueError, SystemExit)) and not type(e).__name__.endswith('Hatasi'):
                    self.q.put(traceback.format_exc())
            self.after(150, self._bitir, baslik, bitince, sonuc, hata, tekrar, anlik)

        threading.Thread(target=calis, daemon=True).start()

    def _bitir(self, baslik, bitince, sonuc, hata, tekrar=None, anlik=None):
        self._kuyrugu_bosalt_hemen()
        self._mesgul(False)
        if anlik is not None and anlik[:3] != self._anlik()[:3]:
            self.hz = None
            self._yaz('\nİş sürerken alanlar değişti; sonuç kullanılmadı. Yeniden başlatın.\n')
            self.durum.configure(text=f'{baslik}: alanlar değiştiği için sonuç kullanılmadı.')
            return
        if hata is not None:
            self.hz = None
        if isinstance(hata, BelirsizSecim):
            self.durum.configure(text=f'{baslik}: seçim gerekiyor.')
            secilen = self._aday_sec(hata)
            if secilen:
                if hata.rol == 'SABLON':
                    self.v_sablon.set(str(secilen))
                else:
                    self.zorla[hata.rol] = Path(secilen)
                self.hz = None
                self._yaz(f'Seçildi ({rol_adi(hata.rol)}): {secilen}\n')
                if tekrar:
                    self.after(50, tekrar)
            return
        if hata is not None:
            mesaj = str(hata) or type(hata).__name__
            self._yaz(f'\nHATA: {mesaj}\n')
            self.durum.configure(text=f'{baslik} yapılamadı.')
            messagebox.showerror(baslik, mesaj)
            return
        bitince(sonuc)

    def _kuyrugu_bosalt_hemen(self):
        try:
            while True:
                self._yaz(self.q.get_nowait())
        except queue.Empty:
            pass

    def _aday_sec(self, hata):
        pen = tk.Toplevel(self)
        pen.title(f'Seçim gerekli — {rol_adi(hata.rol)}')
        pen.transient(self)
        pen.grab_set()
        ttk.Label(pen, text=hata.aciklama, wraplength=640, padding=10).pack(anchor='w')
        lb = tk.Listbox(pen, width=110, height=min(10, len(hata.adaylar)))
        for y in hata.adaylar:
            lb.insert('end', str(y))
        lb.pack(padx=10, fill='both', expand=True)
        lb.selection_set(0)
        sonuc = {}

        def tamam():
            sec = lb.curselection()
            if sec:
                sonuc['yol'] = hata.adaylar[sec[0]]
            pen.destroy()
        cub = ttk.Frame(pen, padding=10)
        cub.pack(fill='x')
        ttk.Button(cub, text='Bunu kullan', command=tamam).pack(side='right')
        ttk.Button(cub, text='Vazgeç', command=pen.destroy).pack(side='right', padx=6)
        self.wait_window(pen)
        return sonuc.get('yol')

    # ---------------------------------------------------------------- tanıma tablosu
    def _tabloyu_doldur(self, hz):
        self.agac.delete(*self.agac.get_children())
        self.satir_yolu = {}
        s = hz.secim
        if hz.sablon:
            iid = self.agac.insert('', 'end', text=f'[önceki ay] {hz.sablon}',
                                   values=('RAPOR', 'Şablon: önceki ayın raporu', s.donem.onceki().tire,
                                           f'✔ şablon ({hz.sablon_kaynagi})', ''), tags=('kullanilan',))
            self.satir_yolu[iid] = Path(hz.sablon)
        for d, konu, ac in hz.tum_notlar():
            if d in ('HATA', 'UYARI'):
                self.agac.insert('', 0 if d == 'HATA' else 'end', text=f'{"✘" if d == "HATA" else "⚠"} {d}: {konu}',
                                 values=('', ac, '', '', ''), tags=('hata' if d == 'HATA' else 'uyari',))
        sira = {r: i for i, r in enumerate(list(KULLANIM))}
        for b in sorted(s.belgeler, key=lambda b: (b.rol == '', sira.get(b.rol, 99), b.tur in ('BILINMEYEN', 'TARANMIS', 'DESTEKLENMEYEN'),
                                                   b.tur, b.goreli)):
            etiket = ('kullanilan',) if b.rol else (('sorun',) if b.tur in ('OKUNAMADI',) else ())
            iid = self.agac.insert('', 'end', text=b.goreli,
                                   values=(b.tur + (' (elle)' if b.elle else ''), ACIKLAMA.get(b.tur, ''),
                                           b.donem.tire if b.donem else '', KULLANIM.get(b.rol, ''), b.notu), tags=etiket)
            self.satir_yolu[iid] = b.yol
        self.defter.select(0)

    def _tur_duzelt(self, olay):
        iid = self.agac.identify_row(olay.y)
        if not iid or iid not in self.satir_yolu or self.calisiyor:
            return
        yol = self.satir_yolu[iid]
        if str(self.agac.item(iid, 'text')).startswith('[önceki ay]'):
            messagebox.showinfo('Şablon', 'Şablonu değiştirmek için "Önceki ayın raporu" alanından başka bir dosya seçin.')
            return
        if self.hz is None:
            messagebox.showinfo('Belge türü', 'Önce "1) Belgeleri tanı"ya basın.')
            return
        ok, ayar = self._ayar_ya_da_hata()
        if not ok:
            return
        duz = TurDuzeltmeleri(duzeltme_dosyasi(ayar))
        klasor = self.hz.klasor
        pen = tk.Toplevel(self)
        pen.title('Belge türünü düzelt')
        pen.transient(self)
        pen.grab_set()
        ttk.Label(pen, text=f'{yol.name}\nBu dosya nedir?', padding=10).pack(anchor='w')
        turler = [OTOMATIK] + TUR_LISTESI
        secenekler = ['Otomatik — içeriğinden belirlensin (elle düzeltmeyi kaldır)'] + [f'{t} — {ACIKLAMA.get(t, "")}' for t in TUR_LISTESI]
        cb = ttk.Combobox(pen, values=secenekler, state='readonly', width=70)
        mevcut = str(self.agac.set(iid, 'tur'))
        if mevcut.endswith(' (elle)'):
            mevcut = mevcut.replace(' (elle)', '')
        cb.current(turler.index(mevcut) if mevcut in turler else 0)
        cb.pack(padx=10)
        sonuc = {}

        def kaydet():
            sonuc['tur'] = turler[cb.current()]
            pen.destroy()
        cub = ttk.Frame(pen, padding=10)
        cub.pack(fill='x')
        ttk.Button(cub, text='Kaydet (sonraki aylarda da hatırla)', command=kaydet).pack(side='right')
        ttk.Button(cub, text='Vazgeç', command=pen.destroy).pack(side='right', padx=6)
        self.wait_window(pen)
        if 'tur' not in sonuc:
            return
        anahtar = duz.anahtar(yol, klasor)
        digerleri = [b.goreli for b in self.hz.secim.belgeler
                     if not Path(b.yol) == Path(yol) and duz.anahtar(b.yol, klasor) == anahtar]
        if digerleri and not messagebox.askyesno(
                'Belge türü', 'Bu düzeltme adı aynı (ay/yıl dışında) olan şu dosyalara da uygulanacak:\n\n'
                + '\n'.join(digerleri[:15]) + '\n\nDevam edilsin mi?'):
            return
        try:
            if sonuc['tur'] == OTOMATIK:
                duz.kaldir(yol, klasor)
            else:
                duz.ayarla(yol, sonuc['tur'], klasor)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror('Belge türü', str(e))
            return
        self._yaz(f'Tür düzeltildi: {yol.name} → {sonuc["tur"]} (kaydedildi: {duz.yol})\n')
        # bu dosya için daha önce yapılmış elle seçimler geçersiz
        self.zorla = {r: y for r, y in self.zorla.items() if Path(y) != Path(yol)}
        self.hz = None
        self.tani()

    # ---------------------------------------------------------------- işlemler
    def _hazirla_isi(self):
        ayar = self._ayar()
        klasor, sablon, zorla = self.v_klasor.get(), self.v_sablon.get() or None, dict(self.zorla)
        return lambda: (ayar, hazirla(klasor, ayar, sablon, zorla, ilerleme=lambda m: print(m)))

    def tani(self):
        if not self._kontrol():
            return
        self.hz = None
        try:
            is_ = self._hazirla_isi()
        except Exception as e:  # noqa: BLE001 - .exe'de konsol yok: her hata pencerede gösterilmeli
            messagebox.showerror('Ayar dosyası', str(e) or type(e).__name__)
            return
        self._hatirla_kaydet()
        self._arka_planda('Belge tanıma', is_, self._tani_bitti, tekrar=self.tani)

    def _tani_bitti(self, sonuc):
        _, hz = sonuc
        self.hz = hz
        self._tabloyu_doldur(hz)
        self._yaz(hz.tablo() + '\n')
        for d, konu, ac in hz.tum_notlar():
            if d in ('HATA', 'UYARI'):
                self._yaz(f'[{d}] {konu}: {ac}\n')
        eksik = hz.secim.eksik_zorunlu()
        uyari = sum(1 for d, _, _ in hz.tum_notlar() if d in ('HATA', 'UYARI'))
        ek = f' {uyari} uyarı var — tablonun başında (sarı satırlar), inceleyin.' if uyari else ''
        if eksik or not hz.sablon:
            parca = [rol_adi(r) for r in eksik] + ([] if hz.sablon else ['Şablon (önceki ayın raporu)'])
            self.durum.configure(text='Tanıma bitti — eksik: ' + ', '.join(parca) + '.' + ek)
        else:
            self.durum.configure(text=f'Tanıma bitti — {hz.secim.donem.tire}: gerekli belgeler ve şablon bulundu.{ek} '
                                      '"2) Taslak oluştur"a basabilirsiniz.')

    def taslak(self):
        if not self._kontrol():
            return
        ok, ayar = self._ayar_ya_da_hata()
        if not ok:
            return
        if self.hz is not None:
            ciddi = [ac for d, konu, ac in self.hz.tum_notlar()
                     if d == 'UYARI' and konu in ('Ayar dosyası', 'Belge dönemi', 'Şablon')]
            if ciddi and not messagebox.askyesno('Taslak', 'Tanımada şu uyarılar var:\n\n• ' + '\n• '.join(ciddi)
                                                 + '\n\nYine de taslak oluşturulsun mu?'):
                return
        self._hatirla_kaydet()
        klasor, sablon, zorla = self.v_klasor.get(), self.v_sablon.get() or None, dict(self.zorla)

        def is_():
            # klasördeki dosyalar son tanımadan sonra değişmiş olabilir: her seferinde yeniden tanınır (önbellek sayesinde hızlı)
            from .taslak import taslak_uret
            hz = hazirla(klasor, ayar, sablon, zorla, ilerleme=lambda m: print(m))
            print(hz.tablo())
            tc = taslak_uret(ayar, None, hz.klasor, hz=hz, ilerleme=lambda m: print(m))
            return hz, tc
        self._arka_planda('Taslak', is_, self._taslak_bitti, tekrar=self.taslak)

    def _taslak_bitti(self, sonuc):
        from .kontroller.cikti import konsol_ozeti
        hz, tc = sonuc
        self.hz = hz
        self._tabloyu_doldur(hz)
        self.son_cikti = tc.docx.parent
        self._yaz(f'\nTaslak : {tc.docx}\nKontrol: {tc.xlsx}\n         {tc.html}\n')
        if tc.farklar:
            self._yaz(f'Farklar: {tc.farklar} (gerçek raporla tutar farkı: {tc.karsilastirma.tutar_farki})\n')
        self._yaz('\n' + konsol_ozeti(tc.sonuc.kontroller) + '\n')
        self.defter.select(1)
        hata = tc.sonuc.kontroller.say('HATA')
        self.durum.configure(text=f'Taslak oluşturuldu: {tc.docx.name}' + (f' — kontrol listesinde {hata} HATA var, inceleyin.' if hata else '.'))
        if messagebox.askyesno('Taslak hazır', f'{tc.docx.name} ve kontrol listesi oluşturuldu.\nKlasör açılsın mı?'):
            klasoru_ac(self.son_cikti)

    def cikti_ac(self):
        yol = self.son_cikti
        if yol is None:
            k = yol_temizle(self.v_klasor.get())
            yol = k / 'CLAUDE TASLAK' if k else None
        if yol is None or not Path(yol).is_dir():
            messagebox.showinfo('Klasör', 'Henüz çıktı yok. Önce "2) Taslak oluştur"a basın.')
            return
        klasoru_ac(yol)

    def _cli(self, argv, baslik):
        def is_():
            kod = cli.main(argv)
            return kod
        self._arka_planda(baslik, is_, lambda kod: self.durum.configure(
            text=f'{baslik} tamamlandı.' if kod in (0, 1) else f'{baslik} yapılamadı (Çıktı sekmesine bakın).'))
        self.defter.select(1)

    def karsilastir(self):
        if self.calisiyor:
            return
        g = filedialog.askopenfilename(title='Ofisin bitmiş raporu', filetypes=[('Word', '*.doc *.docx')],
                                       initialdir=self.v_klasor.get() or None)
        if not g:
            return
        t = filedialog.askopenfilename(title='Programın taslağı', filetypes=[('Word', '*.docx')],
                                       initialdir=str(self.son_cikti) if self.son_cikti else None)
        if not t:
            return
        cikti = Path(t).with_name(Path(t).stem + ' FARKLAR.txt')
        self._cli(['karsilastir', g, t, '--cikti', str(cikti)], 'Karşılaştırma')

    def denetle(self):
        if self.calisiyor:
            return
        r = filedialog.askopenfilename(title='Denetlenecek rapor', filetypes=[('Word', '*.doc *.docx')])
        if r:
            self._cli(['denetle', r], 'Rapor denetimi')

    def geriye_donuk(self):
        if self.calisiyor:
            return
        bas = yol_temizle(self.v_klasor.get())
        k = filedialog.askdirectory(title='Firmanın bütün aylarını içeren klasör',
                                    initialdir=str(bas.parent) if bas and bas.parent.is_dir() else None)
        if not k:
            return
        if not messagebox.askyesno('Geriye dönük test', 'Klasördeki bütün aylar içerikten bulunup her biri için taslak '
                                   'üretilecek ve o ayın bitmiş raporuyla karşılaştırılacak. Bu birkaç dakika sürebilir. Devam?'):
            return
        argv = ['geriye-donuk', k] + (['--ayar', self.v_ayar.get()] if self.v_ayar.get().strip() else [])
        self._cli(argv, 'Geriye dönük test')


def main(klasor=None):
    Uygulama(klasor).mainloop()
