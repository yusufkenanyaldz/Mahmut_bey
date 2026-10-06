"""Basit masaüstü penceresi (tkinter): ayar dosyası + firma klasörü + dönem → taslak ve kontrol listesi.

Komut satırındaki işlemlerin aynısını çalıştırır (`cli.main`), çıktıyı pencereye yazar. Ek kütüphane gerektirmez.
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
from .klasorler import ay_klasoru
from .ortak import Donem

HATIRLA = Path.home() / '.teminat_cozum.json'


def program_klasoru():
    """.exe'nin (ya da kaynak kodun) bulunduğu klasör; firmalar/ burada aranır."""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parents[1]


class _KuyrukYazici:
    def __init__(self, q):
        self.q = q

    def write(self, s):
        self.q.put(s)
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
    def __init__(self):
        super().__init__()
        self.title(f'Teminat Çözüm Raporu — Taslak Programı {__version__}')
        self.geometry('960x640')
        self.minsize(760, 480)
        self.q = queue.Queue()
        self.calisiyor = False
        self.v_ayar, self.v_kok, self.v_donem = tk.StringVar(), tk.StringVar(), tk.StringVar()
        self._hatirla_yukle()
        self._kur()
        self.after(100, self._kuyrugu_bosalt)
        self._yaz('Ayar dosyasını, firma klasörünü (ör. ...\\TEMİNAT ÇÖZÜMÜ\\<FİRMA>) ve dönemi (ör. 2026-03) seçip '
                  '"Taslak oluştur"a basın. Taslak ve kontrol listesi o ayın klasöründeki "CLAUDE TASLAK" klasörüne yazılır; '
                  'orijinal dosyalara dokunulmaz.\n')
        if not soffice_yolu():
            self._yaz('\nNot: LibreOffice bulunamadı. Ofis raporları .doc ise çevrilemez; LibreOffice kurun ya da raporu '
                      'Word ile .docx olarak kaydedin.\n')

    # ---------------------------------------------------------------- yerleşim
    def _kur(self):
        ust = ttk.Frame(self, padding=10)
        ust.pack(fill='x')
        ust.columnconfigure(1, weight=1)
        satirlar = [('Ayar dosyası (ayar.yaml):', self.v_ayar, self._sec_ayar),
                    ('Firma klasörü:', self.v_kok, self._sec_kok)]
        for i, (etiket, deg, komut) in enumerate(satirlar):
            ttk.Label(ust, text=etiket).grid(row=i, column=0, sticky='w', pady=3)
            ttk.Entry(ust, textvariable=deg).grid(row=i, column=1, sticky='ew', padx=6)
            ttk.Button(ust, text='Seç…', command=komut).grid(row=i, column=2)
        ttk.Label(ust, text='Dönem (ör. 2026-03):').grid(row=2, column=0, sticky='w', pady=3)
        ttk.Entry(ust, textvariable=self.v_donem, width=14).grid(row=2, column=1, sticky='w', padx=6)

        dugmeler = ttk.Frame(self, padding=(10, 0))
        dugmeler.pack(fill='x')
        self.dugmeler = []
        for metin, komut in [('Taslak oluştur', self.taslak), ('Çıktı klasörünü aç', self.cikti_ac),
                             ('Gerçek raporla karşılaştır…', self.karsilastir), ('Rapor denetle…', self.denetle),
                             ('Geriye dönük test', self.geriye_donuk)]:
            b = ttk.Button(dugmeler, text=metin, command=komut)
            b.pack(side='left', padx=(0, 6), pady=6)
            self.dugmeler.append(b)

        alt = ttk.Frame(self, padding=(10, 0, 10, 10))
        alt.pack(fill='both', expand=True)
        self.metin = tk.Text(alt, wrap='word', font=('Consolas', 10) if sys.platform.startswith('win') else None)
        kay = ttk.Scrollbar(alt, command=self.metin.yview)
        self.metin.configure(yscrollcommand=kay.set, state='disabled')
        kay.pack(side='right', fill='y')
        self.metin.pack(side='left', fill='both', expand=True)
        self.durum = ttk.Label(self, text='Hazır.', anchor='w', padding=(10, 0, 10, 6))
        self.durum.pack(fill='x')

    # ---------------------------------------------------------------- yardımcılar
    def _hatirla_yukle(self):
        try:
            d = json.loads(HATIRLA.read_text(encoding='utf-8'))
        except (OSError, ValueError):
            d = {}
        self.v_ayar.set(d.get('ayar', ''))
        self.v_kok.set(d.get('kok', ''))
        self.v_donem.set(d.get('donem', ''))
        if not self.v_ayar.get():
            firmalar = program_klasoru() / 'firmalar'
            aday = sorted(p for p in firmalar.glob('*/ayar.yaml') if p.parent.name != 'ornek') if firmalar.is_dir() else []
            if len(aday) == 1:
                self.v_ayar.set(str(aday[0]))

    def _hatirla_kaydet(self):
        try:
            HATIRLA.write_text(json.dumps({'ayar': self.v_ayar.get(), 'kok': self.v_kok.get(), 'donem': self.v_donem.get()},
                                          ensure_ascii=False), encoding='utf-8')
        except OSError:
            pass

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

    def _sec_kok(self):
        p = filedialog.askdirectory(title='Firma klasörü (yıl klasörlerini içeren)', initialdir=self.v_kok.get() or None)
        if p:
            self.v_kok.set(p)

    def _donem(self):
        try:
            return Donem.coz(self.v_donem.get())
        except ValueError as e:
            messagebox.showerror('Dönem', str(e))
            return None

    def _ortak_kontrol(self):
        if self.calisiyor:
            return False
        if not self.v_kok.get() or not Path(self.v_kok.get()).is_dir():
            messagebox.showerror('Firma klasörü', 'Firma klasörünü seçin (ör. ...\\TEMİNAT ÇÖZÜMÜ\\<FİRMA>).')
            return False
        if self.v_ayar.get() and not Path(self.v_ayar.get()).is_file():
            messagebox.showerror('Ayar dosyası', f'Ayar dosyası bulunamadı: {self.v_ayar.get()}')
            return False
        return True

    def _ayar_arg(self):
        return ['--ayar', self.v_ayar.get()] if self.v_ayar.get() else []

    def _calistir(self, argv, baslik):
        self.calisiyor = True
        for b in self.dugmeler:
            b.state(['disabled'])
        self.durum.configure(text=f'{baslik} çalışıyor…')
        self._yaz(f'\n===== {baslik} =====\n')

        def is_parcacigi():
            yazici = _KuyrukYazici(self.q)
            try:
                with redirect_stdout(yazici), redirect_stderr(yazici):
                    kod = cli.main(argv)
            except SystemExit as e:
                kod = e.code if isinstance(e.code, int) else 2
                if not isinstance(e.code, int) and e.code:
                    self.q.put(f'{e.code}\n')
            except Exception:  # noqa: BLE001 - beklenmeyen hata pencerede görünsün
                self.q.put(traceback.format_exc())
                kod = 2
            self.after(0, self._bitti, baslik, kod)

        threading.Thread(target=is_parcacigi, daemon=True).start()

    def _bitti(self, baslik, kod):
        self.calisiyor = False
        for b in self.dugmeler:
            b.state(['!disabled'])
        durum = {0: 'tamamlandı.', 1: 'tamamlandı — kontrol listesinde HATA var, inceleyin.'}.get(kod, 'yapılamadı (yukarıdaki hataya bakın).')
        self.durum.configure(text=f'{baslik} {durum}')
        self._yaz(f'\n{baslik} {durum}\n')

    # ---------------------------------------------------------------- işlemler
    def taslak(self):
        if not self._ortak_kontrol():
            return
        d = self._donem()
        if not d:
            return
        self._hatirla_kaydet()
        self._calistir(['taslak', *self._ayar_arg(), '--kok', self.v_kok.get(), '--donem', d.tire], f'{d.tire} taslak')

    def cikti_ac(self):
        d = self._donem() if self.v_donem.get() else None
        try:
            yol = ay_klasoru(self.v_kok.get(), d) / 'CLAUDE TASLAK' if d else Path(self.v_kok.get())
        except Exception as e:  # noqa: BLE001
            messagebox.showerror('Klasör', str(e))
            return
        if not yol.is_dir():
            messagebox.showinfo('Klasör', f'Henüz çıktı yok: {yol}')
            return
        klasoru_ac(yol)

    def karsilastir(self):
        if self.calisiyor:
            return
        g = filedialog.askopenfilename(title='Ofisin bitmiş raporu', filetypes=[('Word', '*.doc *.docx')])
        if not g:
            return
        t = filedialog.askopenfilename(title='Programın taslağı', initialdir=str(Path(g).parent.parent / 'CLAUDE TASLAK'),
                                       filetypes=[('Word', '*.docx')])
        if not t:
            return
        cikti = Path(t).with_name(Path(t).stem + ' FARKLAR.txt')
        self._calistir(['karsilastir', g, t, '--cikti', str(cikti)], 'Karşılaştırma')

    def denetle(self):
        if self.calisiyor:
            return
        r = filedialog.askopenfilename(title='Denetlenecek rapor', filetypes=[('Word', '*.doc *.docx')])
        if r:
            self._calistir(['denetle', r], 'Rapor denetimi')

    def geriye_donuk(self):
        if not self._ortak_kontrol():
            return
        if not messagebox.askyesno('Geriye dönük test', 'Firma klasöründeki bütün aylar için taslak üretilip gerçek '
                                   'raporlarla karşılaştırılacak. Bu birkaç dakika sürebilir. Devam edilsin mi?'):
            return
        self._hatirla_kaydet()
        self._calistir(['geriye-donuk', *self._ayar_arg(), '--kok', self.v_kok.get()], 'Geriye dönük test')


def main():
    Uygulama().mainloop()
