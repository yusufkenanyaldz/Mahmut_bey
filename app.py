import functools
import os
import sys
from collections import OrderedDict

import customtkinter as ctk
import pandas as pd
from tkinter import filedialog, font as tkfont, messagebox, ttk

from denetim import checks, importers, inceleme
from denetim.export import export_sections  # noqa: F401  (eski içe aktarımlar için app.export_sections korunur)
from denetim.firms import FirmError, FirmRegistry
from denetim.utils import parse_account_list, parse_number

def _app_dir():
    """Uygulama klasörü. PyInstaller ile paketlenmiş .exe'de __file__ her açılışta silinen geçici bir
    klasörü gösterir; bu yüzden .exe'nin bulunduğu klasör kullanılır."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _data_dir(app_dir):
    """Veriler uygulamanın yanındaki `veri/` klasöründe tutulur. Klasör yazılabilir değilse (ör. .exe
    Program Files altındaysa) kullanıcının yerel uygulama verisi klasörüne geçilir."""
    preferred = os.path.join(app_dir, "veri")
    try:
        os.makedirs(preferred, exist_ok=True)
        probe = os.path.join(preferred, ".yazma_testi")
        with open(probe, "w", encoding="utf-8") as fh:
            fh.write("")
        os.remove(probe)
        return preferred
    except OSError:
        base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
        return os.path.join(base, "DenetimSistemi", "veri")


APP_DIR = _app_dir()
DATA_DIR = _data_dir(APP_DIR)
APP_TITLE = "Finansal Denetim ve Analiz Sistemi | SMMM Modülü"
SECTORS = ["İnşaat", "Halı Üretimi", "Uluslararası Taşımacılık", "Otomotiv Satış ve Kiralama", "Muhtelif İmalat"]
PERIOD_TYPES = ["Aylık", "Çeyreklik", "Yıllık"]
MAX_LOG_LINES = 300

pd.set_option("display.width", 250)
pd.set_option("display.max_columns", 30)
pd.set_option("display.max_colwidth", 40)

# --- TEMA: sakin, tek vurgu renkli palet. Her değer (aydınlık, karanlık) çiftidir. ---
C = {
    "bg": ("#F3F4F8", "#14161B"),           # pencere zemini
    "sidebar": ("#FFFFFF", "#1A1D23"),
    "card": ("#FFFFFF", "#1E2128"),
    "card_alt": ("#F7F8FB", "#23262E"),      # kart içi ikincil yüzey
    "border": ("#E3E6ED", "#2C3039"),
    "text": ("#1F2430", "#E6E8EE"),
    "muted": ("#6B7280", "#9AA1AE"),
    "faint": ("#9CA3AF", "#6B7280"),
    "accent": ("#4F6BED", "#6C84F5"),
    "accent_hover": ("#3F59D6", "#5A73E8"),
    "accent_soft": ("#EAEEFD", "#262D45"),   # aktif menü / ikincil buton zemini
    "accent_soft_hover": ("#DCE2FB", "#2E3654"),
    "success": ("#2E9E6A", "#3DB57D"),
    "success_soft": ("#E6F5EE", "#1D3329"),
    "warning": ("#C98A12", "#E0A83A"),
    "warning_soft": ("#FBF2E0", "#3A3020"),
    "danger": ("#D2504B", "#E06A65"),
    "danger_soft": ("#FBE9E8", "#3A2224"),
    "input": ("#F7F8FB", "#171A20"),
}
# Buton türleri: (zemin, üzerine gelince, yazı)
BTN = {
    "primary": (C["accent"], C["accent_hover"], ("#FFFFFF", "#FFFFFF")),
    "secondary": (C["accent_soft"], C["accent_soft_hover"], C["accent"]),
    "success": (C["success"], ("#258757", "#33A06D"), ("#FFFFFF", "#FFFFFF")),
    "warning": (C["warning_soft"], ("#F5E6C5", "#463A26"), C["warning"]),
    "danger": (C["danger_soft"], ("#F7D9D7", "#46282B"), C["danger"]),
    "ghost": ("transparent", C["card_alt"], C["muted"]),
}
FONT = "Segoe UI"


def tbutton(parent, text, command, kind="primary", height=38, width=None, font=None, **kw):
    """Temalı buton (primary / secondary / success / warning / danger / ghost)."""
    fg, hover, txt = BTN[kind]
    opts = dict(text=text, command=command, height=height, corner_radius=10, fg_color=fg, hover_color=hover,
                text_color=txt, font=font or ctk.CTkFont(family=FONT, size=13, weight="bold"), border_width=0)
    if width:
        opts["width"] = width
    opts.update(kw)
    return ctk.CTkButton(parent, **opts)


def initials(title):
    """Firma unvanından avatar için baş harfler: 'Arslan İnşaat Ltd. Şti.' → 'Aİ'."""
    words = [w for w in str(title or "").replace(".", " ").split() if w[:1].isalpha()]
    return ("".join(w[0] for w in words[:2]) or "?").upper()


def card(parent, **kw):
    """Yumuşak köşeli, ince kenarlıklı kart."""
    opts = dict(fg_color=C["card"], corner_radius=14, border_width=1, border_color=C["border"])
    opts.update(kw)
    return ctk.CTkFrame(parent, **opts)


def df_to_text(df, max_rows=500):
    if df.empty:
        return "  (bulgu yok)\n"
    text = df.head(max_rows).to_string(index=False)
    if len(df) > max_rows:
        text += f"\n  ... ve {len(df) - max_rows} satır daha (tamamı için Excel'e aktarın)"
    return text + "\n"


def requires_firm(method):
    """Ekranı yalnızca aktif firma varken açar; yoksa kullanıcıyı firma seçmeye yönlendirir."""
    @functools.wraps(method)
    def wrapper(self, *args, **kwargs):
        if self.db is None:
            return self.show_no_firm_screen()
        return method(self, *args, **kwargs)
    return wrapper


class ColumnMappingDialog(ctk.CTkToplevel):
    """Sütun eşleme sihirbazı: başlık satırı, önizleme ve her standart alan için dosya sütunu seçimi.

    Pencere kapandığında self.result onaylanan importers.ColumnMapping'dir (iptalde None).
    """
    NONE = "— (yok) —"

    def __init__(self, master, kind, raw, initial, file_name, message=""):
        super().__init__(master)
        self.kind, self.raw = kind, raw
        self.schema = importers.SCHEMAS[kind]
        self.required = importers.JOURNAL_REQUIRED if kind == importers.KIND_YEVMIYE else importers.INVOICE_REQUIRED
        self.result = None
        self.configure(fg_color=C["bg"])
        self.title(f"Sütunları Eşle — {file_name}")
        self.geometry("1000x720")
        font = ctk.CTkFont(family="Segoe UI", size=13)
        mono = ctk.CTkFont(family="Consolas", size=12)

        intro = ("Dosyadaki her sütunun hangi alana karşılık geldiğini seçin. * işaretli alanlar zorunludur.")
        if kind == importers.KIND_YEVMIYE:
            intro += ("\nTutar için Borç ve/veya Alacak ya da tek bir Tutar sütunu seçin. Belge No ayrı bir sütun "
                      "olmalıdır; açıklama sütununu Belge No olarak seçmeyin.")
        if message:
            intro = message + "\n\n" + intro
        ctk.CTkLabel(self, text=intro, font=font, justify="left", anchor="w", wraplength=950).pack(
            fill="x", padx=20, pady=(16, 8))

        top = ctk.CTkFrame(self, fg_color="transparent")
        top.pack(fill="x", padx=20)
        ctk.CTkLabel(top, text="Başlık satırı (Excel satır no):", font=font).pack(side="left", padx=(0, 6))
        n_rows = max(1, min(importers.HEADER_SCAN_ROWS, len(raw)))
        self.header_var = ctk.StringVar(value=str(initial.header_row + 1))
        ctk.CTkOptionMenu(top, values=[str(i + 1) for i in range(n_rows)], variable=self.header_var, width=80,
                          command=self.on_header_change).pack(side="left")
        ctk.CTkLabel(self, text=f"Önizleme (başlığın altındaki ilk {importers.PREVIEW_ROWS} satır):", font=font,
                     anchor="w").pack(fill="x", padx=20, pady=(10, 2))
        self.preview_box = ctk.CTkTextbox(self, font=mono, height=230, wrap="none")
        self.preview_box.pack(fill="x", padx=20)

        self.fields_frame = ctk.CTkFrame(self)
        self.fields_frame.pack(fill="both", expand=True, padx=20, pady=10)
        self.vars, self.menus = {}, {}
        for i, std in enumerate(self.fields()):
            label = importers.FIELD_LABELS.get(std, std) + (" *" if std in self.required else "")
            ctk.CTkLabel(self.fields_frame, text=label, font=font, anchor="w").grid(
                row=i // 2, column=(i % 2) * 2, padx=(12, 6), pady=4, sticky="w")
            self.vars[std] = ctk.StringVar(value=self.NONE)
            self.menus[std] = ctk.CTkOptionMenu(self.fields_frame, values=[self.NONE], variable=self.vars[std],
                                                width=260, dynamic_resizing=False)
            self.menus[std].grid(row=i // 2, column=(i % 2) * 2 + 1, padx=(0, 24), pady=4, sticky="w")

        buttons = ctk.CTkFrame(self, fg_color="transparent")
        buttons.pack(fill="x", padx=20, pady=(0, 16))
        tbutton(buttons, "✓  Onayla ve Yükle", self.confirm).pack(side="left", padx=(0, 10))
        tbutton(buttons, "İptal", self.destroy, kind="secondary").pack(side="left")
        self.refresh(initial.header_row, initial.columns)

        self.transient(master)
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.after(50, self._grab)

    def _grab(self):
        try:
            self.grab_set()
            self.focus_force()
        except Exception:  # Pencere henüz görünür değilse modal olmadan devam edilir
            pass

    def fields(self):
        if self.kind == importers.KIND_YEVMIYE:
            return ["Tarih", "Belge_No", "Hesap_Kodu", "Borc", "Alacak", "Tutar", "Aciklama"]
        return list(self.schema)

    def refresh(self, header_row, columns_map):
        self.layout = importers.analyze_layout(self.raw, self.schema, header_row)
        preview = self.layout.preview()
        self.preview_box.configure(state="normal")
        self.preview_box.delete("1.0", "end")
        text = preview.to_string(max_colwidth=28) if not preview.empty else "(başlığın altında veri yok)"
        self.preview_box.insert("1.0", "Sütunlar: " + " | ".join(self.layout.columns) + "\n\n" + text)
        self.preview_box.configure(state="disabled")
        values = [self.NONE] + self.layout.columns
        for std, menu in self.menus.items():
            menu.configure(values=values)
            col = columns_map.get(std)
            self.vars[std].set(col if col in self.layout.columns else self.NONE)

    def on_header_change(self, value):
        header_row = int(value) - 1
        self.refresh(header_row, importers.analyze_layout(self.raw, self.schema, header_row).mapping)

    def selected(self):
        return {std: var.get() for std, var in self.vars.items() if var.get() != self.NONE}

    def confirm(self):
        columns = self.selected()
        errors = importers.validate_mapping(columns, self.kind)
        if errors:
            messagebox.showwarning("Eşleme Tamamlanmadı", "\n\n".join(errors), parent=self)
            return
        self.result = importers.ColumnMapping(self.layout.header_row, columns, self.layout.signature)
        self.destroy()


def hucre_metni(value):
    """Tablo hücresi: boş değerler boş, ondalıklı sayılar Türkçe biçimde (1.234,56)."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, float):
        return f"{value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return str(value)


class BulguPaneli(ctk.CTkFrame):
    """Bulgu tablosu (ttk.Treeview) ve inceleme işaretleme: seçili satırları "İncelendi – Sorun Yok" /
    "Düzeltme İstendi" / "Açık" yapar, notu kaydeder. İş mantığı denetim/inceleme.py'dedir."""
    MAX_SATIR = 5000
    BASLIKLAR = {inceleme.DURUM_COL: "Durum", inceleme.NOT_COL: "İnceleme Notu",
                 inceleme.TARIH_COL: "İnceleme Tarihi"}
    DURUM_ETIKET = {inceleme.DURUM_ACIK: "acik", inceleme.DURUM_SORUN_YOK: "sorun_yok",
                    inceleme.DURUM_DUZELTME: "duzeltme"}

    def __init__(self, app, master):
        super().__init__(master, fg_color="transparent")
        self.app = app
        self.sections = OrderedDict()  # İncelenebilir bulgu bölümleri (ham)
        self.view = OrderedDict()      # İnceleme sütunları uygulanmış
        self.anahtarlar = {}           # Treeview satır kimliği → bulgu anahtarı
        self.etiketler = OrderedDict()
        self.liste_butonlari = {}      # bölüm başlığı → sol listedeki satır
        self.bolum_var = ctk.StringVar(value="")  # Seçili bölümün listedeki etiketi
        font = app.font_label
        small_font = ctk.CTkFont(family=FONT, size=12)

        # Sol: kontrol listesi (her kontrol için açık bulgu rozeti)
        left = card(self, width=270)
        left.pack(side="left", fill="y", padx=(0, 12))
        left.pack_propagate(False)
        ctk.CTkLabel(left, text="KONTROLLER", font=ctk.CTkFont(family=FONT, size=11, weight="bold"),
                     text_color=C["faint"], anchor="w").pack(fill="x", padx=16, pady=(14, 6))
        self.liste = ctk.CTkScrollableFrame(left, fg_color="transparent", corner_radius=0)
        self.liste.pack(fill="both", expand=True, padx=6, pady=(0, 10))

        # Sağ: araç çubuğu, tablo, işlem çubuğu
        right = ctk.CTkFrame(self, fg_color="transparent")
        right.pack(side="left", fill="both", expand=True)
        top = ctk.CTkFrame(right, fg_color="transparent")
        top.pack(fill="x", pady=(0, 8))
        self.baslik_label = ctk.CTkLabel(top, text="", font=ctk.CTkFont(family=FONT, size=15, weight="bold"),
                                         text_color=C["text"], anchor="w")
        self.baslik_label.pack(side="left")
        self.goster_var = ctk.BooleanVar(value=app.setting("inceleme_sorun_yok_goster", "0") == "1")
        ctk.CTkSwitch(top, text="Sorun yok olanları göster", variable=self.goster_var, font=small_font,
                      text_color=C["muted"], progress_color=C["accent"], command=self.on_toggle_goster) \
            .pack(side="right")
        bar = ctk.CTkFrame(right, fg_color="transparent")
        bar.pack(fill="x", pady=(0, 6))
        self.info_label = ctk.CTkLabel(bar, text="Kontrolü çalıştırınca bulgular burada listelenir.",
                                       font=small_font, text_color=C["muted"], anchor="w")
        self.info_label.pack(side="left", fill="x", expand=True)
        tbutton(bar, "Tümünü Seç", self.select_all, kind="ghost", height=26,
                font=ctk.CTkFont(family=FONT, size=12, weight="bold")).pack(side="right")

        table = card(right)
        table.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(table, style="Bulgu.Treeview", show="headings", selectmode="extended", height=4)
        vsb = ctk.CTkScrollbar(table, orientation="vertical", command=self.tree.yview)
        hsb = ctk.CTkScrollbar(table, orientation="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew", padx=(8, 0), pady=(8, 0))
        vsb.grid(row=0, column=1, sticky="ns", pady=(8, 0), padx=(0, 4))
        hsb.grid(row=1, column=0, sticky="ew", padx=(8, 0), pady=(0, 6))
        table.grid_rowconfigure(0, weight=1)
        table.grid_columnconfigure(0, weight=1)
        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        self.tree.bind("<Control-a>", lambda e: (self.select_all(), "break")[1])

        bottom = card(right, fg_color=C["card_alt"])
        bottom.pack(fill="x", pady=(10, 0))
        inner = ctk.CTkFrame(bottom, fg_color="transparent")
        inner.pack(fill="x", padx=10, pady=8)
        self.sel_label = ctk.CTkLabel(inner, text="Seçim yok", font=small_font, text_color=C["muted"], width=80,
                                      anchor="w")
        self.sel_label.pack(side="left", padx=(4, 8))
        small = dict(height=34, font=ctk.CTkFont(family=FONT, size=12, weight="bold"))
        tbutton(inner, "↺  Açığa Al", lambda: self.mark(inceleme.DURUM_ACIK), kind="secondary", width=100,
                **small).pack(side="right", padx=(6, 0))
        tbutton(inner, "✎  Düzeltme İstendi", lambda: self.mark(inceleme.DURUM_DUZELTME), kind="warning",
                width=150, **small).pack(side="right", padx=(6, 0))
        tbutton(inner, "✓  Sorun Yok", lambda: self.mark(inceleme.DURUM_SORUN_YOK), kind="success", width=118,
                **small).pack(side="right", padx=(6, 0))
        self.not_entry = ctk.CTkEntry(inner, height=34, corner_radius=10, border_width=1,
                                      border_color=C["border"], fg_color=C["card"],
                                      placeholder_text="Not (seçili bulgulara yazılır)")
        self.not_entry.pack(side="left", fill="x", expand=True, padx=(0, 4))
        app.apply_tree_style()
        self.build_list(None)

    # ------------------------------------------------------------------ veri
    def set_sections(self, sections):
        """Rapor sonuçlarını yükler; yalnızca incelenebilir bulgu bölümleri tabloda gösterilir."""
        self.sections = OrderedDict((t, df) for t, df in (sections or {}).items()
                                    if inceleme.kontrol_kodu(t) is not None and df is not None)
        secili = self.bolum_basligi()
        self.reload()
        if secili not in self.sections:  # İlk açık bulgusu olan bölüm seçilir
            secili = next((t for t, df in self.view.items() if (df[inceleme.DURUM_COL] == inceleme.DURUM_ACIK).any()),
                          next(iter(self.sections), ""))
        self.set_bolum(secili)

    def reload(self):
        """İnceleme kayıtlarını veritabanından yeniden okuyup bölümlere uygular, menü etiketlerini günceller."""
        self.view = inceleme.bolumlere_uygula(self.sections, inceleme.incelemeleri_oku(self.app.db))
        ozet = inceleme.inceleme_ozeti(self.view).set_index("Kontrol")
        self.etiketler = OrderedDict(
            (f"{t}  —  {inceleme.ozet_satiri(ozet.loc[t])}" if t in ozet.index else t, t) for t in self.view)
        self.build_list(ozet)

    def build_list(self, ozet):
        """Sol listedeki kontrol satırlarını rozetleriyle yeniden çizer (rozet: açık bulgu sayısı)."""
        for w in self.liste.winfo_children():
            w.destroy()
        self.liste_butonlari = {}
        if not self.view:
            ctk.CTkLabel(self.liste, text="Kontrolü çalıştırınca\nbulgular burada listelenir.",
                         font=ctk.CTkFont(family=FONT, size=12), text_color=C["faint"], justify="left") \
                .pack(fill="x", padx=10, pady=10)
            return
        for t in self.view:
            acik = int(ozet.loc[t, "Acik"]) if t in ozet.index else 0
            duz = int(ozet.loc[t, "Duzeltme_Istendi"]) if t in ozet.index else 0
            row = ctk.CTkFrame(self.liste, fg_color="transparent", corner_radius=10, cursor="hand2")
            row.pack(fill="x", pady=1)
            lbl = ctk.CTkLabel(row, text=t, font=ctk.CTkFont(family=FONT, size=12), text_color=C["text"],
                               anchor="w", justify="left", wraplength=175)
            lbl.pack(side="left", fill="x", expand=True, padx=(10, 4), pady=7)
            if acik:
                rozet, renk, zemin = str(acik), C["danger"], C["danger_soft"]
            elif duz:
                rozet, renk, zemin = str(duz), C["warning"], C["warning_soft"]
            else:
                rozet, renk, zemin = "✓", C["success"], C["success_soft"]
            ctk.CTkLabel(row, text=rozet, font=ctk.CTkFont(family=FONT, size=11, weight="bold"), text_color=renk,
                         fg_color=zemin, corner_radius=9, width=34, height=22).pack(side="right", padx=(0, 8))
            for w in (row, lbl):
                w.bind("<Button-1>", lambda _e, b=t: self.set_bolum(b))
            self.liste_butonlari[t] = row

    def bolum_basligi(self):
        return self.etiketler.get(self.bolum_var.get(), "")

    def set_bolum(self, baslik):
        etiket = next((e for e, t in self.etiketler.items() if t == baslik), "")
        self.bolum_var.set(etiket)
        for t, row in self.liste_butonlari.items():
            row.configure(fg_color=C["accent_soft"] if t == baslik else "transparent")
        self.baslik_label.configure(text=baslik or "Bulgular")
        self.refresh()

    def export_sections(self, sections):
        """Excel çıktısı: bulgu bölümlerine durum / not / tarih sütunları ve baştaki özetin ardına İnceleme Özeti."""
        if not sections:
            return sections
        view = inceleme.bolumlere_uygula(sections, inceleme.incelemeleri_oku(self.app.db))
        items = list(view.items())
        pos = 1 if items and items[0][0] == "Özet" else 0
        items.insert(pos, ("İnceleme Özeti", inceleme.inceleme_ozeti(view)))
        return OrderedDict(items)

    # ------------------------------------------------------------------ görünüm
    def refresh(self):
        tree = self.tree
        tree.delete(*tree.get_children())
        self.anahtarlar = {}
        baslik = self.bolum_basligi()
        df = self.view.get(baslik)
        if df is None:
            tree["columns"] = ()
            self.info_label.configure(text="Kontrolü çalıştırınca bulgular burada listelenir." if not self.sections else "Bu kontrolde bulgu yok.")
            self.update_sel_label()
            return
        gorunen = df if self.goster_var.get() else inceleme.sorun_yok_gizle(df)
        gizli = len(df) - len(gorunen)
        veri_cols = [c for c in df.columns if c not in inceleme.INCELEME_COLS]
        cols = [inceleme.DURUM_COL] + veri_cols + [inceleme.NOT_COL, inceleme.TARIH_COL]
        tree["columns"] = cols
        ornek = gorunen.head(200)
        yazi, kalin = tkfont.Font(font=self.app.tree_font), tkfont.Font(font=self.app.tree_heading_font)
        for c in cols:
            baslik_metni = self.BASLIKLAR.get(c, c.replace("_", " "))
            en_uzun = max([hucre_metni(v) for v in ornek[c]] + [""], key=len)
            genislik = max(kalin.measure(baslik_metni), yazi.measure(en_uzun)) + 24
            tree.heading(c, text=baslik_metni, anchor="w")
            tree.column(c, width=max(70, min(360, genislik)), minwidth=50, stretch=False, anchor="w")
        for i, (_, r) in enumerate(gorunen.head(self.MAX_SATIR).iterrows()):
            iid = str(i)
            self.anahtarlar[iid] = r[inceleme.ANAHTAR_COL]
            tree.insert("", "end", iid=iid, values=[hucre_metni(r[c]) for c in cols],
                        tags=(self.DURUM_ETIKET.get(r[inceleme.DURUM_COL], "acik"), "tek" if i % 2 else "cift"))
        text = f"{len(gorunen)} bulgu gösteriliyor"
        if gizli:
            text += f" ({gizli} 'sorun yok' gizli)"
        if len(gorunen) > self.MAX_SATIR:
            text += f" — ilk {self.MAX_SATIR} satır; tamamı için Excel'e aktarın"
        self.info_label.configure(text=text)
        self.update_sel_label()

    def update_sel_label(self):
        n = len(self.tree.selection())
        self.sel_label.configure(text=f"{n} seçili" if n else "Seçim yok",
                                 text_color=C["accent"] if n else C["muted"])

    def on_toggle_goster(self):
        self.app.db.set_setting("inceleme_sorun_yok_goster", "1" if self.goster_var.get() else "0")
        self.refresh()

    def on_select(self, _event=None):
        sec = self.tree.selection()
        if len(sec) == 1:  # Tek satır seçilince mevcut not düzenlemek için kutuya gelir
            df = self.view.get(self.bolum_basligi())
            notlar = df.loc[df[inceleme.ANAHTAR_COL] == self.anahtarlar.get(sec[0]), inceleme.NOT_COL]
            self.not_entry.delete(0, "end")
            if len(notlar) and notlar.iloc[0]:
                self.not_entry.insert(0, notlar.iloc[0])
        self.update_sel_label()

    def select_all(self):
        self.tree.selection_set(self.tree.get_children())

    def mark(self, durum):
        sec = self.tree.selection()
        if not sec:
            messagebox.showwarning("Uyarı", "Önce tablodan bir veya daha fazla bulgu seçin.")
            return
        not_ = self.not_entry.get().strip()
        n = inceleme.isaretle(self.app.db, [self.anahtarlar[i] for i in sec], durum, not_ or None)
        baslik = self.bolum_basligi()
        self.reload()
        self.set_bolum(baslik)
        self.info_label.configure(text=f"{n} bulgu '{durum}' olarak işaretlendi.  " + self.info_label.cget("text"))


# --- ARAYÜZ (GUI) - AYDINLIK/KARANLIK MOD DESTEKLİ ---
class AuditApp(ctk.CTk):
    def __init__(self, data_dir=DATA_DIR, legacy_dirs=None):
        super().__init__()
        self.registry = FirmRegistry(data_dir)
        self.legacy_dirs = legacy_dirs if legacy_dirs is not None else list(dict.fromkeys([APP_DIR, os.getcwd()]))
        self.firm = None
        self.db = None
        self.analysis_df = None
        self.recon_sections = None
        self.audit_sections = None
        self.recon_haric = self.audit_haric = None
        self.analysis_result = self.audit_fiyat_haric = self.recon_kur = self.audit_kur_farki = None
        self.recon_ek, self.audit_ek = [], []

        self.title(APP_TITLE)
        self.geometry("1360x860")
        self.minsize(1100, 700)

        ctk.set_appearance_mode(self.registry.get_pref("appearance", "Light"))
        ctk.set_default_color_theme("blue")
        self.configure(fg_color=C["bg"])

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.font_title = ctk.CTkFont(family=FONT, size=24, weight="bold")
        self.font_subtitle = ctk.CTkFont(family=FONT, size=13)
        self.font_btn = ctk.CTkFont(family=FONT, size=13, weight="bold")
        self.font_label = ctk.CTkFont(family=FONT, size=13)
        self.font_small = ctk.CTkFont(family=FONT, size=12)
        self.font_section = ctk.CTkFont(family=FONT, size=11, weight="bold")
        self.font_console = ctk.CTkFont(family="Consolas", size=12)
        self.tree_font = (FONT, 10)
        self.tree_heading_font = (FONT, 10, "bold")
        self.check_style = dict(font=self.font_small, text_color=C["text"], fg_color=C["accent"],
                                hover_color=C["accent_hover"], border_color=C["faint"], corner_radius=6,
                                checkbox_width=20, checkbox_height=20, border_width=2)
        self.nav_buttons = {}
        self._row_parent = None  # create_button_row satırlarının ekleneceği kap (ayar kartı açıkken)

        self.create_sidebar()
        self.create_main_frame()
        self.startup()

    # ------------------------------------------------------------------ firma yönetimi
    def startup(self):
        """Eski tek veritabanını (varsa) aktarır, son seçilen firmayı açar."""
        legacy = self.registry.find_legacy_db(*self.legacy_dirs)
        if self.registry.legacy_pending(legacy):
            old_title, _ = self.registry.read_legacy_company(legacy)
            title = old_title or self.ask_legacy_title()
            try:
                firm = self.registry.import_legacy(legacy, title=title)
                self.registry.set_last_firm(firm.code)
                messagebox.showinfo("Veri Aktarımı",
                                    f"Önceki sürümün veritabanı '{firm.title}' adlı firma olarak aktarıldı.\n\n"
                                    f"Firma bilgilerini 'Firmalar' ekranından düzenleyebilirsiniz.\n"
                                    f"Eski dosya silinmedi: {legacy}")
            except Exception as e:
                messagebox.showerror("Hata", f"Önceki sürümün veritabanı aktarılamadı:\n{e}")
        last = self.registry.get_last_firm()
        if last:
            self.select_firm(last.code)
        else:
            self.show_welcome_screen()

    def ask_legacy_title(self):
        dialog = ctk.CTkInputDialog(title="Önceki Verilerin Aktarımı",
                                    text="Önceki sürümde kayıtlı veriler bulundu ve bir firma olarak aktarılacak.\n"
                                         "Bu verilerin ait olduğu firmanın unvanını girin\n"
                                         "(boş bırakırsanız 'Aktarılan Firma' adı kullanılır):")
        return (dialog.get_input() or "").strip()

    def select_firm(self, code):
        try:
            db = self.registry.open_db(code)
        except Exception as e:
            messagebox.showerror("Hata", f"Firma veritabanı açılamadı:\n{e}")
            return
        self.firm, self.db = self.registry.get(code), db
        self.registry.set_last_firm(self.firm.code)
        self.analysis_df = self.recon_sections = self.audit_sections = self.recon_haric = self.audit_haric = None
        self.analysis_result = self.audit_fiyat_haric = self.recon_kur = self.audit_kur_farki = None
        self.refresh_firm_display()
        self.show_welcome_screen()

    def close_firm(self):
        self.firm = self.db = None
        self.analysis_df = self.recon_sections = self.audit_sections = self.recon_haric = self.audit_haric = None
        self.analysis_result = self.audit_fiyat_haric = self.recon_kur = self.audit_kur_farki = None
        self.refresh_firm_display()

    def refresh_firm_display(self):
        if self.firm:
            self.title(f"{APP_TITLE} | {self.firm.title}")
            vkn = f"VKN {self.firm.vkn}" if self.firm.vkn else "VKN girilmedi"
            self.firm_avatar.configure(text=initials(self.firm.title), fg_color=C["accent_soft"],
                                       text_color=C["accent"])
            self.firm_name_label.configure(text=self.firm.title, text_color=C["text"])
            self.firm_info_label.configure(text=f"{self.firm.code}\n{vkn}")
        else:
            self.title(APP_TITLE)
            self.firm_avatar.configure(text="?", fg_color=C["warning_soft"], text_color=C["warning"])
            self.firm_name_label.configure(text="Firma seçilmedi", text_color=C["warning"])
            self.firm_info_label.configure(text="Çalışmak için bir firma seçin")

    # ------------------------------------------------------------------ ayar yardımcıları
    def setting(self, key, default=""):
        return self.db.get_setting(key, default)

    def read_threshold(self, entry):
        try:
            value = parse_number(entry.get())
        except ValueError:
            value = None
        if value is None or value <= 0 or value >= 1000:
            messagebox.showwarning("Uyarı", "Sapma eşiği 0 ile 1000 arasında bir sayı olmalıdır (ör. 15).")
            return None
        self.db.set_setting("threshold", f"{value:g}")
        return value

    def read_tolerance(self, entry):
        try:
            value = parse_number(entry.get())
        except ValueError:
            value = None
        if value is None or value < 0:
            messagebox.showwarning("Uyarı", "Tutar toleransı 0 veya pozitif bir sayı olmalıdır (ör. 0,01).")
            return None
        self.db.set_setting("tolerance", f"{value:g}")
        return value

    def read_kur_toleransi(self, entry):
        """Dövizli faturalar için yüzde kur toleransı (firma ayarı; 0 → kapalı)."""
        try:
            value = parse_number(entry.get())
        except ValueError:
            value = None
        if value is None or value < 0 or value >= 100:
            messagebox.showwarning("Uyarı", "Kur toleransı 0 ile 100 arasında bir yüzde olmalıdır (ör. "
                                            f"{checks.KUR_TOLERANSI_VARSAYILAN:g}; 0 → yüzde tolerans yok).")
            return None
        self.db.set_setting("kur_toleransi", f"{value:g}")
        return value

    def add_tolerans_row(self):
        """Tutar toleransları satırı: TL faturalar için sabit TL, dövizli faturalar için yüzde kur toleransı.
        Dönüş: (TL tolerans kutusu, kur toleransı kutusu)"""
        row = self.create_button_row()
        tl = self.add_labeled_entry(row, "Tolerans (TL)", self.setting("tolerance", "0.01"), 70)
        kur = self.add_labeled_entry(row, "Kur toleransı (%, dövizli)",
                                     self.setting("kur_toleransi", f"{checks.KUR_TOLERANSI_VARSAYILAN:g}"), 55)
        ctk.CTkLabel(row, text="(TL faturalara yüzde uygulanmaz)", font=self.font_label,
                     text_color=C["faint"], anchor="w").pack(side="left")
        return tl, kur

    def read_accounts(self, entry):
        accounts = parse_account_list(entry.get())
        if not accounts:
            messagebox.showwarning("Uyarı", "Mutabakat için en az bir hesap kodu girin (ör. 153, 770).")
            return None
        self.db.set_setting("accounts", entry.get().strip())
        return accounts

    def haric_onek_text(self):
        """Firmanın faturasız kayıt kontrolünde hariç tuttuğu belge no önekleri (hiç kaydedilmediyse varsayılan)."""
        return self.setting("haric_onekler", checks.VARSAYILAN_HARIC_ONEKLER)

    def read_haric_onekler(self, entry):
        """Önek listesini firma ayarına kaydeder; boş bırakılırsa önek filtresi uygulanmaz."""
        text = entry.get().strip()
        self.db.set_setting("haric_onekler", text)
        return checks.parse_onek_listesi(text)

    def add_haric_onek_row(self, with_excel_option=True):
        """Hariç önek listesi satırı: giriş kutusu, Varsayılan butonu ve (istenirse) Excel seçeneği."""
        row = self.create_button_row()
        entry = self.add_labeled_entry(row, "Hariç belge önekleri", self.haric_onek_text(), 380,
                                       "boş: önek filtresi yok")

        def reset():
            entry.delete(0, "end")
            entry.insert(0, checks.VARSAYILAN_HARIC_ONEKLER)

        tbutton(row, "Varsayılan", reset, kind="secondary", width=96, height=32).pack(side="left", padx=(0, 18))
        var = None
        if with_excel_option:
            var = ctk.BooleanVar(value=self.setting("haric_excel", "0") == "1")
            ctk.CTkCheckBox(row, text="Hariç tutulanları Excel'e ekle", variable=var, **self.check_style,
                            command=lambda: self.db.set_setting("haric_excel", "1" if var.get() else "0")) \
                .pack(side="left")
        return entry, var

    @staticmethod
    def with_haric_sheet(sections, haric, var):
        """Excel çıktısına, seçiliyse faturasız listeden hariç tutulan kayıtların bilgi sayfasını ekler."""
        if not sections or haric is None or var is None or not var.get():
            return sections
        return OrderedDict(list(sections.items()) + [("Faturasız Listeden Hariç Tutulanlar", haric)])

    def add_fiyat_kural_rows(self):
        """Fiyat analizi hariç tutma kuralları: tevkifat ve anahtar kelime kuralları (ayrı ayrı açılıp kapatılır),
        kelime listesi (Varsayılan butonuyla), en az alım sayısı ve analiz dışı satırların Excel seçeneği."""
        row = self.create_button_row()
        tevkifat = ctk.BooleanVar(value=self.setting("fiyat_tevkifat_haric", "1") == "1")
        ctk.CTkCheckBox(row, text="Tevkifatlı faturaları hariç tut", variable=tevkifat, **self.check_style) \
            .pack(side="left", padx=(0, 18))
        kelime = ctk.BooleanVar(value=self.setting("fiyat_kelime_haric", "1") == "1")
        ctk.CTkCheckBox(row, text="Anahtar kelimeyle hariç tut", variable=kelime, **self.check_style) \
            .pack(side="left", padx=(0, 18))
        min_alim = self.add_labeled_entry(row, "En az alım (dönemde)",
                                          self.setting("fiyat_min_alim", str(checks.FIYAT_MIN_ALIM)), 50)
        excel = ctk.BooleanVar(value=self.setting("fiyat_haric_excel", "1") == "1")
        ctk.CTkCheckBox(row, text="Analiz dışı satırları Excel'e ekle", variable=excel, **self.check_style,
                        command=lambda: self.db.set_setting("fiyat_haric_excel", "1" if excel.get() else "0")) \
            .pack(side="left")
        row2 = self.create_button_row()
        kelimeler = self.add_labeled_entry(
            row2, "Hariç ürün / hizmet kelimeleri",
            self.setting("fiyat_haric_kelimeler", checks.VARSAYILAN_FIYAT_HARIC_KELIMELER), 460,
            "boş: kelime filtresi yok")

        def reset():
            kelimeler.delete(0, "end")
            kelimeler.insert(0, checks.VARSAYILAN_FIYAT_HARIC_KELIMELER)

        tbutton(row2, "Varsayılan", reset, kind="secondary", width=96, height=32).pack(side="left")
        return {"tevkifat": tevkifat, "kelime": kelime, "min_alim": min_alim, "kelimeler": kelimeler, "excel": excel}

    def read_fiyat_kurallari(self, w):
        """Fiyat analizi kurallarını doğrular, firma ayarlarına kaydeder ve checks.FiyatKurallari döndürür."""
        try:
            min_alim = parse_number(w["min_alim"].get())
        except ValueError:
            min_alim = None
        if min_alim is None or min_alim < 1 or min_alim != int(min_alim) or min_alim > 1000:
            messagebox.showwarning("Uyarı", "En az alım sayısı 1 veya daha büyük bir tam sayı olmalıdır (ör. "
                                            f"{checks.FIYAT_MIN_ALIM}; 1 → yetersiz veri kuralı kapalı).")
            return None
        text = w["kelimeler"].get().strip()
        self.db.set_setting("fiyat_tevkifat_haric", "1" if w["tevkifat"].get() else "0")
        self.db.set_setting("fiyat_kelime_haric", "1" if w["kelime"].get() else "0")
        self.db.set_setting("fiyat_haric_kelimeler", text)
        self.db.set_setting("fiyat_min_alim", str(int(min_alim)))
        return checks.FiyatKurallari(w["tevkifat"].get(), w["kelime"].get(), checks.parse_kelime_listesi(text),
                                     int(min_alim))

    def add_vergi_rows(self):
        """Vergi ve satış mutabakatı ayarları: KDV (191), tevkifat (360), gelir (600–602) ve hesaplanan KDV (391)
        hesapları, maliyete eklenen vergi türü kodları (Varsayılan butonuyla)."""
        row0 = self.create_button_row()
        kdv = self.add_labeled_entry(row0, "İndirilecek KDV",
                                     self.setting("kdv_hesaplari", checks.VARSAYILAN_KDV_HESAPLARI), 90,
                                     "boş: KDV kontrolü yok")
        tev = self.add_labeled_entry(row0, "Tevkifat",
                                     self.setting("tevkifat_hesaplari", checks.VARSAYILAN_TEVKIFAT_HESAPLARI), 90,
                                     "boş: tevkifat kontrolü yok")
        gelir = self.add_labeled_entry(row0, "Gelir (satış)",
                                       self.setting("gelir_hesaplari", checks.VARSAYILAN_GELIR_HESAPLARI), 120,
                                       "boş: satış mutabakatı yok")
        skdv = self.add_labeled_entry(row0, "Hesaplanan KDV",
                                      self.setting("satis_kdv_hesaplari", checks.VARSAYILAN_SATIS_KDV_HESAPLARI), 70,
                                      "boş: yok")
        row = self.create_button_row()
        kodlar = self.add_labeled_entry(
            row, "Maliyete eklenen vergi kodları",
            self.setting("maliyet_vergi_kodlari", checks.VARSAYILAN_MALIYET_VERGI_KODLARI), 420,
            "boş: vergi eklenmez (ör. 9077 ÖTV II, 0071 ÖTV I, 4080 ÖİV)")

        def reset():
            kodlar.delete(0, "end")
            kodlar.insert(0, checks.VARSAYILAN_MALIYET_VERGI_KODLARI)

        tbutton(row, "Varsayılan", reset, kind="secondary", width=96, height=32).pack(side="left")
        return {"maliyet": kodlar, "kdv": kdv, "tevkifat": tev, "gelir": gelir, "satis_kdv": skdv}

    def read_vergi_ayarlari(self, w):
        """Vergi ayarlarını firma ayarlarına kaydeder ve checks.VergiAyarlari döndürür."""
        metin = {k: e.get().strip() for k, e in w.items()}
        self.db.set_setting("maliyet_vergi_kodlari", metin["maliyet"])
        self.db.set_setting("kdv_hesaplari", metin["kdv"])
        self.db.set_setting("tevkifat_hesaplari", metin["tevkifat"])
        self.db.set_setting("gelir_hesaplari", metin["gelir"])
        self.db.set_setting("satis_kdv_hesaplari", metin["satis_kdv"])
        return checks.VergiAyarlari(checks.parse_vergi_kodlari(metin["maliyet"]), parse_account_list(metin["kdv"]),
                                    parse_account_list(metin["tevkifat"]), parse_account_list(metin["gelir"]),
                                    parse_account_list(metin["satis_kdv"]))

    def maliyet_vergili_faturalar(self, vergi):
        """Faturalar + maliyete eklenen vergi tutarları (beklenen maliyet için)."""
        return checks.maliyet_vergisi_ekle(self.db.get_invoices_df(), self.db.get_invoice_taxes_df(),
                                           vergi.maliyet_kodlari)

    @staticmethod
    def with_fiyat_haric_sheet(sections, haric, w):
        """Excel çıktısına, seçiliyse fiyat analizi dışında kalan satırların bilgi sayfasını ekler."""
        if not sections or haric is None or w is None or not w["excel"].get():
            return sections
        return OrderedDict(list(sections.items()) + [("Fiyat Analizi Dışı Satırlar", haric)])

    # ------------------------------------------------------------------ iskelet
    def create_sidebar(self):
        self.sidebar = ctk.CTkFrame(self, width=264, corner_radius=0, fg_color=C["sidebar"], border_width=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_propagate(False)
        self.sidebar.grid_columnconfigure(0, weight=1)

        logo = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        logo.grid(row=0, column=0, padx=22, pady=(26, 18), sticky="ew")
        ctk.CTkLabel(logo, text="◆", font=ctk.CTkFont(family=FONT, size=22), text_color=C["accent"]) \
            .pack(side="left", padx=(0, 10))
        names = ctk.CTkFrame(logo, fg_color="transparent")
        names.pack(side="left")
        ctk.CTkLabel(names, text="Denetim", font=ctk.CTkFont(family=FONT, size=19, weight="bold"),
                     text_color=C["text"], anchor="w").pack(fill="x")
        ctk.CTkLabel(names, text="Finansal analiz sistemi", font=self.font_small, text_color=C["muted"],
                     anchor="w").pack(fill="x")

        # Aktif firma kartı (tıklanınca firmalar ekranı)
        firm = ctk.CTkFrame(self.sidebar, fg_color=C["card_alt"], corner_radius=14, border_width=1,
                            border_color=C["border"], cursor="hand2")
        firm.grid(row=1, column=0, padx=16, pady=(0, 18), sticky="ew")
        self.firm_avatar = ctk.CTkLabel(firm, text="?", width=40, height=40, corner_radius=12,
                                        font=ctk.CTkFont(family=FONT, size=14, weight="bold"))
        self.firm_avatar.grid(row=0, column=0, rowspan=2, padx=(12, 10), pady=12)
        self.firm_name_label = ctk.CTkLabel(firm, text="", font=ctk.CTkFont(family=FONT, size=13, weight="bold"),
                                            anchor="w", justify="left", wraplength=150)
        self.firm_name_label.grid(row=0, column=1, sticky="sw", pady=(12, 0), padx=(0, 10))
        self.firm_info_label = ctk.CTkLabel(firm, text="", font=ctk.CTkFont(family=FONT, size=11),
                                            text_color=C["muted"], anchor="w", justify="left", wraplength=150)
        self.firm_info_label.grid(row=1, column=1, sticky="nw", pady=(0, 12), padx=(0, 10))
        firm.grid_columnconfigure(1, weight=1)
        for w in (firm, self.firm_avatar, self.firm_name_label, self.firm_info_label):
            w.bind("<Button-1>", lambda _e: self.show_firms_frame())

        nav = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        nav.grid(row=2, column=0, padx=12, sticky="new")
        groups = [
            (None, [("home", "⌂", "Ana Sayfa", self.show_welcome_screen)]),
            ("VERİ YÜKLE", [("xml", "⇪", "e-Fatura (XML)", self.show_import_frame),
                            ("excel", "▦", "Fatura (Excel)", self.show_excel_import_frame),
                            ("yevmiye", "☰", "Yevmiye (Excel)", self.show_journal_import_frame)]),
            ("ANALİZ", [("fiyat", "↗", "Fiyat Risk Analizi", self.show_analysis_frame),
                        ("mutabakat", "⚖", "Muhasebe Mutabakatı", self.show_reconciliation_frame),
                        ("rapor", "✓", "Genel Denetim Raporu", self.show_audit_frame)]),
            ("YÖNETİM", [("firmalar", "◫", "Firmalar", self.show_firms_frame),
                         ("ayarlar", "⚙", "Firma Ayarları", self.show_settings_frame)]),
        ]
        for title, items in groups:
            if title:
                ctk.CTkLabel(nav, text=title, font=self.font_section, text_color=C["faint"], anchor="w") \
                    .pack(fill="x", padx=12, pady=(14, 4))
            for key, icon, text, command in items:
                btn = ctk.CTkButton(nav, text=f"{icon}   {text}", command=command, anchor="w", height=38,
                                    corner_radius=10, fg_color="transparent", hover_color=C["card_alt"],
                                    text_color=C["text"], font=ctk.CTkFont(family=FONT, size=13))
                btn.pack(fill="x", pady=1)
                self.nav_buttons[key] = btn

        self.sidebar.grid_rowconfigure(3, weight=1)
        bottom = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        bottom.grid(row=4, column=0, padx=22, pady=(10, 18), sticky="ew")
        dark = ctk.get_appearance_mode() == "Dark"
        self.switch_var = ctk.StringVar(value="on" if dark else "off")
        self.mode_switch = ctk.CTkSwitch(bottom, text="Karanlık mod" if dark else "Aydınlık mod",
                                         command=self.toggle_mode, variable=self.switch_var, onvalue="on",
                                         offvalue="off", font=self.font_small, text_color=C["muted"],
                                         progress_color=C["accent"])
        self.mode_switch.pack(anchor="w")
        ctk.CTkLabel(bottom, text="Sürüm 2.1  ·  Çevrimdışı", font=ctk.CTkFont(family=FONT, size=11),
                     text_color=C["faint"], anchor="w").pack(fill="x", pady=(8, 0))

    def set_active_nav(self, key):
        """Yan menüde bulunulan sayfayı vurgular."""
        for k, btn in self.nav_buttons.items():
            active = k == key
            btn.configure(fg_color=C["accent_soft"] if active else "transparent",
                          text_color=C["accent"] if active else C["text"],
                          font=ctk.CTkFont(family=FONT, size=13, weight="bold" if active else "normal"))

    def toggle_mode(self):
        mode = "Dark" if self.switch_var.get() == "on" else "Light"
        ctk.set_appearance_mode(mode)
        self.mode_switch.configure(text="Karanlık mod" if mode == "Dark" else "Aydınlık mod")
        self.registry.set_pref("appearance", mode)
        self.apply_tree_style()

    def apply_tree_style(self):
        """Bulgu tablosunun (ttk.Treeview) renklerini aydınlık / karanlık moda uyarlar."""
        i = 1 if ctk.get_appearance_mode() == "Dark" else 0
        bg, fg, head_bg = C["card"][i], C["text"][i], C["card_alt"][i]
        zebra, sel = ("#F9FAFC", "#DCE2FB") if i == 0 else ("#22252D", "#2E3654")
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Bulgu.Treeview", background=bg, fieldbackground=bg, foreground=fg, rowheight=30,
                        borderwidth=0, font=self.tree_font)
        style.configure("Bulgu.Treeview.Heading", background=head_bg, foreground=C["muted"][i], relief="flat",
                        font=self.tree_heading_font, padding=(6, 6), borderwidth=0)
        style.map("Bulgu.Treeview", background=[("selected", sel)], foreground=[("selected", fg)])
        style.layout("Bulgu.Treeview", [("Bulgu.Treeview.treearea", {"sticky": "nswe"})])  # Kenarlık yok
        style.map("Bulgu.Treeview.Heading", background=[("active", C["accent_soft"][i])])
        for panel in (getattr(self, "bulgu_paneli", None),):
            if panel is not None and panel.winfo_exists():
                panel.tree.tag_configure("cift", background=bg)
                panel.tree.tag_configure("tek", background=zebra)
                panel.tree.tag_configure("sorun_yok", foreground=C["faint"][i])
                panel.tree.tag_configure("duzeltme", foreground=C["warning"][i])
                panel.tree.tag_configure("acik", foreground=fg)

    def create_main_frame(self):
        self.main_frame = ctk.CTkFrame(self, corner_radius=0, fg_color=C["bg"])
        self.main_frame.grid(row=0, column=1, sticky="nsew")

    def clear_main_frame(self, nav=None):
        for widget in self.main_frame.winfo_children():
            widget.destroy()
        self._row_parent = None
        self.set_active_nav(nav)

    def create_header(self, title, subtitle, details=None):
        """Sayfa başlığı; uzun açıklama (details) 'Nasıl çalışır?' bağlantısıyla açılıp kapanır."""
        header = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        header.pack(fill="x", padx=36, pady=(28, 14))
        ctk.CTkLabel(header, text=title, font=self.font_title, anchor="w", text_color=C["text"]).pack(fill="x")
        line = ctk.CTkFrame(header, fg_color="transparent")
        line.pack(fill="x", pady=(4, 0))
        ctk.CTkLabel(line, text=subtitle, font=self.font_subtitle, text_color=C["muted"], anchor="w",
                     justify="left", wraplength=880).pack(side="left")
        if details:
            box = ctk.CTkLabel(header, text=details, font=self.font_small, text_color=C["muted"], anchor="w",
                               justify="left", wraplength=920, fg_color=C["card_alt"], corner_radius=10)
            link = ctk.CTkLabel(line, text="ⓘ  Nasıl çalışır?", font=ctk.CTkFont(family=FONT, size=12,
                                                                                    weight="bold"),
                                text_color=C["accent"], cursor="hand2")
            link.pack(side="left", padx=(12, 0))

            def toggle(_e=None):
                if box.winfo_ismapped():
                    box.pack_forget()
                else:
                    box.pack(fill="x", pady=(10, 0), ipadx=12, ipady=10)
            link.bind("<Button-1>", toggle)
        return header

    def create_button_row(self, parent=None):
        frame = ctk.CTkFrame(parent or self._row_parent or self.main_frame, fg_color="transparent")
        inside = (parent or self._row_parent) is not None
        frame.pack(fill="x", padx=(20 if inside else 36), pady=(6 if inside else 8))
        return frame

    def settings_card(self, summary):
        """Katlanabilir 'Kontrol Ayarları' kartı. Kart açıkken create_button_row satırları kartın içine eklenir;
        end_settings_card() ile kapatılır. Kapalıyken yalnızca özet satırı görünür."""
        box = card(self.main_frame)
        box.pack(fill="x", padx=36, pady=(0, 10))
        head = ctk.CTkFrame(box, fg_color="transparent", cursor="hand2")
        head.pack(fill="x", padx=18, pady=12)
        arrow = ctk.CTkLabel(head, text="▸", font=ctk.CTkFont(family=FONT, size=14), text_color=C["muted"],
                             width=16)
        arrow.pack(side="left")
        title = ctk.CTkLabel(head, text="Kontrol ayarları", font=ctk.CTkFont(family=FONT, size=13, weight="bold"),
                             text_color=C["text"])
        title.pack(side="left", padx=(6, 12))
        ozet = ctk.CTkLabel(head, text=summary, font=self.font_small, text_color=C["muted"], anchor="w")
        ozet.pack(side="left", fill="x", expand=True)
        hint = ctk.CTkLabel(head, text="Düzenle", font=ctk.CTkFont(family=FONT, size=12, weight="bold"),
                            text_color=C["accent"])
        hint.pack(side="right")
        body = ctk.CTkFrame(box, fg_color="transparent")

        def toggle(_e=None):
            if body.winfo_ismapped():
                body.pack_forget()
                arrow.configure(text="▸")
                hint.configure(text="Düzenle")
            else:
                body.pack(fill="x", pady=(0, 12))
                arrow.configure(text="▾")
                hint.configure(text="Gizle")
        for w in (head, arrow, title, ozet, hint):
            w.bind("<Button-1>", toggle)
        self._row_parent = body
        self.settings_toggle = toggle
        return box

    def settings_group(self, text):
        """Ayar kartı içinde küçük grup başlığı."""
        ctk.CTkLabel(self._row_parent, text=text.upper(), font=self.font_section, text_color=C["faint"],
                     anchor="w").pack(fill="x", padx=20, pady=(10, 0))

    def ayar_ozeti(self, fiyat=False):
        """Kapalı ayar kartında gösterilen tek satırlık özet."""
        parts = [self.setting("period_type", "Aylık")]
        if fiyat:
            parts.append(f"eşik %{self.setting('threshold', '15')}")
        acc = self.setting("accounts", "")
        parts.append(f"hesaplar {acc}" if acc else "hesap kodu girilmedi")
        parts.append(f"tolerans {self.setting('tolerance', '0.01')} TL / "
                     f"%{self.setting('kur_toleransi', f'{checks.KUR_TOLERANSI_VARSAYILAN:g}')} kur")
        return "  ·  ".join(parts)

    def end_settings_card(self):
        self._row_parent = None

    def add_button(self, parent, text, command, kind="primary", **_legacy):
        """Temalı buton ekler. Eski çağrılardaki fg_color / hover_color yok sayılır; görünümü `kind` belirler."""
        btn = tbutton(parent, text, command, kind=kind, height=40)
        btn.pack(side="left", padx=(0, 10))
        return btn

    def add_labeled_entry(self, parent, label, value, width=120, placeholder=""):
        ctk.CTkLabel(parent, text=label, font=self.font_small, text_color=C["muted"]).pack(side="left",
                                                                                          padx=(0, 6))
        entry = ctk.CTkEntry(parent, width=width, height=34, corner_radius=10, border_width=1,
                             border_color=C["border"], fg_color=C["input"], text_color=C["text"],
                             placeholder_text=placeholder)
        if value not in (None, ""):
            entry.insert(0, str(value))
        entry.pack(side="left", padx=(0, 18))
        return entry

    def add_period_menu(self, parent):
        ctk.CTkLabel(parent, text="Dönem:", font=self.font_small, text_color=C["muted"]).pack(side="left",
                                                                                             padx=(0, 6))
        var = ctk.StringVar(value=self.setting("period_type", "Aylık"))
        ctk.CTkSegmentedButton(parent, values=PERIOD_TYPES, variable=var, height=34, corner_radius=10,
                               font=self.font_small, selected_color=C["accent"], selected_hover_color=C["accent_hover"],
                               unselected_color=C["card_alt"], unselected_hover_color=C["accent_soft"],
                               fg_color=C["card_alt"], text_color=(C["text"][0], C["text"][1]),
                               command=lambda v: self.db.set_setting("period_type", v)) \
            .pack(side="left", padx=(0, 18))
        return var

    def create_console_box(self, parent=None, padx=36, pady=(6, 28), title=None):
        outer = parent or self.main_frame
        if parent is None:
            outer = card(self.main_frame)
            outer.pack(fill="both", expand=True, padx=padx, pady=pady)
            if title:
                ctk.CTkLabel(outer, text=title, font=ctk.CTkFont(family=FONT, size=13, weight="bold"),
                             text_color=C["text"], anchor="w").pack(fill="x", padx=18, pady=(14, 0))
            padx, pady = 10, 10
        box = ctk.CTkTextbox(outer, font=self.font_console, corner_radius=10, wrap="none", border_width=0,
                             fg_color=C["card"], text_color=C["text"])
        box.pack(fill="both", expand=True, padx=padx, pady=pady)
        box.tag_config("hata", foreground=C["danger"][1] if ctk.get_appearance_mode() == "Dark" else C["danger"][0])
        box.tag_config("uyari", foreground=C["warning"][1] if ctk.get_appearance_mode() == "Dark"
                       else C["warning"][0])
        box.tag_config("ok", foreground=C["success"][1] if ctk.get_appearance_mode() == "Dark" else C["success"][0])
        box.tag_config("baslik", foreground=C["accent"][1] if ctk.get_appearance_mode() == "Dark"
                       else C["accent"][0])
        return box

    def create_results_area(self):
        """Sonuç alanı: "Bulgular" sekmesinde inceleme işaretlenebilen tablo, "Özet (metin)" sekmesinde konsol.
        Dönüş: (BulguPaneli, konsol kutusu)"""
        tabs = ctk.CTkTabview(self.main_frame, corner_radius=14, height=260, fg_color="transparent",
                              segmented_button_fg_color=C["card_alt"], segmented_button_selected_color=C["accent"],
                              segmented_button_selected_hover_color=C["accent_hover"],
                              segmented_button_unselected_color=C["card_alt"],
                              segmented_button_unselected_hover_color=C["accent_soft"], text_color=C["text"],
                              anchor="w")
        tabs.pack(fill="both", expand=True, padx=30, pady=(0, 18))
        tabs.add("Bulgular")
        tabs.add("Özet (metin)")
        self.results_tabs = tabs
        self.bulgu_paneli = BulguPaneli(self, tabs.tab("Bulgular"))
        self.bulgu_paneli.pack(fill="both", expand=True)
        self.apply_tree_style()
        wrap = card(tabs.tab("Özet (metin)"))
        wrap.pack(fill="both", expand=True)
        box = self.create_console_box(wrap, padx=10, pady=10)
        return self.bulgu_paneli, box

    def upload_card(self, icon, title, desc):
        """Veri yükleme ekranlarının büyük kartı. Dönüş: butonların ekleneceği satır."""
        box = card(self.main_frame)
        box.pack(fill="x", padx=36, pady=(0, 12))
        ctk.CTkLabel(box, text=icon, width=56, height=56, corner_radius=16, fg_color=C["accent_soft"],
                     text_color=C["accent"], font=ctk.CTkFont(family=FONT, size=24)).pack(side="left", padx=22,
                                                                                         pady=22)
        texts = ctk.CTkFrame(box, fg_color="transparent")
        texts.pack(side="left", fill="x", expand=True, pady=18)
        ctk.CTkLabel(texts, text=title, font=ctk.CTkFont(family=FONT, size=15, weight="bold"),
                     text_color=C["text"], anchor="w").pack(fill="x")
        ctk.CTkLabel(texts, text=desc, font=self.font_small, text_color=C["muted"], anchor="w") \
            .pack(fill="x", pady=(2, 10))
        row = ctk.CTkFrame(texts, fg_color="transparent")
        row.pack(fill="x")
        return row

    def page_actions(self):
        """Sayfanın ana işlem butonları satırı."""
        row = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        row.pack(fill="x", padx=36, pady=(4, 12))
        return row

    def show_results(self, sections):
        """Bulgu tablosunu doldurur; bulgu bölümü yoksa (veri yok uyarısı) metin özetine geçer.

        Not: CTkTabview.set() diğer sekmeleri 100 ms sonra gizler; art arda iki set() çağrısı yapılmamalıdır."""
        self.bulgu_paneli.set_sections(sections)
        self.apply_tree_style()
        hedef = "Bulgular" if self.bulgu_paneli.sections else "Özet (metin)"
        if self.results_tabs.get() != hedef:
            self.results_tabs.set(hedef)

    def log_review_summary(self, box, sections):
        """Her kontrol için bulgu ve inceleme (açık / sorun yok / düzeltme istendi) sayıları."""
        view = inceleme.bolumlere_uygula(sections, inceleme.incelemeleri_oku(self.db))
        ozet = inceleme.inceleme_ozeti(view)
        self.log(box, "ÖZET (bulgu sayısı — inceleme durumu)", "baslik")
        for _, r in ozet.iterrows():
            tag = "ok" if r["Acik"] == 0 and r["Duzeltme_Istendi"] == 0 else ("hata" if r["Acik"] else "uyari")
            self.log(box, f"  {'✔' if tag == 'ok' else '✖'} {r['Kontrol']}: {r['Bulgu_Sayisi']}  "
                          f"({inceleme.ozet_satiri(r)})", tag)

    def log(self, box, text, tag=None):
        box.insert("end", text + "\n", tag)
        box.see("end")

    def log_messages(self, box, messages, tag):
        for msg in messages[:MAX_LOG_LINES]:
            self.log(box, f"  • {msg}", tag)
        if len(messages) > MAX_LOG_LINES:
            self.log(box, f"  ... ve {len(messages) - MAX_LOG_LINES} mesaj daha", tag)

    def log_section(self, box, title, df):
        self.log(box, "-" * 70 + f"\n  {title}  [{len(df)} bulgu]\n" + "-" * 70, "baslik")
        self.log(box, df_to_text(df), "ok" if df.empty else None)

    def confirm_reimport(self, kind, digest, file_name):
        prev = self.db.find_import(kind, digest)
        if not prev:
            return True
        msg = f"'{file_name}' dosyası daha önce yüklenmiş ({prev[1]}).\n\nYine de yüklensin mi?"
        if kind == "YEVMIYE_EXCEL":
            msg += "\n\nDİKKAT: Yevmiye satırları tekrar eklenir ve tutarlar ÇİFT sayılır!"
        return messagebox.askyesno("Mükerrer Dosya", msg, icon="warning")

    def save_template(self, df, default_name):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=default_name,
                                            filetypes=[("Excel Dosyası", "*.xlsx")])
        if path:
            try:
                df.to_excel(path, index=False)
                messagebox.showinfo("Başarılı", f"Şablon kaydedildi:\n{path}")
            except Exception as e:
                messagebox.showerror("Hata", f"Şablon kaydedilemedi:\n{e}")

    def ask_column_mapping(self, kind, raw, res, file_name, message=""):
        """Eşleme penceresini açar ve kullanıcı kapatana kadar bekler. Dönüş: ColumnMapping ya da None."""
        dialog = ColumnMappingDialog(self, kind, raw, res.mapping, file_name, message)
        self.wait_window(dialog)
        return dialog.result

    def read_excel_with_mapping(self, box, kind, reader, data, file_name, force_wizard=False):
        """Excel'i firmanın kayıtlı eşlemesi / otomatik tespit ile okur; gerekirse eşleme penceresini açar.

        Onaylanan eşleme firmanın veritabanına dosyanın başlık imzasıyla kaydedilir. Kullanıcı vazgeçerse None.
        """
        raw = importers.read_raw_excel(data)
        saved = importers.parse_saved_mappings(self.db.get_column_mappings(kind))
        res = reader(raw, saved_mappings=saved)
        if not (force_wizard or res.needs_mapping):
            return res
        message = ""
        if res.needs_mapping:
            message = "Zorunlu sütunlar otomatik bulunamadı.\n" + "\n".join(res.errors)
            self.log(box, "[UYARI] Zorunlu sütunlar otomatik bulunamadı; sütun eşleme penceresi açıldı.", "uyari")
        mapping = self.ask_column_mapping(kind, raw, res, file_name, message)
        if mapping is None:
            if res.needs_mapping:  # Eksik sütun hatası (belge no ayrı sütun açıklaması dahil) gösterilir
                self.log(box, "[YÜKLENMEDİ] Zorunlu sütunlar eşlenmediği için dosya yüklenmedi.", "hata")
                self.log_messages(box, res.errors, "hata")
            else:
                self.log(box, "[İPTAL] Sütun eşleme penceresi kapatıldı, dosya yüklenmedi.", "uyari")
            return None
        res = reader(raw, mapping=mapping)
        if not res.needs_mapping:
            self.db.save_column_mapping(kind, res.mapping.signature, res.mapping.to_json())
            res.infos.append("Eşleme bu firma için kaydedildi; aynı biçimdeki dosyalarda sorulmadan uygulanacak")
        return res

    def log_import_result(self, box, res):
        if res.infos:
            self.log(box, f"\n[BİLGİ] ({len(res.infos)})", "baslik")
            self.log_messages(box, res.infos, "baslik")
        if res.errors:
            self.log(box, f"\n[HATALAR] ({len(res.errors)})", "hata")
            self.log_messages(box, res.errors, "hata")
        if res.warnings:
            self.log(box, f"\n[UYARILAR] ({len(res.warnings)})", "uyari")
            self.log_messages(box, res.warnings, "uyari")

    def stat_card(self, parent, value, label, color="text"):
        box = card(parent)
        ctk.CTkLabel(box, text=value, font=ctk.CTkFont(family=FONT, size=26, weight="bold"), text_color=C[color],
                     anchor="w").pack(fill="x", padx=20, pady=(16, 0))
        ctk.CTkLabel(box, text=label, font=self.font_small, text_color=C["muted"], anchor="w") \
            .pack(fill="x", padx=20, pady=(0, 16))
        return box

    def show_welcome_screen(self):
        if self.db is None:
            return self.show_no_firm_screen(welcome=True)
        self.clear_main_frame("home")
        f = self.firm
        self.create_header(f"Merhaba, {f.title}",
                           "Verileriniz bu bilgisayarda, her firma için ayrı bir veritabanında saklanır."
                           + ("" if f.vkn else "  ·  VKN girilmedi: alıcı VKN ve satış kontrolleri atlanır."))
        counts = self.db.counts()
        invoices = self.db.get_invoices_df()
        n_satis = len(checks.satis_faturalari(invoices, f.vkn)) if not invoices.empty else 0
        n_alis = counts["fatura"] - n_satis

        stats = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        stats.pack(fill="x", padx=36, pady=(0, 14))
        for i, (value, label) in enumerate([
                (f"{n_alis:,}".replace(",", "."), "Alış faturası"),
                (f"{n_satis:,}".replace(",", "."), "Satış faturası"),
                (f"{counts['yevmiye_satiri']:,}".replace(",", "."), "Yevmiye satırı"),
                (f"{counts['dosya']:,}".replace(",", "."), "Yüklenen dosya")]):
            self.stat_card(stats, value, label).grid(row=0, column=i, sticky="ew", padx=(0 if i == 0 else 12, 0))
            stats.grid_columnconfigure(i, weight=1, uniform="stat")

        steps = card(self.main_frame)
        steps.pack(fill="x", padx=36, pady=(0, 14))
        ctk.CTkLabel(steps, text="Denetim akışı", font=ctk.CTkFont(family=FONT, size=15, weight="bold"),
                     text_color=C["text"], anchor="w").pack(fill="x", padx=22, pady=(18, 2))
        ctk.CTkLabel(steps, text="Adımları sırayla tamamlayın; rapor tüm kontrolleri tek seferde çalıştırır.",
                     font=self.font_small, text_color=C["muted"], anchor="w").pack(fill="x", padx=22, pady=(0, 8))
        has_inv, has_jou = counts["fatura"] > 0, counts["yevmiye_satiri"] > 0
        for n, (done, title, desc, actions) in enumerate([
                (has_inv, "Faturaları yükleyin", "e-Fatura / e-Arşiv XML (ZIP veya klasör) ya da Excel",
                 [("XML Yükle", self.show_import_frame), ("Excel Yükle", self.show_excel_import_frame)]),
                (has_jou, "Yevmiye kayıtlarını yükleyin", "Muhasebe programından alınan Excel dökümü",
                 [("Yevmiye Yükle", self.show_journal_import_frame)]),
                (False, "Genel denetim raporunu çalıştırın", "Tüm kontroller, bulgular ve Excel raporu",
                 [("Raporu Aç", self.show_audit_frame)])], start=1):
            row = ctk.CTkFrame(steps, fg_color=C["card_alt"], corner_radius=12)
            row.pack(fill="x", padx=16, pady=4)
            ctk.CTkLabel(row, text="✓" if done else str(n), width=32, height=32, corner_radius=16,
                         fg_color=C["success_soft"] if done else C["accent_soft"],
                         text_color=C["success"] if done else C["accent"],
                         font=ctk.CTkFont(family=FONT, size=13, weight="bold")).pack(side="left", padx=14, pady=12)
            texts = ctk.CTkFrame(row, fg_color="transparent")
            texts.pack(side="left", fill="x", expand=True)
            ctk.CTkLabel(texts, text=title, font=ctk.CTkFont(family=FONT, size=13, weight="bold"),
                         text_color=C["text"], anchor="w").pack(fill="x")
            ctk.CTkLabel(texts, text=desc + ("  ·  tamamlandı" if done else ""), font=self.font_small,
                         text_color=C["muted"], anchor="w").pack(fill="x")
            for i, (label, cmd) in enumerate(reversed(actions)):
                tbutton(row, label, cmd, kind="secondary" if (done or i) else "primary", height=34) \
                    .pack(side="right", padx=(0, 14 if i == 0 else 8))
        row.pack_configure(pady=(4, 16))

    def show_no_firm_screen(self, welcome=False):
        self.clear_main_frame("home" if welcome else None)
        has_firms = bool(self.registry.list_firms())
        title = "Hoş geldiniz" if welcome else "Önce bir firma seçin"
        text = ("Veri yükleme ve analiz ekranları seçili firmanın verileriyle çalışır."
                if has_firms else "Başlamak için ilk firmanızı oluşturun. Her firmanın verileri ayrı saklanır.")
        self.create_header(title, text)
        box = card(self.main_frame)
        box.pack(fill="x", padx=36, pady=(4, 0))
        ctk.CTkLabel(box, text="◫", font=ctk.CTkFont(family=FONT, size=40), text_color=C["accent"]) \
            .pack(pady=(28, 4))
        ctk.CTkLabel(box, text="Firma seçilmedi", font=ctk.CTkFont(family=FONT, size=16, weight="bold"),
                     text_color=C["text"]).pack()
        ctk.CTkLabel(box, text="Denetleyeceğiniz firmayı seçin ya da yeni bir firma ekleyin.", font=self.font_small,
                     text_color=C["muted"]).pack(pady=(2, 14))
        row = ctk.CTkFrame(box, fg_color="transparent")
        row.pack(pady=(0, 28))
        if has_firms:
            tbutton(row, "Firma Seç", self.show_firms_frame, kind="secondary").pack(side="left", padx=6)
        tbutton(row, "＋  Yeni Firma", lambda: self.show_firm_form(None)).pack(side="left", padx=6)

    def show_firms_frame(self):
        self.clear_main_frame("firmalar")
        self.create_header("Firmalar", "Her firmanın faturaları, yevmiye kayıtları ve ayarları ayrı saklanır. "
                                       "Son seçilen firma açılışta otomatik açılır.")
        row = self.page_actions()
        self.add_button(row, "＋  Yeni Firma", lambda: self.show_firm_form(None))
        firms = self.registry.list_firms()
        if not firms:
            ctk.CTkLabel(self.main_frame, text="Henüz kayıtlı firma yok.", font=self.font_label,
                         text_color=C["muted"], anchor="w").pack(fill="x", padx=36, pady=10)
            return
        lst = ctk.CTkScrollableFrame(self.main_frame, fg_color="transparent", corner_radius=0)
        lst.pack(fill="both", expand=True, padx=26, pady=(0, 20))
        for firm in firms:
            active = self.firm is not None and firm.code == self.firm.code
            item = card(lst, border_color=C["accent"] if active else C["border"])
            item.pack(fill="x", padx=10, pady=5)
            ctk.CTkLabel(item, text=initials(firm.title), width=44, height=44, corner_radius=12,
                         fg_color=C["accent_soft"], text_color=C["accent"],
                         font=ctk.CTkFont(family=FONT, size=15, weight="bold")).pack(side="left", padx=16, pady=14)
            texts = ctk.CTkFrame(item, fg_color="transparent")
            texts.pack(side="left", fill="x", expand=True)
            name = ctk.CTkFrame(texts, fg_color="transparent")
            name.pack(fill="x")
            ctk.CTkLabel(name, text=firm.title, font=ctk.CTkFont(family=FONT, size=14, weight="bold"),
                         text_color=C["text"], anchor="w").pack(side="left")
            if active:
                ctk.CTkLabel(name, text="Aktif", font=ctk.CTkFont(family=FONT, size=11, weight="bold"),
                             fg_color=C["accent_soft"], text_color=C["accent"], corner_radius=8, height=20,
                             width=46).pack(side="left", padx=10)
            meta = "  ·  ".join(x for x in [firm.code, f"VKN {firm.vkn}" if firm.vkn else "VKN yok",
                                            firm.sector, (firm.created_at or "")[:10]] if x)
            ctk.CTkLabel(texts, text=meta, font=self.font_small, text_color=C["muted"], anchor="w").pack(fill="x")
            btns = ctk.CTkFrame(item, fg_color="transparent")
            btns.pack(side="right", padx=14)
            small = dict(height=32, width=84)
            tbutton(btns, "Sil", lambda c=firm.code: self.delete_firm(c), kind="danger", **small) \
                .pack(side="right", padx=3)
            tbutton(btns, "Düzenle", lambda c=firm.code: self.show_firm_form(c), kind="secondary", **small) \
                .pack(side="right", padx=3)
            if not active:
                tbutton(btns, "Seç", lambda c=firm.code: self.select_firm(c), **small).pack(side="right", padx=3)

    def show_firm_form(self, code=None):
        """Yeni firma (code=None) ya da mevcut firmayı düzenleme formu."""
        firm = self.registry.get(code) if code else None
        self.clear_main_frame("firmalar")
        if firm:
            self.create_header("Firma bilgilerini düzenle",
                               "Kod değiştirilse de firmanın verileri korunur.")
        else:
            self.create_header("Yeni firma", "Kod kısa ve benzersiz bir addır (ör. ABC_INSAAT).")
        box = card(self.main_frame)
        box.pack(fill="x", padx=36, pady=(0, 10))
        form = ctk.CTkFrame(box, fg_color="transparent")
        form.pack(fill="x", padx=24, pady=20)
        entries = {}
        fields = [("code", "Firma kodu / kısa ad *", firm.code if firm else "", "ör. ABC_INSAAT", ""),
                  ("title", "Unvan *", firm.title if firm else "", "ör. ABC İnşaat Taahhüt A.Ş.", ""),
                  ("vkn", "VKN / TCKN", firm.vkn if firm else "", "10 veya 11 hane",
                   "Satış faturalarını tanımak ve alıcı VKN kontrolü için gerekir.")]
        for r, (key, label, value, placeholder, hint) in enumerate(fields):
            ctk.CTkLabel(form, text=label, font=self.font_small, text_color=C["muted"], anchor="w") \
                .grid(row=r * 2, column=0, sticky="w", pady=(8 if r else 0, 2))
            entry = ctk.CTkEntry(form, width=420, height=38, corner_radius=10, border_width=1,
                                 border_color=C["border"], fg_color=C["input"], text_color=C["text"],
                                 placeholder_text=placeholder)
            if value:
                entry.insert(0, value)
            entry.grid(row=r * 2 + 1, column=0, sticky="w")
            if hint:
                ctk.CTkLabel(form, text=hint, font=ctk.CTkFont(family=FONT, size=11), text_color=C["faint"]) \
                    .grid(row=r * 2 + 1, column=1, sticky="w", padx=14)
            entries[key] = entry
        ctk.CTkLabel(form, text="Sektör", font=self.font_small, text_color=C["muted"], anchor="w") \
            .grid(row=6, column=0, sticky="w", pady=(8, 2))
        sector = ctk.CTkComboBox(form, width=420, height=38, corner_radius=10, border_width=1,
                                 border_color=C["border"], fg_color=C["input"], button_color=C["accent"],
                                 values=SECTORS)
        sector.set(firm.sector if firm else "")
        sector.grid(row=7, column=0, sticky="w")
        entries["sector"] = sector
        self.firm_form_entries = entries
        row = self.page_actions()
        self.add_button(row, "Kaydet", lambda: self.save_firm_form(firm.code if firm else None))
        self.add_button(row, "Vazgeç", self.show_firms_frame, kind="secondary")

    def save_firm_form(self, code=None):
        values = {k: e.get().strip() for k, e in self.firm_form_entries.items()}
        try:
            if code is None:
                firm = self.registry.create(values["code"], values["title"], values["vkn"], values["sector"])
            else:
                firm = self.registry.update(code, new_code=values["code"], title=values["title"],
                                            vkn=values["vkn"], sector=values["sector"])
        except FirmError as e:
            messagebox.showwarning("Uyarı", str(e))
            return
        if code is None:
            messagebox.showinfo("Başarılı", f"'{firm.title}' firması oluşturuldu ve seçildi.")
            self.select_firm(firm.code)
            return
        if self.firm is not None and self.firm.code == code:
            self.firm = firm
            self.refresh_firm_display()
        messagebox.showinfo("Başarılı", "Firma bilgileri kaydedildi.")
        self.show_firms_frame()

    def delete_firm(self, code):
        firm = self.registry.get(code)
        if firm is None:
            return
        counts = self.registry.open_db(code).counts()
        if not messagebox.askyesno(
                "Firmayı Sil",
                f"'{firm.display_name}' firması ve TÜM verileri ({counts['fatura']} fatura, "
                f"{counts['yevmiye_satiri']} yevmiye satırı, analiz ayarları) kalıcı olarak silinecek.\n\n"
                "BU İŞLEM GERİ ALINAMAZ.\n\nDevam edilsin mi?", icon="warning"):
            return
        try:
            self.registry.delete(code)
        except Exception as e:
            messagebox.showerror("Hata", f"Firma silinemedi:\n{e}")
            return
        if self.firm is not None and self.firm.code == firm.code:
            self.close_firm()
        messagebox.showinfo("Tamam", f"'{firm.title}' firması silindi.")
        self.show_firms_frame()

    # ------------------------------------------------------------------ XML YÜKLEME
    @requires_firm
    def show_import_frame(self):
        self.clear_main_frame("xml")
        self.create_header("e-Fatura aktarımı", "e-Fatura / e-Arşiv XML dosyalarını sisteme işleyin.",
                           "Birden fazla XML dosyası, içinde XML bulunan ZIP arşivleri ya da bir klasörün tamamı "
                           "seçilebilir. Daha önce yüklenmiş dosyalar ve aynı tedarikçinin aynı numaralı faturaları "
                           "atlanır. Satıcı VKN'si firma VKN'si olan faturalar satış faturası sayılır.")
        row = self.upload_card("⇪", "XML veya ZIP dosyalarını seçin", "GİB portalından ya da entegratörden "
                               "indirilen gelen / giden faturalar")
        self.add_button(row, "Dosya Seç", self.select_xml_files)
        self.add_button(row, "Klasör Seç", self.select_xml_folder, kind="secondary")
        self.xml_log_box = self.create_console_box(title="İşlem günlüğü")
        self.log(self.xml_log_box, "Hazır. Dosya seçilmesi bekleniyor.")

    def select_xml_files(self):
        paths = filedialog.askopenfilenames(title="UBL-TR XML Seç",
                                            filetypes=[("XML / ZIP", "*.xml *.zip"), ("Tüm Dosyalar", "*.*")])
        if paths:
            self.import_xml_paths(list(paths))

    def select_xml_folder(self):
        folder = filedialog.askdirectory(title="XML Klasörü Seç")
        if not folder:
            return
        paths = []
        for root, _, files in os.walk(folder):
            paths += [os.path.join(root, f) for f in files if f.lower().endswith((".xml", ".zip"))]
        if not paths:
            self.log(self.xml_log_box, "[UYARI] Klasörde XML/ZIP dosyası bulunamadı.", "uyari")
            return
        self.import_xml_paths(sorted(paths))

    def import_xml_paths(self, paths):
        box = self.xml_log_box
        added = dup = failed = skipped = 0
        errors, warnings = [], []
        self.log(box, f"> {len(paths)} dosya işleniyor...")
        for i, path in enumerate(paths, start=1):
            try:
                payloads = list(importers.iter_xml_payloads(path))
            except Exception as e:
                errors.append(f"{os.path.basename(path)}: dosya açılamadı ({e})")
                failed += 1
                continue
            for name, data in payloads:
                digest = importers.file_hash(data)
                if self.db.find_import("XML", digest):
                    skipped += 1
                    continue
                res = importers.read_xml(data, name)
                errors += res.errors
                warnings += res.warnings
                if not res.items:
                    failed += 1
                    continue
                n_added, dups = self.db.save_invoices(res.items, "XML", name, digest)
                added += n_added
                for h in dups:
                    dup += 1
                    warnings.append(f"{name}: {h['invoice_no']} nolu fatura (VKN {h['supplier_vkn']}) zaten kayıtlı")
            if i % 25 == 0:
                self.log(box, f"  ... {i}/{len(paths)}")
                self.update()
        self.log(box, f"[TAMAMLANDI] Eklenen: {added}  |  Mükerrer fatura: {dup}  |  "
                      f"Daha önce yüklenmiş dosya: {skipped}  |  Hatalı: {failed}", "ok")
        if errors:
            self.log(box, f"\n[HATALAR] ({len(errors)})", "hata")
            self.log_messages(box, errors, "hata")
        if warnings:
            self.log(box, f"\n[UYARILAR] ({len(warnings)})", "uyari")
            self.log_messages(box, warnings, "uyari")

    # ------------------------------------------------------------------ EXCEL FATURA YÜKLEME
    @requires_firm
    def show_excel_import_frame(self):
        self.clear_main_frame("excel")
        self.create_header("Fatura aktarımı (Excel)", "XML'i olmayan faturaları Excel şablonuyla yükleyin.",
                           "Zorunlu sütunlar: Fatura_No, Tarih, Tedarikci_VKN, Tedarikci_Ad, Urun_Adi, Miktar, Fiyat\n"
                           "İsteğe bağlı: Birim, Iskonto, KDV_Orani, Para_Birimi, Kur, OTV_Tutari, Fatura_Tipi (SATIS, "
                           "IADE, TEVKIFAT, ISTISNA, IHRACAT ...), Tevkifat_Orani (4/10), Yon (Alış / Satış).  Fiyat "
                           "KDV HARİÇ birim fiyattır. Aynı Fatura_No + VKN'li satırlar tek faturanın kalemleri olarak "
                           "kaydedilir. Satış satırlarında Tedarikci_VKN / Tedarikci_Ad alanlarına müşteri yazılır.")
        row = self.upload_card("▦", "Fatura Excel dosyasını seçin", "Sütunlar otomatik tanınır; tanınmazsa "
                               "eşleme penceresi açılır")
        self.add_button(row, "Excel Dosyası Seç", self.select_and_read_excel)
        self.add_button(row, "Sütunları Eşle", lambda: self.select_and_read_excel(force_wizard=True),
                        kind="secondary")
        self.add_button(row, "Boş Şablon İndir", lambda: self.save_template(importers.invoice_template(),
                                                                           "Fatura_Sablonu.xlsx"), kind="ghost")
        self.excel_log_box = self.create_console_box(title="İşlem günlüğü")
        self.log(self.excel_log_box, "Hazır. Dosya seçilmesi bekleniyor.")

    def select_and_read_excel(self, force_wizard=False):
        file_path = filedialog.askopenfilename(title="Fatura Excel Seç", filetypes=[("Excel Dosyaları", "*.xlsx")])
        if not file_path:
            return
        box, name = self.excel_log_box, os.path.basename(file_path)
        self.log(box, f"> Okunuyor: {name}...")
        self.update()
        try:
            with open(file_path, "rb") as fh:
                data = fh.read()
            digest = importers.file_hash(data)
            if not self.confirm_reimport("FATURA_EXCEL", digest, name):
                self.log(box, "[İPTAL] Dosya yüklenmedi.", "uyari")
                return
            reader = functools.partial(importers.read_invoice_excel, source_file=name, company_vkn=self.firm.vkn)
            res = self.read_excel_with_mapping(box, importers.KIND_FATURA, reader, data, name, force_wizard)
        except Exception as e:
            self.log(box, f"[HATA] Dosya okunamadı: {e}", "hata")
            return
        if res is None:
            return
        added, dups = (self.db.save_invoices(res.items, "FATURA_EXCEL", name, digest) if res.items else (0, []))
        lines = sum(len(lns) for h, lns in res.items if h not in dups)
        self.log(box, f"[TAMAMLANDI] {added} fatura ({lines} kalem) eklendi.  |  Mükerrer: {len(dups)}  |  "
                      f"Hatalı satır: {len(res.errors)}", "ok" if not res.errors else "uyari")
        for h in dups:
            res.warnings.append(f"{h['invoice_no']} nolu fatura (VKN {h['supplier_vkn']}) zaten kayıtlı, atlandı")
        self.log_import_result(box, res)

    # ------------------------------------------------------------------ YEVMİYE YÜKLEME
    @requires_firm
    def show_journal_import_frame(self):
        self.clear_main_frame("yevmiye")
        self.create_header("Yevmiye aktarımı", "Muhasebe programından alınan yevmiye dökümünü yükleyin.",
                           "Zorunlu sütunlar: Tarih, Belge_No, Hesap_Kodu ve (Borc + Alacak) ya da Tutar.  "
                           "İsteğe bağlı: Aciklama.  Borç/Alacak kullanılırsa tutar = Borç − Alacak olarak saklanır.\n"
                           "Başlık satırı ve yaygın sütun adları (Evrak No, Borç Tutarı, Fiş Tarihi ...) otomatik "
                           "bulunur; bulunamazsa sütun eşleme penceresi açılır ve eşleme firma için hatırlanır. "
                           "Belge numarası ayrı bir sütunda olmalıdır (açıklamanın içinden okunmaz).")
        row = self.upload_card("☰", "Yevmiye Excel dosyasını seçin", "Belge numarası ayrı bir sütunda olmalıdır "
                               "(Evrak No / Belge No)")
        self.add_button(row, "Yevmiye Excel Seç", self.select_and_read_journal)
        self.add_button(row, "Sütunları Eşle", lambda: self.select_and_read_journal(force_wizard=True),
                        kind="secondary")
        self.add_button(row, "Boş Şablon İndir", lambda: self.save_template(importers.journal_template(),
                                                                           "Yevmiye_Sablonu.xlsx"), kind="ghost")
        self.journal_log_box = self.create_console_box(title="İşlem günlüğü")
        self.log(self.journal_log_box, "Hazır. Dosya seçilmesi bekleniyor.")

    def select_and_read_journal(self, force_wizard=False):
        file_path = filedialog.askopenfilename(title="Yevmiye Excel Seç", filetypes=[("Excel Dosyaları", "*.xlsx")])
        if not file_path:
            return
        box, name = self.journal_log_box, os.path.basename(file_path)
        self.log(box, f"> Dosya okunuyor: {name}...")
        self.update()
        try:
            with open(file_path, "rb") as fh:
                data = fh.read()
            digest = importers.file_hash(data)
            if not self.confirm_reimport("YEVMIYE_EXCEL", digest, name):
                self.log(box, "[İPTAL] Dosya yüklenmedi.", "uyari")
                return
            res = self.read_excel_with_mapping(box, importers.KIND_YEVMIYE, importers.read_journal_excel, data,
                                               name, force_wizard)
        except Exception as e:
            self.log(box, f"[HATA] Dosya okunamadı: {e}", "hata")
            return
        if res is None:
            return
        saved = self.db.save_journal(res.items, name, digest) if res.items else 0
        self.log(box, f"[TAMAMLANDI] {saved} yevmiye satırı işlendi.  |  Hatalı satır: {len(res.errors)}",
                 "ok" if not res.errors else "uyari")
        self.log_import_result(box, res)

    # ------------------------------------------------------------------ RİSK ANALİZİ
    @requires_firm
    def show_analysis_frame(self):
        self.clear_main_frame("fiyat")
        self.create_header("Fiyat risk analizi", "Aynı ürünün dönem ortalamasından belirgin sapan alış fiyatları.",
                           "Her dönem içinde aynı ürün + birim + para birimi için ağırlıklı ortalama birim fiyat "
                           "(AOBF, belge para biriminde, KDV hariç) hesaplanır; eşiği aşan sapmalar listelenir. İade "
                           "faturaları, (seçiliyse) tevkifatlı faturalar ve adında hariç kelime geçen hizmet / hakediş "
                           "kalemleri analize alınmaz. Dönemde en az alım sayısından az alımı olan üründe sapma "
                           "riskli sayılmaz, bilgi olarak gösterilir.")
        self.settings_card(self.ayar_ozeti(fiyat=True))
        self.settings_group("Genel")
        opts = self.create_button_row()
        self.period_var = self.add_period_menu(opts)
        self.threshold_entry = self.add_labeled_entry(opts, "Sapma eşiği (%)", self.setting("threshold", "15"), 70)
        self.settings_group("Analiz dışı bırakılanlar")
        self.analysis_rules = self.add_fiyat_kural_rows()
        self.end_settings_card()
        row = self.page_actions()
        self.add_button(row, "▶  Analizi Çalıştır", self.run_analysis)
        self.add_button(row, "Excel'e Aktar", self.export_to_excel, kind="secondary")
        self.analysis_panel, self.result_box = self.create_results_area()
        self.log(self.result_box, "Analiz bekleniyor.")

    def run_analysis(self):
        threshold = self.read_threshold(self.threshold_entry)
        kurallar = self.read_fiyat_kurallari(self.analysis_rules)
        if threshold is None or kurallar is None:
            return
        box = self.result_box
        box.delete("1.0", "end")
        self.log(box, "> Analiz motoru başlatıldı...")
        self.update()
        lines = checks.alis_faturalari(self.db.get_lines_df(), self.firm.vkn)  # Satış faturaları analiz edilmez
        res = checks.fiyat_analizi(lines, self.period_var.get(), threshold, kurallar)
        self.analysis_result = res
        self.analysis_df = res.satirlar
        if res.ozet["toplam"] == 0:
            self.analysis_df = self.analysis_result = None
            self.show_results({})
            return self.log(box, "[BİLGİ] Analiz edilecek fatura satırı bulunamadı.", "uyari")
        self.log(box, f"> {checks.fiyat_ozeti_metni(res.ozet)}\n")
        sections = OrderedDict([(f"Fiyat Anomalileri (±%{threshold:g})", res.riskli)])
        self.log_review_summary(box, sections)
        self.log(box, "  Riskli satırlar 'Bulgular' sekmesinde; inceleme durumu orada işaretlenir.\n")
        cols = ["Donem", "Fatura_No", "Tedarikci", "Urun_Adi", "Birim", "Para_Birimi", "Miktar", "Birim_Fiyat",
                "Birim_Fiyat_TL", "AOBF", "Fark_Yuzdesi"]
        bilgi = res.satirlar[res.satirlar["Risk_Durumu"] == checks.RISK_YETERSIZ]
        if not bilgi.empty:
            self.log_section(box, "BİLGİ: EŞİK ÜSTÜ AMA YETERSİZ VERİ (RİSKLİ SAYILMADI)",
                             bilgi[cols + ["Donemdeki_Alim_Sayisi"]])
        self.log(box, f"Toplam {len(res.satirlar)} satır incelendi, {len(res.riskli)} satır riskli.")
        self.show_results(sections)

    def export_to_excel(self):
        if self.analysis_result is None:
            messagebox.showwarning("Uyarı", "Lütfen önce analizi çalıştırın.")
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="Risk_Raporu.xlsx",
                                                 title="Excel Olarak Kaydet", filetypes=[("Excel Dosyası", "*.xlsx")])
        if file_path:
            res = self.analysis_result
            riskli = inceleme.durum_uygula("Fiyat Anomalileri", res.riskli, inceleme.incelemeleri_oku(self.db))
            self._export(file_path, self.with_fiyat_haric_sheet(
                OrderedDict([("Fiyat Analizi Özeti", checks.fiyat_ozeti_df(res.ozet)),
                             ("İnceleme Özeti", inceleme.inceleme_ozeti({"Fiyat Anomalileri": riskli})),
                             ("Riskli Satırlar", riskli), ("Tüm Satırlar", res.satirlar)]), res.haric,
                self.analysis_rules))

    def _export(self, file_path, sections):
        try:
            export_sections(file_path, sections)
            messagebox.showinfo("Başarılı", f"Rapor başarıyla kaydedildi:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Hata", f"Excel kaydedilirken hata oluştu:\n{e}")

    # ------------------------------------------------------------------ MUHASEBE MUTABAKAT
    @requires_firm
    def show_reconciliation_frame(self):
        self.clear_main_frame("mutabakat")
        self.create_header("Muhasebe mutabakatı", "Faturaların yevmiye kayıtlarıyla tutar, hesap, dönem ve "
                           "vergi karşılaştırması.",
                           "Faturaların KDV HARİÇ tutarları, girdiğiniz hesap kodlarındaki yevmiye kayıtlarıyla "
                           "karşılaştırılır. Fatura no ↔ belge no sırasıyla: Tam eşleşme, Seri+Sıra (ABC123, "
                           "ABC-2024-123 gibi kısaltılmış yazımlar) ve son çare olarak tek adaylı Tutar+Tarih "
                           f"(±{checks.TARIH_PENCERESI_GUN} gün, düşük güven) ile eşleştirilir. Hesap kodları önek olarak "
                           "eşleşir (153 → 153.01, 153.02 ...). Aynı belgenin karşı hesaplarını (ör. 153 ile 320) "
                           "birlikte girmeyin; toplamlar birbirini sıfırlar. Faturasız kayıt listesine yalnızca "
                           "borç yönlü belgeler alınır; belge no'su hariç öneklerden biriyle başlayanlar (bordro, "
                           "amortisman, mahsup ...) listelenmez. Alış faturalarının KDV'si (191) ve tevkifatı (360), "
                           "satış faturaları (satıcı VKN'si firma VKN'si olanlar) gelir hesapları ve hesaplanan KDV "
                           "(391) ile ayrıca karşılaştırılır.")
        self.settings_card(self.ayar_ozeti())
        self.settings_group("Genel")
        opts = self.create_button_row()
        self.accounts_entry = self.add_labeled_entry(opts, "Maliyet / stok hesapları", self.setting("accounts", ""),
                                                     220, "ör. 153, 770")
        self.recon_period_var = self.add_period_menu(opts)
        self.tolerance_entry, self.recon_kur_entry = self.add_tolerans_row()
        self.settings_group("Faturasız kayıt kontrolü")
        self.recon_onek_entry, self.recon_haric_var = self.add_haric_onek_row()
        self.settings_group("Vergi ve satış hesapları")
        self.recon_vergi = self.add_vergi_rows()
        self.end_settings_card()
        if not self.setting("accounts", ""):
            self.settings_toggle()  # Hesap kodu girilmeden mutabakat yapılamaz; ayarlar açık gelir
        row = self.page_actions()
        self.add_button(row, "▶  Mutabakatı Çalıştır", self.run_reconciliation)
        self.add_button(row, "Excel'e Aktar", lambda: self.export_sections_dialog(
            self.recon_panel.export_sections(self.with_satis_sheets(self.with_kur_sheet(
                self.with_haric_sheet(self.recon_sections, self.recon_haric, self.recon_haric_var), self.recon_kur),
                self.recon_ek, self.recon_haric_var)),
            "Mutabakat_Raporu.xlsx"), kind="secondary")
        self.recon_panel, self.recon_box = self.create_results_area()
        self.log(self.recon_box, "Mutabakat bekleniyor.")

    def run_reconciliation(self):
        accounts = self.read_accounts(self.accounts_entry)
        tolerance = self.read_tolerance(self.tolerance_entry)
        kur = self.read_kur_toleransi(self.recon_kur_entry) if tolerance is not None else None
        if accounts is None or kur is None:
            return
        box = self.recon_box
        box.delete("1.0", "end")
        self.log(box, "> Veritabanı taranıyor...")
        self.update()
        vergi = self.read_vergi_ayarlari(self.recon_vergi)
        invoices, journal = self.maliyet_vergili_faturalar(vergi), self.db.get_journal_df()
        if invoices.empty or journal.empty:
            self.recon_sections = self.recon_haric = self.recon_kur = None
            self.recon_ek = []
            self.show_results({})
            self.log(box, "[UYARI] İşlem yapılamadı. Hem Fatura hem de Yevmiye kayıtlarının yüklü olduğundan "
                          "emin olun.", "uyari")
            return
        onekler = self.read_haric_onekler(self.recon_onek_entry)
        alis = checks.alis_faturalari(invoices, self.firm.vkn)
        satis = checks.satis_faturalari(invoices, self.firm.vkn)
        res = checks.reconcile(alis, journal, accounts, tolerance, self.recon_period_var.get(), onekler, kur,
                               vergi.kdv_hesaplari, vergi.tevkifat_hesaplari)
        sres = None
        if not satis.empty and vergi.gelir_hesaplari:
            sres = checks.satis_mutabakati(satis, journal, vergi.gelir_hesaplari, vergi.satis_kdv_hesaplari, tolerance,
                                           self.recon_period_var.get(), onekler, kur)
        bulgular = OrderedDict(list(res.items()) + (list(sres.items()) if sres is not None else []))
        head = [("Eşleşme Özeti", checks.eslesme_ozeti_df(res.eslesme_ozeti)),
                ("Faturasız Kayıt Özeti", checks.faturasiz_ozeti_df(res.faturasiz_ozeti))]
        self.recon_ek = []
        if sres is not None:
            head += self.satis_ozet_sayfalari(sres.eslesme_ozeti, sres.faturasiz_ozeti)
            self.recon_ek = self.satis_bilgi_sayfalari(sres.kur_farki, sres.faturasiz_haric)
        self.recon_sections = OrderedDict(head + list(bulgular.items()))
        self.recon_haric, self.recon_kur = res.faturasiz_haric, res.kur_farki
        self.log(box, f"> Hesaplar: {', '.join(accounts)}  |  Tolerans: {tolerance:g} TL  |  Kur toleransı "
                      f"(dövizli): %{kur:g}  |  Alış faturası: {len(alis)}  |  Satış faturası: {len(satis)}")
        self.log(box, f"> {checks.eslesme_ozeti_metni(res.eslesme_ozeti)}")
        self.log(box, f"> {checks.faturasiz_ozeti_metni(res.faturasiz_ozeti)}",
                 None if res.faturasiz_ozeti["isaretli"] else "uyari")
        self.log(box, f"> {checks.kur_ozeti_metni(res.kur_ozeti)}")
        self.log(box, f"> {checks.maliyet_vergisi_metni(checks.maliyet_vergisi_ozeti(alis), vergi.maliyet_kodlari)}")
        self.log(box, f"> {checks.vergi_ozeti_metni(res.vergi_ozeti) or 'Vergi mutabakatı kapalı (hesap kodu yok)'}")
        self.log_satis(box, sres.eslesme_ozeti if sres is not None else None,
                       sres.faturasiz_ozeti if sres is not None else None, sres.kur_ozeti if sres is not None else None,
                       sres.vergi_ozeti if sres is not None else None, len(satis), vergi)
        self.log(box, "")
        self.log_review_summary(box, bulgular)
        self.show_results(bulgular)

    @staticmethod
    def satis_ozet_sayfalari(eslesme, faturasiz):
        return [("Satış Eşleşme Özeti", checks.eslesme_ozeti_df(eslesme)),
                ("Faturasız Gelir Özeti", checks.faturasiz_ozeti_df(faturasiz))]

    @staticmethod
    def satis_bilgi_sayfalari(kur_farki, haric):
        """Satış mutabakatının Excel'e eklenen bilgi sayfaları: (başlık, df, hariç seçeneğine bağlı mı)"""
        return [("Satış Kur Farkı (Tolerans İçi)", kur_farki, False),
                ("Faturasız Gelir Listesinden Hariç Tutulanlar", haric, True)]

    @staticmethod
    def with_satis_sheets(sections, ek, haric_var):
        if not sections or not ek:
            return sections
        return OrderedDict(list(sections.items()) + [(t, df) for t, df, haricli in ek
                                                     if not haricli or (haric_var is not None and haric_var.get())])

    def log_satis(self, box, eslesme, faturasiz, kur_ozeti, kdv_ozeti, n_satis, vergi):
        """Satış mutabakatı özet satırları."""
        if eslesme is None:
            if n_satis and not vergi.gelir_hesaplari:
                self.log(box, "> Satış mutabakatı: gelir hesap kodu girilmediği için atlandı.", "uyari")
            elif not n_satis:
                self.log(box, "> Satış mutabakatı: satış faturası yok (satıcı VKN'si firma VKN'si olan ya da Excel'de "
                              "Yon = Satış verilen fatura bulunamadı)." + ("" if self.firm.vkn else
                                                                         " Firma VKN'si girilmemiş."))
            return
        self.log(box, f"> SATIŞ ({n_satis} fatura; gelir: {', '.join(vergi.gelir_hesaplari)}) — "
                      f"{checks.eslesme_ozeti_metni(eslesme)}", "baslik")
        self.log(box, f"> {checks.faturasiz_ozeti_metni(faturasiz)}")
        self.log(box, f"> Satış {checks.kur_ozeti_metni(kur_ozeti)}")
        self.log(box, f"> {checks.satis_kdv_ozeti_metni(kdv_ozeti) or 'Satış KDV kontrolü kapalı (hesap kodu yok)'}")

    @staticmethod
    def with_kur_sheet(sections, kur_farki):
        """Excel çıktısına yüzde kur toleransı içinde kalan dövizli faturaların bilgi sayfasını ekler."""
        if not sections or kur_farki is None:
            return sections
        return OrderedDict(list(sections.items()) + [("Kur Farkı (Tolerans İçi)", kur_farki)])

    def export_sections_dialog(self, sections, default_name):
        if not sections:
            messagebox.showwarning("Uyarı", "Lütfen önce kontrolü çalıştırın.")
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=default_name,
                                                 title="Excel Olarak Kaydet", filetypes=[("Excel Dosyası", "*.xlsx")])
        if file_path:
            self._export(file_path, sections)

    # ------------------------------------------------------------------ GENEL DENETİM RAPORU
    @requires_firm
    def show_audit_frame(self):
        self.clear_main_frame("rapor")
        self.create_header("Genel denetim raporu", "Tüm kontroller tek tıkla; bulguları inceleyip işaretleyin.",
                           "Tüm kontroller tek seferde çalıştırılır: fiyat anomalileri (alışlar), alış mutabakatı "
                           "(muhasebeleşmemiş, yanlış hesap, tutar farkı — ÖTV gibi maliyete eklenen vergiler dahil —, "
                           "dönem farkı, faturasız kayıt, KDV (191) ve tevkifat (360)), satış mutabakatı (gelir "
                           "hesapları, hesaplanan KDV, faturasız gelir), olası mükerrer faturalar, fatura hesaplama "
                           "tutarsızlıkları ve alıcı VKN kontrolü. Satıcı VKN'si firma VKN'si olan faturalar satıştır.")
        self.settings_card(self.ayar_ozeti(fiyat=True))
        self.settings_group("Genel")
        opts = self.create_button_row()
        self.audit_period_var = self.add_period_menu(opts)
        self.audit_threshold = self.add_labeled_entry(opts, "Sapma eşiği (%)", self.setting("threshold", "15"), 60)
        self.audit_accounts = self.add_labeled_entry(opts, "Maliyet / stok hesapları", self.setting("accounts", ""),
                                                     180, "ör. 153, 770")
        self.audit_tolerance, self.audit_kur = self.add_tolerans_row()
        self.settings_group("Fiyat analizi")
        self.audit_rules = self.add_fiyat_kural_rows()
        self.settings_group("Faturasız kayıt kontrolü")
        self.audit_onek_entry, self.audit_haric_var = self.add_haric_onek_row()
        self.settings_group("Vergi ve satış hesapları")
        self.audit_vergi = self.add_vergi_rows()
        self.end_settings_card()
        row = self.page_actions()
        self.add_button(row, "▶  Tüm Kontrolleri Çalıştır", self.run_full_audit)
        self.add_button(row, "Raporu Excel'e Aktar", lambda: self.export_sections_dialog(
            self.audit_panel.export_sections(self.with_satis_sheets(self.with_fiyat_haric_sheet(
                self.with_kur_sheet(self.with_haric_sheet(self.audit_sections, self.audit_haric, self.audit_haric_var),
                                    self.audit_kur_farki),
                self.audit_fiyat_haric, self.audit_rules), self.audit_ek, self.audit_haric_var)),
            "Denetim_Raporu.xlsx"), kind="secondary")
        self.audit_panel, self.audit_box = self.create_results_area()
        self.log(self.audit_box, "Rapor bekleniyor.")

    def run_full_audit(self):
        threshold = self.read_threshold(self.audit_threshold)
        tolerance = self.read_tolerance(self.audit_tolerance)
        kur = self.read_kur_toleransi(self.audit_kur) if threshold is not None and tolerance is not None else None
        kurallar = self.read_fiyat_kurallari(self.audit_rules) if kur is not None else None
        if kurallar is None:
            return
        accounts = parse_account_list(self.audit_accounts.get())
        self.db.set_setting("accounts", self.audit_accounts.get().strip())
        onekler = self.read_haric_onekler(self.audit_onek_entry)
        vergi = self.read_vergi_ayarlari(self.audit_vergi)
        box = self.audit_box
        box.delete("1.0", "end")
        self.log(box, "> Tüm kontroller çalıştırılıyor...")
        self.update()
        summary, sections, notes, counts = checks.run_full_audit(
            self.db, self.audit_period_var.get(), threshold, accounts, tolerance, self.firm.vkn, onekler, kurallar,
            kur, vergi)
        self.audit_haric = counts["faturasiz_haric"]
        self.audit_fiyat_haric = counts["fiyat_haric"]
        self.audit_kur_farki = counts["kur_farki"]
        st = counts["satis"]
        self.audit_ek = self.satis_bilgi_sayfalari(st["kur_farki"], st["faturasiz_haric"]) if st else []
        if counts["fatura"] == 0:
            self.audit_sections = None
            self.show_results({})
            self.log(box, "[BİLGİ] Veritabanında fatura bulunamadı.", "uyari")
            return
        head = [("Özet", summary), ("Fiyat Analizi Özeti", checks.fiyat_ozeti_df(counts["fiyat_ozeti"]))]
        if counts["eslesme"]:
            head.append(("Eşleşme Özeti", checks.eslesme_ozeti_df(counts["eslesme"])))
            head.append(("Faturasız Kayıt Özeti", checks.faturasiz_ozeti_df(counts["faturasiz_ozeti"])))
        if st:
            head += self.satis_ozet_sayfalari(st["eslesme"], st["faturasiz_ozeti"])
        self.audit_sections = OrderedDict(head + list(sections.items()))
        self.log(box, f"> {counts['fatura']} fatura (alış {counts['alis_fatura']}, satış {counts['satis_fatura']}), "
                      f"{counts['yevmiye']} yevmiye satırı incelendi.")
        self.log(box, f"> {checks.fiyat_ozeti_metni(counts['fiyat_ozeti'])}")
        if counts["eslesme"]:
            self.log(box, f"> {checks.eslesme_ozeti_metni(counts['eslesme'])}")
            self.log(box, f"> {checks.faturasiz_ozeti_metni(counts['faturasiz_ozeti'])}")
            self.log(box, f"> {checks.kur_ozeti_metni(counts['kur_ozeti'])}")
            if counts["vergi_ozeti"]:
                self.log(box, f"> {checks.vergi_ozeti_metni(counts['vergi_ozeti'])}")
        self.log(box, f"> {checks.maliyet_vergisi_metni(counts['maliyet_vergisi'], vergi.maliyet_kodlari)}")
        if st:
            self.log_satis(box, st["eslesme"], st["faturasiz_ozeti"], st["kur_ozeti"], st["vergi_ozeti"], st["fatura"],
                           vergi)
        self.log(box, "")
        self.log_review_summary(box, sections)
        for note in notes:
            self.log(box, f"  ! {note}", "uyari")
        self.show_results(sections)

    # ------------------------------------------------------------------ AYARLAR
    @requires_firm
    def show_settings_frame(self):
        self.clear_main_frame("ayarlar")
        self.create_header("Firma ayarları", "Bu firmaya özel bilgiler, kontrol ayarları ve veriler.")
        page = ctk.CTkScrollableFrame(self.main_frame, fg_color="transparent", corner_radius=0)
        page.pack(fill="both", expand=True, padx=26, pady=(0, 16))

        def section(title, desc):
            box = card(page)
            box.pack(fill="x", padx=10, pady=6)
            ctk.CTkLabel(box, text=title, font=ctk.CTkFont(family=FONT, size=15, weight="bold"),
                         text_color=C["text"], anchor="w").pack(fill="x", padx=20, pady=(16, 0))
            ctk.CTkLabel(box, text=desc, font=self.font_small, text_color=C["muted"], anchor="w", justify="left",
                         wraplength=860).pack(fill="x", padx=20, pady=(2, 6))
            body = ctk.CTkFrame(box, fg_color="transparent")
            body.pack(fill="x", pady=(0, 12))
            self._row_parent = body
            return body

        f = self.firm
        section("Firma bilgileri", "VKN; satış faturalarını tanımak ve alıcı VKN kontrolü için kullanılır.")
        info = self.create_button_row()
        for label, value in [("Kod", f.code), ("Unvan", f.title), ("VKN / TCKN", f.vkn or "girilmedi"),
                             ("Sektör", f.sector or "-"), ("Oluşturulma", (f.created_at or "-")[:10])]:
            col = ctk.CTkFrame(info, fg_color="transparent")
            col.pack(side="left", padx=(0, 28))
            ctk.CTkLabel(col, text=label, font=ctk.CTkFont(family=FONT, size=11), text_color=C["faint"],
                         anchor="w").pack(fill="x")
            ctk.CTkLabel(col, text=value, font=ctk.CTkFont(family=FONT, size=13, weight="bold"),
                         text_color=C["text"], anchor="w").pack(fill="x")
        row = self.create_button_row()
        self.add_button(row, "Bilgileri Düzenle", lambda: self.show_firm_form(self.firm.code), kind="secondary")

        section("Faturasız kayıt kontrolü", "Belge no'su bu öneklerle başlayan kayıtlar (bordro, amortisman, "
                                            "mahsup ...) faturasız kayıt listesine alınmaz.")
        onek_entry, _ = self.add_haric_onek_row(with_excel_option=False)

        def save_onekler():
            n = len(self.read_haric_onekler(onek_entry))
            messagebox.showinfo("Tamam", f"Hariç önek listesi kaydedildi ({n} önek)." if n else
                                "Hariç önek listesi boş kaydedildi; faturasız kontrolde önek filtresi uygulanmayacak.")

        self.add_button(onek_entry.master, "Kaydet", save_onekler)

        section("Vergi ve satış hesapları", "KDV (191), tevkifat (360), gelir (600–602) ve hesaplanan KDV (391) "
                                            "mutabakatında kullanılan hesaplar; maliyete eklenen vergi türleri.")
        vergi_w = self.add_vergi_rows()

        def save_vergi():
            v = self.read_vergi_ayarlari(vergi_w)
            messagebox.showinfo("Tamam", f"Vergi ayarları kaydedildi (KDV hesapları: {', '.join(v.kdv_hesaplari) or '-'}"
                                         f", tevkifat hesapları: {', '.join(v.tevkifat_hesaplari) or '-'}, maliyete "
                                         f"eklenen vergi kodu: {len(v.maliyet_kodlari)}).")

        self.add_button(vergi_w["maliyet"].master, "Kaydet", save_vergi)

        counts = self.db.counts()
        section("Veriler", f"{counts['fatura']} fatura  ·  {counts['fatura_satiri']} fatura satırı  ·  "
                           f"{counts['yevmiye_satiri']} yevmiye satırı  ·  {counts['dosya']} yüklenen dosya")
        row2 = self.create_button_row()
        self.add_button(row2, "Bu Firmanın Tüm Verilerini Sil", self.clear_data, kind="danger")
        ctk.CTkLabel(row2, text="Yalnızca bu firma etkilenir; işlem geri alınamaz.", font=self.font_small,
                     text_color=C["faint"]).pack(side="left", padx=6)
        self._row_parent = None

    def clear_data(self):
        if not messagebox.askyesno("Onay", f"'{self.firm.title}' firmasının tüm fatura ve yevmiye kayıtları "
                                           "silinecek (diğer firmalar etkilenmez). Bu işlem geri alınamaz.\n\n"
                                           "Devam edilsin mi?", icon="warning"):
            return
        self.db.clear_data()
        self.analysis_df = self.recon_sections = self.audit_sections = self.recon_haric = self.audit_haric = None
        self.analysis_result = self.audit_fiyat_haric = self.recon_kur = self.audit_kur_farki = None
        messagebox.showinfo("Tamam", "Veriler silindi.")
        self.show_settings_frame()


if __name__ == "__main__":
    app = AuditApp()
    app.mainloop()
