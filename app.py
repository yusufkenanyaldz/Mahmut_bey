import functools
import os
from collections import OrderedDict

import customtkinter as ctk
import pandas as pd
from tkinter import filedialog, messagebox

from denetim import checks, importers
from denetim.export import export_sections  # noqa: F401  (eski içe aktarımlar için app.export_sections korunur)
from denetim.firms import FirmError, FirmRegistry
from denetim.utils import parse_account_list, parse_number, tr_upper

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(APP_DIR, "veri")
APP_TITLE = "Finansal Denetim ve Analiz Sistemi | SMMM Modülü"
SECTORS = ["İnşaat", "Halı Üretimi", "Uluslararası Taşımacılık", "Otomotiv Satış ve Kiralama", "Muhtelif İmalat"]
PERIOD_TYPES = ["Aylık", "Çeyreklik", "Yıllık"]
MAX_LOG_LINES = 300

pd.set_option("display.width", 250)
pd.set_option("display.max_columns", 30)
pd.set_option("display.max_colwidth", 40)


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
        ctk.CTkButton(buttons, text="✔ Onayla ve Yükle", command=self.confirm).pack(side="left", padx=(0, 10))
        ctk.CTkButton(buttons, text="İptal", command=self.destroy, fg_color=("gray55", "gray30"),
                      hover_color=("gray45", "gray25")).pack(side="left")
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
        self.analysis_result = self.audit_fiyat_haric = None

        self.title(APP_TITLE)
        self.geometry("1280x800")

        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("blue")

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.font_title = ctk.CTkFont(family="Segoe UI", size=22, weight="bold")
        self.font_subtitle = ctk.CTkFont(family="Segoe UI", size=14)
        self.font_btn = ctk.CTkFont(family="Segoe UI", size=14, weight="bold")
        self.font_label = ctk.CTkFont(family="Segoe UI", size=13)
        self.font_console = ctk.CTkFont(family="Consolas", size=13)

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
        self.analysis_result = self.audit_fiyat_haric = None
        self.refresh_firm_display()
        self.show_welcome_screen()

    def close_firm(self):
        self.firm = self.db = None
        self.analysis_df = self.recon_sections = self.audit_sections = self.recon_haric = self.audit_haric = None
        self.analysis_result = self.audit_fiyat_haric = None
        self.refresh_firm_display()

    def refresh_firm_display(self):
        if self.firm:
            self.title(f"{APP_TITLE} | {self.firm.title}")
            vkn = f"VKN: {self.firm.vkn}" if self.firm.vkn else "VKN girilmedi"
            self.firm_name_label.configure(text=self.firm.title, text_color=("black", "white"))
            self.firm_info_label.configure(text=f"{self.firm.code}  •  {vkn}")
        else:
            self.title(APP_TITLE)
            self.firm_name_label.configure(text="Firma seçilmedi", text_color=("#b58900", "#d4a000"))
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
        entry = self.add_labeled_entry(row, "Faturasız Kontrolde Hariç Önekler:", self.haric_onek_text(), 380,
                                       "boş: önek filtresi yok")

        def reset():
            entry.delete(0, "end")
            entry.insert(0, checks.VARSAYILAN_HARIC_ONEKLER)

        ctk.CTkButton(row, text="Varsayılan", width=90, height=28, command=reset).pack(side="left", padx=(0, 18))
        var = None
        if with_excel_option:
            var = ctk.BooleanVar(value=self.setting("haric_excel", "0") == "1")
            ctk.CTkCheckBox(row, text="Hariç tutulanları Excel'e ekle", variable=var, font=self.font_label,
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
        ctk.CTkCheckBox(row, text="Tevkifatlı faturaları hariç tut", variable=tevkifat, font=self.font_label) \
            .pack(side="left", padx=(0, 18))
        kelime = ctk.BooleanVar(value=self.setting("fiyat_kelime_haric", "1") == "1")
        ctk.CTkCheckBox(row, text="Anahtar kelimeyle hariç tut", variable=kelime, font=self.font_label) \
            .pack(side="left", padx=(0, 18))
        min_alim = self.add_labeled_entry(row, "En Az Alım (dönemde):",
                                          self.setting("fiyat_min_alim", str(checks.FIYAT_MIN_ALIM)), 50)
        excel = ctk.BooleanVar(value=self.setting("fiyat_haric_excel", "1") == "1")
        ctk.CTkCheckBox(row, text="Analiz dışı satırları Excel'e ekle", variable=excel, font=self.font_label,
                        command=lambda: self.db.set_setting("fiyat_haric_excel", "1" if excel.get() else "0")) \
            .pack(side="left")
        row2 = self.create_button_row()
        kelimeler = self.add_labeled_entry(
            row2, "Fiyat Analizinde Hariç Ürün/Hizmet Kelimeleri:",
            self.setting("fiyat_haric_kelimeler", checks.VARSAYILAN_FIYAT_HARIC_KELIMELER), 460,
            "boş: kelime filtresi yok")

        def reset():
            kelimeler.delete(0, "end")
            kelimeler.insert(0, checks.VARSAYILAN_FIYAT_HARIC_KELIMELER)

        ctk.CTkButton(row2, text="Varsayılan", width=90, height=28, command=reset).pack(side="left")
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

    @staticmethod
    def with_fiyat_haric_sheet(sections, haric, w):
        """Excel çıktısına, seçiliyse fiyat analizi dışında kalan satırların bilgi sayfasını ekler."""
        if not sections or haric is None or w is None or not w["excel"].get():
            return sections
        return OrderedDict(list(sections.items()) + [("Fiyat Analizi Dışı Satırlar", haric)])

    # ------------------------------------------------------------------ iskelet
    def create_sidebar(self):
        self.sidebar = ctk.CTkFrame(self, width=250, corner_radius=0, fg_color=("gray85", "#1e1e21"))
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_rowconfigure(10, weight=1)

        logo_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        logo_frame.grid(row=0, column=0, padx=20, pady=(30, 20), sticky="ew")
        ctk.CTkLabel(logo_frame, text="DENETİM", font=ctk.CTkFont(family="Segoe UI", size=24, weight="bold"),
                     text_color=("#1f538d", "#3a7ebf")).pack()
        ctk.CTkLabel(logo_frame, text="Masaüstü Analiz Sistemi", font=ctk.CTkFont(family="Segoe UI", size=12)).pack()

        firm_frame = ctk.CTkFrame(self.sidebar, corner_radius=8, fg_color=("gray78", "#2a2a2e"))
        firm_frame.grid(row=1, column=0, padx=20, pady=(0, 14), sticky="ew")
        ctk.CTkLabel(firm_frame, text="AKTİF FİRMA", font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"),
                     text_color=("gray35", "gray60"), anchor="w").pack(fill="x", padx=12, pady=(8, 0))
        self.firm_name_label = ctk.CTkLabel(firm_frame, text="", font=ctk.CTkFont(family="Segoe UI", size=14,
                                                                                    weight="bold"),
                                            anchor="w", justify="left", wraplength=190)
        self.firm_name_label.pack(fill="x", padx=12)
        self.firm_info_label = ctk.CTkLabel(firm_frame, text="", font=ctk.CTkFont(family="Segoe UI", size=11),
                                            text_color=("gray30", "gray65"), anchor="w", justify="left",
                                            wraplength=190)
        self.firm_info_label.pack(fill="x", padx=12)
        ctk.CTkButton(firm_frame, text="🏢  Firma Seç / Yönet", command=self.show_firms_frame, height=30,
                      font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold")) \
            .pack(fill="x", padx=12, pady=(6, 10))

        def create_nav_button(row, text, command, fg_color=None, hover_color=None):
            btn = ctk.CTkButton(self.sidebar, text=text, command=command, font=self.font_btn,
                                height=42, corner_radius=8, anchor="w",
                                fg_color=fg_color if fg_color else "transparent",
                                text_color=("black", "white") if not fg_color else "white",
                                hover_color=hover_color if hover_color else ("gray70", "#2c2c30"),
                                border_width=1 if not fg_color else 0,
                                border_color=("#3a7ebf", "#3a7ebf") if not fg_color else ("gray85", "#1e1e21"))
            btn.grid(row=row, column=0, padx=20, pady=6, sticky="ew")
            return btn

        create_nav_button(2, "📂  UBL-TR (XML) Yükle", self.show_import_frame)
        create_nav_button(3, "📊  Fatura (Excel) Yükle", self.show_excel_import_frame)
        create_nav_button(4, "📒  Yevmiye (Excel) Yükle", self.show_journal_import_frame,
                          fg_color=("#d4a000", "#b58900"), hover_color=("#b58900", "#856500"))
        create_nav_button(5, "🔍  Fiyat Risk Analizi", self.show_analysis_frame,
                          fg_color=("#2a70bf", "#1f538d"), hover_color=("#1f538d", "#14375e"))
        create_nav_button(6, "⚖️  Muhasebe Mutabakatı", self.show_reconciliation_frame,
                          fg_color=("#e83e8f", "#d33682"), hover_color=("#d33682", "#a32a65"))
        create_nav_button(7, "🧾  Genel Denetim Raporu", self.show_audit_frame,
                          fg_color=("#388e3c", "#2e7d32"), hover_color=("#2e7d32", "#1b5e20"))
        create_nav_button(8, "⚙️  Firma ve Veri Ayarları", self.show_settings_frame)

        self.switch_var = ctk.StringVar(value="on")
        self.mode_switch = ctk.CTkSwitch(self.sidebar, text="Karanlık Mod", command=self.toggle_mode,
                                         variable=self.switch_var, onvalue="on", offvalue="off",
                                         font=ctk.CTkFont(size=12, weight="bold"))
        self.mode_switch.grid(row=11, column=0, padx=20, pady=(10, 5), sticky="s")

        ctk.CTkLabel(self.sidebar, text="V 2.0 (Çevrimdışı)", font=ctk.CTkFont(size=10), text_color="gray") \
            .grid(row=12, column=0, pady=(0, 20), sticky="s")

    def toggle_mode(self):
        if self.switch_var.get() == "on":
            ctk.set_appearance_mode("Dark")
            self.mode_switch.configure(text="Karanlık Mod")
        else:
            ctk.set_appearance_mode("Light")
            self.mode_switch.configure(text="Aydınlık Mod")

    def create_main_frame(self):
        self.main_frame = ctk.CTkFrame(self, corner_radius=12, fg_color=("gray95", "#242427"))
        self.main_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")

    def clear_main_frame(self):
        for widget in self.main_frame.winfo_children():
            widget.destroy()

    def create_header(self, title, subtitle):
        header_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        header_frame.pack(fill="x", padx=40, pady=(30, 15))
        ctk.CTkLabel(header_frame, text=title, font=self.font_title, anchor="w",
                     text_color=("black", "white")).pack(fill="x")
        ctk.CTkLabel(header_frame, text=subtitle, font=self.font_subtitle, text_color=("gray30", "gray70"),
                     anchor="w", justify="left", wraplength=900).pack(fill="x", pady=(5, 0))
        return header_frame

    def create_button_row(self):
        frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        frame.pack(fill="x", padx=40, pady=8)
        return frame

    def add_button(self, parent, text, command, fg_color=None, hover_color=None):
        kwargs = {}
        if fg_color:
            kwargs.update(fg_color=fg_color, hover_color=hover_color)
        btn = ctk.CTkButton(parent, text=text, font=self.font_btn, height=40, command=command, **kwargs)
        btn.pack(side="left", padx=(0, 10))
        return btn

    def add_labeled_entry(self, parent, label, value, width=120, placeholder=""):
        ctk.CTkLabel(parent, text=label, font=self.font_label).pack(side="left", padx=(0, 6))
        entry = ctk.CTkEntry(parent, width=width, placeholder_text=placeholder)
        if value not in (None, ""):
            entry.insert(0, str(value))
        entry.pack(side="left", padx=(0, 18))
        return entry

    def add_period_menu(self, parent):
        ctk.CTkLabel(parent, text="Dönem:", font=self.font_label).pack(side="left", padx=(0, 6))
        var = ctk.StringVar(value=self.setting("period_type", "Aylık"))
        ctk.CTkOptionMenu(parent, values=PERIOD_TYPES, variable=var, width=120,
                          command=lambda v: self.db.set_setting("period_type", v)).pack(side="left", padx=(0, 18))
        return var

    def create_console_box(self):
        box = ctk.CTkTextbox(self.main_frame, font=self.font_console, corner_radius=8, wrap="none",
                             border_width=1, border_color=("gray75", "#3c3c3c"),
                             fg_color=("white", "#18181a"), text_color=("black", "#d4d4d4"))
        box.pack(fill="both", expand=True, padx=40, pady=(10, 30))
        box.tag_config("hata", foreground="#e5534b")
        box.tag_config("uyari", foreground="#d4a000")
        box.tag_config("ok", foreground="#3fb950")
        box.tag_config("baslik", foreground="#3a7ebf")
        return box

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

    def show_welcome_screen(self):
        if self.db is None:
            return self.show_no_firm_screen(welcome=True)
        self.clear_main_frame()
        self.create_header(f"Sisteme Hoş Geldiniz — {self.firm.title}",
                           "Sol menüden yapmak istediğiniz işlemi seçin. Verileriniz yerel diskte (offline) "
                           "ve her firma için ayrı bir veritabanında saklanmaktadır.\n\nÖnerilen akış: Firma Seç → "
                           "Fatura (XML/Excel) → Yevmiye → Genel Denetim Raporu")
        counts = self.db.counts()
        box = self.create_console_box()
        self.log(box, f"> Aktif firma: {self.firm.display_name}" + (f"  |  VKN: {self.firm.vkn}" if self.firm.vkn
                                                                    else "  |  VKN girilmedi (alıcı VKN kontrolü "
                                                                         "atlanır)"))
        self.log(box, f"> Kayıtlı fatura: {counts['fatura']}  |  Fatura satırı: {counts['fatura_satiri']}  |  "
                      f"Yevmiye satırı: {counts['yevmiye_satiri']}  |  Yüklenen dosya: {counts['dosya']}")

    def show_no_firm_screen(self, welcome=False):
        self.clear_main_frame()
        has_firms = bool(self.registry.list_firms())
        title = "Sisteme Hoş Geldiniz" if welcome else "Önce Bir Firma Seçin"
        text = ("Veri yükleme ve analiz ekranları, seçili firmanın verileriyle çalışır. "
                + ("Devam etmek için listeden bir firma seçin ya da yeni bir firma oluşturun."
                   if has_firms else "Başlamak için ilk firmanızı oluşturun. Her firmanın verileri ayrı saklanır."))
        self.create_header(title, text)
        row = self.create_button_row()
        if has_firms:
            self.add_button(row, "🏢  Firma Seç", self.show_firms_frame)
        self.add_button(row, "➕  Yeni Firma", lambda: self.show_firm_form(None),
                        fg_color=("#388e3c", "#2e7d32"), hover_color=("#2e7d32", "#1b5e20"))

    def show_firms_frame(self):
        self.clear_main_frame()
        self.create_header("Firmalar",
                           "Denetlediğiniz firmalar. Her firmanın faturaları, yevmiye kayıtları ve analiz ayarları "
                           "ayrı bir veritabanında tutulur. Son seçilen firma bir sonraki açılışta otomatik açılır.")
        row = self.create_button_row()
        self.add_button(row, "➕  Yeni Firma", lambda: self.show_firm_form(None),
                        fg_color=("#388e3c", "#2e7d32"), hover_color=("#2e7d32", "#1b5e20"))
        firms = self.registry.list_firms()
        if not firms:
            ctk.CTkLabel(self.main_frame, text="Henüz kayıtlı firma yok.", font=self.font_label,
                         text_color=("gray30", "gray70"), anchor="w").pack(fill="x", padx=40, pady=10)
            return
        table = ctk.CTkScrollableFrame(self.main_frame, corner_radius=8, fg_color=("white", "#18181a"))
        table.pack(fill="both", expand=True, padx=40, pady=(10, 30))
        table.grid_columnconfigure(1, weight=1)
        bold = ctk.CTkFont(family="Segoe UI", size=13, weight="bold")
        for col, text in enumerate(["Kod", "Unvan", "VKN/TCKN", "Sektör", "Oluşturulma", ""]):
            ctk.CTkLabel(table, text=text, font=bold, anchor="w").grid(row=0, column=col, padx=8, pady=(6, 4),
                                                                      sticky="w")
        for i, firm in enumerate(firms, start=1):
            active = self.firm is not None and firm.code == self.firm.code
            color = ("#1f538d", "#3a7ebf") if active else ("black", "#d4d4d4")
            values = [firm.code, firm.title + ("  (aktif)" if active else ""), firm.vkn or "-", firm.sector or "-",
                      (firm.created_at or "")[:10]]
            for col, text in enumerate(values):
                ctk.CTkLabel(table, text=text, font=bold if active else self.font_label, text_color=color,
                             anchor="w").grid(row=i, column=col, padx=8, pady=3, sticky="w")
            btns = ctk.CTkFrame(table, fg_color="transparent")
            btns.grid(row=i, column=5, padx=8, pady=3, sticky="e")
            small = dict(height=28, width=80, font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"))
            ctk.CTkButton(btns, text="Seç", state="disabled" if active else "normal",
                          command=lambda c=firm.code: self.select_firm(c), **small).pack(side="left", padx=3)
            ctk.CTkButton(btns, text="Düzenle", fg_color=("gray55", "gray30"), hover_color=("gray45", "gray25"),
                          command=lambda c=firm.code: self.show_firm_form(c), **small).pack(side="left", padx=3)
            ctk.CTkButton(btns, text="Sil", fg_color=("#c62828", "#b71c1c"), hover_color=("#b71c1c", "#7f0000"),
                          command=lambda c=firm.code: self.delete_firm(c), **small).pack(side="left", padx=3)

    def show_firm_form(self, code=None):
        """Yeni firma (code=None) ya da mevcut firmayı düzenleme formu."""
        firm = self.registry.get(code) if code else None
        self.clear_main_frame()
        if firm:
            self.create_header("Firma Bilgilerini Düzenle",
                               "Firma VKN/TCKN'si, alıcısı bu firma olmayan XML faturaların raporlanmasında "
                               "kullanılır. Kod değiştirilse de firmanın verileri korunur.")
        else:
            self.create_header("Yeni Firma",
                               "Firma kodu kısa ve benzersiz bir addır (ör. ABC_INSAAT). VKN/TCKN girilirse alıcı "
                               "VKN kontrolü yapılır. Sektör isteğe bağlıdır.")
        form = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        form.pack(fill="x", padx=40, pady=8)
        entries = {}
        fields = [("code", "Firma Kodu / Kısa Ad: *", firm.code if firm else "", "ör. ABC_INSAAT"),
                  ("title", "Unvan: *", firm.title if firm else "", "ör. ABC İnşaat Taahhüt A.Ş."),
                  ("vkn", "VKN/TCKN:", firm.vkn if firm else "", "10 veya 11 hane")]
        for r, (key, label, value, placeholder) in enumerate(fields):
            ctk.CTkLabel(form, text=label, font=self.font_label, anchor="w").grid(row=r, column=0, padx=(0, 12),
                                                                                  pady=6, sticky="w")
            entry = ctk.CTkEntry(form, width=360, placeholder_text=placeholder)
            if value:
                entry.insert(0, value)
            entry.grid(row=r, column=1, pady=6, sticky="w")
            entries[key] = entry
        ctk.CTkLabel(form, text="Sektör:", font=self.font_label, anchor="w").grid(row=3, column=0, padx=(0, 12),
                                                                                  pady=6, sticky="w")
        sector = ctk.CTkComboBox(form, width=360, values=SECTORS)
        sector.set(firm.sector if firm else "")
        sector.grid(row=3, column=1, pady=6, sticky="w")
        entries["sector"] = sector
        self.firm_form_entries = entries
        row = self.create_button_row()
        self.add_button(row, "💾 Kaydet", lambda: self.save_firm_form(firm.code if firm else None),
                        fg_color=("#388e3c", "#2e7d32"), hover_color=("#2e7d32", "#1b5e20"))
        self.add_button(row, "Vazgeç", self.show_firms_frame,
                        fg_color=("gray55", "gray30"), hover_color=("gray45", "gray25"))

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
        self.clear_main_frame()
        self.create_header("UBL-TR Fatura Aktarımı",
                           "e-Fatura / e-Arşiv XML dosyalarını (veya XML içeren ZIP arşivlerini) seçin. "
                           "Birden fazla dosya seçebilir ya da bir klasörün tamamını aktarabilirsiniz.")
        row = self.create_button_row()
        self.add_button(row, "XML / ZIP Dosyası Seç", self.select_xml_files)
        self.add_button(row, "Klasör Seç", self.select_xml_folder)
        self.xml_log_box = self.create_console_box()
        self.log(self.xml_log_box, "> Sistem hazır. İşlem bekliyor...")

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
        self.clear_main_frame()
        self.create_header("Excel'den Toplu Fatura Aktarımı",
                           "Zorunlu sütunlar: Fatura_No, Tarih, Tedarikci_VKN, Tedarikci_Ad, Urun_Adi, Miktar, Fiyat\n"
                           "İsteğe bağlı: Birim, Iskonto, KDV_Orani, Para_Birimi, Kur.  Fiyat KDV HARİÇ birim fiyattır. "
                           "Aynı Fatura_No + VKN'li satırlar tek faturanın kalemleri olarak kaydedilir.")
        row = self.create_button_row()
        self.add_button(row, "Excel Dosyası Seç", self.select_and_read_excel)
        self.add_button(row, "🧭 Sütunları Eşle", lambda: self.select_and_read_excel(force_wizard=True),
                        fg_color=("#2f7d4f", "#2a6b45"), hover_color=("#25633f", "#1f5034"))
        self.add_button(row, "📄 Boş Şablon İndir", lambda: self.save_template(importers.invoice_template(),
                                                                              "Fatura_Sablonu.xlsx"),
                        fg_color=("gray55", "gray30"), hover_color=("gray45", "gray25"))
        self.excel_log_box = self.create_console_box()
        self.log(self.excel_log_box, "> Sistem hazır. İşlem bekliyor...")

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
            reader = functools.partial(importers.read_invoice_excel, source_file=name)
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
        self.clear_main_frame()
        self.create_header("Muhasebe Yevmiye Kayıtları Yükle",
                           "Zorunlu sütunlar: Tarih, Belge_No, Hesap_Kodu ve (Borc + Alacak) ya da Tutar.  "
                           "İsteğe bağlı: Aciklama.  Borç/Alacak kullanılırsa tutar = Borç − Alacak olarak saklanır.\n"
                           "Başlık satırı ve yaygın sütun adları (Evrak No, Borç Tutarı, Fiş Tarihi ...) otomatik "
                           "bulunur; bulunamazsa sütun eşleme penceresi açılır ve eşleme firma için hatırlanır. "
                           "Belge numarası ayrı bir sütunda olmalıdır (açıklamanın içinden okunmaz).")
        row = self.create_button_row()
        self.add_button(row, "Yevmiye Excel Seç", self.select_and_read_journal,
                        fg_color=("#d4a000", "#b58900"), hover_color=("#b58900", "#856500"))
        self.add_button(row, "🧭 Sütunları Eşle", lambda: self.select_and_read_journal(force_wizard=True),
                        fg_color=("#2f7d4f", "#2a6b45"), hover_color=("#25633f", "#1f5034"))
        self.add_button(row, "📄 Boş Şablon İndir", lambda: self.save_template(importers.journal_template(),
                                                                              "Yevmiye_Sablonu.xlsx"),
                        fg_color=("gray55", "gray30"), hover_color=("gray45", "gray25"))
        self.journal_log_box = self.create_console_box()
        self.log(self.journal_log_box, "> Sistem hazır. İşlem bekliyor...")

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
        self.clear_main_frame()
        self.create_header("Fatura Bazlı Risk ve Anomali Analizi",
                           "Her dönem içinde aynı ürün + birim + para birimi için ağırlıklı ortalama birim fiyat "
                           "(AOBF, belge para biriminde, KDV hariç) hesaplanır; eşiği aşan sapmalar listelenir. İade "
                           "faturaları, (seçiliyse) tevkifatlı faturalar ve adında hariç kelime geçen hizmet / hakediş "
                           "kalemleri analize alınmaz. Dönemde en az alım sayısından az alımı olan üründe sapma "
                           "riskli sayılmaz, bilgi olarak gösterilir.")
        opts = self.create_button_row()
        self.period_var = self.add_period_menu(opts)
        self.threshold_entry = self.add_labeled_entry(opts, "Sapma Eşiği (%):", self.setting("threshold", "15"), 80)
        self.analysis_rules = self.add_fiyat_kural_rows()
        row = self.create_button_row()
        self.add_button(row, "▶ Analizi Çalıştır", self.run_analysis)
        self.add_button(row, "📥 Excel'e Aktar", self.export_to_excel,
                        fg_color=("#388e3c", "#2e7d32"), hover_color=("#2e7d32", "#1b5e20"))
        self.result_box = self.create_console_box()
        self.log(self.result_box, "> Rapor bekleniyor...")

    def run_analysis(self):
        threshold = self.read_threshold(self.threshold_entry)
        kurallar = self.read_fiyat_kurallari(self.analysis_rules)
        if threshold is None or kurallar is None:
            return
        box = self.result_box
        box.delete("1.0", "end")
        self.log(box, "> Analiz motoru başlatıldı...")
        self.update()
        res = checks.fiyat_analizi(self.db.get_lines_df(), self.period_var.get(), threshold, kurallar)
        self.analysis_result = res
        self.analysis_df = res.satirlar
        if res.ozet["toplam"] == 0:
            self.analysis_df = self.analysis_result = None
            return self.log(box, "[BİLGİ] Analiz edilecek fatura satırı bulunamadı.", "uyari")
        self.log(box, f"> {checks.fiyat_ozeti_metni(res.ozet)}\n")
        cols = ["Donem", "Fatura_No", "Tedarikci", "Urun_Adi", "Birim", "Para_Birimi", "Miktar", "Birim_Fiyat",
                "Birim_Fiyat_TL", "AOBF", "Fark_Yuzdesi"]
        self.log_section(box, f"RİSKLİ FATURA SATIRLARI ({self.period_var.get()} dönem, ±%{threshold:g} sapma)",
                         res.riskli[cols])
        bilgi = res.satirlar[res.satirlar["Risk_Durumu"] == checks.RISK_YETERSIZ]
        if not bilgi.empty:
            self.log_section(box, "BİLGİ: EŞİK ÜSTÜ AMA YETERSİZ VERİ (RİSKLİ SAYILMADI)",
                             bilgi[cols + ["Donemdeki_Alim_Sayisi"]])
        self.log(box, f"Toplam {len(res.satirlar)} satır incelendi, {len(res.riskli)} satır riskli.")

    def export_to_excel(self):
        if self.analysis_result is None:
            messagebox.showwarning("Uyarı", "Lütfen önce analizi çalıştırın.")
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="Risk_Raporu.xlsx",
                                                 title="Excel Olarak Kaydet", filetypes=[("Excel Dosyası", "*.xlsx")])
        if file_path:
            res = self.analysis_result
            self._export(file_path, self.with_fiyat_haric_sheet(
                OrderedDict([("Fiyat Analizi Özeti", checks.fiyat_ozeti_df(res.ozet)), ("Riskli Satırlar", res.riskli),
                             ("Tüm Satırlar", res.satirlar)]), res.haric, self.analysis_rules))

    def _export(self, file_path, sections):
        try:
            export_sections(file_path, sections)
            messagebox.showinfo("Başarılı", f"Rapor başarıyla kaydedildi:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Hata", f"Excel kaydedilirken hata oluştu:\n{e}")

    # ------------------------------------------------------------------ MUHASEBE MUTABAKAT
    @requires_firm
    def show_reconciliation_frame(self):
        self.clear_main_frame()
        self.create_header("Fatura ve Yevmiye Mutabakatı",
                           "Faturaların KDV HARİÇ tutarları, girdiğiniz hesap kodlarındaki yevmiye kayıtlarıyla "
                           "karşılaştırılır. Fatura no ↔ belge no sırasıyla: Tam eşleşme, Seri+Sıra (ABC123, "
                           "ABC-2024-123 gibi kısaltılmış yazımlar) ve son çare olarak tek adaylı Tutar+Tarih "
                           f"(±{checks.TARIH_PENCERESI_GUN} gün, düşük güven) ile eşleştirilir. Hesap kodları önek olarak "
                           "eşleşir (153 → 153.01, 153.02 ...). Aynı belgenin karşı hesaplarını (ör. 153 ile 320) "
                           "birlikte girmeyin; toplamlar birbirini sıfırlar. Faturasız kayıt listesine yalnızca "
                           "borç yönlü belgeler alınır; belge no'su hariç öneklerden biriyle başlayanlar (bordro, "
                           "amortisman, mahsup ...) listelenmez.")
        opts = self.create_button_row()
        self.accounts_entry = self.add_labeled_entry(opts, "Hesap Kodları:", self.setting("accounts", ""), 220,
                                                     "ör. 153, 770")
        self.tolerance_entry = self.add_labeled_entry(opts, "Tolerans (TL):", self.setting("tolerance", "0.01"), 80)
        self.recon_period_var = self.add_period_menu(opts)
        self.recon_onek_entry, self.recon_haric_var = self.add_haric_onek_row()
        row = self.create_button_row()
        self.add_button(row, "▶ Mutabakat Kontrolü Yap", self.run_reconciliation,
                        fg_color=("#e83e8f", "#d33682"), hover_color=("#d33682", "#a32a65"))
        self.add_button(row, "📥 Excel'e Aktar", lambda: self.export_sections_dialog(
            self.with_haric_sheet(self.recon_sections, self.recon_haric, self.recon_haric_var),
            "Mutabakat_Raporu.xlsx"),
                        fg_color=("#388e3c", "#2e7d32"), hover_color=("#2e7d32", "#1b5e20"))
        self.recon_box = self.create_console_box()
        self.log(self.recon_box, "> Mutabakat bekleniyor...")

    def run_reconciliation(self):
        accounts = self.read_accounts(self.accounts_entry)
        tolerance = self.read_tolerance(self.tolerance_entry)
        if accounts is None or tolerance is None:
            return
        box = self.recon_box
        box.delete("1.0", "end")
        self.log(box, "> Veritabanı taranıyor...")
        self.update()
        invoices, journal = self.db.get_invoices_df(), self.db.get_journal_df()
        if invoices.empty or journal.empty:
            self.recon_sections = self.recon_haric = None
            self.log(box, "[UYARI] İşlem yapılamadı. Hem Fatura hem de Yevmiye kayıtlarının yüklü olduğundan "
                          "emin olun.", "uyari")
            return
        onekler = self.read_haric_onekler(self.recon_onek_entry)
        res = checks.reconcile(invoices, journal, accounts, tolerance, self.recon_period_var.get(), onekler)
        self.recon_sections = OrderedDict([("Eşleşme Özeti", checks.eslesme_ozeti_df(res.eslesme_ozeti)),
                                           ("Faturasız Kayıt Özeti", checks.faturasiz_ozeti_df(res.faturasiz_ozeti))]
                                          + list(res.items()))
        self.recon_haric = res.faturasiz_haric
        self.log(box, f"> Hesaplar: {', '.join(accounts)}  |  Tolerans: {tolerance:g} TL")
        self.log(box, f"> {checks.eslesme_ozeti_metni(res.eslesme_ozeti)}")
        self.log(box, f"> {checks.faturasiz_ozeti_metni(res.faturasiz_ozeti)}\n",
                 None if res.faturasiz_ozeti["isaretli"] else "uyari")
        for title, df in res.items():
            self.log_section(box, tr_upper(title), df)

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
        self.clear_main_frame()
        self.create_header("Genel Denetim Raporu",
                           "Tüm kontroller tek seferde çalıştırılır: fiyat anomalileri, mutabakat (muhasebeleşmemiş, "
                           "yanlış hesap, tutar farkı, dönem farkı, faturasız kayıt), olası mükerrer faturalar, "
                           "fatura hesaplama tutarsızlıkları ve alıcı VKN kontrolü.")
        opts = self.create_button_row()
        self.audit_period_var = self.add_period_menu(opts)
        self.audit_threshold = self.add_labeled_entry(opts, "Sapma Eşiği (%):", self.setting("threshold", "15"), 70)
        self.audit_accounts = self.add_labeled_entry(opts, "Hesap Kodları:", self.setting("accounts", ""), 180,
                                                     "ör. 153, 770")
        self.audit_tolerance = self.add_labeled_entry(opts, "Tolerans (TL):", self.setting("tolerance", "0.01"), 70)
        self.audit_rules = self.add_fiyat_kural_rows()
        self.audit_onek_entry, self.audit_haric_var = self.add_haric_onek_row()
        row = self.create_button_row()
        self.add_button(row, "▶ Tüm Kontrolleri Çalıştır", self.run_full_audit,
                        fg_color=("#388e3c", "#2e7d32"), hover_color=("#2e7d32", "#1b5e20"))
        self.add_button(row, "📥 Raporu Excel'e Aktar", lambda: self.export_sections_dialog(
            self.with_fiyat_haric_sheet(
                self.with_haric_sheet(self.audit_sections, self.audit_haric, self.audit_haric_var),
                self.audit_fiyat_haric, self.audit_rules),
            "Denetim_Raporu.xlsx"))
        self.audit_box = self.create_console_box()
        self.log(self.audit_box, "> Rapor bekleniyor...")

    def run_full_audit(self):
        threshold = self.read_threshold(self.audit_threshold)
        tolerance = self.read_tolerance(self.audit_tolerance)
        kurallar = self.read_fiyat_kurallari(self.audit_rules) if threshold is not None and tolerance is not None \
            else None
        if kurallar is None:
            return
        accounts = parse_account_list(self.audit_accounts.get())
        self.db.set_setting("accounts", self.audit_accounts.get().strip())
        onekler = self.read_haric_onekler(self.audit_onek_entry)
        box = self.audit_box
        box.delete("1.0", "end")
        self.log(box, "> Tüm kontroller çalıştırılıyor...")
        self.update()
        summary, sections, notes, counts = checks.run_full_audit(
            self.db, self.audit_period_var.get(), threshold, accounts, tolerance, self.firm.vkn, onekler, kurallar)
        self.audit_haric = counts["faturasiz_haric"]
        self.audit_fiyat_haric = counts["fiyat_haric"]
        if counts["fatura"] == 0:
            self.audit_sections = None
            self.log(box, "[BİLGİ] Veritabanında fatura bulunamadı.", "uyari")
            return
        head = [("Özet", summary), ("Fiyat Analizi Özeti", checks.fiyat_ozeti_df(counts["fiyat_ozeti"]))]
        if counts["eslesme"]:
            head.append(("Eşleşme Özeti", checks.eslesme_ozeti_df(counts["eslesme"])))
            head.append(("Faturasız Kayıt Özeti", checks.faturasiz_ozeti_df(counts["faturasiz_ozeti"])))
        self.audit_sections = OrderedDict(head + list(sections.items()))
        self.log(box, f"> {counts['fatura']} fatura, {counts['yevmiye']} yevmiye satırı incelendi.")
        self.log(box, f"> {checks.fiyat_ozeti_metni(counts['fiyat_ozeti'])}")
        if counts["eslesme"]:
            self.log(box, f"> {checks.eslesme_ozeti_metni(counts['eslesme'])}")
            self.log(box, f"> {checks.faturasiz_ozeti_metni(counts['faturasiz_ozeti'])}")
        self.log(box, "")
        self.log(box, "ÖZET", "baslik")
        for _, r in summary.iterrows():
            tag = "ok" if r["Bulgu_Sayisi"] == 0 else "hata"
            self.log(box, f"  {'✔' if r['Bulgu_Sayisi'] == 0 else '✖'} {r['Kontrol']}: {r['Bulgu_Sayisi']}", tag)
        for note in notes:
            self.log(box, f"  ! {note}", "uyari")
        self.log(box, "")
        for title, df in sections.items():
            if not df.empty:
                self.log_section(box, tr_upper(title), df)

    # ------------------------------------------------------------------ AYARLAR
    @requires_firm
    def show_settings_frame(self):
        self.clear_main_frame()
        self.create_header("Firma ve Veri Ayarları",
                           "Firma VKN'si girilirse, alıcısı bu firma olmayan XML faturalar raporlanır. "
                           "Analiz ayarları (eşik, hesap kodları, tolerans, dönem, faturasız kontrolde hariç tutulan "
                           "belge no önekleri, fiyat analizinin hariç tutma kuralları ve kelimeleri) bu firmaya özel "
                           "saklanır.")
        info = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        info.pack(fill="x", padx=40, pady=(10, 8))
        f = self.firm
        ctk.CTkLabel(info, font=self.font_label, justify="left", anchor="w",
                     text=f"Firma kodu: {f.code}\nUnvan: {f.title}\nVKN/TCKN: {f.vkn or '(girilmedi)'}\n"
                          f"Sektör: {f.sector or '-'}\nOluşturulma: {f.created_at or '-'}").pack(fill="x")
        row = self.create_button_row()
        self.add_button(row, "✏️  Firma Bilgilerini Düzenle", lambda: self.show_firm_form(self.firm.code))

        onek_entry, _ = self.add_haric_onek_row(with_excel_option=False)

        def save_onekler():
            n = len(self.read_haric_onekler(onek_entry))
            messagebox.showinfo("Tamam", f"Hariç önek listesi kaydedildi ({n} önek)." if n else
                                "Hariç önek listesi boş kaydedildi; faturasız kontrolde önek filtresi uygulanmayacak.")

        self.add_button(onek_entry.master, "💾 Kaydet", save_onekler)

        counts = self.db.counts()
        info2 = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        info2.pack(fill="x", padx=40, pady=(25, 8))
        ctk.CTkLabel(info2, font=self.font_label, justify="left", anchor="w",
                     text=f"Kayıtlı fatura: {counts['fatura']}    Fatura satırı: {counts['fatura_satiri']}    "
                          f"Yevmiye satırı: {counts['yevmiye_satiri']}    Yüklenen dosya: {counts['dosya']}") \
            .pack(fill="x")
        row2 = self.create_button_row()
        self.add_button(row2, "🗑  Tüm Verileri Sil", self.clear_data,
                        fg_color=("#c62828", "#b71c1c"), hover_color=("#b71c1c", "#7f0000"))

    def clear_data(self):
        if not messagebox.askyesno("Onay", f"'{self.firm.title}' firmasının tüm fatura ve yevmiye kayıtları "
                                           "silinecek (diğer firmalar etkilenmez). Bu işlem geri alınamaz.\n\n"
                                           "Devam edilsin mi?", icon="warning"):
            return
        self.db.clear_data()
        self.analysis_df = self.recon_sections = self.audit_sections = self.recon_haric = self.audit_haric = None
        self.analysis_result = self.audit_fiyat_haric = None
        messagebox.showinfo("Tamam", "Veriler silindi.")
        self.show_settings_frame()


if __name__ == "__main__":
    app = AuditApp()
    app.mainloop()
