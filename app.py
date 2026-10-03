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
        self.analysis_df = self.recon_sections = self.audit_sections = None
        self.refresh_firm_display()
        self.show_welcome_screen()

    def close_firm(self):
        self.firm = self.db = None
        self.analysis_df = self.recon_sections = self.audit_sections = None
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
        self.add_button(row, "📄 Boş Şablon İndir", lambda: self.save_template(importers.invoice_template(),
                                                                              "Fatura_Sablonu.xlsx"),
                        fg_color=("gray55", "gray30"), hover_color=("gray45", "gray25"))
        self.excel_log_box = self.create_console_box()
        self.log(self.excel_log_box, "> Sistem hazır. İşlem bekliyor...")

    def select_and_read_excel(self):
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
            res = importers.read_invoice_excel(data, source_file=name)
        except Exception as e:
            self.log(box, f"[HATA] Dosya okunamadı: {e}", "hata")
            return
        added, dups = (self.db.save_invoices(res.items, "FATURA_EXCEL", name, digest) if res.items else (0, []))
        lines = sum(len(lns) for h, lns in res.items if h not in dups)
        self.log(box, f"[TAMAMLANDI] {added} fatura ({lines} kalem) eklendi.  |  Mükerrer: {len(dups)}  |  "
                      f"Hatalı satır: {len(res.errors)}", "ok" if not res.errors else "uyari")
        for h in dups:
            res.warnings.append(f"{h['invoice_no']} nolu fatura (VKN {h['supplier_vkn']}) zaten kayıtlı, atlandı")
        if res.errors:
            self.log(box, f"\n[HATALAR] ({len(res.errors)})", "hata")
            self.log_messages(box, res.errors, "hata")
        if res.warnings:
            self.log(box, f"\n[UYARILAR] ({len(res.warnings)})", "uyari")
            self.log_messages(box, res.warnings, "uyari")

    # ------------------------------------------------------------------ YEVMİYE YÜKLEME
    @requires_firm
    def show_journal_import_frame(self):
        self.clear_main_frame()
        self.create_header("Muhasebe Yevmiye Kayıtları Yükle",
                           "Zorunlu sütunlar: Tarih, Belge_No, Hesap_Kodu ve (Borc + Alacak) ya da Tutar.  "
                           "İsteğe bağlı: Aciklama.  Borç/Alacak kullanılırsa tutar = Borç − Alacak olarak saklanır.")
        row = self.create_button_row()
        self.add_button(row, "Yevmiye Excel Seç", self.select_and_read_journal,
                        fg_color=("#d4a000", "#b58900"), hover_color=("#b58900", "#856500"))
        self.add_button(row, "📄 Boş Şablon İndir", lambda: self.save_template(importers.journal_template(),
                                                                              "Yevmiye_Sablonu.xlsx"),
                        fg_color=("gray55", "gray30"), hover_color=("gray45", "gray25"))
        self.journal_log_box = self.create_console_box()
        self.log(self.journal_log_box, "> Sistem hazır. İşlem bekliyor...")

    def select_and_read_journal(self):
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
            res = importers.read_journal_excel(data)
        except Exception as e:
            self.log(box, f"[HATA] Dosya okunamadı: {e}", "hata")
            return
        saved = self.db.save_journal(res.items, name, digest) if res.items else 0
        self.log(box, f"[TAMAMLANDI] {saved} yevmiye satırı işlendi.  |  Hatalı satır: {len(res.errors)}",
                 "ok" if not res.errors else "uyari")
        if res.errors:
            self.log(box, f"\n[HATALAR] ({len(res.errors)})", "hata")
            self.log_messages(box, res.errors, "hata")
        if res.warnings:
            self.log(box, f"\n[UYARILAR] ({len(res.warnings)})", "uyari")
            self.log_messages(box, res.warnings, "uyari")

    # ------------------------------------------------------------------ RİSK ANALİZİ
    @requires_firm
    def show_analysis_frame(self):
        self.clear_main_frame()
        self.create_header("Fatura Bazlı Risk ve Anomali Analizi",
                           "Her dönem içinde aynı ürün + birim için ağırlıklı ortalama birim fiyat (AOBF, TL, KDV "
                           "hariç) hesaplanır; eşiği aşan sapmalar listelenir. İade faturaları hariç tutulur.")
        opts = self.create_button_row()
        self.period_var = self.add_period_menu(opts)
        self.threshold_entry = self.add_labeled_entry(opts, "Sapma Eşiği (%):", self.setting("threshold", "15"), 80)
        row = self.create_button_row()
        self.add_button(row, "▶ Analizi Çalıştır", self.run_analysis)
        self.add_button(row, "📥 Excel'e Aktar", self.export_to_excel,
                        fg_color=("#388e3c", "#2e7d32"), hover_color=("#2e7d32", "#1b5e20"))
        self.result_box = self.create_console_box()
        self.log(self.result_box, "> Rapor bekleniyor...")

    def run_analysis(self):
        threshold = self.read_threshold(self.threshold_entry)
        if threshold is None:
            return
        box = self.result_box
        box.delete("1.0", "end")
        self.log(box, "> Analiz motoru başlatıldı...")
        self.update()
        df = checks.price_anomalies(self.db.get_lines_df(), self.period_var.get(), threshold)
        if df.empty:
            self.analysis_df = None
            return self.log(box, "[BİLGİ] Analiz edilecek fatura satırı bulunamadı.", "uyari")
        self.analysis_df = df
        risky = df[df["Risk_Durumu"] == "YÜKSEK RİSK"]
        self.log_section(box, f"RİSKLİ FATURA SATIRLARI ({self.period_var.get()} dönem, ±%{threshold:g} sapma)",
                         risky[["Donem", "Fatura_No", "Tedarikci", "Urun_Adi", "Birim", "Miktar", "Birim_Fiyat_TL",
                                "AOBF_TL", "Fark_Yuzdesi"]])
        self.log(box, f"Toplam {len(df)} satır incelendi, {len(risky)} satır riskli.")

    def export_to_excel(self):
        if self.analysis_df is None or self.analysis_df.empty:
            messagebox.showwarning("Uyarı", "Lütfen önce analizi çalıştırın.")
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="Risk_Raporu.xlsx",
                                                 title="Excel Olarak Kaydet", filetypes=[("Excel Dosyası", "*.xlsx")])
        if file_path:
            df = self.analysis_df
            self._export(file_path, OrderedDict([("Riskli Satırlar", df[df["Risk_Durumu"] == "YÜKSEK RİSK"]),
                                                 ("Tüm Satırlar", df)]))

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
                           "birlikte girmeyin; toplamlar birbirini sıfırlar.")
        opts = self.create_button_row()
        self.accounts_entry = self.add_labeled_entry(opts, "Hesap Kodları:", self.setting("accounts", ""), 220,
                                                     "ör. 153, 770")
        self.tolerance_entry = self.add_labeled_entry(opts, "Tolerans (TL):", self.setting("tolerance", "0.01"), 80)
        self.recon_period_var = self.add_period_menu(opts)
        row = self.create_button_row()
        self.add_button(row, "▶ Mutabakat Kontrolü Yap", self.run_reconciliation,
                        fg_color=("#e83e8f", "#d33682"), hover_color=("#d33682", "#a32a65"))
        self.add_button(row, "📥 Excel'e Aktar", lambda: self.export_sections_dialog(self.recon_sections,
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
            self.recon_sections = None
            self.log(box, "[UYARI] İşlem yapılamadı. Hem Fatura hem de Yevmiye kayıtlarının yüklü olduğundan "
                          "emin olun.", "uyari")
            return
        res = checks.reconcile(invoices, journal, accounts, tolerance, self.recon_period_var.get())
        self.recon_sections = OrderedDict([("Eşleşme Özeti", checks.eslesme_ozeti_df(res.eslesme_ozeti))]
                                          + list(res.items()))
        self.log(box, f"> Hesaplar: {', '.join(accounts)}  |  Tolerans: {tolerance:g} TL")
        self.log(box, f"> {checks.eslesme_ozeti_metni(res.eslesme_ozeti)}\n")
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
        row = self.create_button_row()
        self.add_button(row, "▶ Tüm Kontrolleri Çalıştır", self.run_full_audit,
                        fg_color=("#388e3c", "#2e7d32"), hover_color=("#2e7d32", "#1b5e20"))
        self.add_button(row, "📥 Raporu Excel'e Aktar", lambda: self.export_sections_dialog(self.audit_sections,
                                                                                           "Denetim_Raporu.xlsx"))
        self.audit_box = self.create_console_box()
        self.log(self.audit_box, "> Rapor bekleniyor...")

    def run_full_audit(self):
        threshold = self.read_threshold(self.audit_threshold)
        tolerance = self.read_tolerance(self.audit_tolerance)
        if threshold is None or tolerance is None:
            return
        accounts = parse_account_list(self.audit_accounts.get())
        self.db.set_setting("accounts", self.audit_accounts.get().strip())
        box = self.audit_box
        box.delete("1.0", "end")
        self.log(box, "> Tüm kontroller çalıştırılıyor...")
        self.update()
        summary, sections, notes, counts = checks.run_full_audit(
            self.db, self.audit_period_var.get(), threshold, accounts, tolerance, self.firm.vkn)
        if counts["fatura"] == 0:
            self.audit_sections = None
            self.log(box, "[BİLGİ] Veritabanında fatura bulunamadı.", "uyari")
            return
        head = [("Özet", summary)]
        if counts["eslesme"]:
            head.append(("Eşleşme Özeti", checks.eslesme_ozeti_df(counts["eslesme"])))
        self.audit_sections = OrderedDict(head + list(sections.items()))
        self.log(box, f"> {counts['fatura']} fatura, {counts['yevmiye']} yevmiye satırı incelendi.")
        if counts["eslesme"]:
            self.log(box, f"> {checks.eslesme_ozeti_metni(counts['eslesme'])}")
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
                           "Analiz ayarları (eşik, hesap kodları, tolerans, dönem) bu firmaya özel saklanır.")
        info = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        info.pack(fill="x", padx=40, pady=(10, 8))
        f = self.firm
        ctk.CTkLabel(info, font=self.font_label, justify="left", anchor="w",
                     text=f"Firma kodu: {f.code}\nUnvan: {f.title}\nVKN/TCKN: {f.vkn or '(girilmedi)'}\n"
                          f"Sektör: {f.sector or '-'}\nOluşturulma: {f.created_at or '-'}").pack(fill="x")
        row = self.create_button_row()
        self.add_button(row, "✏️  Firma Bilgilerini Düzenle", lambda: self.show_firm_form(self.firm.code))

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
        self.analysis_df = self.recon_sections = self.audit_sections = None
        messagebox.showinfo("Tamam", "Veriler silindi.")
        self.show_settings_frame()


if __name__ == "__main__":
    app = AuditApp()
    app.mainloop()
