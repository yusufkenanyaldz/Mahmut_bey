import os
import re
from collections import OrderedDict

import customtkinter as ctk
import pandas as pd
from tkinter import filedialog, messagebox

from denetim import checks, importers
from denetim.database import DatabaseManager
from denetim.utils import normalize_vkn, parse_account_list, parse_number, tr_upper

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


def export_sections(file_path, sections):
    """OrderedDict(başlık → DataFrame) yapısını çok sayfalı Excel'e yazar."""
    used = set()
    with pd.ExcelWriter(file_path) as writer:
        for title, df in sections.items():
            name = re.sub(r"[\[\]:*?/\\]", "", title)[:31] or "Sayfa"
            base, n = name, 2
            while name in used:
                suffix = f" {n}"
                name = base[:31 - len(suffix)] + suffix
                n += 1
            used.add(name)
            df.to_excel(writer, sheet_name=name, index=False)


# --- ARAYÜZ (GUI) - AYDINLIK/KARANLIK MOD DESTEKLİ ---
class AuditApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.db = DatabaseManager()
        self.analysis_df = None
        self.recon_sections = None
        self.audit_sections = None

        self.title("Finansal Denetim ve Analiz Sistemi | SMMM Modülü")
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
        self.sidebar.grid_rowconfigure(9, weight=1)

        logo_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        logo_frame.grid(row=0, column=0, padx=20, pady=(30, 20), sticky="ew")
        ctk.CTkLabel(logo_frame, text="DENETİM", font=ctk.CTkFont(family="Segoe UI", size=24, weight="bold"),
                     text_color=("#1f538d", "#3a7ebf")).pack()
        ctk.CTkLabel(logo_frame, text="Masaüstü Analiz Sistemi", font=ctk.CTkFont(family="Segoe UI", size=12)).pack()

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

        create_nav_button(1, "📂  UBL-TR (XML) Yükle", self.show_import_frame)
        create_nav_button(2, "📊  Fatura (Excel) Yükle", self.show_excel_import_frame)
        create_nav_button(3, "📒  Yevmiye (Excel) Yükle", self.show_journal_import_frame,
                          fg_color=("#d4a000", "#b58900"), hover_color=("#b58900", "#856500"))
        create_nav_button(4, "🔍  Fiyat Risk Analizi", self.show_analysis_frame,
                          fg_color=("#2a70bf", "#1f538d"), hover_color=("#1f538d", "#14375e"))
        create_nav_button(5, "⚖️  Muhasebe Mutabakatı", self.show_reconciliation_frame,
                          fg_color=("#e83e8f", "#d33682"), hover_color=("#d33682", "#a32a65"))
        create_nav_button(6, "🧾  Genel Denetim Raporu", self.show_audit_frame,
                          fg_color=("#388e3c", "#2e7d32"), hover_color=("#2e7d32", "#1b5e20"))
        create_nav_button(7, "⚙️  Firma ve Veri Ayarları", self.show_settings_frame)

        self.switch_var = ctk.StringVar(value="on")
        self.mode_switch = ctk.CTkSwitch(self.sidebar, text="Karanlık Mod", command=self.toggle_mode,
                                         variable=self.switch_var, onvalue="on", offvalue="off",
                                         font=ctk.CTkFont(size=12, weight="bold"))
        self.mode_switch.grid(row=10, column=0, padx=20, pady=(10, 5), sticky="s")

        ctk.CTkLabel(self.sidebar, text="V 2.0 (Çevrimdışı)", font=ctk.CTkFont(size=10), text_color="gray") \
            .grid(row=11, column=0, pady=(0, 20), sticky="s")

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
        self.show_welcome_screen()

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
        self.clear_main_frame()
        self.create_header("Sisteme Hoş Geldiniz",
                           "Sol menüden yapmak istediğiniz işlemi seçin. Verileriniz yerel diskte (offline) "
                           "saklanmaktadır.\n\nÖnerilen akış: Firma Ayarları → Fatura (XML/Excel) → Yevmiye → "
                           "Genel Denetim Raporu")
        counts = self.db.counts()
        box = self.create_console_box()
        self.log(box, f"> Kayıtlı fatura: {counts['fatura']}  |  Fatura satırı: {counts['fatura_satiri']}  |  "
                      f"Yevmiye satırı: {counts['yevmiye_satiri']}  |  Yüklenen dosya: {counts['dosya']}")

    # ------------------------------------------------------------------ XML YÜKLEME
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
    def show_reconciliation_frame(self):
        self.clear_main_frame()
        self.create_header("Fatura ve Yevmiye Mutabakatı",
                           "Faturaların KDV HARİÇ tutarları, girdiğiniz hesap kodlarındaki yevmiye kayıtlarıyla "
                           "Belge_No = Fatura_No eşleşmesi üzerinden karşılaştırılır. Hesap kodları önek olarak "
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
        self.recon_sections = checks.reconcile(invoices, journal, accounts, tolerance, self.recon_period_var.get())
        self.log(box, f"> Hesaplar: {', '.join(accounts)}  |  Tolerans: {tolerance:g} TL\n")
        for title, df in self.recon_sections.items():
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
            self.db, self.audit_period_var.get(), threshold, accounts, tolerance, self.setting("company_vkn", ""))
        if counts["fatura"] == 0:
            self.audit_sections = None
            self.log(box, "[BİLGİ] Veritabanında fatura bulunamadı.", "uyari")
            return
        self.audit_sections = OrderedDict([("Özet", summary)] + list(sections.items()))
        self.log(box, f"> {counts['fatura']} fatura, {counts['yevmiye']} yevmiye satırı incelendi.\n")
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
    def show_settings_frame(self):
        self.clear_main_frame()
        self.create_header("Firma ve Veri Ayarları",
                           "Firma VKN'si girilirse, alıcısı bu firma olmayan XML faturalar raporlanır.")
        row = self.create_button_row()
        self.company_vkn_entry = self.add_labeled_entry(row, "Firma VKN/TCKN:", self.setting("company_vkn"), 140)
        self.company_title_entry = self.add_labeled_entry(row, "Unvan:", self.setting("company_title"), 300)
        self.add_button(row, "💾 Kaydet", self.save_company)

        counts = self.db.counts()
        info = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        info.pack(fill="x", padx=40, pady=(25, 8))
        ctk.CTkLabel(info, font=self.font_label, justify="left", anchor="w",
                     text=f"Kayıtlı fatura: {counts['fatura']}    Fatura satırı: {counts['fatura_satiri']}    "
                          f"Yevmiye satırı: {counts['yevmiye_satiri']}    Yüklenen dosya: {counts['dosya']}") \
            .pack(fill="x")
        row2 = self.create_button_row()
        self.add_button(row2, "🗑  Tüm Verileri Sil", self.clear_data,
                        fg_color=("#c62828", "#b71c1c"), hover_color=("#b71c1c", "#7f0000"))

    def save_company(self):
        vkn = normalize_vkn(self.company_vkn_entry.get())
        if vkn and len(vkn) not in (10, 11):
            messagebox.showwarning("Uyarı", "VKN 10, TCKN 11 haneli olmalıdır.")
            return
        self.db.set_setting("company_vkn", vkn)
        self.db.set_setting("company_title", self.company_title_entry.get().strip())
        messagebox.showinfo("Başarılı", "Firma bilgileri kaydedildi.")

    def clear_data(self):
        if not messagebox.askyesno("Onay", "Tüm fatura ve yevmiye kayıtları silinecek. Bu işlem geri alınamaz.\n\n"
                                           "Devam edilsin mi?", icon="warning"):
            return
        self.db.clear_data()
        self.analysis_df = self.recon_sections = self.audit_sections = None
        messagebox.showinfo("Tamam", "Veriler silindi.")
        self.show_settings_frame()


if __name__ == "__main__":
    app = AuditApp()
    app.mainloop()
