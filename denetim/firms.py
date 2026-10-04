"""Firma kayıt defteri: her firmanın ayrı SQLite veritabanı ve firma listesi.

Klasör yapısı (data_dir, ör. uygulama klasöründe `veri/`):
    firmalar.db      Firma listesi ve uygulama durumu (son seçilen firma vb.)
    <kod>.db         Firma başına denetim veritabanı (DatabaseManager şeması)
"""
import os
import re
import sqlite3
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime

from .database import DatabaseManager
from .utils import normalize_vkn

REGISTRY_FILE = "firmalar.db"
LEGACY_DB_NAME = "audit_data.db"
DEFAULT_LEGACY_TITLE = "Aktarılan Firma"
DEFAULT_LEGACY_CODE = "AKTARILAN"
MAX_CODE_LEN = 30

REGISTRY_SCHEMA = """
CREATE TABLE IF NOT EXISTS firms (
    code TEXT PRIMARY KEY COLLATE NOCASE,   -- Kısa ad / firma kodu (ör. ABC_INSAAT)
    title TEXT NOT NULL,                    -- Unvan
    vkn TEXT DEFAULT '',                    -- VKN (10) / TCKN (11)
    sector TEXT DEFAULT '',
    db_file TEXT NOT NULL UNIQUE,           -- data_dir içindeki veritabanı dosyası
    created_at TEXT
);
CREATE TABLE IF NOT EXISTS app_state (
    key TEXT PRIMARY KEY,
    value TEXT
);
"""

_TR_ASCII = str.maketrans("çğıöşüÇĞİÖŞÜ", "cgiosuCGIOSU")


class FirmError(ValueError):
    """Kullanıcıya gösterilebilecek firma işlemi hatası."""


@dataclass
class Firm:
    code: str
    title: str
    vkn: str
    sector: str
    db_file: str
    created_at: str

    @property
    def display_name(self):
        return f"{self.title} ({self.code})"


def validate_vkn(value):
    """VKN/TCKN'yi normalize eder; boş bırakılabilir, doluysa 10 ya da 11 hane olmalıdır."""
    vkn = normalize_vkn(value)
    if vkn and len(vkn) not in (10, 11):
        raise FirmError("VKN 10, TCKN 11 haneli olmalıdır.")
    return vkn


def validate_code(value):
    code = re.sub(r"\s+", " ", str(value or "")).strip()
    if not code:
        raise FirmError("Firma kodu / kısa adı boş olamaz.")
    if len(code) > MAX_CODE_LEN:
        raise FirmError(f"Firma kodu en fazla {MAX_CODE_LEN} karakter olabilir.")
    return code


def code_to_filename(code):
    """Firma kodundan güvenli dosya adı kökü üretir (Türkçe karakterler ASCII'ye çevrilir)."""
    stem = re.sub(r"[^A-Za-z0-9_-]+", "_", str(code).translate(_TR_ASCII)).strip("_")
    return stem.upper() or "FIRMA"


class FirmRegistry:
    def __init__(self, data_dir):
        self.data_dir = os.path.abspath(data_dir)
        os.makedirs(self.data_dir, exist_ok=True)
        self.registry_path = os.path.join(self.data_dir, REGISTRY_FILE)
        with self._conn() as conn:
            conn.executescript(REGISTRY_SCHEMA)

    @contextmanager
    def _conn(self):
        conn = sqlite3.connect(self.registry_path)
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    # ------------------------------------------------------------------ okuma
    def list_firms(self):
        with self._conn() as conn:
            rows = conn.execute("SELECT code, title, vkn, sector, db_file, created_at FROM firms "
                                "ORDER BY title COLLATE NOCASE, code COLLATE NOCASE").fetchall()
        return [Firm(*r) for r in rows]

    def get(self, code):
        with self._conn() as conn:
            row = conn.execute("SELECT code, title, vkn, sector, db_file, created_at FROM firms WHERE code = ?",
                               (str(code).strip(),)).fetchone()
        return Firm(*row) if row else None

    def db_path(self, firm):
        return os.path.join(self.data_dir, firm.db_file)

    def open_db(self, code):
        """Firmanın veritabanını açar (yoksa şemayla oluşturur)."""
        firm = self.get(code)
        if firm is None:
            raise FirmError(f"'{code}' kodlu firma bulunamadı.")
        return DatabaseManager(self.db_path(firm))

    # ------------------------------------------------------------------ yazma
    def _unique_db_file(self, conn, code):
        stem = code_to_filename(code)
        used = {r[0].lower() for r in conn.execute("SELECT db_file FROM firms")}
        name, n = f"{stem}.db", 2
        while (name.lower() in used or name == REGISTRY_FILE
               or os.path.exists(os.path.join(self.data_dir, name))):
            name = f"{stem}_{n}.db"
            n += 1
        return name

    def create(self, code, title, vkn="", sector=""):
        code = validate_code(code)
        title = str(title or "").strip()
        if not title:
            raise FirmError("Firma unvanı boş olamaz.")
        vkn = validate_vkn(vkn)
        with self._conn() as conn:
            if conn.execute("SELECT 1 FROM firms WHERE code = ?", (code,)).fetchone():
                raise FirmError(f"'{code}' kodlu bir firma zaten var.")
            db_file = self._unique_db_file(conn, code)
            conn.execute("INSERT INTO firms (code, title, vkn, sector, db_file, created_at) VALUES (?,?,?,?,?,?)",
                         (code, title, vkn, str(sector or "").strip(), db_file,
                          datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        firm = self.get(code)
        DatabaseManager(self.db_path(firm))  # Boş veritabanını şemayla oluştur
        return firm

    def update(self, code, new_code=None, title=None, vkn=None, sector=None):
        """Firma bilgilerini günceller. Kod değişse de veritabanı dosyası aynı kalır."""
        firm = self.get(code)
        if firm is None:
            raise FirmError(f"'{code}' kodlu firma bulunamadı.")
        new_code = firm.code if new_code is None else validate_code(new_code)
        title = firm.title if title is None else str(title).strip()
        if not title:
            raise FirmError("Firma unvanı boş olamaz.")
        vkn = firm.vkn if vkn is None else validate_vkn(vkn)
        sector = firm.sector if sector is None else str(sector).strip()
        with self._conn() as conn:
            if new_code.lower() != firm.code.lower() and \
                    conn.execute("SELECT 1 FROM firms WHERE code = ?", (new_code,)).fetchone():
                raise FirmError(f"'{new_code}' kodlu bir firma zaten var.")
            conn.execute("UPDATE firms SET code = ?, title = ?, vkn = ?, sector = ? WHERE code = ?",
                         (new_code, title, vkn, sector, firm.code))
            conn.execute("UPDATE app_state SET value = ? WHERE key = 'last_firm' AND value = ?",
                         (new_code, firm.code))
        return self.get(new_code)

    def delete(self, code):
        """Firmayı listeden çıkarır ve veritabanı dosyasını KALICI olarak siler."""
        firm = self.get(code)
        if firm is None:
            raise FirmError(f"'{code}' kodlu firma bulunamadı.")
        with self._conn() as conn:
            conn.execute("DELETE FROM firms WHERE code = ?", (firm.code,))
            conn.execute("DELETE FROM app_state WHERE key = 'last_firm' AND value = ?", (firm.code,))
        path = self.db_path(firm)
        for p in (path, path + "-journal", path + "-wal", path + "-shm"):
            if os.path.exists(p):
                os.remove(p)

    # ------------------------------------------------------------------ uygulama durumu
    def _get_state(self, key, default=None):
        with self._conn() as conn:
            row = conn.execute("SELECT value FROM app_state WHERE key = ?", (key,)).fetchone()
        return row[0] if row else default

    def _set_state(self, key, value):
        with self._conn() as conn:
            if value is None:
                conn.execute("DELETE FROM app_state WHERE key = ?", (key,))
            else:
                conn.execute("INSERT OR REPLACE INTO app_state(key, value) VALUES (?, ?)", (key, str(value)))

    def get_last_firm(self):
        """Son seçilen firma (hâlâ kayıtlıysa)."""
        code = self._get_state("last_firm")
        return self.get(code) if code else None

    def set_last_firm(self, code):
        self._set_state("last_firm", code)

    def get_pref(self, key, default=None):
        """Firmadan bağımsız uygulama tercihi (ör. aydınlık / karanlık görünüm)."""
        return self._get_state(f"pref:{key}", default)

    def set_pref(self, key, value):
        self._set_state(f"pref:{key}", value)

    # ------------------------------------------------------------------ eski sürüm (tek veritabanı) aktarımı
    @staticmethod
    def find_legacy_db(*folders):
        """Verilen klasörlerdeki ilk eski `audit_data.db` dosyasını döndürür."""
        for folder in folders:
            path = os.path.join(folder, LEGACY_DB_NAME)
            if os.path.isfile(path):
                return os.path.abspath(path)
        return None

    def legacy_pending(self, legacy_path):
        """Eski veritabanı daha önce aktarılmadıysa True."""
        if not legacy_path or not os.path.isfile(legacy_path):
            return False
        return self._get_state("legacy_imported:" + os.path.realpath(legacy_path)) is None

    @staticmethod
    def read_legacy_company(legacy_path):
        """Eski veritabanının ayarlarındaki (unvan, VKN) bilgisini okur."""
        conn = sqlite3.connect(f"file:{legacy_path}?mode=ro", uri=True)
        try:
            rows = dict(conn.execute("SELECT key, value FROM settings WHERE key IN "
                                     "('company_title', 'company_vkn')").fetchall())
        except sqlite3.Error:
            rows = {}
        finally:
            conn.close()
        return (rows.get("company_title") or "").strip(), normalize_vkn(rows.get("company_vkn"))

    def import_legacy(self, legacy_path, title=None, code=None):
        """Eski tek veritabanını yeni bir firma olarak kopyalar. Eski dosya silinmez, değiştirilmez.

        Unvan verilmezse eski ayarlardaki unvan, o da yoksa "Aktarılan Firma" kullanılır.
        """
        old_title, old_vkn = self.read_legacy_company(legacy_path)
        title = (title or "").strip() or old_title or DEFAULT_LEGACY_TITLE
        try:
            old_vkn = validate_vkn(old_vkn)
        except FirmError:
            old_vkn = ""
        base = validate_code(code) if code else DEFAULT_LEGACY_CODE
        code, n = base, 2
        while self.get(code):
            code = f"{base}_{n}"
            n += 1
        with self._conn() as conn:
            db_file = self._unique_db_file(conn, code)
        target = os.path.join(self.data_dir, db_file)
        src = sqlite3.connect(f"file:{legacy_path}?mode=ro", uri=True)
        dst = sqlite3.connect(target)
        try:
            src.backup(dst)
        finally:
            src.close()
            dst.close()
        try:
            with self._conn() as conn:
                conn.execute("INSERT INTO firms (code, title, vkn, sector, db_file, created_at) "
                             "VALUES (?,?,?,?,?,?)",
                             (code, title, old_vkn, "", db_file, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            DatabaseManager(target)  # V1 şemasıysa kopya üzerinde taşıma yapılır
        except Exception:
            with self._conn() as conn:
                conn.execute("DELETE FROM firms WHERE code = ?", (code,))
            if os.path.exists(target):
                os.remove(target)
            raise
        self._set_state("legacy_imported:" + os.path.realpath(legacy_path), code)
        return self.get(code)
