"""SQLite veritabanı katmanı."""
import sqlite3
from contextlib import contextmanager
from datetime import datetime

import pandas as pd

from .utils import normalize_account, normalize_doc_no, normalize_text

SCHEMA_VERSION = 2

SCHEMA = """
CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT
);
CREATE TABLE IF NOT EXISTS imports (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    kind TEXT NOT NULL,              -- XML / FATURA_EXCEL / YEVMIYE_EXCEL
    file_name TEXT,
    file_hash TEXT NOT NULL,
    imported_at TEXT,
    record_count INTEGER,
    UNIQUE(kind, file_hash)
);
CREATE TABLE IF NOT EXISTS invoices (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    invoice_no TEXT NOT NULL,
    invoice_no_norm TEXT NOT NULL,
    issue_date TEXT NOT NULL,
    supplier_vkn TEXT NOT NULL,
    supplier_name TEXT,
    customer_vkn TEXT,
    customer_name TEXT,
    invoice_type TEXT,               -- SATIS, IADE, TEVKIFAT, ISTISNA ...
    profile TEXT,                    -- TEMELFATURA, TICARIFATURA, EARSIVFATURA ...
    currency TEXT DEFAULT 'TRY',
    exchange_rate REAL DEFAULT 1,
    line_extension_amount REAL,      -- Satır net toplamları (belgede yazan)
    allowance_total REAL DEFAULT 0,  -- Belge geneli iskonto
    total_amount REAL NOT NULL,      -- KDV HARİÇ net tutar (belge para biriminde)
    vat_amount REAL DEFAULT 0,       -- KDV toplamı
    payable_amount REAL,             -- Ödenecek tutar
    source TEXT,                     -- XML / EXCEL
    source_file TEXT,
    import_id INTEGER,
    UNIQUE(supplier_vkn, invoice_no_norm)
);
CREATE TABLE IF NOT EXISTS invoice_lines (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    invoice_id INTEGER NOT NULL,
    line_no INTEGER,
    item_name TEXT,
    item_norm TEXT,
    quantity REAL,
    uom TEXT,
    unit_price REAL,                 -- Liste birim fiyatı (iskonto öncesi)
    line_net REAL,                   -- Satır net tutarı (iskonto sonrası, KDV hariç)
    unit_price_net REAL,             -- line_net / quantity
    vat_rate REAL,
    vat_amount REAL,
    FOREIGN KEY(invoice_id) REFERENCES invoices(id) ON DELETE CASCADE
);
CREATE TABLE IF NOT EXISTS journal_entries (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entry_date TEXT,
    document_no TEXT,
    document_no_norm TEXT,
    account_code TEXT,
    account_norm TEXT,
    amount REAL,                     -- Borç - Alacak (ya da şablondaki Tutar)
    description TEXT,
    source_row INTEGER,
    import_id INTEGER
);
CREATE INDEX IF NOT EXISTS ix_lines_invoice ON invoice_lines(invoice_id);
CREATE INDEX IF NOT EXISTS ix_journal_doc ON journal_entries(document_no_norm);
CREATE INDEX IF NOT EXISTS ix_invoice_no ON invoices(invoice_no_norm);
"""


class DatabaseManager:
    def __init__(self, db_name="audit_data.db"):
        self.db_name = db_name
        self.setup_database()

    @contextmanager
    def connection(self):
        conn = sqlite3.connect(self.db_name)
        conn.execute("PRAGMA foreign_keys = ON")
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    # ------------------------------------------------------------------ kurulum
    def setup_database(self):
        with self.connection() as conn:
            if self._is_legacy_schema(conn):
                self._migrate_v1(conn)
            conn.executescript(SCHEMA)
            conn.execute("INSERT OR REPLACE INTO settings(key, value) VALUES ('schema_version', ?)",
                         (str(SCHEMA_VERSION),))

    @staticmethod
    def _columns(conn, table):
        return {row[1] for row in conn.execute(f"PRAGMA table_info({table})")}

    def _is_legacy_schema(self, conn):
        cols = self._columns(conn, "invoices")
        return bool(cols) and "invoice_no_norm" not in cols

    def _migrate_v1(self, conn):
        """V1.1 veritabanındaki verileri yeni şemaya taşır, eski tabloları yedek olarak saklar."""
        old_inv = pd.read_sql_query("SELECT * FROM invoices", conn)
        old_lines = pd.read_sql_query("SELECT * FROM invoice_lines", conn)
        old_jou = pd.read_sql_query("SELECT * FROM journal_entries", conn) \
            if self._columns(conn, "journal_entries") else pd.DataFrame()
        for table in ("invoices", "invoice_lines", "journal_entries"):
            if self._columns(conn, table):
                conn.execute(f"DROP TABLE IF EXISTS {table}_v1_yedek")
                conn.execute(f"ALTER TABLE {table} RENAME TO {table}_v1_yedek")
        conn.executescript(SCHEMA)

        id_map = {}
        for _, r in old_inv.iterrows():
            no_norm = normalize_doc_no(r["invoice_no"])
            vkn = str(r["supplier_vkn"] or "")
            cur = conn.execute(
                "INSERT OR IGNORE INTO invoices (invoice_no, invoice_no_norm, issue_date, supplier_vkn, supplier_name,"
                " currency, exchange_rate, line_extension_amount, total_amount, source) VALUES (?,?,?,?,?,?,1,?,?,'V1')",
                (str(r["invoice_no"]), no_norm, str(r["issue_date"])[:10], vkn, r["supplier_name"],
                 r["currency"] or "TRY", r["total_amount"], r["total_amount"]))
            if cur.rowcount:
                id_map[r["id"]] = cur.lastrowid
        for _, r in old_lines.iterrows():
            new_id = id_map.get(r["invoice_id"])
            if new_id is None:
                continue
            qty, price = float(r["quantity"] or 0), float(r["unit_price_net"] or 0)
            conn.execute(
                "INSERT INTO invoice_lines (invoice_id, item_name, item_norm, quantity, uom, unit_price, line_net,"
                " unit_price_net) VALUES (?,?,?,?,?,?,?,?)",
                (new_id, r["item_name"], normalize_text(r["item_name"]), qty, r["uom"], price, qty * price, price))
        for _, r in old_jou.iterrows():
            conn.execute(
                "INSERT INTO journal_entries (entry_date, document_no, document_no_norm, account_code, account_norm,"
                " amount) VALUES (?,?,?,?,?,?)",
                (str(r["entry_date"])[:10], r["document_no"], normalize_doc_no(r["document_no"]),
                 r["account_code"], normalize_account(r["account_code"]), r["amount"]))

    # ------------------------------------------------------------------ ayarlar
    def get_setting(self, key, default=None):
        with self.connection() as conn:
            row = conn.execute("SELECT value FROM settings WHERE key = ?", (key,)).fetchone()
        return row[0] if row else default

    def set_setting(self, key, value):
        with self.connection() as conn:
            conn.execute("INSERT OR REPLACE INTO settings(key, value) VALUES (?, ?)", (key, str(value)))

    # ------------------------------------------------------------------ sütun eşlemeleri
    # Onaylanan eşlemeler ayarlar tablosunda "sutun_esleme:<TÜR>:<başlık imzası>" anahtarıyla JSON olarak saklanır.
    def get_column_mappings(self, kind):
        """Kayıtlı eşlemelerin JSON metinleri (en son kaydedilen önce)."""
        with self.connection() as conn:
            rows = conn.execute("SELECT value FROM settings WHERE key LIKE ? ORDER BY rowid DESC",
                                (f"sutun_esleme:{kind}:%",)).fetchall()
        return [r[0] for r in rows]

    def save_column_mapping(self, kind, signature, mapping_json):
        with self.connection() as conn:
            key = f"sutun_esleme:{kind}:{signature}"
            conn.execute("DELETE FROM settings WHERE key = ?", (key,))  # Yeniden kayıt en yeni sayılsın
            conn.execute("INSERT INTO settings(key, value) VALUES (?, ?)", (key, mapping_json))

    # ------------------------------------------------------------------ dosya takibi
    def find_import(self, kind, file_hash):
        with self.connection() as conn:
            return conn.execute("SELECT file_name, imported_at FROM imports WHERE kind = ? AND file_hash = ?",
                                (kind, file_hash)).fetchone()

    def _record_import(self, conn, kind, file_name, file_hash, count):
        cur = conn.execute(
            "INSERT OR REPLACE INTO imports (kind, file_name, file_hash, imported_at, record_count) VALUES (?,?,?,?,?)",
            (kind, file_name, file_hash, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), count))
        return cur.lastrowid

    # ------------------------------------------------------------------ faturalar
    def save_invoices(self, invoices, kind, file_name, file_hash):
        """Faturaları tek işlemde kaydeder.

        invoices: [(header_dict, [line_dict, ...]), ...]
        Dönüş: (eklenen_sayısı, [mükerrer header, ...])
        """
        added, duplicates = 0, []
        with self.connection() as conn:
            import_id = self._record_import(conn, kind, file_name, file_hash, 0)
            for header, lines in invoices:
                exists = conn.execute(
                    "SELECT 1 FROM invoices WHERE supplier_vkn = ? AND invoice_no_norm = ?",
                    (header["supplier_vkn"], normalize_doc_no(header["invoice_no"]))).fetchone()
                if exists:
                    duplicates.append(header)
                    continue
                self._insert_invoice(conn, header, lines, import_id)
                added += 1
            conn.execute("UPDATE imports SET record_count = ? WHERE id = ?", (added, import_id))
        return added, duplicates

    @staticmethod
    def _insert_invoice(conn, h, lines, import_id):
        cur = conn.execute(
            "INSERT INTO invoices (invoice_no, invoice_no_norm, issue_date, supplier_vkn, supplier_name, customer_vkn,"
            " customer_name, invoice_type, profile, currency, exchange_rate, line_extension_amount, allowance_total,"
            " total_amount, vat_amount, payable_amount, source, source_file, import_id)"
            " VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (h["invoice_no"], normalize_doc_no(h["invoice_no"]), h["issue_date"], h["supplier_vkn"],
             h.get("supplier_name"), h.get("customer_vkn"), h.get("customer_name"), h.get("invoice_type"),
             h.get("profile"), h.get("currency") or "TRY", h.get("exchange_rate") or 1.0,
             h.get("line_extension_amount"), h.get("allowance_total") or 0.0, h["total_amount"],
             h.get("vat_amount") or 0.0, h.get("payable_amount"), h.get("source"), h.get("source_file"), import_id))
        invoice_id = cur.lastrowid
        for ln in lines:
            qty = ln["quantity"]
            conn.execute(
                "INSERT INTO invoice_lines (invoice_id, line_no, item_name, item_norm, quantity, uom, unit_price,"
                " line_net, unit_price_net, vat_rate, vat_amount) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                (invoice_id, ln.get("line_no"), ln["item_name"], normalize_text(ln["item_name"]), qty,
                 ln.get("uom") or "ADET", ln.get("unit_price"), ln["line_net"],
                 ln["line_net"] / qty if qty else None, ln.get("vat_rate"), ln.get("vat_amount")))
        return invoice_id

    # ------------------------------------------------------------------ yevmiye
    def save_journal(self, rows, file_name, file_hash):
        """rows: [{'entry_date','document_no','account_code','amount','description','source_row'}]"""
        with self.connection() as conn:
            import_id = self._record_import(conn, "YEVMIYE_EXCEL", file_name, file_hash, len(rows))
            conn.executemany(
                "INSERT INTO journal_entries (entry_date, document_no, document_no_norm, account_code, account_norm,"
                " amount, description, source_row, import_id) VALUES (?,?,?,?,?,?,?,?,?)",
                [(r["entry_date"], r["document_no"], normalize_doc_no(r["document_no"]), r["account_code"],
                  normalize_account(r["account_code"]), r["amount"], r.get("description"), r.get("source_row"),
                  import_id) for r in rows])
        return len(rows)

    # ------------------------------------------------------------------ okuma
    def get_invoices_df(self):
        with self.connection() as conn:
            return pd.read_sql_query("SELECT * FROM invoices", conn)

    def get_lines_df(self):
        query = """
            SELECT l.*, i.invoice_no, i.issue_date, i.supplier_vkn, i.supplier_name, i.currency, i.exchange_rate,
                   i.invoice_type
            FROM invoice_lines l JOIN invoices i ON l.invoice_id = i.id
        """
        with self.connection() as conn:
            return pd.read_sql_query(query, conn)

    def get_journal_df(self):
        with self.connection() as conn:
            return pd.read_sql_query("SELECT * FROM journal_entries", conn)

    def counts(self):
        with self.connection() as conn:
            return {
                "fatura": conn.execute("SELECT COUNT(*) FROM invoices").fetchone()[0],
                "fatura_satiri": conn.execute("SELECT COUNT(*) FROM invoice_lines").fetchone()[0],
                "yevmiye_satiri": conn.execute("SELECT COUNT(*) FROM journal_entries").fetchone()[0],
                "dosya": conn.execute("SELECT COUNT(*) FROM imports").fetchone()[0],
            }

    def clear_data(self):
        """Ayarlar hariç tüm verileri siler."""
        with self.connection() as conn:
            for table in ("invoice_lines", "invoices", "journal_entries", "imports"):
                conn.execute(f"DELETE FROM {table}")
