"""
database_web.py — SQLite layer for the Streamlit web app.

Key differences from desktop database.py:
  • CSV raw bytes are stored as BLOBs in a new `csv_data` table so the app
    is fully self-contained — no filesystem paths needed.
  • get_csv_data(tid) returns raw bytes; load_csv can accept them directly.
  • Works from a single .db file that the user uploads/downloads each session.
"""

import io
import sqlite3
import json
from datetime import datetime


class WebDatabase:
    def __init__(self, db_bytes: bytes = None):
        """
        Open (or create) an in-memory SQLite database.
        If db_bytes is supplied (user uploaded their .db), load it.
        """
        self._mem = sqlite3.connect(':memory:', check_same_thread=False)
        self._mem.row_factory = sqlite3.Row

        if db_bytes:
            # Restore from uploaded bytes by copying page-by-page
            disk = sqlite3.connect(':memory:')
            disk.deserialize(db_bytes)
            disk.backup(self._mem)
            disk.close()

        self._create_tables()
        self._migrate()

    # ── Schema ────────────────────────────────────────────────────────────────
    def _create_tables(self):
        self._mem.executescript("""
        CREATE TABLE IF NOT EXISTS source_folders (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            name         TEXT    NOT NULL,
            display_name TEXT    DEFAULT NULL,
            path         TEXT    NOT NULL UNIQUE,
            imported_at  TEXT    NOT NULL
        );
        CREATE TABLE IF NOT EXISTS tests (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            source_folder_id INTEGER NOT NULL,
            test_folder_name TEXT    NOT NULL,
            display_name     TEXT    DEFAULT NULL,
            test_number      INTEGER,
            csv_path         TEXT    NOT NULL DEFAULT '',
            analyzed         INTEGER DEFAULT 0,
            condition_id     INTEGER DEFAULT NULL,
            sample_width_mm  REAL    DEFAULT NULL,
            created_at       TEXT    NOT NULL,
            FOREIGN KEY (source_folder_id) REFERENCES source_folders(id)
        );
        CREATE TABLE IF NOT EXISTS csv_data (
            test_id  INTEGER PRIMARY KEY,
            filename TEXT,
            data     BLOB,
            FOREIGN KEY (test_id) REFERENCES tests(id)
        );
        CREATE TABLE IF NOT EXISTS analyses (
            id                INTEGER PRIMARY KEY AUTOINCREMENT,
            test_id           INTEGER NOT NULL UNIQUE,
            smoothing_window  INTEGER DEFAULT 51,
            whole_mean        REAL, whole_max  REAL,
            whole_min         REAL, whole_std  REAL,
            whole_mean_smooth REAL, whole_max_smooth REAL,
            whole_min_smooth  REAL, whole_std_smooth REAL,
            regions_json      TEXT  DEFAULT '[]',
            analyzed_at       TEXT,
            FOREIGN KEY (test_id) REFERENCES tests(id)
        );
        CREATE TABLE IF NOT EXISTS conditions (
            id    INTEGER PRIMARY KEY AUTOINCREMENT,
            name  TEXT NOT NULL,
            color TEXT NOT NULL DEFAULT '#3b82f6'
        );
        CREATE TABLE IF NOT EXISTS saved_figures (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            name        TEXT    NOT NULL,
            config_json TEXT    NOT NULL,
            created_at  TEXT    NOT NULL,
            updated_at  TEXT    NOT NULL
        );
        """)
        self._mem.commit()

    def _migrate(self):
        cur = self._mem.execute("PRAGMA table_info(tests)")
        cols = {r['name'] for r in cur.fetchall()}
        for col, defn in [
            ('display_name',    'TEXT DEFAULT NULL'),
            ('sample_width_mm', 'REAL DEFAULT NULL'),
        ]:
            if col not in cols:
                self._mem.execute(f"ALTER TABLE tests ADD COLUMN {col} {defn}")
        cur = self._mem.execute("PRAGMA table_info(source_folders)")
        cols = {r['name'] for r in cur.fetchall()}
        if 'display_name' not in cols:
            self._mem.execute(
                "ALTER TABLE source_folders ADD COLUMN display_name TEXT DEFAULT NULL")
        self._mem.commit()

    # ── Serialise to bytes (for download) ────────────────────────────────────
    def to_bytes(self) -> bytes:
        return self._mem.serialize()

    # ── Source folders ────────────────────────────────────────────────────────
    def add_source_folder(self, name: str, path: str = None) -> int:
        path = path or f"web://{name}"
        try:
            c = self._mem.execute(
                "INSERT INTO source_folders (name,path,imported_at) VALUES (?,?,?)",
                (name, path, datetime.now().isoformat()))
            self._mem.commit()
            return c.lastrowid
        except sqlite3.IntegrityError:
            row = self._mem.execute(
                "SELECT id FROM source_folders WHERE path=?", (path,)).fetchone()
            return row['id']

    def get_source_folders(self):
        return [dict(r) for r in self._mem.execute(
            "SELECT * FROM source_folders ORDER BY imported_at DESC").fetchall()]

    def rename_source_folder(self, sfid: int, display_name: str):
        val = display_name.strip() or None
        self._mem.execute(
            "UPDATE source_folders SET display_name=? WHERE id=?", (val, sfid))
        self._mem.commit()

    def delete_source_folder(self, sfid: int):
        tests = self.get_tests_for_folder(sfid)
        for t in tests:
            self._mem.execute("DELETE FROM csv_data WHERE test_id=?", (t['id'],))
            self._mem.execute("DELETE FROM analyses WHERE test_id=?", (t['id'],))
            self._mem.execute("DELETE FROM tests WHERE id=?", (t['id'],))
        self._mem.execute("DELETE FROM source_folders WHERE id=?", (sfid,))
        self._mem.commit()

    # ── Tests ─────────────────────────────────────────────────────────────────
    def add_test(self, sfid: int, name: str, num: int,
                 csv_bytes: bytes, filename: str) -> int:
        self._mem.execute(
            "INSERT OR IGNORE INTO tests "
            "(source_folder_id,test_folder_name,test_number,csv_path,analyzed,created_at) "
            "VALUES (?,?,?,'',0,?)",
            (sfid, name, num, datetime.now().isoformat()))
        self._mem.commit()
        row = self._mem.execute(
            "SELECT id FROM tests "
            "WHERE source_folder_id=? AND test_folder_name=? AND test_number=?",
            (sfid, name, num)).fetchone()
        tid = row['id']
        self._mem.execute(
            "INSERT OR REPLACE INTO csv_data (test_id,filename,data) VALUES (?,?,?)",
            (tid, filename, csv_bytes))
        self._mem.commit()
        return tid

    def get_csv_data(self, tid: int) -> bytes:
        row = self._mem.execute(
            "SELECT data FROM csv_data WHERE test_id=?", (tid,)).fetchone()
        return bytes(row['data']) if row else b''

    def get_tests_for_folder(self, sfid: int):
        return [dict(r) for r in self._mem.execute(
            "SELECT t.*, a.id as analysis_id "
            "FROM tests t LEFT JOIN analyses a ON t.id=a.test_id "
            "WHERE t.source_folder_id=? ORDER BY t.test_number",
            (sfid,)).fetchall()]

    def get_all_tests(self):
        return [dict(r) for r in self._mem.execute(
            "SELECT t.*, a.id as analysis_id "
            "FROM tests t LEFT JOIN analyses a ON t.id=a.test_id "
            "ORDER BY t.source_folder_id, t.test_number").fetchall()]

    def rename_test(self, tid: int, display_name: str):
        self._mem.execute(
            "UPDATE tests SET display_name=? WHERE id=?",
            (display_name.strip() or None, tid))
        self._mem.commit()

    def delete_test(self, tid: int):
        self._mem.execute("DELETE FROM csv_data WHERE test_id=?", (tid,))
        self._mem.execute("DELETE FROM analyses WHERE test_id=?", (tid,))
        self._mem.execute("DELETE FROM tests WHERE id=?", (tid,))
        self._mem.commit()

    def set_test_condition(self, tid: int, cid):
        self._mem.execute(
            "UPDATE tests SET condition_id=? WHERE id=?", (cid, tid))
        self._mem.commit()

    def set_test_width(self, tid: int, width_mm: float):
        self._mem.execute(
            "UPDATE tests SET sample_width_mm=? WHERE id=?", (width_mm, tid))
        self._mem.commit()

    # ── Analyses ──────────────────────────────────────────────────────────────
    def get_analysis(self, test_id: int):
        row = self._mem.execute(
            "SELECT * FROM analyses WHERE test_id=?", (test_id,)).fetchone()
        return dict(row) if row else None

    def save_analysis(self, tid: int, sw: int, ws: dict, wss: dict, regions: list):
        self._mem.execute(
            "INSERT OR REPLACE INTO analyses "
            "(test_id,smoothing_window,"
            " whole_mean,whole_max,whole_min,whole_std,"
            " whole_mean_smooth,whole_max_smooth,whole_min_smooth,whole_std_smooth,"
            " regions_json,analyzed_at) "
            "VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
            (tid, sw,
             ws.get('mean'), ws.get('max'), ws.get('min'), ws.get('std'),
             wss.get('mean'), wss.get('max'), wss.get('min'), wss.get('std'),
             json.dumps(regions), datetime.now().isoformat()))
        self._mem.execute("UPDATE tests SET analyzed=1 WHERE id=?", (tid,))
        self._mem.commit()

    # ── Conditions ────────────────────────────────────────────────────────────
    def get_conditions(self):
        return [dict(r) for r in self._mem.execute(
            "SELECT * FROM conditions ORDER BY id").fetchall()]

    def add_condition(self, name: str, color: str) -> int:
        c = self._mem.execute(
            "INSERT INTO conditions (name,color) VALUES (?,?)", (name, color))
        self._mem.commit()
        return c.lastrowid

    def update_condition(self, cid: int, name: str, color: str):
        self._mem.execute(
            "UPDATE conditions SET name=?,color=? WHERE id=?", (name, color, cid))
        self._mem.commit()

    def delete_condition(self, cid: int):
        self._mem.execute("DELETE FROM conditions WHERE id=?", (cid,))
        self._mem.execute(
            "UPDATE tests SET condition_id=NULL WHERE condition_id=?", (cid,))
        self._mem.commit()

    # ── Saved figures ─────────────────────────────────────────────────────────
    def get_saved_figures(self):
        return [dict(r) for r in self._mem.execute(
            "SELECT * FROM saved_figures ORDER BY updated_at DESC").fetchall()]

    def save_figure(self, name: str, config: dict) -> int:
        now = datetime.now().isoformat()
        c = self._mem.execute(
            "INSERT INTO saved_figures (name,config_json,created_at,updated_at) "
            "VALUES (?,?,?,?)", (name, json.dumps(config), now, now))
        self._mem.commit()
        return c.lastrowid

    def update_figure(self, fig_id: int, name: str, config: dict):
        self._mem.execute(
            "UPDATE saved_figures SET name=?,config_json=?,updated_at=? WHERE id=?",
            (name, json.dumps(config), datetime.now().isoformat(), fig_id))
        self._mem.commit()

    def delete_figure(self, fig_id: int):
        self._mem.execute("DELETE FROM saved_figures WHERE id=?", (fig_id,))
        self._mem.commit()

    def get_figure(self, fig_id: int):
        row = self._mem.execute(
            "SELECT * FROM saved_figures WHERE id=?", (fig_id,)).fetchone()
        return dict(row) if row else None

    def close(self):
        self._mem.close()
