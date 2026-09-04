"""
Data layer: the only module that opens a SQLite connection directly.

Fixes applied here (see project README for the full list):
  - Every connection sets journal_mode=DELETE explicitly (never WAL --
    WAL is documented by SQLite as unreliable on network/cloud-sync
    filesystems) and a busy_timeout, so a sync client briefly touching
    the file causes a short wait-and-retry instead of a raw
    "database is locked" exception reaching the UI.
  - backup() copies the live DB file to data/backups/ with a timestamp.
    Callers use this before any bulk write (archive, import) so a bad
    write has a same-day rollback point.
  - Corrections are scoped by (fund_name, description), not description
    alone -- a correction made while working Fund II can no longer get
    silently applied to an identical-looking Fund III transaction.
  - The training corpus for fuzzy matching is read straight from
    transaction_history. There is no separate "master training Excel
    file" to remember to reload -- an archive or a saved correction is
    live for the very next mapping run, in the same session.
"""

from __future__ import annotations

import shutil
import sqlite3
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

from config import AppConfig, BACKUP_DIR
from logging_setup import get_logger

_logger = get_logger("db")

SCHEMA = """
CREATE TABLE IF NOT EXISTS transaction_history (
    id                INTEGER PRIMARY KEY AUTOINCREMENT,
    fund_name         TEXT NOT NULL,
    property_code     TEXT,
    sheet_name        TEXT,
    bank_name         TEXT,
    account_number    TEXT,
    trans_date        TEXT,
    post_month        TEXT,
    amount            REAL,
    bank_desc         TEXT,
    bank_desc2        TEXT,
    combined_desc     TEXT,
    mapped_gl         TEXT,
    mapped_gl_name    TEXT,
    bank_gl           TEXT,
    archived_on       TEXT,
    source_month      TEXT
);
CREATE INDEX IF NOT EXISTS idx_th_combined_desc ON transaction_history(combined_desc);
CREATE INDEX IF NOT EXISTS idx_th_fund          ON transaction_history(fund_name);
CREATE INDEX IF NOT EXISTS idx_th_source_month  ON transaction_history(source_month);
CREATE INDEX IF NOT EXISTS idx_th_dedupe        ON transaction_history(fund_name, trans_date, amount, bank_desc);

CREATE TABLE IF NOT EXISTS corrections (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    fund_name    TEXT NOT NULL,
    description  TEXT NOT NULL,
    gl_code      TEXT,
    gl_name      TEXT,
    corrected_at TEXT
);
CREATE UNIQUE INDEX IF NOT EXISTS idx_corr_fund_desc ON corrections(fund_name, description);
"""


class Database:

    def __init__(self, config: AppConfig):
        self.config = config
        self.path: Path = config.db_path
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self._connect() as conn:
            conn.executescript(SCHEMA)

    @contextmanager
    def _connect(self):
        conn = sqlite3.connect(str(self.path), timeout=10.0)
        try:
            conn.execute("PRAGMA journal_mode=DELETE")
            yield conn
            conn.commit()
        finally:
            conn.close()

    # -- maintenance --------------------------------------------------
    def backup(self, tag: str) -> Optional[Path]:
        """Best-effort snapshot before a risky write. Never raises --
        a failed backup just means this one write has no extra safety
        net, it should not block the operation itself."""
        if not self.path.exists():
            return None
        BACKUP_DIR.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        dest = BACKUP_DIR / f"{self.path.stem}_{tag}_{stamp}{self.path.suffix}"
        try:
            shutil.copy2(self.path, dest)
            _logger.info("Backed up database to %s", dest)
            return dest
        except Exception:
            _logger.exception("Database backup failed (tag=%s) -- proceeding without one", tag)
            return None

    # -- corrections (fund-scoped) -------------------------------------
    def get_corrections(self, fund_name: str) -> Dict[str, Tuple[str, str]]:
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT description, gl_code, gl_name FROM corrections "
                "WHERE fund_name = ?", (fund_name,)
            ).fetchall()
        return {r[0]: (r[1], r[2]) for r in rows}

    def save_correction(self, fund_name: str, description: str,
                         gl_code: str, gl_name: str) -> None:
        with self._connect() as conn:
            conn.execute(
                "INSERT INTO corrections (fund_name, description, gl_code, "
                "gl_name, corrected_at) VALUES (?,?,?,?,?) "
                "ON CONFLICT(fund_name, description) DO UPDATE SET "
                "gl_code=excluded.gl_code, gl_name=excluded.gl_name, "
                "corrected_at=excluded.corrected_at",
                (fund_name, description.lower().strip(), gl_code, gl_name,
                 datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            )

    # -- training corpus, read live from history ------------------------
    def get_training_corpus(self, fund_name: str) -> List[Tuple[str, str, str]]:
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT combined_desc, mapped_gl, mapped_gl_name "
                "FROM transaction_history "
                "WHERE fund_name = ? AND combined_desc != '' AND mapped_gl != ''",
                (fund_name,)
            ).fetchall()
        return rows

    # -- archiving ------------------------------------------------------
    def is_duplicate(self, fund_name: str, trans_date: str,
                      bank_desc: str, amount: float) -> bool:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT 1 FROM transaction_history WHERE fund_name=? AND "
                "trans_date=? AND bank_desc=? AND amount=?",
                (fund_name, trans_date, bank_desc, amount)
            ).fetchone()
        return row is not None

    def insert_transactions(self, fund_name: str, property_code: str,
                             source_month: str,
                             transactions: Iterable[dict]) -> Tuple[int, int]:
        inserted = skipped = 0
        with self._connect() as conn:
            for txn in transactions:
                if txn.get("status") == "SKIP":
                    continue
                trans_date = str(txn.get("date", ""))
                bank_desc = str(txn.get("bank_desc", ""))
                amount = txn.get("amount", 0)
                exists = conn.execute(
                    "SELECT 1 FROM transaction_history WHERE fund_name=? AND "
                    "trans_date=? AND bank_desc=? AND amount=?",
                    (fund_name, trans_date, bank_desc, amount)
                ).fetchone()
                if exists:
                    skipped += 1
                    continue
                conn.execute(
                    "INSERT INTO transaction_history (fund_name, property_code, "
                    "sheet_name, bank_name, account_number, trans_date, "
                    "post_month, amount, bank_desc, bank_desc2, combined_desc, "
                    "mapped_gl, mapped_gl_name, bank_gl, archived_on, source_month) "
                    "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                    (fund_name, property_code, txn.get("sheet", ""),
                     txn.get("bank_name", ""), txn.get("account_number", ""),
                     trans_date, txn.get("post_month", ""), amount, bank_desc,
                     str(txn.get("bank_desc2", "")),
                     str(txn.get("combined_desc", "")),
                     str(txn.get("mapped_gl", "")),
                     str(txn.get("mapped_name", "")),
                     str(txn.get("bank_gl", "")),
                     datetime.now().strftime("%Y-%m-%d"), source_month)
                )
                inserted += 1
        return inserted, skipped

    def archived_months(self, fund_name: str) -> set:
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT DISTINCT source_month FROM transaction_history "
                "WHERE fund_name = ? AND source_month != ''", (fund_name,)
            ).fetchall()
        return {r[0] for r in rows}

    # -- stats / browsing -------------------------------------------
    def get_archived_keys(self, fund_name: str) -> set:
        """(trans_date, amount, bank_desc) already archived for this
        fund. Used to filter a freshly re-read cashbook sheet down to
        rows that actually need review -- without this, a completed
        month reappears every time the sheet is read again, since the
        cashbook itself is meant to keep every month forever."""
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT trans_date, amount, bank_desc FROM transaction_history "
                "WHERE fund_name = ?", (fund_name,)
            ).fetchall()
        return {
            (str(d), round(a, 2) if a is not None else None, str(b or "").strip().lower())
            for d, a, b in rows
        }

    def get_transactions_for_export(self, fund_name: str, source_month: str = "All") -> list:
        """Rows shaped for ETLExporter, sourced directly from the
        database rather than whatever happens to be loaded in the
        on-screen review table. Exists specifically for the case where
        a month is already archived but its ETL was never generated,
        was lost, or needs to be reissued after a correction."""
        query = ("SELECT trans_date, amount, bank_gl, mapped_gl, combined_desc "
                 "FROM transaction_history WHERE fund_name = ?")
        params: list = [fund_name]
        if source_month != "All":
            query += " AND source_month = ?"
            params.append(source_month)
        with self._connect() as conn:
            rows = conn.execute(query, params).fetchall()
        return [
            {"date": d, "amount": a, "bank_gl": bg, "mapped_gl": mg, "combined_desc": cd}
            for d, a, bg, mg, cd in rows
        ]

    def get_stats(self) -> dict:
        with self._connect() as conn:
            total = conn.execute("SELECT COUNT(*) FROM transaction_history").fetchone()[0]
            months = conn.execute("SELECT COUNT(DISTINCT source_month) FROM transaction_history").fetchone()[0]
            funds = conn.execute("SELECT COUNT(DISTINCT fund_name) FROM transaction_history").fetchone()[0]
            accts = conn.execute(
                "SELECT COUNT(DISTINCT mapped_gl) FROM transaction_history WHERE mapped_gl != ''"
            ).fetchone()[0]
            dr = conn.execute("SELECT MIN(trans_date), MAX(trans_date) FROM transaction_history").fetchone()
            corr = conn.execute("SELECT COUNT(*) FROM corrections").fetchone()[0]
            fund_rows = conn.execute(
                "SELECT fund_name, COUNT(*) FROM transaction_history "
                "GROUP BY fund_name ORDER BY COUNT(*) DESC"
            ).fetchall()
            top_gl = conn.execute(
                "SELECT mapped_gl, COUNT(*) FROM transaction_history "
                "WHERE mapped_gl != '' GROUP BY mapped_gl ORDER BY COUNT(*) DESC LIMIT 8"
            ).fetchall()
        return {
            "total": total, "months": months, "funds": funds, "accts": accts,
            "earliest": dr[0] or "N/A", "latest": dr[1] or "N/A",
            "corrections": corr, "fund_rows": fund_rows, "top_gl": top_gl,
        }

    def query_transactions(self, fund_name: str = "All", month: str = "All",
                            search: str = "") -> List[tuple]:
        query = ("SELECT id, fund_name, sheet_name, bank_name, account_number, "
                 "trans_date, post_month, amount, bank_desc, bank_desc2, "
                 "mapped_gl, mapped_gl_name, archived_on, source_month "
                 "FROM transaction_history WHERE 1=1")
        params: list = []
        if fund_name != "All":
            query += " AND fund_name = ?"
            params.append(fund_name)
        if month != "All":
            query += " AND source_month = ?"
            params.append(month)
        query += " ORDER BY trans_date DESC"
        with self._connect() as conn:
            rows = conn.execute(query, params).fetchall()
        if search:
            s = search.strip().lower()
            rows = [r for r in rows if any(s in str(c).lower() for c in r)]
        return rows

    def list_source_months(self) -> List[str]:
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT DISTINCT source_month FROM transaction_history "
                "WHERE source_month IS NOT NULL ORDER BY source_month DESC"
            ).fetchall()
        return [r[0] for r in rows if r[0]]

    def delete_transaction(self, row_id: int) -> None:
        with self._connect() as conn:
            conn.execute("DELETE FROM transaction_history WHERE id=?", (row_id,))

    def get_transaction(self, row_id: int) -> Optional[tuple]:
        with self._connect() as conn:
            return conn.execute(
                "SELECT id, fund_name, property_code, sheet_name, bank_name, "
                "account_number, trans_date, post_month, amount, bank_desc, "
                "bank_desc2, combined_desc, mapped_gl, mapped_gl_name, bank_gl, "
                "archived_on, source_month FROM transaction_history WHERE id=?",
                (row_id,)
            ).fetchone()

    def update_transaction(self, row_id: int, fields: Dict) -> bool:
        """Update specific columns of an existing transaction_history row.
        If fund_name is among the changed fields, property_code is
        recomputed from config automatically so the two can't drift
        apart -- editing the fund without the property code updating
        with it would silently misfile the row for ETL/export purposes."""
        allowed = {
            "fund_name", "bank_desc", "bank_desc2", "combined_desc",
            "mapped_gl", "mapped_gl_name", "bank_gl", "amount",
            "trans_date", "source_month",
        }
        updates = {k: v for k, v in fields.items() if k in allowed}
        if not updates:
            return False
        if "fund_name" in updates:
            try:
                updates["property_code"] = self.config.property_code(updates["fund_name"])
            except KeyError:
                pass
        set_clause = ", ".join(f"{col} = ?" for col in updates)
        params = list(updates.values()) + [row_id]
        with self._connect() as conn:
            cur = conn.execute(
                f"UPDATE transaction_history SET {set_clause} WHERE id = ?", params)
            return cur.rowcount > 0
