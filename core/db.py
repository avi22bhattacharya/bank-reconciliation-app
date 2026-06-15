"""Database layer: SQLite for local dev, PostgreSQL (Supabase) for production.

Detects backend via st.secrets["postgres"]["url"]. Falls back to SQLite at
data/bank_rec.db when the secret is absent.

Compat layer:
  _Conn / _Cursor wrap both backends behind the same interface.
  - All SQL uses ? placeholders (translated to %s for PostgreSQL).
  - INSERTs that need the inserted id use RETURNING <col>; the cursor sets
    .lastrowid from that result on both backends.
  - Row access is always dict-style: row["column_name"].

All amounts are stored as integer cents, dates as ISO YYYY-MM-DD strings.
"""

from __future__ import annotations

import hashlib
import json
import re
import sqlite3
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path


# ── Schemas ──────────────────────────────────────────────────────────────────

_MIGRATION_PG = """
ALTER TABLE runs ADD COLUMN IF NOT EXISTS results_data TEXT;
"""

_MIGRATION_SQLITE = """
ALTER TABLE runs ADD COLUMN results_data TEXT;
"""

_SCHEMA_PG = """
CREATE TABLE IF NOT EXISTS properties (
    property_code   TEXT PRIMARY KEY,
    property_name   TEXT,
    gl_type         TEXT NOT NULL CHECK (gl_type IN ('yardi','ma')),
    account_info    TEXT,
    gl_account_id   TEXT,
    prop_label_bank TEXT,
    prop_display    TEXT
);

CREATE TABLE IF NOT EXISTS runs (
    run_id        BIGSERIAL PRIMARY KEY,
    property_code TEXT NOT NULL REFERENCES properties(property_code),
    gl_type       TEXT NOT NULL,
    period        TEXT NOT NULL,
    run_at        TEXT NOT NULL,
    status        TEXT NOT NULL CHECK (status IN ('running','complete','failed','superseded')),
    prior_source  TEXT CHECK (prior_source IN ('db','bootstrap','none')),
    workdir       TEXT,
    results_path  TEXT,
    output_path   TEXT,
    stats_json    TEXT,
    results_data  TEXT
);

CREATE TABLE IF NOT EXISTS matches (
    match_id         BIGSERIAL PRIMARY KEY,
    property_code    TEXT NOT NULL,
    run_id           INTEGER REFERENCES runs(run_id),
    match_rule       TEXT NOT NULL,
    is_manual        INTEGER NOT NULL DEFAULT 0,
    matched_at       TEXT NOT NULL,
    bank_total_cents INTEGER,
    gl_total_cents   INTEGER
);

CREATE TABLE IF NOT EXISTS bank_txns (
    txn_hash          TEXT PRIMARY KEY,
    property_code     TEXT NOT NULL,
    date              TEXT,
    amount_cents      INTEGER NOT NULL,
    check_number      TEXT,
    description       TEXT,
    status            TEXT NOT NULL CHECK (status IN ('unmatched','matched','internal')),
    match_id          INTEGER REFERENCES matches(match_id),
    source_period     TEXT NOT NULL,
    first_seen_run_id INTEGER,
    matched_run_id    INTEGER,
    engine_ref        TEXT
);

CREATE TABLE IF NOT EXISTS gl_txns (
    txn_hash          TEXT PRIMARY KEY,
    property_code     TEXT NOT NULL,
    property_label    TEXT,
    date              TEXT,
    control           TEXT,
    reference         TEXT,
    description       TEXT,
    remarks           TEXT,
    debit_cents       INTEGER NOT NULL DEFAULT 0,
    credit_cents      INTEGER NOT NULL DEFAULT 0,
    deposit_number    TEXT,
    tenant_code       TEXT,
    status            TEXT NOT NULL CHECK (status IN ('unmatched','matched','internal')),
    match_id          INTEGER REFERENCES matches(match_id),
    source_period     TEXT NOT NULL,
    first_seen_run_id INTEGER,
    matched_run_id    INTEGER,
    engine_ref        TEXT
);

CREATE INDEX IF NOT EXISTS idx_bank_open ON bank_txns(property_code, status);
CREATE INDEX IF NOT EXISTS idx_gl_open   ON gl_txns(property_code, status);
"""

# SQLite uses INTEGER PRIMARY KEY (rowid alias) instead of BIGSERIAL
_SCHEMA_SQLITE = _SCHEMA_PG.replace("BIGSERIAL PRIMARY KEY", "INTEGER PRIMARY KEY")


# ── Compat layer ─────────────────────────────────────────────────────────────

class _Cursor:
    """Uniform cursor over sqlite3 or psycopg2."""

    def __init__(self, raw, is_pg: bool):
        self._c = raw
        self._is_pg = is_pg
        self.lastrowid = None

    def execute(self, sql: str, params=None):
        actual = sql.replace("?", "%s") if self._is_pg else sql
        self._c.execute(actual, params or ())
        if self._is_pg and "RETURNING" in sql.upper():
            row = self._c.fetchone()
            self.lastrowid = list(row.values())[0] if row else None
        else:
            self.lastrowid = getattr(self._c, "lastrowid", None)
        return self

    @property
    def rowcount(self):
        return self._c.rowcount

    def fetchone(self):
        return self._c.fetchone()

    def fetchall(self):
        return self._c.fetchall()

    def __iter__(self):
        return iter(self._c.fetchall())


class _Conn:
    """Uniform connection over sqlite3 or psycopg2."""

    def __init__(self, raw, is_pg: bool, schema: str):
        self._raw = raw
        self._is_pg = is_pg
        self._apply_schema(schema)

    def _apply_schema(self, schema: str):
        if self._is_pg:
            cur = self._raw.cursor()
            for stmt in schema.split(";"):
                s = stmt.strip()
                if s:
                    cur.execute(s)
            # Migration: add results_data column to existing tables
            try:
                cur.execute("ALTER TABLE runs ADD COLUMN IF NOT EXISTS results_data TEXT")
            except Exception:
                pass
            self._raw.commit()
        else:
            self._raw.executescript(schema)
            # Migration: add results_data column to existing SQLite table
            try:
                self._raw.execute("ALTER TABLE runs ADD COLUMN results_data TEXT")
                self._raw.commit()
            except Exception:
                pass

    def execute(self, sql: str, params=None) -> _Cursor:
        if self._is_pg:
            import psycopg2.extras
            raw_cur = self._raw.cursor(
                cursor_factory=psycopg2.extras.RealDictCursor)
        else:
            raw_cur = self._raw.cursor()
        return _Cursor(raw_cur, self._is_pg).execute(sql, params)

    def commit(self):
        self._raw.commit()

    def rollback(self):
        self._raw.rollback()

    def close(self):
        self._raw.close()


# ── Entry point ───────────────────────────────────────────────────────────────

DB_PATH = Path(__file__).resolve().parent.parent / "data" / "bank_rec.db"


def connect(db_path: str | Path | None = None) -> _Conn:
    """Return a _Conn backed by PostgreSQL (if secrets configured) or SQLite."""
    try:
        import streamlit as st
        url = st.secrets["postgres"]["url"]
        import psycopg2
        pg = psycopg2.connect(url)
        return _Conn(pg, is_pg=True, schema=_SCHEMA_PG)
    except Exception:
        pass

    # SQLite fallback for local dev
    path = Path(db_path if db_path is not None else DB_PATH)
    path.parent.mkdir(parents=True, exist_ok=True)
    raw = sqlite3.connect(str(path), check_same_thread=False)
    raw.row_factory = sqlite3.Row
    raw.execute("PRAGMA foreign_keys = ON")
    return _Conn(raw, is_pg=False, schema=_SCHEMA_SQLITE)


# ── Normalisation + hashing ───────────────────────────────────────────────────

def norm(value) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip()).lower()


def to_cents(amount) -> int:
    if amount is None or amount == "":
        return 0
    return round(float(amount) * 100)


def iso_date(value) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    s = str(value).strip()
    m = re.match(r"^(\d{4}-\d{2}-\d{2})", s)
    if m:
        return m.group(1)
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})", s)
    if m:
        return f"{m.group(3)}-{int(m.group(1)):02d}-{int(m.group(2)):02d}"
    return s


def _h(*fields) -> str:
    return hashlib.sha256("|".join(str(f) for f in fields).encode()).hexdigest()


def bank_content_key(property_code, date_val, amount_cents, check_number, description):
    return (norm(property_code), iso_date(date_val), int(amount_cents),
            norm(check_number), norm(description))


def bank_hash(property_code, source_period, occ, date_val, amount_cents,
              check_number, description) -> str:
    return _h(source_period, occ,
              *bank_content_key(property_code, date_val, amount_cents,
                                check_number, description))


def gl_content_key(property_code, date_val, control, reference, debit_cents,
                   credit_cents, description, remarks):
    return (norm(property_code), iso_date(date_val), norm(control), norm(reference),
            int(debit_cents), int(credit_cents), norm(description), norm(remarks))


def gl_hash(property_code, source_period, occ, date_val, control, reference,
            debit_cents, credit_cents, description, remarks) -> str:
    return _h(source_period, occ,
              *gl_content_key(property_code, date_val, control, reference,
                              debit_cents, credit_cents, description, remarks))


# ── Properties ────────────────────────────────────────────────────────────────

@dataclass
class PropertyMeta:
    property_code: str
    property_name: str = ""
    gl_type: str = "ma"
    account_info: str = ""
    gl_account_id: str = ""
    prop_label_bank: str = ""
    prop_display: str = ""


def get_property(con, property_code: str) -> PropertyMeta | None:
    row = con.execute("SELECT * FROM properties WHERE property_code = ?",
                      (property_code,)).fetchone()
    if row is None:
        return None
    return PropertyMeta(**{k: row[k] if row[k] is not None else "" for k in row.keys()})


def upsert_property(con, meta: PropertyMeta):
    con.execute("""
        INSERT INTO properties (property_code, property_name, gl_type, account_info,
                                gl_account_id, prop_label_bank, prop_display)
        VALUES (?,?,?,?,?,?,?)
        ON CONFLICT (property_code) DO UPDATE SET
            property_name=excluded.property_name, gl_type=excluded.gl_type,
            account_info=excluded.account_info, gl_account_id=excluded.gl_account_id,
            prop_label_bank=excluded.prop_label_bank,
            prop_display=excluded.prop_display
    """, (meta.property_code, meta.property_name, meta.gl_type, meta.account_info,
          meta.gl_account_id, meta.prop_label_bank, meta.prop_display))
    con.commit()


def list_properties(con) -> list[PropertyMeta]:
    rows = con.execute("SELECT * FROM properties ORDER BY property_code").fetchall()
    return [PropertyMeta(**{k: r[k] if r[k] is not None else "" for k in r.keys()})
            for r in rows]


# ── Queries used by the UI ────────────────────────────────────────────────────

def unmatched_bank(con, property_code: str) -> list:
    return con.execute("""
        SELECT * FROM bank_txns WHERE property_code = ? AND status = 'unmatched'
        ORDER BY date, amount_cents
    """, (property_code,)).fetchall()


def unmatched_gl(con, property_code: str) -> list:
    return con.execute("""
        SELECT * FROM gl_txns WHERE property_code = ? AND status = 'unmatched'
        ORDER BY date, control
    """, (property_code,)).fetchall()


def latest_run(con, property_code: str):
    return con.execute("""
        SELECT * FROM runs WHERE property_code = ? AND status = 'complete'
        ORDER BY run_id DESC LIMIT 1
    """, (property_code,)).fetchone()


def create_manual_match(con, property_code: str, bank_hashes: list[str],
                        gl_hashes: list[str], run_id: int | None) -> int:
    now = datetime.now().isoformat(timespec="seconds")
    bank_total = con.execute(
        "SELECT COALESCE(SUM(amount_cents), 0) AS total FROM bank_txns WHERE txn_hash IN "
        f"({','.join('?' * len(bank_hashes))})",
        bank_hashes).fetchone()["total"] if bank_hashes else 0
    gl_total = con.execute(
        "SELECT COALESCE(SUM(debit_cents - credit_cents), 0) AS total FROM gl_txns WHERE txn_hash IN "
        f"({','.join('?' * len(gl_hashes))})",
        gl_hashes).fetchone()["total"] if gl_hashes else 0
    cur = con.execute("""
        INSERT INTO matches (property_code, run_id, match_rule, is_manual, matched_at,
                             bank_total_cents, gl_total_cents)
        VALUES (?,?,?,1,?,?,?) RETURNING match_id
    """, (property_code, run_id, "MANUAL", now, bank_total, gl_total))
    match_id = cur.lastrowid
    for h in bank_hashes:
        con.execute("""UPDATE bank_txns SET status='matched', match_id=?, matched_run_id=?
                       WHERE txn_hash=? AND status='unmatched'""", (match_id, run_id, h))
    for h in gl_hashes:
        con.execute("""UPDATE gl_txns SET status='matched', match_id=?, matched_run_id=?
                       WHERE txn_hash=? AND status='unmatched'""", (match_id, run_id, h))
    con.commit()
    return match_id


def matched_view(con, property_code: str | None = None, period: str | None = None,
                 manual_only: bool = False) -> list:
    where, params = ["1=1"], []
    if property_code:
        where.append("m.property_code = ?")
        params.append(property_code)
    if manual_only:
        where.append("m.is_manual = 1")
    sql_period_b = sql_period_g = ""
    if period:
        sql_period_b = " AND t.source_period = ?"
        sql_period_g = " AND t.source_period = ?"
    q = f"""
        SELECT m.match_id, m.match_rule, m.is_manual, m.matched_at,
               'bank' AS side, t.date, t.amount_cents AS amount_cents,
               t.check_number AS ref, t.description AS description, '' AS remarks,
               t.source_period
        FROM matches m JOIN bank_txns t ON t.match_id = m.match_id
        WHERE {' AND '.join(where)}{sql_period_b}
        UNION ALL
        SELECT m.match_id, m.match_rule, m.is_manual, m.matched_at,
               'gl' AS side, t.date, (t.debit_cents - t.credit_cents) AS amount_cents,
               t.reference AS ref, t.description AS description, t.remarks,
               t.source_period
        FROM matches m JOIN gl_txns t ON t.match_id = m.match_id
        WHERE {' AND '.join(where)}{sql_period_g}
        ORDER BY 1, 5
    """
    all_params = params + ([period] if period else []) + params + ([period] if period else [])
    return con.execute(q, all_params).fetchall()


def run_stats(row) -> dict:
    try:
        return json.loads(row["stats_json"] or "{}")
    except (TypeError, json.JSONDecodeError):
        return {}
