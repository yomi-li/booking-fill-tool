# -*- coding: utf-8 -*-
"""SQLite 数据层：客户 / 费用档案 / 客户习惯库。

零依赖（Python 标准库 sqlite3），跨电脑免装数据库，文件即 data/app.db。
schema 首次访问自动创建，并播种一个「默认模板」客户，便于 yomi 编辑全局默认费用。
"""
from __future__ import annotations

import sqlite3
import threading
from datetime import datetime, timezone
from pathlib import Path

from bol_forecast.config import DB_PATH

DEFAULT_CODE = "*DEFAULT*"
DEFAULT_NAME = "（默认模板）"

_local = threading.local()


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def get_conn() -> sqlite3.Connection:
    """每线程一个连接（sqlite3 连接非线程安全）。"""
    conn = getattr(_local, "conn", None)
    if conn is None:
        DB_PATH.parent.mkdir(parents=True, exist_ok=True)
        conn = sqlite3.connect(str(DB_PATH), timeout=30)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA foreign_keys=ON")
        _local.conn = conn
    return conn


def _ensure_default_customer(conn: sqlite3.Connection) -> int:
    row = conn.execute(
        "SELECT id FROM customers WHERE code=?", (DEFAULT_CODE,)).fetchone()
    if row:
        return int(row["id"])
    cur = conn.execute(
        "INSERT INTO customers(name, code, created_at, updated_at) VALUES(?,?,?,?)",
        (DEFAULT_NAME, DEFAULT_CODE, _now(), _now()))
    cid = int(cur.lastrowid)
    # 播种默认费用（工厂 + INVOICE），让 yomi 可在界面编辑全局默认
    from bol_forecast.generators.charges import clone_defaults
    for kind in ("factory", "invoice"):
        _seed_profiles(conn, cid, kind, clone_defaults(kind))
    conn.commit()
    return cid


def _seed_profiles(conn: sqlite3.Connection, cid: int, kind: str, rows: list[dict]) -> None:
    for pos, r in enumerate(rows):
        conn.execute(
            """INSERT INTO charge_profiles
               (customer_id, kind, name, calc_mode, unit, unit_price, qty_var,
                formula, min_amount, max_amount, position, currency, exchange_rate)
               VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (cid, kind, r.get("name", ""), r.get("calc_mode", "FIXED"),
             r.get("unit", ""), float(r.get("unit_price", 0) or 0),
             r.get("qty_var", ""), r.get("formula", ""),
             _none_float(r.get("min_amount")), _none_float(r.get("max_amount")),
             pos, (r.get("currency") or "USD").upper(),
             float(r.get("exchange_rate") or 1.0)))


def _none_float(v):
    if v in (None, ""):
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def init_db() -> int:
    """建表 + 播种默认客户，返回默认客户 id。幂等。"""
    conn = get_conn()
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS customers (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            name      TEXT NOT NULL,
            code      TEXT,
            created_at TEXT,
            updated_at TEXT,
            UNIQUE(code)
        );
        -- 兼容旧库：增量加列（幂等，列已存在则忽略）
        CREATE TABLE IF NOT EXISTS shipments (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            customer_id INTEGER REFERENCES customers(id) ON DELETE SET NULL,
            invoice_customer_id INTEGER REFERENCES customers(id) ON DELETE SET NULL,
            bl_no       TEXT,
            vessel      TEXT,
            atd         TEXT,
            eta         TEXT,
            ctns        REAL,
            gw          REAL,
            cbm         REAL,
            docs        TEXT,
            created_at  TEXT
        );
        CREATE TABLE IF NOT EXISTS charge_profiles (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            customer_id INTEGER NOT NULL REFERENCES customers(id) ON DELETE CASCADE,
            kind        TEXT NOT NULL,
            name        TEXT,
            calc_mode   TEXT,
            unit        TEXT,
            unit_price  REAL,
            qty_var     TEXT,
            formula     TEXT,
            min_amount  REAL,
            max_amount  REAL,
            position    INTEGER
        );
        CREATE TABLE IF NOT EXISTS habits (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            customer_id INTEGER NOT NULL REFERENCES customers(id) ON DELETE CASCADE,
            key         TEXT NOT NULL,
            value       TEXT,
            UNIQUE(customer_id, key)
        );
        -- 预报话术模板：customer_id 为 NULL 表示「所有客户共用」，
        -- 非 NULL 表示该客户专属；两者都支持多版本。
        CREATE TABLE IF NOT EXISTS speech_templates (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            customer_id INTEGER REFERENCES customers(id) ON DELETE CASCADE,
            name        TEXT NOT NULL,
            body        TEXT NOT NULL,
            is_default  INTEGER DEFAULT 0,
            position    INTEGER DEFAULT 0,
            created_at  TEXT,
            updated_at  TEXT
        );
        CREATE INDEX IF NOT EXISTS idx_cp_cust_kind
            ON charge_profiles(customer_id, kind);
        CREATE INDEX IF NOT EXISTS idx_habit_cust ON habits(customer_id);
        CREATE INDEX IF NOT EXISTS idx_speech_cust ON speech_templates(customer_id);
        """
    )
    # 兼容旧库：增量加列（幂等）
    _ALTERS = [
        "ALTER TABLE customers ADD COLUMN consignee_name TEXT",
        "ALTER TABLE customers ADD COLUMN consignee_addr TEXT",
        "ALTER TABLE customers ADD COLUMN consignee_contact TEXT",
        "ALTER TABLE customers ADD COLUMN shipper_name TEXT",
        "ALTER TABLE customers ADD COLUMN shipper_addr TEXT",
        "ALTER TABLE customers ADD COLUMN shipper_contact TEXT",
        "ALTER TABLE customers ADD COLUMN customer_type TEXT",
        "ALTER TABLE customers ADD COLUMN billing_title TEXT",
        "ALTER TABLE customers ADD COLUMN invoice_title TEXT",
        "ALTER TABLE customers ADD COLUMN csr_name TEXT",
        "ALTER TABLE customers ADD COLUMN credit_days INTEGER DEFAULT 0",
        "ALTER TABLE customers ADD COLUMN settlement_node TEXT",
        "ALTER TABLE customers ADD COLUMN settlement_days INTEGER DEFAULT 0",
        # 需求2（2026-08-18）：INVOICE 多币别 —— 费用币别 / 汇率
        "ALTER TABLE charge_profiles ADD COLUMN currency TEXT DEFAULT 'USD'",
        "ALTER TABLE charge_profiles ADD COLUMN exchange_rate REAL DEFAULT 1.0",
    ]
    for a in _ALTERS:
        try:
            conn.execute(a)
        except sqlite3.OperationalError:
            pass  # 列已存在
    conn.commit()
    cid = _ensure_default_customer(conn)
    _migrate_speech_templates(conn)
    return cid


# ---------------- 话术模板：JSON → SQLite 一次性迁移 ----------------
_BUILTIN_SPEECH = (
    "开船预报-标准版",
    "{{bl_no}} - {{pkg_desc}}，此票已开船，请查收附件开船提单以及费用确认书。\n\n"
    "ATD：{{atd}}\nETA：{{eta}}\n\n"
    "收到请核对，有疑问请在3个工作日内提出。\n"
    "核对无误请及时安排付款并提供水单和开票资料。谢谢",
)


def _migrate_speech_templates(conn: sqlite3.Connection) -> None:
    """首次运行时把 data/speech_templates.json 导入为「共用」话术。幂等。"""
    n = conn.execute("SELECT COUNT(*) c FROM speech_templates").fetchone()["c"]
    if n:
        return
    rows: list[tuple[str, str, int]] = []
    try:
        import json
        from bol_forecast.config import DATA_DIR
        p = Path(DATA_DIR) / "speech_templates.json"
        if p.exists():
            data = json.loads(p.read_text(encoding="utf-8"))
            for i, t in enumerate(data.get("templates", [])):
                name = (t.get("name") or f"版本{i + 1}").strip()
                body = t.get("body") or ""
                if body:
                    rows.append((name, body, 1 if t.get("default") else 0))
    except Exception:
        rows = []
    if not rows:
        rows = [(_BUILTIN_SPEECH[0], _BUILTIN_SPEECH[1], 1)]
    for pos, (name, body, dflt) in enumerate(rows):
        conn.execute(
            "INSERT INTO speech_templates"
            "(customer_id, name, body, is_default, position, created_at, updated_at)"
            " VALUES(NULL,?,?,?,?,?,?)",
            (name, body, dflt, pos, _now(), _now()))
    conn.commit()


if __name__ == "__main__":
    print("default customer id =", init_db())
    print("db at", DB_PATH)
