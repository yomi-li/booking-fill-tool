# -*- coding: utf-8 -*-
"""客户 / 费用档案 / 习惯库 的 CRUD 操作。

所有函数均对不存在的客户/费用做安全回退，绝不抛错中断生成流程。
"""
from __future__ import annotations

import threading

from .db import get_conn, init_db, _now, _none_float, DEFAULT_CODE

_lock = threading.Lock()


# ---------------- 客户 ----------------
_CUST_COLS = ("id,name,code,created_at,updated_at,customer_type,billing_title,consignee_name,"
              "consignee_addr,consignee_contact,shipper_name,shipper_addr,shipper_contact,"
              "invoice_title,csr_name,credit_days,settlement_node,settlement_days")


def list_customers() -> list[dict]:
    conn = get_conn()
    rows = conn.execute(
        f"SELECT {_CUST_COLS} FROM customers "
        "ORDER BY (code=?) DESC, id ASC", (DEFAULT_CODE,)).fetchall()
    return [dict(r) for r in rows]


def get_customer(cid: int) -> dict | None:
    conn = get_conn()
    r = conn.execute(
        f"SELECT {_CUST_COLS} FROM customers WHERE id=?", (cid,)).fetchone()
    return dict(r) if r else None


def create_customer(name: str, code: str | None = None,
                    extra: dict | None = None,
                    customer_type: str | None = None) -> dict:
    name = (name or "").strip()
    if not name:
        raise ValueError("客户名称不能为空")
    extra = extra or {}
    conn = get_conn()
    with _lock:
        # 默认客户不允许重复创建
        if code == DEFAULT_CODE:
            return get_customer(get_default_customer_id())
        cur = conn.execute(
            "INSERT INTO customers(name, code, created_at, updated_at, customer_type,"
            " billing_title, consignee_name, consignee_addr, consignee_contact,"
            " shipper_name, shipper_addr, shipper_contact, invoice_title, csr_name,"
            " credit_days, settlement_node, settlement_days)"
            " VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (name, code, _now(), _now(), customer_type, extra.get("billing_title"),
             extra.get("consignee_name"), extra.get("consignee_addr"),
             extra.get("consignee_contact"), extra.get("shipper_name"),
             extra.get("shipper_addr"), extra.get("shipper_contact"),
             extra.get("invoice_title"), extra.get("csr_name"),
             int(extra.get("credit_days") or 0),
             extra.get("settlement_node"),
             int(extra.get("settlement_days") or 0)))
        conn.commit()
        return get_customer(int(cur.lastrowid))


def update_customer(cid: int, data: dict) -> dict:
    """整体更新客户资料（名称/代码/收货人/发货人）。不存在则抛错。"""
    conn = get_conn()
    allowed = {"name", "code", "customer_type", "billing_title",
               "consignee_name", "consignee_addr", "consignee_contact",
               "shipper_name", "shipper_addr", "shipper_contact",
               "invoice_title", "csr_name", "credit_days",
               "settlement_node", "settlement_days"}
    sets = {k: v for k, v in (data or {}).items() if k in allowed}
    if not sets:
        return get_customer(cid)
    if "name" in sets and not (sets["name"] or "").strip():
        raise ValueError("客户名称不能为空")
    sets["updated_at"] = _now()
    cols = ", ".join(f"{k}=?" for k in sets)
    with _lock:
        conn.execute(
            f"UPDATE customers SET {cols} WHERE id=?",
            list(sets.values()) + [cid])
        conn.commit()
    return get_customer(cid)


def get_default_customer_id() -> int:
    conn = get_conn()
    r = conn.execute("SELECT id FROM customers WHERE code=?", (DEFAULT_CODE,)).fetchone()
    if r:
        return int(r["id"])
    return init_db()  # 兜底：自动播种


# ---------------- 费用档案 ----------------
def list_charge_profiles(customer_id: int, kind: str) -> list[dict]:
    conn = get_conn()
    rows = conn.execute(
        """SELECT id, customer_id, kind, name, calc_mode, unit, unit_price,
                  qty_var, formula, min_amount, max_amount, position,
                  currency, exchange_rate
           FROM charge_profiles
           WHERE customer_id=? AND kind=?
           ORDER BY position ASC, id ASC""",
        (customer_id, kind)).fetchall()
    return [_row_to_charge(r) for r in rows]


def save_charge_profiles(customer_id: int, kind: str, profiles: list[dict]) -> int:
    """整体替换某客户某类费用（先删后插），返回写入条数。"""
    conn = get_conn()
    with _lock:
        conn.execute(
            "DELETE FROM charge_profiles WHERE customer_id=? AND kind=?",
            (customer_id, kind))
        for pos, r in enumerate(profiles or []):
            conn.execute(
                """INSERT INTO charge_profiles
                   (customer_id, kind, name, calc_mode, unit, unit_price, qty_var,
                    formula, min_amount, max_amount, position,
                    currency, exchange_rate)
                   VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (customer_id, kind, (r.get("name") or "").strip(),
                 (r.get("calc_mode") or "FIXED").upper(),
                 r.get("unit", "") or "",
                 float(r.get("unit_price") or 0),
                 r.get("qty_var", "") or "",
                 r.get("formula", "") or "",
                 _none_float(r.get("min_amount")),
                 _none_float(r.get("max_amount")),
                 pos,
                 (r.get("currency") or "USD").strip().upper() or "USD",
                 _none_float(r.get("exchange_rate")) or 1.0))
        conn.commit()
    return len(profiles or [])


def _row_to_charge(r: dict) -> dict:
    return {
        "id": r["id"],
        "name": r["name"],
        "calc_mode": r["calc_mode"],
        "unit": r["unit"] or "",
        "unit_price": r["unit_price"] or 0,
        "qty_var": r["qty_var"] or "",
        "formula": r["formula"] or "",
        "min_amount": r["min_amount"],
        "max_amount": r["max_amount"],
        # sqlite3.Row 不支持 .get()，用 in + [] 取值
        "currency": r["currency"] if "currency" in r.keys() and r["currency"] else "USD",
        "exchange_rate": float(r["exchange_rate"]) if "exchange_rate" in r.keys() and r["exchange_rate"] is not None else 1.0,
    }


# ---------------- 习惯库 ----------------
def list_habits(customer_id: int) -> dict:
    conn = get_conn()
    rows = conn.execute(
        "SELECT key, value FROM habits WHERE customer_id=?", (customer_id,)).fetchall()
    return {r["key"]: r["value"] for r in rows}


def get_habit(customer_id: int, key: str) -> str | None:
    conn = get_conn()
    r = conn.execute(
        "SELECT value FROM habits WHERE customer_id=? AND key=?",
        (customer_id, key)).fetchone()
    return r["value"] if r else None


def set_habit(customer_id: int, key: str, value: str) -> None:
    key = (key or "").strip()
    if not key:
        raise ValueError("habit key 不能为空")
    conn = get_conn()
    with _lock:
        conn.execute(
            """INSERT INTO habits(customer_id, key, value) VALUES(?,?,?)
               ON CONFLICT(customer_id, key) DO UPDATE SET value=excluded.value""",
            (customer_id, key, value))
        conn.commit()


def delete_habit(customer_id: int, key: str) -> None:
    conn = get_conn()
    conn.execute("DELETE FROM habits WHERE customer_id=? AND key=?",
                 (customer_id, key))
    conn.commit()


def set_habits_bulk(customer_id: int, kv: dict) -> int:
    """批量写入客户习惯字段（起运港/目的港/柜号/包装种类等）。

    值为空串/None 的键视为「清除该习惯」，避免空值覆盖后无法删除。
    返回实际写入（含删除）的键数。
    """
    if not customer_id:
        raise ValueError("customer_id 不能为空")
    conn = get_conn()
    n = 0
    with _lock:
        for k, v in (kv or {}).items():
            k = (k or "").strip()
            if not k:
                continue
            sv = "" if v is None else str(v).strip()
            if sv:
                conn.execute(
                    """INSERT INTO habits(customer_id, key, value) VALUES(?,?,?)
                       ON CONFLICT(customer_id, key)
                       DO UPDATE SET value=excluded.value""",
                    (customer_id, k, sv))
            else:
                conn.execute("DELETE FROM habits WHERE customer_id=? AND key=?",
                             (customer_id, k))
            n += 1
        conn.commit()
    return n


# ---------------- 预报话术模板 ----------------
_SPEECH_COLS = "id, customer_id, name, body, is_default, position, updated_at"


def _row_to_speech(r) -> dict:
    d = dict(r)
    d["is_default"] = bool(d.get("is_default"))
    d["scope"] = "shared" if d.get("customer_id") is None else "customer"
    return d


def list_speech_templates(customer_id: int | None = None,
                          include_shared: bool = True) -> list[dict]:
    """列出可用话术。

    customer_id 为 None      → 仅共用话术
    customer_id 有值         → 该客户专属 + 共用（专属排在前）
    include_shared=False     → 仅该客户专属（用于管理面板的分组展示）
    """
    conn = get_conn()
    if customer_id:
        if include_shared:
            rows = conn.execute(
                f"SELECT {_SPEECH_COLS} FROM speech_templates "
                "WHERE customer_id=? OR customer_id IS NULL "
                "ORDER BY (customer_id IS NULL) ASC, is_default DESC,"
                " position ASC, id ASC", (customer_id,)).fetchall()
        else:
            rows = conn.execute(
                f"SELECT {_SPEECH_COLS} FROM speech_templates WHERE customer_id=?"
                " ORDER BY is_default DESC, position ASC, id ASC",
                (customer_id,)).fetchall()
    else:
        rows = conn.execute(
            f"SELECT {_SPEECH_COLS} FROM speech_templates WHERE customer_id IS NULL"
            " ORDER BY is_default DESC, position ASC, id ASC").fetchall()
    return [_row_to_speech(r) for r in rows]


def get_speech_template(tid: int) -> dict | None:
    conn = get_conn()
    r = conn.execute(
        f"SELECT {_SPEECH_COLS} FROM speech_templates WHERE id=?", (tid,)).fetchone()
    return _row_to_speech(r) if r else None


def save_speech_template(rec: dict) -> dict:
    """新建或更新一条话术。rec 含 id 则更新。

    customer_id 为空 → 共用话术；有值 → 该客户专属。
    is_default 置 1 时，自动清除同作用域内其它默认标记（保证唯一默认）。
    """
    name = (rec.get("name") or "").strip()
    body = rec.get("body") or ""
    if not name:
        raise ValueError("话术名称不能为空")
    if not body.strip():
        raise ValueError("话术正文不能为空")
    cid_raw = rec.get("customer_id")
    cid = int(cid_raw) if cid_raw not in (None, "", "shared", 0, "0") else None
    is_def = 1 if rec.get("is_default") else 0
    pos = int(rec.get("position") or 0)
    tid = rec.get("id")

    conn = get_conn()
    with _lock:
        if tid:
            conn.execute(
                "UPDATE speech_templates SET customer_id=?, name=?, body=?,"
                " is_default=?, position=?, updated_at=? WHERE id=?",
                (cid, name, body, is_def, pos, _now(), int(tid)))
        else:
            cur = conn.execute(
                "INSERT INTO speech_templates"
                "(customer_id, name, body, is_default, position, created_at, updated_at)"
                " VALUES(?,?,?,?,?,?,?)",
                (cid, name, body, is_def, pos, _now(), _now()))
            tid = int(cur.lastrowid)
        if is_def:
            if cid is None:
                conn.execute(
                    "UPDATE speech_templates SET is_default=0"
                    " WHERE customer_id IS NULL AND id<>?", (tid,))
            else:
                conn.execute(
                    "UPDATE speech_templates SET is_default=0"
                    " WHERE customer_id=? AND id<>?", (cid, tid))
        conn.commit()
    return get_speech_template(int(tid))


def delete_speech_template(tid: int) -> bool:
    """删除话术。若删除后共用话术为空，保留最后一条以免生成时无模板可用。"""
    conn = get_conn()
    row = conn.execute(
        "SELECT customer_id FROM speech_templates WHERE id=?", (tid,)).fetchone()
    if not row:
        return False
    if row["customer_id"] is None:
        left = conn.execute(
            "SELECT COUNT(*) c FROM speech_templates WHERE customer_id IS NULL"
        ).fetchone()["c"]
        if left <= 1:
            raise ValueError("至少保留一条共用话术，不能全部删除")
    with _lock:
        conn.execute("DELETE FROM speech_templates WHERE id=?", (tid,))
        conn.commit()
    return True


# ---------------- 历史记录（不存生成文件，仅存元数据） ----------------
def add_shipment(rec: dict) -> int:
    """记录一次生成历史（客户/提单号/关键字段/生成范围）。返回 id。"""
    conn = get_conn()
    with _lock:
        cur = conn.execute(
            """INSERT INTO shipments
               (customer_id, invoice_customer_id, bl_no, vessel, atd, eta,
                ctns, gw, cbm, docs, created_at)
               VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
            (rec.get("customer_id"), rec.get("invoice_customer_id"),
             rec.get("bl_no"), rec.get("vessel"), rec.get("atd"),
             rec.get("eta"), _none_float(rec.get("ctns")),
             _none_float(rec.get("gw")), _none_float(rec.get("cbm")),
             rec.get("docs"), _now()))
        conn.commit()
        return int(cur.lastrowid)


def list_shipments(customer_id: int | None = None, limit: int = 50) -> list[dict]:
    conn = get_conn()
    if customer_id:
        rows = conn.execute(
            "SELECT * FROM shipments WHERE customer_id=? OR invoice_customer_id=?"
            " ORDER BY created_at DESC LIMIT ?",
            (customer_id, customer_id, limit)).fetchall()
    else:
        rows = conn.execute(
            "SELECT * FROM shipments ORDER BY created_at DESC LIMIT ?",
            (limit,)).fetchall()
    return [dict(r) for r in rows]
