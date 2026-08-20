# -*- coding: utf-8 -*-
"""预报话术渲染与变量库。

模板存储在 SQLite `speech_templates` 表：
  - customer_id IS NULL  → 所有客户共用（可多版本）
  - customer_id 有值     → 该客户专属（可多版本）
选择顺序：指定版本 > 该客户专属默认 > 该客户专属第一条 > 共用默认 > 共用第一条。

占位符语法 {{var}}；变量缺失时渲染为空串并记录在 `missing`，
保证任何情况下都能产出可读文本，绝不因缺字段抛错。
"""
from __future__ import annotations

import re
from datetime import date
from typing import Any

from bol_forecast.data import models as _m
from bol_forecast.generators.map_bl import derive_pkg_unit

_VAR_RE = re.compile(r"\{\{\s*([A-Za-z_][A-Za-z0-9_]*)\s*\}\}")

# ── 可插入变量库（前端「插入字段」按钮据此渲染） ──
SPEECH_VARS: list[dict[str, str]] = [
    {"key": "bl_no", "label": "提单号", "sample": "OOLU2171716520D"},
    {"key": "order_no", "label": "订单号/合同号", "sample": "PO2026-0714"},
    {"key": "customer_po", "label": "客户 PO", "sample": "PO2026-0714-01"},
    {"key": "pkg_desc", "label": "件数描述", "sample": "11CTNS"},
    {"key": "ctns", "label": "件数", "sample": "11"},
    {"key": "pkg_unit", "label": "包装单位", "sample": "CTNS"},
    {"key": "packing_kind", "label": "包装种类", "sample": "纸箱"},
    {"key": "gw", "label": "毛重 KGS", "sample": "2065"},
    {"key": "cbm", "label": "体积 CBM", "sample": "12.87"},
    {"key": "cw", "label": "计费重 KGS", "sample": "2145"},
    {"key": "atd", "label": "开船时间 ATD", "sample": "2026-08-07"},
    {"key": "eta", "label": "预计到港 ETA", "sample": "2026-09-12"},
    {"key": "vessel", "label": "船名航次", "sample": "COSCO FAITH/077E"},
    {"key": "pol", "label": "起运港", "sample": "SHANGHAI"},
    {"key": "pod", "label": "目的港", "sample": "LOS ANGELES"},
    {"key": "container_no", "label": "柜号", "sample": "CCLU7750098"},
    {"key": "seal_no", "label": "封条号", "sample": "SL123456"},
    {"key": "shipper", "label": "发货人", "sample": "上海XX国际物流"},
    {"key": "consignee", "label": "收货人", "sample": "ABC TRADING INC."},
    {"key": "customer_name", "label": "客户名称", "sample": "万年青"},
    {"key": "today", "label": "今天日期", "sample": date.today().isoformat()},
]
_VALID_KEYS = {v["key"] for v in SPEECH_VARS}


def _s(v: Any) -> str:
    """None / 空 → 空串；数字去掉多余尾零。"""
    if v is None:
        return ""
    if isinstance(v, float):
        return f"{v:.4f}".rstrip("0").rstrip(".")
    return str(v).strip()


def build_speech_vars(order: dict, overrides: dict | None = None,
                      customer: dict | None = None) -> dict[str, str]:
    """构造全部可用变量。overrides（界面手改值）优先于 order。"""
    ov = overrides or {}
    order = order or {}

    def pick(*keys):
        for k in keys:
            if ov.get(k) not in (None, ""):
                return ov.get(k)
        for k in keys:
            if order.get(k) not in (None, ""):
                return order.get(k)
        return None

    ctns = pick("ctns")
    unit = derive_pkg_unit(pick("packing_kind"), order.get("remark"))
    # ATD 铁律：仅取轨迹/手填，绝不回退报关单申报日期
    atd = pick("atd", "shipped_date")
    eta = pick("eta")

    v = {
        "bl_no": _s(pick("bl_no")),
        "order_no": _s(pick("order_no", "contract_no")),
        "customer_po": _s(pick("customer_po")),
        "pkg_desc": _s(ov.get("pkg_desc")) or (f"{_s(ctns)}{unit}" if ctns else ""),
        "ctns": _s(ctns),
        "pkg_unit": unit,
        "packing_kind": _s(pick("packing_kind")),
        "gw": _s(pick("gw")),
        "cbm": _s(pick("cbm")),
        "cw": _s(pick("cw", "chargeable_weight")),
        "atd": _s(atd),
        "eta": _s(eta),
        "vessel": _s(pick("picked_vessel", "vessel")),
        "pol": _s(pick("pol")),
        "pod": _s(pick("pod")),
        "container_no": _s(pick("container_no")),
        "seal_no": _s(pick("seal_no")),
        "shipper": _s(pick("shipper_cn", "shipper")),
        "consignee": _s(pick("consignee")),
        "customer_name": _s((customer or {}).get("name")),
        "today": date.today().isoformat(),
    }
    return v


def _pick_template(customer_id: int | None, version_id: Any) -> dict | None:
    """按业务优先级挑选模板。"""
    if version_id not in (None, "", "null"):
        try:
            t = _m.get_speech_template(int(version_id))
            if t:
                return t
        except (TypeError, ValueError):
            pass
    if customer_id:
        own = _m.list_speech_templates(customer_id, include_shared=False)
        if own:
            return next((t for t in own if t["is_default"]), own[0])
    shared = _m.list_speech_templates(None)
    if shared:
        return next((t for t in shared if t["is_default"]), shared[0])
    return None


def render_speech(order: dict, overrides: dict | None = None,
                  version_id: Any = None,
                  customer_id: int | None = None) -> str:
    """渲染话术文本。模板缺失/变量缺失都不抛错。"""
    r = render_speech_detail(order, overrides, version_id, customer_id)
    return r["text"]


def render_speech_detail(order: dict, overrides: dict | None = None,
                         version_id: Any = None,
                         customer_id: int | None = None) -> dict:
    """渲染并返回诊断信息 ``{text, template_id, template_name, missing[], unknown[]}``。"""
    overrides = overrides or {}
    customer = None
    if customer_id:
        try:
            customer = _m.get_customer(int(customer_id))
        except Exception:
            customer = None

    tpl = _pick_template(customer_id, version_id)
    if not tpl:
        return {"text": "（暂无可用话术模板，请在客户信息库中新建）",
                "template_id": None, "template_name": None,
                "missing": [], "unknown": []}

    vars_ = build_speech_vars(order, overrides, customer)
    missing: list[str] = []
    unknown: list[str] = []

    def sub(m: re.Match) -> str:
        k = m.group(1)
        if k not in _VALID_KEYS:
            unknown.append(k)
            return ""
        val = vars_.get(k, "")
        if not val:
            missing.append(k)
        return val

    text = _VAR_RE.sub(sub, tpl.get("body", ""))
    return {
        "text": text,
        "template_id": tpl.get("id"),
        "template_name": tpl.get("name"),
        "missing": sorted(set(missing)),
        "unknown": sorted(set(unknown)),
    }


def list_templates(customer_id: int | None = None) -> list[dict]:
    """供前端下拉使用：专属在前，共用在后。"""
    rows = _m.list_speech_templates(customer_id) if customer_id \
        else _m.list_speech_templates(None)
    return [{"id": t["id"], "name": t["name"],
             "scope": t["scope"], "default": t["is_default"]} for t in rows]
