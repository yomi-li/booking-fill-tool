# -*- coding: utf-8 -*-
"""默认费用清单（工厂账单 / INVOICE）。

用于 M3 产出可见效果；M4 会改为从客户习惯库持久化读取，此处保留一份
合理默认值。每条费用：
  name      费用名称（写入 A 列）
  calc_mode FIXED | UNIT_PRICE | FORMULA
  unit      单位（工厂账单 C 列，可为空）
  unit_price 单价（UNIT_PRICE 用）
  qty_var   数量变量（UNIT_PRICE 用，CBM/GW/CTNS...）
  formula   公式（FORMULA 用，变量见 formula.ALLOWED_VARS）
  min_amount / max_amount 保底 / 封顶
"""
from __future__ import annotations

import copy

from bol_forecast.data import db as _db
from bol_forecast.data import models as _models

# 工厂账单（RMB）—— 散货 LCL 常见费用
DEFAULT_FACTORY_CHARGES = [
    {"name": "海运费 Sea Freight", "calc_mode": "UNIT_PRICE",
     "unit": "CBM", "unit_price": 100.0, "qty_var": "CBM",
     "min_amount": 300.0},
    {"name": "拼箱费 CFS Charges", "calc_mode": "UNIT_PRICE",
     "unit": "CBM", "unit_price": 18.0, "qty_var": "CBM"},
    {"name": "码头操作费 THC", "calc_mode": "FIXED",
     "unit": "票", "unit_price": 200.0},
    {"name": "报关费 Customs Clearance", "calc_mode": "FIXED",
     "unit": "票", "unit_price": 150.0},
    {"name": "文件费 DOC", "calc_mode": "FIXED",
     "unit": "票", "unit_price": 100.0},
    {"name": "订舱费 Booking Fee", "calc_mode": "FIXED",
     "unit": "票", "unit_price": 80.0},
]

# INVOICE（USD）
DEFAULT_INVOICE_CHARGES = [
    {"name": "FREIGHT", "calc_mode": "UNIT_PRICE",
     "unit": "CBM", "unit_price": 35.0, "qty_var": "CBM",
     "min_amount": 50.0,
     "currency": "USD", "exchange_rate": 1.0},
    {"name": "THC", "calc_mode": "FIXED",
     "unit": "SHIPMENT", "unit_price": 120.0,
     "currency": "USD", "exchange_rate": 1.0},
    {"name": "DOC FEE", "calc_mode": "FIXED",
     "unit": "SHIPMENT", "unit_price": 50.0,
     "currency": "USD", "exchange_rate": 1.0},
]


def clone_defaults(kind: str) -> list[dict]:
    src = DEFAULT_FACTORY_CHARGES if kind == "factory" else DEFAULT_INVOICE_CHARGES
    return copy.deepcopy(src)


def load_charges(kind: str, customer_id: int | None = None) -> list[dict]:
    """优先读取指定客户的费用档案；无则回退默认客户；再无则回退硬编码默认。

    用于 M4 费用引擎持久化：客户习惯库里的费用优先于写死的 DEFAULT_*。
    """
    _db.init_db()
    cid = customer_id if customer_id else _models.get_default_customer_id()
    rows = _models.list_charge_profiles(cid, kind)
    if rows:
        return rows
    # 回退到默认客户的档案（yomi 可能改过全局默认）
    default_rows = _models.list_charge_profiles(_models.get_default_customer_id(), kind)
    if default_rows:
        return default_rows
    return clone_defaults(kind)
