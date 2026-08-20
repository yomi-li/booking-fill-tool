# -*- coding: utf-8 -*-
"""合并报关单 + 托书 + 收货数据 + 轨迹，产出统一订单对象。

报关单提供：BL 抬头（提单号、收发件人、船名、港口、成交方式、柜号…）
托书提供：件/重/体聚合（ctns / gw / nw / cbm）、PO 列表、逐箱品名
收货数据（可选）：件数/实重/体积，**优先于**托书（实际承运口径）
轨迹（可选）：ATD / ETA / 船名航次

合并原则：
  - 抬头字段以报关单为准；托书缺则留空（不污染）。
  - 件重体：收货数据 > 托书 > 报关单（口径优先级）。
  - 船名：轨迹 > 界面勾选 > 报关单候选推荐。
  - ATD/ETA：仅来自轨迹数据。
  - 装船时间 shipped_date == ATD（铁律）：**禁止回退报关单申报日期**，
    无 ATD 则留空，由用户在界面手填。
"""
from __future__ import annotations

from typing import Any


def merge_order(
    customs: dict,
    booking: dict,
    picked_vessel: str | None = None,
    receiving: dict | None = None,
    tracking: dict | None = None,
) -> dict:
    cf = customs.get("fields", {})
    bf = booking.get("fields", {})
    rcv = receiving or {}
    trk = tracking or {}

    # ---- 收货数据（优先于托书） ----
    rcv_ctns = rcv.get("pkg_count")
    rcv_gw = rcv.get("gross_weight")
    rcv_cbm = rcv.get("volume_cbm")
    rcv_cw = rcv.get("chargeable_weight")

    # ---- 轨迹数据 ----
    trk_atd = trk.get("atd")
    trk_eta = trk.get("eta")
    trk_vessel = trk.get("vessel_voyage")

    # ---- 船名航次：报关单优先于轨迹（Req6），最终导出以字段编辑栏为准 ----
    candidates = customs.get("vessel_candidates", [])
    vessel = None
    if candidates:
        rec = next((c for c in candidates if c.get("recommend")), candidates[0])
        vessel = rec.get("value")
    if not vessel:
        vessel = trk_vessel or picked_vessel

    # consignee == notify party（铁律）
    consignee = cf.get("consignee") or cf.get("consignee_raw")

    order = {
        "bl_no": cf.get("bl_no"),
        "shipper_cn": cf.get("shipper_cn"),
        "consignee": consignee,
        "notify_party": consignee,           # FOB 条款：consignee = notify
        "vessel": vessel,
        "vessel_candidates": candidates,
        "pol": cf.get("pol"),
        "pod": cf.get("pod"),
        "trade_term": cf.get("trade_term"),
        "departure_port": cf.get("departure_port"),
        "dest_country": cf.get("dest_country"),
        "export_date": cf.get("export_date"),
        "export_date_iso": cf.get("export_date_iso"),
        "contract_no": cf.get("contract_no"),
        "container_no": cf.get("container_no"),
        "cntr_qty": cf.get("cntr_qty"),
        "cntr_desc": cf.get("cntr_desc"),
        "customs_no": cf.get("customs_no"),
        "packing_kind": cf.get("packing_kind"),
        "remark": cf.get("remark"),
        "order_no": cf.get("contract_no") or cf.get("bl_no"),
        # 件重体（优先级：收货数据 > 托书）
        "ctns": rcv_ctns if rcv_ctns is not None else bf.get("ctns"),
        "gw": rcv_gw if rcv_gw is not None else bf.get("gw"),
        "nw": bf.get("nw"),  # 净重仅托书有
        "cbm": rcv_cbm if rcv_cbm is not None else bf.get("cbm"),
        "chargeable_weight": rcv_cw,  # 计费重（收货数据独有）
        # 轨迹
        "atd": trk_atd,
        "eta": trk_eta,
        "atd_rule": trk.get("atd_rule"),
        "eta_rule": trk.get("eta_rule"),
        # 装船时间：仅等于 ATD；无 ATD 则为 None（绝不用申报日期兜底）
        "shipped_date": trk_atd,
        # 收货原始数据（供 UI 展示/编辑）
        "receiving_raw": {k: v for k, v in rcv.items() if not k.startswith("_")},
        "tracking_raw": {k: v for k, v in trk.items() if not k.startswith("_")},
        "po_list": bf.get("po_list", []),
        # 商品明细（两源都保留，标注来源）
        "goods_customs": customs.get("goods", []),
        "goods_booking": booking.get("goods", []),
        "warnings": (list(customs.get("warnings", []))
                     + list(booking.get("warnings", []))
                     + [f"收货数据：{w}" for w in rcv.get("warnings", [])]
                     + [f"轨迹：{w}" for w in trk.get("warnings", [])]),
    }
    if not trk_atd:
        order["warnings"].append(
            "未识别到 ATD，装船时间留空（按规则不使用报关单申报日期兜底），请手动填写。")
    return order
