# -*- coding: utf-8 -*-
"""散货提单和预报生成 · 后端路由（命名空间 /api/bol）。

作为「单证提取填充工具」的功能模块挂载：所有接口统一前缀 /api/bol，
与目标系统既有 /api/* 完全隔离、零冲突。鉴权、上传/下载、错误格式
均复用目标系统的 BasicAuthMiddleware / FileResponse / JSONResponse 约定。
"""
from __future__ import annotations

import json
import logging
import os
import time
import traceback
import uuid
from pathlib import Path

from fastapi import APIRouter, File, Form, UploadFile, Request
from fastapi.responses import JSONResponse

from bol_forecast.config import JOBS_DIR
from bol_forecast.parsers.customs_pdf import parse_customs_pdf
from bol_forecast.parsers.booking_xlsx import parse_booking_xlsx
from bol_forecast.parsers.receiving_image import parse_receiving_image, parse_receiving_text
from bol_forecast.parsers.tracking import parse_tracking
from bol_forecast.parsers.merge import merge_order
from bol_forecast.generators.map_bl import generate_bill_of_lading
from bol_forecast.generators.map_telex import generate_telex
from bol_forecast.generators.map_factory import generate_factory
from bol_forecast.generators.map_invoice import generate_invoice
from bol_forecast.generators.speech import (render_speech_detail, list_templates,
                                            SPEECH_VARS)
from bol_forecast.generators import charges as charge_lib
from bol_forecast.data import models as data_models

router = APIRouter(prefix="/api/bol", tags=["散货提单和预报生成"])


def _job_dir() -> Path:
    d = JOBS_DIR / f"web_{int(time.time())}_{uuid.uuid4().hex[:4]}"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _to_int(x):
    """安全转 int；空/非法返回 None。"""
    try:
        return int(x) if x not in (None, "", "null") else None
    except (TypeError, ValueError):
        return None


@router.post("/parse")
async def parse(
    customs: UploadFile | None = File(None),
    booking: UploadFile | None = File(None),
    receiving: UploadFile | None = File(None),
    tracking: UploadFile | None = File(None),
    tracking_text: str | None = Form(None),
    receiving_text: str | None = Form(None),
    customer_id: str | None = Form(None),
):
    """上传报关单 PDF + 托书 XLSX + (可选)收货截图 + 轨迹截图/文本，返回合并后的订单字段。"""
    job = _job_dir()

    customs_data: dict = {"fields": {}, "warnings": ["未上传报关单"], "goods": [], "containers": []}
    if customs and customs.filename:
        cpath = job / "customs.pdf"
        await customs.seek(0)
        cpath.write_bytes(await customs.read())
        if customs.filename.lower().endswith(".xls") and not customs.filename.lower().endswith(".xlsx"):
            return JSONResponse(
                {"error": "报关单需为 .pdf 格式。请检查上传文件。"},
                status_code=400)
        try:
            customs_data = parse_customs_pdf(str(cpath))
        except Exception as e:
            logging.getLogger(__name__).warning("报关单解析跳过: %s", e)
            customs_data = {"fields": {}, "warnings": [f"报关单解析失败: {e}"], "goods": [], "containers": []}

    booking_data: dict = {"fields": {}, "warnings": ["未上传托书"], "goods": []}
    if booking and booking.filename:
        bpath = job / "booking.xlsx"
        await booking.seek(0)
        bbytes = await booking.read()
        bpath.write_bytes(bbytes)
        if booking.filename.lower().endswith(".xls") \
                and not booking.filename.lower().endswith(".xlsx"):
            return JSONResponse(
                {"error": "托书需为 .xlsx 格式（不支持旧版 .xls）。请用 Excel 另存为 .xlsx 后重试。"},
                status_code=400)
        try:
            booking_data = parse_booking_xlsx(str(bpath))
        except Exception as e:
            logging.getLogger(__name__).warning("托书解析跳过: %s", e)
            booking_data = {"fields": {}, "warnings": [f"托书解析失败: {e}"], "goods": []}

    rcv_data: dict = {}
    if receiving_text and receiving_text.strip():
        try:
            rcv_data = parse_receiving_text(receiving_text.strip())
        except Exception as e:
            logging.getLogger(__name__).warning("收货文本解析跳过: %s", e)
            rcv_data = {"_source": "error", "warnings": [f"收货文本解析失败: {e}"]}
    elif receiving and receiving.filename:
        rpath = job / f"receiving_{receiving.filename}"
        rpath.write_bytes(await receiving.read())
        try:
            rcv_data = parse_receiving_image(str(rpath))
        except Exception as e:
            logging.getLogger(__name__).warning("收货数据解析跳过: %s", e)
            rcv_data = {"_source": "error", "warnings": [f"收货截图解析失败: {e}"]}

    trk_data: dict = {}
    if tracking_text and tracking_text.strip():
        try:
            trk_data = parse_tracking(tracking_text.strip(), is_text=True)
        except Exception as e:
            logging.getLogger(__name__).warning("轨迹文本解析跳过: %s", e)
            trk_data = {"_source": "error", "warnings": [str(e)]}
    elif tracking and tracking.filename:
        tpath = job / f"tracking_{tracking.filename}"
        tpath.write_bytes(await tracking.read())
        try:
            trk_data = parse_tracking(str(tpath), is_text=False)
        except Exception as e:
            logging.getLogger(__name__).warning("轨迹图片解析跳过: %s", e)
            trk_data = {"_source": "error", "warnings": [str(e)]}

    try:
        order = merge_order(customs_data, booking_data,
                            receiving=rcv_data, tracking=trk_data)
    except Exception as e:
        traceback.print_exc()
        logging.getLogger(__name__).exception("解析失败")
        return JSONResponse(
            {"error": f"解析失败：{type(e).__name__} {e}。请检查文件是否为有效的报关单PDF与托书xlsx。"},
            status_code=400)

    warns = list(order.get("warnings", []))
    if not order.get("bl_no"):
        warns.append("未提取到提单号(B/L NO.)，请在字段确认页手动填写。")
    if not order.get("ctns"):
        warns.append("未提取到件数，请在收货数据区手动填写。")
    for k, label in (("pol", "起运港"), ("pod", "目的港"),
                     ("container_no", "柜号")):
        if not order.get(k):
            warns.append(f"未识别到{label}，可在字段确认区手动填写并保存。")
    # ── Req2：收货数据优先级透明提示（截图 > 托书，逐字段回退） ──
    if receiving and receiving.filename:
        rcv_ok = any(k in rcv_data for k in
                     ("pkg_count", "gross_weight", "volume_cbm", "chargeable_weight"))
        if not rcv_ok:
            warns.append("收货截图未识别到有效收货数据，件/重/体已回退使用托书。")
    else:
        warns.append("未上传收货截图，件/重/体直接取自托书（如需以收货截图为准请粘贴截图）。")

    order["warnings"] = warns

    cid = _to_int(customer_id)
    habits = {}
    if cid:
        try:
            habits = data_models.list_habits(cid)
            for k in ("pol", "pod", "container_no", "packing_kind"):
                if not order.get(k) and habits.get(k):
                    order[k] = habits[k]
                    warns.append(f"{k} 取自客户习惯值「{habits[k]}」，请确认。")
        except Exception as e:
            logging.getLogger(__name__).warning("习惯值回填失败: %s", e)

    return JSONResponse({
        "order": order,
        "vessel_candidates": order.get("vessel_candidates", []),
        "warnings": warns,
        "speech_templates": list_templates(cid),
        "speech_vars": SPEECH_VARS,
        "habits": habits,
        "receiving": {k: v for k, v in rcv_data.items() if not k.startswith("_")},
        "tracking": {k: v for k, v in trk_data.items() if not k.startswith("_")},
        "customers": data_models.list_customers(),
    })


# ---------------- 客户 / 费用档案 / 习惯库 API ----------------
@router.get("/customers")
def api_list_customers():
    return JSONResponse({"customers": data_models.list_customers()})


@router.post("/customers")
async def api_create_customer(request: Request):
    body = await request.json()
    try:
        cust = data_models.create_customer(
            body.get("name", ""), body.get("code"),
            customer_type=body.get("customer_type"),
            extra={"billing_title": body.get("billing_title")})
    except ValueError as e:
        return JSONResponse({"error": str(e)}, status_code=400)
    return JSONResponse({"customer": cust})


@router.get("/charge-profiles")
def api_get_charge_profiles(customer_id: int, kind: str):
    rows = charge_lib.load_charges(kind, customer_id)
    return JSONResponse({"profiles": rows})


@router.put("/charge-profiles")
async def api_save_charge_profiles(request: Request):
    body = await request.json()
    cid = int(body.get("customer_id"))
    kind = body.get("kind")
    profiles = body.get("profiles", [])
    if kind not in ("factory", "invoice"):
        return JSONResponse({"error": "kind 必须为 factory 或 invoice"},
                            status_code=400)
    n = data_models.save_charge_profiles(cid, kind, profiles)
    return JSONResponse({"ok": True, "count": n})


@router.get("/habits")
def api_get_habits(customer_id: int):
    return JSONResponse({"habits": data_models.list_habits(customer_id)})


@router.post("/habits")
async def api_set_habit(request: Request):
    body = await request.json()
    try:
        data_models.set_habit(int(body.get("customer_id")),
                              body.get("key", ""), str(body.get("value", "")))
    except (ValueError, TypeError) as e:
        return JSONResponse({"error": str(e)}, status_code=400)
    return JSONResponse({"ok": True})


@router.post("/habits/bulk")
async def api_set_habits_bulk(request: Request):
    body = await request.json()
    cid = _to_int(body.get("customer_id"))
    if not cid:
        return JSONResponse({"error": "请先选择客户后再保存字段"}, status_code=400)
    try:
        n = data_models.set_habits_bulk(cid, body.get("values") or {})
    except (ValueError, TypeError) as e:
        return JSONResponse({"error": str(e)}, status_code=400)
    return JSONResponse({"ok": True, "count": n,
                         "habits": data_models.list_habits(cid)})


# ---------------- 预报话术管理 ----------------
@router.get("/speech-vars")
def api_speech_vars():
    return JSONResponse({"vars": SPEECH_VARS})


@router.get("/speech-templates")
def api_list_speech(customer_id: str | None = None, scope: str = "all"):
    cid = _to_int(customer_id)
    if scope == "shared":
        rows = data_models.list_speech_templates(None)
    elif scope == "customer":
        if not cid:
            return JSONResponse({"templates": []})
        rows = data_models.list_speech_templates(cid, include_shared=False)
    else:
        rows = data_models.list_speech_templates(cid)
    return JSONResponse({"templates": rows})


@router.put("/speech-templates")
async def api_save_speech(request: Request):
    body = await request.json()
    try:
        t = data_models.save_speech_template(body)
    except (ValueError, TypeError) as e:
        return JSONResponse({"error": str(e)}, status_code=400)
    return JSONResponse({"template": t})


@router.delete("/speech-templates/{tid}")
def api_delete_speech(tid: int):
    try:
        ok = data_models.delete_speech_template(tid)
    except ValueError as e:
        return JSONResponse({"error": str(e)}, status_code=400)
    if not ok:
        return JSONResponse({"error": "话术不存在"}, status_code=404)
    return JSONResponse({"ok": True})


@router.post("/speech-preview")
async def api_speech_preview(request: Request):
    from bol_forecast.generators.speech import build_speech_vars, _VAR_RE, _VALID_KEYS
    body = await request.json()
    order = body.get("order") or {}
    ov = body.get("overrides") or {}
    cid = _to_int(body.get("customer_id"))
    cust = data_models.get_customer(cid) if cid else None
    vars_ = build_speech_vars(order, ov, cust)
    samples = {v["key"]: v["sample"] for v in SPEECH_VARS}
    use_sample = bool(body.get("use_sample"))
    missing: list[str] = []

    def sub(m):
        k = m.group(1)
        if k not in _VALID_KEYS:
            return f"[未知变量:{k}]"
        val = vars_.get(k) or ""
        if not val:
            missing.append(k)
            return samples.get(k, "") if use_sample else ""
        return val

    text = _VAR_RE.sub(sub, body.get("body") or "")
    return JSONResponse({"text": text, "missing": sorted(set(missing))})


# ---------------- 客户全量资料 + 历史记录 ----------------
@router.put("/customers")
async def api_upsert_customer(request: Request):
    body = await request.json()
    cid = body.get("id")
    if cid:
        try:
            cust = data_models.update_customer(int(cid), body)
        except ValueError as e:
            return JSONResponse({"error": str(e)}, status_code=400)
        return JSONResponse({"customer": cust})
    try:
        cust = data_models.create_customer(
            body.get("name", ""), body.get("code"),
            customer_type=body.get("customer_type"),
            extra={
                "billing_title": body.get("billing_title"),
                "consignee_name": body.get("consignee_name"),
                "consignee_addr": body.get("consignee_addr"),
                "consignee_contact": body.get("consignee_contact"),
                "shipper_name": body.get("shipper_name"),
                "shipper_addr": body.get("shipper_addr"),
                "shipper_contact": body.get("shipper_contact"),
                "invoice_title": body.get("invoice_title"),
                "csr_name": body.get("csr_name"),
            })
    except ValueError as e:
        return JSONResponse({"error": str(e)}, status_code=400)
    return JSONResponse({"customer": cust})


@router.get("/shipments")
def api_list_shipments(customer_id: int | None = None):
    rows = data_models.list_shipments(customer_id)
    return JSONResponse({"shipments": rows})


@router.delete("/customers/{cid}")
def api_delete_customer(cid: int):
    if cid == data_models.get_default_customer_id():
        return JSONResponse({"error": "默认模板不可删除"}, status_code=400)
    from bol_forecast.data import db as _db
    conn = _db.get_conn()
    conn.execute("DELETE FROM customers WHERE id=?", (cid,))
    conn.commit()
    return JSONResponse({"ok": True})


@router.post("/generate")
async def generate(request: Request):
    """接收订单字段 + 覆盖项 + 勾选范围，生成所选单证，返回下载链接。

    铁律：导出文件必须与「字段确认栏位」最新编辑值完全一致 —— 任何前端可编辑
    且后端生成器会读取的字段，都要先把 overrides 写回 order 再交给生成器。
    这样确保：
      • OCR/parse 阶段冻的值再晚到，也会被用户编辑覆盖；
      • 字段栏位临时改值（件数/计费重/收货人等），不会无声丢失。
    """
    body = await request.json()
    order = body.get("order", {})
    ov = body.get("overrides", {})
    sel = body.get("selections", ["bl"])
    speech_version = body.get("speech_version")

    def _has(k):
        """仅当 overrides 显式给出非空值时写回 order。"""
        v = ov.get(k)
        if v is None:
            return False
        if isinstance(v, str) and v.strip() == "":
            return False
        return True

    # —— 抬头：bl_no / atd / eta / pol / pod / container_no ——
    # （需求3：包装种类栏位已取消，不再接受 overrides 写回；
    #   packing_kind 仍由解析阶段提取，仅用于推导件数单位。）
    for k in ("bl_no", "atd", "eta",
              "pol", "pod", "container_no"):
        if _has(k):
            order[k] = ov[k]
    # —— 船名航次（来自界面船名输入框，键名 picked_vessel）——
    if _has("picked_vessel"):
        order["vessel"] = ov["picked_vessel"]
    # —— 收发货人 / 成交方式 / 订单号 / 起运时间 / 封条 / 包装描述 / 品名 ——
    for k in ("consignee", "shipper_cn", "trade_term", "order_no",
              "shipped_date", "seal_no", "pkg_desc", "goods_name",
              "customer_po", "customer_po_on_bl"):
        if _has(k):
            order[k] = ov[k]
    # FOB 铁律：consignee = notify party — 任何一方被覆盖都同步
    if _has("consignee") and not _has("notify_party"):
        order["notify_party"] = ov["consignee"]
    # —— 件重体（这是 OCR/界面编辑最频繁的字段，原 collect 漏 chargeable_weight）——
    for k in ("ctns", "gw", "cbm", "chargeable_weight"):
        if _has(k):
            order[k] = ov[k]

    bl = ov.get("bl_no") or order.get("bl_no") or "BL"
    # 装船时间铁律：手填 > 轨迹 ATD；绝不回退报关单申报日期，无则留空
    shipped = (ov.get("shipped_date") or ov.get("atd")
               or order.get("shipped_date") or order.get("atd") or "")
    order["shipped_date"] = shipped
    seal = ov.get("seal_no")
    fac_cid = _to_int(ov.get("factory_customer_id") or order.get("factory_customer_id")
                       or ov.get("customer_id"))
    inv_cid = _to_int(ov.get("invoice_customer_id") or order.get("invoice_customer_id")
                       or ov.get("customer_id"))
    primary = fac_cid or inv_cid
    if primary:
        cust = data_models.get_customer(primary)
        if cust:
            # 用户已编辑过收发货人时，不覆盖；否则从客户默认值带出
            if cust.get("consignee_name") and not _has("consignee"):
                order["consignee"] = cust["consignee_name"]
                order["notify_party"] = cust["consignee_name"]
            if cust.get("shipper_name") and not _has("shipper_cn"):
                order["shipper_cn"] = cust["shipper_name"]
            # Req4：工厂客户的开票抬头 / 客服名字 → 工厂账单 TO / FROM
            if cust.get("invoice_title"):
                order["invoice_title"] = cust["invoice_title"]
            if cust.get("csr_name"):
                order["csr_name"] = cust["csr_name"]
            # Attn 联系方式：取客户的 consignee_contact（工厂联系人），供工厂账单 B8 列
            if cust.get("consignee_contact"):
                order["attn_contact"] = cust["consignee_contact"]
            elif cust.get("shipper_contact"):
                order["attn_contact"] = cust["shipper_contact"]
    # 需求6：INVOICE 的 Due Date 以每票手动录入为准，不再从客户库结算账期推导。
    # 前端在 INVOICE 客户下方手录 due_date，直接作为覆盖写入。
    if ov.get("due_date"):
        order["due_date_override"] = ov["due_date"]
    # 需求2（2026-08-18）：账单总币别以每票手录为准（默认 USD）
    if ov.get("invoice_currency"):
        order["invoice_currency"] = str(ov["invoice_currency"]).strip().upper()
    charges_fac = charge_lib.load_charges("factory", fac_cid)
    # 需求6：INVOICE 费用标准以每票手动录入为准（前端 INVOICE 客户下方手录区）；
    # 前端将 invoice_profiles 放在 overrides 内提交，兼容 body 顶层两种位置。
    inv_profiles = body.get("invoice_profiles") or ov.get("invoice_profiles")
    charges_inv = (inv_profiles if isinstance(inv_profiles, list) and inv_profiles
                   else charge_lib.load_charges("invoice", inv_cid))

    job = _job_dir()
    files = []
    try:
        # ★2026-08-18：4 个生成器统一走进程内单例 Excel（launch_excel 幂等，
        # 生成器内部经 XlsxWriter 自动获取）。不再在此传 excel 参数——
        # 忙态自愈（com_retry 杀实例重启）后，新实例由生成器内部重新获取，
        # 外部句柄已失效。仅 COM 渲染器需要 Excel，html 模式无需。
        from bol_forecast.rendering.renderer import renderer_mode as _rmode
        if _rmode() == "com":
            from bol_forecast.generators.com_session import launch_excel
            launch_excel()   # 预热单例（幂等；自愈后首次调用即重建）
        if "bl" in sel:
            r = generate_bill_of_lading(order, str(job),
                                        bl_no_override=bl, shipped_date=shipped,
                                        seal_no=seal,
                                        customer_po_on_bl=order.get("customer_po_on_bl", True))
            files.append(("提单核对件", r["xlsx"], r["pdf"]))
        if "telex" in sel:
            r = generate_telex(order, str(job), bl_no_override=bl,
                               shipped_date=shipped, seal_no=seal)
            files.append(("电放保函", r["xlsx"], r["pdf"]))
        if "factory" in sel:
            r = generate_factory(order, str(job), charges=charges_fac,
                                 bl_no_override=bl, shipped_date=shipped)
            files.append(("工厂账单", r["xlsx"], r["pdf"]))
        if "invoice" in sel:
            r = generate_invoice(order, str(job), charges=charges_inv,
                                 bl_no_override=bl)
            files.append(("INVOICE", r["xlsx"], r["pdf"]))
    except Exception as e:
        log.exception("生成异常: %s", e)
        return JSONResponse({"error": f"生成失败: {e}"}, status_code=500)
    except BaseException as e:
        # ExcelBusyError 继承 BaseException，避免被 ASGI 框架转成 HTML 500 页面
        # 导致前端 JSON.parse 失败。此处显式捕获并返回 JSON。
        log.exception("生成严重异常(BaseException): %s", e)
        return JSONResponse({"error": f"生成失败: {e}"}, status_code=500)

    out = []
    for label, xlsx, pdf in files:
        item = {"label": label}
        # HTML 渲染器只产 PDF（xlsx 为 None）；COM 路径产出 xlsx（可能带 pdf）
        if xlsx:
            item["xlsx"] = "/bol-files/" + os.path.relpath(xlsx, str(JOBS_DIR)).replace("\\", "/")
        if pdf:
            item["pdf"] = "/bol-files/" + os.path.relpath(pdf, str(JOBS_DIR)).replace("\\", "/")
        out.append(item)

    speech_text, speech_info = None, None
    if "speech" in sel:
        d = render_speech_detail(order, ov, speech_version, primary)
        speech_text = d["text"]
        speech_info = {"template_name": d["template_name"],
                       "missing": d["missing"], "unknown": d["unknown"]}

    try:
        data_models.add_shipment({
            "customer_id": fac_cid,
            "invoice_customer_id": inv_cid,
            "bl_no": bl,
            "vessel": order.get("vessel"),
            "atd": shipped or order.get("atd"),
            "eta": order.get("eta"),
            "ctns": order.get("ctns"),
            "gw": order.get("gw"),
            "cbm": order.get("cbm"),
            "docs": ",".join(sel),
        })
    except Exception as e:
        logging.getLogger(__name__).warning("历史记录写入失败(忽略): %s", e)

    return JSONResponse({"files": out, "speech": speech_text,
                         "speech_info": speech_info})
