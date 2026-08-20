# -*- coding: utf-8 -*-
"""电放保函生成：merged order -> 电放提单模版.xlsx 单元格。

模板原是对开船提单的外链公式（A8=开船提单!A8 等）。母表/开船提单副本都不
保证在场，故直接把电放所需字段写成字面值，实现自我包含。

字段映射（模板公式 -> 实际来源）：
  A8  收货人 Consignee        <- consignee
  A12 = A8（同收货人）
  B21 提单号 B/L NO.          <- bl_no
  B22 柜号/封条 Container/Seal <- container_no (+ seal)
  B23 件数/包装 No. of Pkg    <- "{ctns}{unit}"
  B24 发货人 Shipper          <- shipper_cn
  B25 = A12（收货人）
  B26 起运港 POL（值）         <- pol        （A26 为标签，保持不动）
  B27 起运时间 On Board       <- shipped_date
  B28 目的港 POD（值）         <- pod        （A28 为标签，保持不动）
  B29 船名航次 Vessel（值）    <- vessel     （A29 为标签，保持不动）

注：A26/A28/A29 是模板原有标签单元格（"起运港 /POL：" 等），
**保持原样不写、不重排**；对应值写入右侧 B26/B28/B29。
"""
from __future__ import annotations

import logging
import os
from typing import Any

from bol_forecast.config import customer_po_tag, doc_label, safe_filename, template_path
from bol_forecast.generators.map_bl import derive_pkg_unit
from .com_session import com_retry
from .writer_base import XlsxWriter

log = logging.getLogger(__name__)

TLEX_SHEET = "Sheet1"


def _name_only(text: str) -> str:
    """Req1：电放保函收发货人仅显示名称（第一行），去除地址/电话等联系信息。"""
    if not text:
        return ""
    # 多行文本：取第一行（名称）；单行：原样返回
    first = text.split("\n")[0].strip()
    return first


def build_telex_cells(order: dict, *,
                      bl_no_override: str | None = None,
                      shipped_date: str | None = None,
                      seal_no: str | None = None) -> dict[str, Any]:
    ctns = order.get("ctns") or 0
    unit = derive_pkg_unit(order.get("packing_kind"), order.get("remark"))
    bl = bl_no_override or order.get("bl_no") or ""
    # Req1：电放保函只保留名称（第一行）—— address/contact stripped
    shipper = _name_only(order.get("shipper_cn") or "")
    consignee = _name_only(order.get("consignee") or "")
    container = order.get("container_no") or ""
    # 需求3：封条默认生成值 //（与提单 C28 语义一致）；也回退 order 内已有值
    seal = (seal_no or order.get("seal_no") or "//").strip() or "//"
    # B22 柜号/封条：有柜号则 "柜号 / 封条"，无柜号则仅封条
    b22 = f"{container} / {seal}" if container else seal

    # 注意：A26/A28/A29 是标签，保持不动；值写入 B26/B28/B29
    return {
        "A8": consignee,
        "A12": consignee,
        "B21": bl,
        "B22": b22,
        "B23": f"{ctns}{unit}",
        "B24": shipper,
        "B25": consignee,
        "B26": order.get("pol") or "",
        # 起运时间：仅取 ATD（界面手填优先）；无 ATD 留空，禁止用申报日期兜底
        "B27": shipped_date or order.get("shipped_date") or order.get("atd") or "",
        "B28": order.get("pod") or "",
        "B29": order.get("vessel") or "",
    }


def build_telex_context(order: dict, *,
                        bl_no_override: str | None = None,
                        shipped_date: str | None = None,
                        seal_no: str | None = None) -> dict[str, Any]:
    """电放保函 HTML 渲染 context：命名键 dict，模板直接消费。"""
    ctns = order.get("ctns") or 0
    unit = derive_pkg_unit(order.get("packing_kind"), order.get("remark"))
    bl = bl_no_override or order.get("bl_no") or ""
    shipper = _name_only(order.get("shipper_cn") or "")
    consignee = _name_only(order.get("consignee") or "")
    container = order.get("container_no") or ""
    seal = (seal_no or order.get("seal_no") or "//").strip() or "//"
    b22 = f"{container} / {seal}" if container else seal
    return {
        "bl_no": bl,
        "container": b22,
        "package": f"{ctns}{unit}",
        "shipper": shipper,
        "consignee": consignee,
        "pol": order.get("pol") or "",
        "shipped_date": shipped_date or order.get("shipped_date") or order.get("atd") or "",
        "pod": order.get("pod") or "",
        "vessel": order.get("vessel") or "",
    }


@com_retry()
def generate_telex(order: dict, out_dir: str, *,
                   bl_no_override: str | None = None,
                   shipped_date: str | None = None,
                   seal_no: str | None = None,
                   export_pdf: bool = True,
                   excel: object | None = None) -> dict:
    """生成电放保函（EXCEL + 可选 PDF）。html 模式只产 PDF。"""
    bl = bl_no_override or order.get("bl_no") or "BL"
    po_tag = customer_po_tag(order)
    out_name = f"{safe_filename(str(bl))}{po_tag}+{safe_filename(doc_label('telex'))}.xlsx"
    out_xlsx = os.path.join(out_dir, out_name)
    result: dict[str, Any] = {"xlsx": out_xlsx, "pdf": None}

    from bol_forecast.rendering.renderer import renderer_mode
    if renderer_mode() == "html":
        from bol_forecast.rendering.fit_specs import apply_sizes
        from bol_forecast.rendering.html_render import render_doc
        ctx = build_telex_context(order, bl_no_override=bl_no_override,
                                  shipped_date=shipped_date, seal_no=seal_no)
        apply_sizes("telex", ctx)
        out_pdf = os.path.join(out_dir, f"{safe_filename(str(bl))}{po_tag}+{safe_filename(doc_label('telex'))}.pdf")
        try:
            render_doc("telex", ctx, out_pdf)
            result["xlsx"] = None
            result["pdf"] = out_pdf
            log.info("电放保函 HTML 渲染完成: %s", out_pdf)
        except Exception as e:
            log.warning("电放保函 HTML 渲染失败: %s", e)
        return result

    cells = build_telex_cells(order, bl_no_override=bl_no_override,
                              shipped_date=shipped_date, seal_no=seal_no)
    with XlsxWriter(template_path("telex"), out_xlsx, sheet=TLEX_SHEET, excel=excel) as w:
        w.set_cells(cells)
        # ★2026-08-20：WrapText 行高适配 —— 收货人/发货人/柜号等
        for coord in ("A8", "A12", "B24", "B22", "B21", "B23", "B26", "B27",
                      "B28", "B29", "B25"):
            try:
                w.fit_text(coord, cells.get(coord), min_row_height=14.0)
            except Exception as e:
                log.warning("电放 %s wrap行高适配失败(继续): %s", coord, e)
        if export_pdf:
            out_pdf = os.path.join(
                out_dir, f"{safe_filename(str(bl))}{po_tag}+{safe_filename(doc_label('telex'))}.pdf")
            try:
                w.export_pdf(out_pdf)
                result["pdf"] = out_pdf
            except Exception as e:
                log.warning("电放保函 PDF 导出失败: %s", e)
    log.info("电放保函已生成: %s (pdf=%s)", out_xlsx, result["pdf"])
    return result
