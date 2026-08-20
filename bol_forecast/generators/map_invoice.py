# -*- coding: utf-8 -*-
"""INVOICE 生成：merged order + 费用清单 -> INVOICE模版.xlsx。

模板原是对开船提单 / 母表（BILL DRAFT）的外链公式。自我包含：头部与费用
行直接写字面值；SUBTOTAL(G23)/TOTAL(G25) 由 Python 计算写定值。

费用行：模板 A18~G18 表头（ITEM/DESCRIPTION/CURRENCY/AMOUNT），数据区第
19~21 行。G23=SUM(G19:G21)，G25=G23-G24（G24 折扣，默认 0）。

头部关键单元格（依据模板 INVOICE模版.xlsx 坐标核对，2026-08-17 重构）：
  G2  INVOICE NO.    <- bl_no + "A"     （标签 F2）
  G3  ISSUE DATE     <- 今天            （标签 F3）
  G4  DUE DATE       <- 每票手动录入    （标签 F4；无则留空）
  E8  ORDER NUMBER   <- 订单号          （标签 E7）
  G8  BILL NUMBER    <- bl_no           （标签 G7）
  A7  BILL TO 标签 / A9-A11 = BILL TO 3 行值
  E9  CONTAINER NUMBER 标签 / E10 值
  G9  VESSEL 标签        / G10 值
  E11 POL 标签           / E12 值        （★2026-08-17 修复：此前误写 E11 覆盖标签）
  G11 POD 标签           / G12 值        （★2026-08-17 修复：此前误写 G11 覆盖标签）
  E13 QUANTITY 标签 / E14 值
  G13 UNIT 标签     / G14 值
  E15 VOLUME 标签   / E16 值
  G15 GROSS WEIGHT 标签 / G16 值
  费用行 19-21（A/C/F/G 列写值，不变）
  F23 SUBTOTAL 标签 / G23 汇总值
  F25 TOTAL 标签    / G25 汇总值
"""
from __future__ import annotations

import logging
import os
from datetime import date, datetime as dt
from typing import Any

from bol_forecast.config import CFG, customer_po_tag, doc_label, safe_filename, template_path
from bol_forecast.core.formula import build_env, compute_amount
from bol_forecast.generators.charges import clone_defaults
from bol_forecast.generators.map_bl import derive_pkg_unit
from .com_session import com_retry
from .writer_base import XlsxWriter

log = logging.getLogger(__name__)

INV_SHEET = "Sheet1"
FIRST_ROW = 19
LAST_ROW = 21           # 费用行结束（模板容量：19~21 共 3 行）
INV_MAX_ROWS = LAST_ROW - FIRST_ROW + 1
CURRENCY = CFG["charge"].get("invoice_currency", "USD")
# 需求2（2026-08-18）：AMOUNT 栏位货币符号按账单总币别（USD $ / EUR € / GBP £）
CURRENCY_SYMBOLS = {"USD": "$", "EUR": "€", "GBP": "£"}
DEFAULT_CURRENCY_SYMBOL = "$"
# 打印区域：模板实际内容边界 A1:G35（含右侧 H~R 边框/底纹、NOTES/账户信息），
# 固定完整区域 + export_pdf 单页约束(FitToPagesWide=1/Tall=1)
# 保证整页缩放、绝不分页或截断（需求 4）。
# ★2026-08-20 修复：此前 A1:R45 包含空白列 H-Q，导致 FitToPagesWide=1 把内容
# 缩得太小，右侧和下方大量留白。现改为 A1:G35 紧贴实际内容边界。
INV_PRINT_AREA = "A1:G35"

# 列宽缩窄（2026-08-18）：竖向 A4 下内容宽于可打印宽度会右侧贴边/截断，
# 故将 A-G 列宽缩窄（G 列加宽至 16 以容纳提单号/船名长文本），
# 使整表落入 A4 纵向可打印范围。仅调整格子大小，不改格子位置与内容（用户要求竖向排版）。
# ★2026-08-20 修复：此前总宽 91 单位，过窄导致内容缩得太小。
# 调整为模板原始宽度（A=21, B=9.62, C=9.38, D=8.25, E=12.88, F=14, G=20.62），
# 总宽 95.75 单位，占 A4 可打印宽度的 143%，FitToPagesWide=1 会自动缩放到合适大小。
INV_COL_WIDTHS = {
    "A": 21.0,  # BILL TO / 公司名
    "B": 9.6,
    "C": 9.4,   # DESCRIPTION (C:E 合并)
    "D": 8.3,
    "E": 12.9,  # ORDER/POL/QUANTITY 等
    "F": 14.0,  # 标签列
    "G": 20.6,  # 值列 / AMOUNT / BILL NUMBER / VESSEL（需放 16+ 字符长文本）
}


def _split_consignee(consignee: str) -> list[str]:
    """将多行收货人信息按换行拆分为最多 3 行（名称 / 地址 / 联系方式）。"""
    if not consignee:
        return ["", "", ""]
    lines = [l.strip() for l in consignee.split("\n") if l.strip()]
    while len(lines) < 3:
        lines.append("")
    return lines[:3]


def _to_date(val) -> str:
    """归一化手动录入的 Due Date 为 'YYYY-MM-DD'，非法/空返回 ''。"""
    if not val:
        return ""
    s = str(val).strip()[:10]
    try:
        return dt.strptime(s, "%Y-%m-%d").date().isoformat()
    except ValueError:
        return ""


def build_invoice_cells(order: dict, charges: list[dict] | None = None, *,
                        bl_no_override: str | None = None) -> dict[str, Any]:
    bl = bl_no_override or order.get("bl_no") or ""
    ctns = order.get("ctns") or 0
    gw = order.get("gw") or 0
    cbm = order.get("cbm") or 0
    unit = derive_pkg_unit(order.get("packing_kind"), order.get("remark"))
    today = date.today()
    # Due Date：以每票手动录入为准（需求 6：不再与客户绑定存储）
    due_date = _to_date(order.get("due_date_override"))

    if charges is None:
        charges = clone_defaults("invoice")
    env = build_env(order)
    cons_lines = _split_consignee(order.get("consignee") or "")

    # 账单总币别（默认 config invoice_currency，可被前端覆盖）
    total_currency = (order.get("invoice_currency") or CURRENCY).strip().upper() or CURRENCY

    cells: dict[str, Any] = {
        "G2": f"{bl}A",
        "G3": today.isoformat(),
        "G4": due_date,
        "E8": order.get("order_no") or bl,
        "G8": bl,
        "A9": cons_lines[0],
        "A10": cons_lines[1],
        "A11": cons_lines[2],
        "E10": order.get("container_no") or "",
        "G10": order.get("vessel") or "",
        # ★2026-08-17：POL 值位是 E12、POD 值位是 G12（标签在 E11/G11），
        # 此前误写 E11/G11 会把标签覆盖掉且内容错位。
        "E12": order.get("pol") or "",
        "G12": order.get("pod") or "",
        "E14": ctns,
        "G14": unit,
        "E16": round(float(cbm), 3),
        "G16": round(float(gw), 2),
    }

    subtotal = 0.0
    r = FIRST_ROW
    n_used = 0
    for ch in charges:
        # 需求（2026-08-18）：不再设行数上限——超出模板固定区的费用行由
        # generate_invoice 插入新行承接（见 insert_rows），保证 Excel 可手动编辑。
        amt = compute_amount(ch, env)
        cur = (ch.get("currency") or "USD").strip().upper() or "USD"
        rate = float(ch.get("exchange_rate") or 1.0)
        # 需求2：金额×汇率折算为账单总币别金额，填入 AMOUNT（G 列）
        amt_ccy = round(amt * rate, 2)
        cells[f"A{r}"] = ch.get("name", "")
        # 需求4：币别与总账单币别一致时不显示原币别/汇率；
        # 不一致时作为「新增信息」追加到原本 Description（不覆盖）。
        desc_base = (ch.get("desc") or ch.get("name") or "").strip()
        if cur != total_currency:
            extra = f" Orig {amt:,.2f} {cur}"
            if rate != 1.0:
                extra += f" @ {rate}"
            desc_base = f"{desc_base} |{extra}" if desc_base else extra.lstrip()
        cells[f"C{r}"] = desc_base
        cells[f"F{r}"] = cur                      # CURRENCY：该费用币别
        cells[f"G{r}"] = amt_ccy                  # 折算后金额（总币别）
        subtotal += amt_ccy
        r += 1
        n_used += 1

    # 动态行号：SUBTOTAL 紧接费用行之后，TOTAL 在其下 2 行（与模板 R23/R25 对齐，n=3 时）
    subtotal_row = FIRST_ROW + n_used + 1      # 19+n+1 = n+20（n=3 → 23）
    total_row = subtotal_row + 2               # n+22（n=3 → 25）
    # 注意：不写 A17/G17（用户要求不显示在导出的 INVOICE 中）
    cells[f"F{subtotal_row}"] = f"SUBTOTAL({total_currency})"
    cells[f"G{subtotal_row}"] = round(subtotal, 2)    # SUBTOTAL
    cells[f"F{total_row}"] = f"TOTAL({total_currency})"
    cells[f"G{total_row}"] = round(subtotal, 2)       # TOTAL（无折扣）
    # 元数据：COM 插入行、金额格式、print_area 依据
    cells["_inv_subtotal_row"] = subtotal_row
    cells["_inv_total_row"] = total_row
    cells["_inv_charge_count"] = n_used
    cells["_total_currency"] = total_currency
    return cells


def build_invoice_context(order: dict, charges: list[dict] | None, *,
                          bl_no_override: str | None = None,
                          today: date | None = None,
                          due_date: str | None = None) -> dict[str, Any]:
    """INVOICE HTML 渲染 context。"""
    from bol_forecast.generators.charges import clone_defaults
    from bol_forecast.core.formula import build_env, compute_amount
    if today is None:
        today = date.today()
    if due_date is None:
        due_date = today.isoformat()

    bl = bl_no_override or order.get("bl_no") or ""
    bill_no = f"{bl}A"
    if charges is None:
        charges = clone_defaults("invoice")
    env = build_env(order)
    cons_lines = _split_consignee(order.get("consignee") or "")
    ctns = order.get("ctns") or 0
    gw = order.get("gw") or 0
    cbm = order.get("cbm") or 0
    unit = derive_pkg_unit(order.get("packing_kind"), order.get("remark"))

    charge_rows: list[dict[str, Any]] = []
    subtotal = 0.0
    n = 0
    total_currency = (order.get("invoice_currency") or CURRENCY).strip().upper() or CURRENCY
    for ch in charges:
        # 需求（2026-08-18）：无行数上限，HTML 表格自然多行
        amt = round(compute_amount(ch, env), 2)
        cur = (ch.get("currency") or "USD").strip().upper() or "USD"
        rate = float(ch.get("exchange_rate") or 1.0)
        amt_ccy = round(amt * rate, 2)          # 折算为账单总币别金额
        # 需求4：币别与总币别一致时不显示原币别/汇率；不一致时追加到 Description
        desc_base = (ch.get("desc") or ch.get("name") or "").strip()
        desc_en = desc_base
        if cur != total_currency:
            extra = f" Orig {amt:,.2f} {cur}"
            if rate != 1.0:
                extra += f" @ {rate}"
            desc_en = f"{desc_base} |{extra}" if desc_base else extra.lstrip()
        charge_rows.append({
            "name":       ch.get("desc", ch.get("name", "")),
            "desc_en":    desc_en,              # Description（含追加的币别/汇率）
            "unit":       ch.get("unit", ""),
            "unit_price": ch.get("unit_price", 0),
            "qty":        ch.get("qty", 0),
            "currency":   cur,
            "rate":       rate,
            "amount":     amt_ccy,
        })
        subtotal += amt_ccy
        n += 1
    subtotal = round(subtotal, 2)

    return {
        "bill_no":        bill_no,
        "bill_date":      today.isoformat(),
        "due_date":       due_date,
        "bl_no":          bl,
        "container":      order.get("container_no") or "",
        "vessel":         order.get("vessel") or "",
        "pol":            order.get("pol") or "",
        "pod":            order.get("pod") or "",
        "ctns":           ctns,
        "goods_name":     order.get("goods_name") or unit,        # 品名（回退包装单位）
        "gw":             round(float(gw), 2),
        "cbm":            round(float(cbm), 3),
        "bill_to_name":   cons_lines[0] or order.get("consignee") or "",
        "bill_to_contact": "\n".join(filter(None, cons_lines[1:])),
        "charge_rows":    charge_rows,
        "subtotal":       subtotal,
        "total":          subtotal,
        "currency":       CURRENCY,
        "total_currency": total_currency,        # 账单总币别
    }


@com_retry()
def generate_invoice(order: dict, out_dir: str,
                     charges: list[dict] | None = None, *,
                     bl_no_override: str | None = None,
                     export_pdf: bool = True,
                     excel: object | None = None) -> dict:
    """生成 INVOICE（EXCEL + 可选 PDF）。html 模式只产 PDF。"""
    bl = bl_no_override or order.get("bl_no") or "BL"
    po_tag = customer_po_tag(order)
    out_name = f"{safe_filename(str(bl))}{po_tag}+{safe_filename(doc_label('invoice'))}.xlsx"
    out_xlsx = os.path.join(out_dir, out_name)
    result: dict[str, Any] = {"xlsx": out_xlsx, "pdf": None}

    from bol_forecast.rendering.renderer import renderer_mode
    if renderer_mode() == "html":
        from bol_forecast.rendering.fit_specs import apply_sizes
        from bol_forecast.rendering.html_render import render_doc
        ctx = build_invoice_context(order, charges, bl_no_override=bl_no_override)
        apply_sizes("invoice", ctx)
        out_pdf = os.path.join(out_dir, f"{safe_filename(str(bl))}{po_tag}+{safe_filename(doc_label('invoice'))}.pdf")
        try:
            render_doc("invoice", ctx, out_pdf)
            result["xlsx"] = None
            result["pdf"] = out_pdf
            result["subtotal"] = ctx["subtotal"]
            log.info("INVOICE HTML 渲染完成: %s subtotal=%s", out_pdf, ctx["subtotal"])
        except Exception as e:
            log.warning("INVOICE HTML 渲染失败: %s", e)
        return result

    cells = build_invoice_cells(order, charges, bl_no_override=bl_no_override)
    result = {"xlsx": out_xlsx, "pdf": None}
    # 动态行号（build_invoice_cells 元数据）
    subtotal_row = int(cells.get("_inv_subtotal_row") or (FIRST_ROW + 3 + 1))
    total_row = int(cells.get("_inv_total_row") or (FIRST_ROW + 3 + 3))
    charge_count = int(cells.get("_inv_charge_count") or 0)
    total_currency = cells.get("_total_currency") or CURRENCY
    with XlsxWriter(template_path("invoice"), out_xlsx, sheet=INV_SHEET, excel=excel) as w:
        # 需求3：费用行数不设上限——超出模板固定区（19~21 共 3 行）时，
        # 在其下方插入新行承接（下方 SUBTOTAL/账户区随行下移），Excel 仍可手动编辑。
        extra_rows = max(0, charge_count - INV_MAX_ROWS)
        if extra_rows:
            w.insert_rows(FIRST_ROW + INV_MAX_ROWS, extra_rows)   # 在 R22 起插入
        w.set_cells(cells)
        w.set_col_widths(INV_COL_WIDTHS)   # 竖向 A4：缩窄列宽防右侧截断
        # 需求2：AMOUNT 栏位（G 列）货币符号按账单总币别（USD $ / EUR € / GBP £）
        sym = CURRENCY_SYMBOLS.get(total_currency, DEFAULT_CURRENCY_SYMBOL)
        money_fmt = f'"{sym}"#,##0.00'
        for rr in range(FIRST_ROW, FIRST_ROW + charge_count):
            w.set_number_format(f"G{rr}", money_fmt)
        w.set_number_format(f"G{subtotal_row}", money_fmt)   # SUBTOTAL
        w.set_number_format(f"G{total_row}", money_fmt)      # TOTAL
        # 数值列自动缩字：件数/体积/重量/金额/汇总（防 ###，保持完整可读）
        for coord in ("E14", "E16", "G16"):
            try:
                w.fit_cell(coord, cells.get(coord))
            except Exception as e:
                log.warning("INVOICE %s 缩字失败(继续): %s", coord, e)
        for r in range(FIRST_ROW, FIRST_ROW + charge_count):
            try:
                w.fit_cell(f"G{r}", cells.get(f"G{r}"))
            except Exception as e:
                log.warning("INVOICE G%s 缩字失败(继续): %s", r, e)
        for coord in (f"G{subtotal_row}", f"G{total_row}"):
            try:
                w.fit_cell(coord, cells.get(coord))
            except Exception as e:
                log.warning("INVOICE %s 缩字失败(继续): %s", coord, e)
        # 文本溢出检测：BILL NUMBER / VESSEL 等长文本列（2026-08-17）
        for coord in ("G8", "G10"):
            try:
                w.fit_cell(coord, cells.get(coord))
            except Exception as e:
                log.warning("INVOICE %s 文本溢出缩字失败(继续): %s", coord, e)
        # ★2026-08-20：WrapText 行高适配 —— COM 不自适应行高，需显式撑高
        # 避免多行收货人/描述被下一行相邻单元格遮盖。
        for coord in ("A9", "A10", "A11", "C19", "C20", "C21"):
            try:
                w.fit_text(coord, cells.get(coord), min_row_height=14.0)
            except Exception as e:
                log.warning("INVOICE %s wrap行高适配失败(继续): %s", coord, e)
        # 收货人 3 行合并区 A9:D16：确保 A 列高度足够（B/D 同行的 A 列）
        try:
            w.set_row_height(16, 32.0)
        except Exception:
            pass
        if export_pdf:
            out_pdf = os.path.join(
                out_dir, f"{safe_filename(str(bl))}{po_tag}+{safe_filename(doc_label('invoice'))}.pdf")
            try:
                # 需求3：行数增多时整页缩放至单页（FitToPagesTall=1）；
                # print_area 随插入行动态扩展，保证账户/备注区完整入图。
                max_r = max(35, total_row + 6)
                w.export_pdf(out_pdf, print_area=f"A1:G{max_r}",
                             orientation=1, single_page=True)
                result["pdf"] = out_pdf
            except Exception as e:
                # Excel COM 忙（0x800AC472）时 PDF 导出失败 —— 用 HTML 渲染器兜底，
                # 保证 PDF 一定产出（xlsx 已由 COM 生成、内容完整可编辑）。
                log.warning("INVOICE COM PDF 导出失败，改走 HTML 渲染兜底: %s", e)
                try:
                    from bol_forecast.rendering.fit_specs import apply_sizes
                    from bol_forecast.rendering.html_render import render_doc
                    ctx = build_invoice_context(order, charges, bl_no_override=bl_no_override)
                    apply_sizes("invoice", ctx)
                    render_doc("invoice", ctx, out_pdf)
                    result["pdf"] = out_pdf
                except Exception as e2:
                    log.warning("INVOICE HTML 兜底 PDF 也失败: %s", e2)
    log.info("INVOICE 已生成: %s (pdf=%s) rows=%d", out_xlsx, result["pdf"], charge_count)
    return result
