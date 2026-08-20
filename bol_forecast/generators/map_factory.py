# -*- coding: utf-8 -*-
"""工厂账单（运费确认书）生成：merged order + 费用清单 -> 工厂账单模版.xlsx。

模板原是对开船提单 / 母表的外链公式。自我包含：头部字段与费用行直接写
字面值；汇总 G31 由 Python 计算后写定值，不依赖 Excel 公式重算。

费用行：模板 A15~I15 为表头（费用名称/单位/金额/数量/币别/总额/备注），
数据区为第 16~30 行。G31 = SUM(G16:H30) -> 写为 Python 汇总定值。
"""
from __future__ import annotations

import logging
import os
from datetime import date
from typing import Any

from bol_forecast.config import CFG, customer_po_tag, doc_label, safe_filename, template_path
from bol_forecast.core.formula import build_env, compute_amount
from bol_forecast.generators.charges import clone_defaults
from bol_forecast.generators.map_bl import derive_pkg_unit
from .com_session import com_retry
from .writer_base import XlsxWriter

log = logging.getLogger(__name__)

FAC_SHEET = "Sheet1"
FIRST_ROW = 16          # 费用行起始
LAST_ROW = 30           # 费用行结束（模板容量：16~30 共 15 行）
FAC_MAX_ROWS = LAST_ROW - FIRST_ROW + 1
CURRENCY = CFG["charge"].get("factory_currency", "RMB")
# 打印区域：模板全部内容到 K51（含右侧 J/K 边框、账户信息、注意事项、账单日期），
# 固定完整区域 + export_pdf 单页约束(FitToPagesWide=1/Tall=1)
# 保证整页缩放、绝不分页或截断（需求 4）。此前 A1:I51 会截断右侧边框。
FAC_PRINT_AREA = "A1:K51"

# 列宽缩窄（2026-08-18）：竖向 A4 下内容宽于可打印宽度会右侧贴边/截断，
# 故将 A-K 列宽从模板原值缩窄至约 70.5 字符，使整表落入 A4 纵向可打印范围。
# 仅调整格子大小，不改格子位置与内容（用户要求竖向排版）。
FAC_COL_WIDTHS = {
    "A": 7.0,   # was 7.60  — 费用名称(左半)
    "B": 7.0,   # was 8.43  — 费用名称(右半, A:B 合并)
    "C": 6.0,   # was 7.40  — 单位
    "D": 7.0,   # was 8.43  — 金额
    "E": 6.5,   # was 8.43  — 数量
    "F": 6.0,   # was 8.43  — 币别
    "G": 7.0,   # was 8.60  — 总额(左半)
    "H": 5.5,   # was 6.18  — 总额(右半, G:H 合并)
    "I": 7.0,   # was 8.43  — 备注(左)
    "J": 5.0,   # was 5.40  — 备注(中, I:K 合并)
    "K": 6.5,   # was 8.43  — 备注(右, I:K 合并)
}


def _fmt_gw(gw: float) -> str:
    s = f"{float(gw):.2f}".rstrip("0").rstrip(".")
    return s if s else "0"


def _marks(order: dict) -> str:
    ctns = order.get("ctns") or 0
    gw = order.get("gw") or 0
    cbm = order.get("cbm") or 0
    unit = derive_pkg_unit(order.get("packing_kind"), order.get("remark"))
    return f"{ctns}{unit}/{_fmt_gw(gw)}KGS/{float(cbm):.3f}CBM"


def build_factory_cells(order: dict, charges: list[dict] | None = None, *,
                        bl_no_override: str | None = None,
                        shipped_date: str | None = None) -> tuple[dict, float]:
    """返回 (cell_map, total_amount)。"""
    bl = bl_no_override or order.get("bl_no") or ""
    work_no = order.get("order_no") or bl
    ctns = order.get("ctns") or 0
    unit = derive_pkg_unit(order.get("packing_kind"), order.get("remark"))

    if charges is None:
        charges = clone_defaults("factory")
    env = build_env(order)

    cells: dict[str, Any] = {
        # 头部
        # 模板：A7=TO, B7=发货人公司名 | G7=From, H7=客服名字 | A8=Attn, B8=联系人
        "B7": order.get("invoice_title") or order.get("shipper_cn") or "",
        "B8": order.get("attn_contact") or "",               # Attn = 工厂联系方式
        "H7": order.get("csr_name") or "",                    # From = 客服名字
        "B9": order.get("shipper_cn") or "",                  # 原发货人公司名
        "B10": work_no,                                        # 工作号
        "H8": f"{work_no}B",                                   # Invoice No.
        "B11": bl,                                             # 提单号
        "H10": _marks(order),                                  # 货量
        "B12": order.get("pol") or "",                         # 起运港
        # 开船日期：仅取 ATD（界面手填优先）；无 ATD 留空，禁止用申报日期兜底
        "H12": shipped_date or order.get("shipped_date") or order.get("atd") or "",
        "B13": order.get("vessel") or "",                      # 船名航次
        "H13": order.get("container_no") or "",                # 柜号
        "H11": order.get("pod") or "",                         # 目的港
        "I51": date.today().isoformat(),                       # 账单日期（覆盖 TODAY()）
    }

    # 费用行（需求 2026-08-18：不设行数上限，超出模板固定区由 generate_factory 插入行承接）
    total = 0.0
    r = FIRST_ROW
    n_used = 0
    for ch in charges:
        amt = compute_amount(ch, env)
        cells[f"A{r}"] = ch.get("name", "")
        cells[f"C{r}"] = ch.get("unit", "")
        cells[f"D{r}"] = ch.get("unit_price", 0)
        cells[f"E{r}"] = ch.get("qty", 0)
        cells[f"F{r}"] = ch.get("currency", CURRENCY)
        cells[f"G{r}"] = round(amt, 2)
        total += amt
        r += 1
        n_used += 1

    total = round(total, 2)
    # 动态汇总行：紧接费用行之后（模板 G31 固定区，n=15 时对齐）
    total_row = FIRST_ROW + n_used       # 16+n（n=15 → 31）
    cells[f"G{total_row}"] = total       # ★汇总
    # 元数据：COM 插入行 / print_area 依据
    cells["_fac_total_row"] = total_row
    cells["_fac_charge_count"] = n_used
    return cells, total


def build_factory_context(order: dict, charges: list[dict] | None, *,
                          bl_no_override: str | None = None,
                          shipped_date: str | None = None) -> dict[str, Any]:
    """工厂账单 HTML 渲染 context。

    业务逻辑与 build_factory_cells 同步：包装单位推导、SAY 行省略、费用
    计算复用 compute_amount；超 FAC_MAX_ROWS 行截断同原行为。
    """
    bl = bl_no_override or order.get("bl_no") or ""
    work_no = order.get("order_no") or bl
    from bol_forecast.core.formula import build_env, compute_amount
    env = build_env({
        "ctns": order.get("ctns") or 0,
        "gw":   order.get("gw") or 0,
        "cbm":  order.get("cbm") or 0,
    })

    charge_rows: list[dict[str, Any]] = []
    total = 0.0
    n = 0
    for ch in (charges or []):
        # 需求（2026-08-18）：无行数上限，HTML 表格自然多行
        amt = round(compute_amount(ch, env), 2)
        charge_rows.append({
            "name":       ch.get("name", ""),
            "unit":       ch.get("unit", ""),
            "unit_price": ch.get("unit_price", 0),
            "qty":        ch.get("qty", 0),
            "currency":   ch.get("currency", CURRENCY),
            "amount":     amt,
            "remark":     ch.get("remark", ""),
        })
        total += amt
        n += 1
    total = round(total, 2)

    return {
        # 头部（与 build_factory_cells 的 B7/B8/H7/B9 对应）
        "to_name":     order.get("invoice_title") or order.get("shipper_cn") or "",
        "attn":        order.get("attn_contact") or "",
        "from_name":   order.get("csr_name") or "深圳市嘀嗒嘀物流科技有限公司",
        "shipper":     order.get("shipper_cn") or "",
        # 编号信息
        "work_no":     work_no,
        "invoice_no":  f"{work_no}B",
        "bl_no":       bl,
        "marks":       _marks(order),   # H10 货量
        "pol":         order.get("pol") or "",
        "shipped_date": shipped_date or order.get("shipped_date") or order.get("atd") or "",
        "vessel":      order.get("vessel") or "",
        "pod":         order.get("pod") or "",
        "container":   order.get("container_no") or "",
        "bill_date":   date.today().isoformat(),   # I51 账单日期
        "charge_rows": charge_rows,
        "total":       total,
        "currency":    CURRENCY,
    }


@com_retry()
def generate_factory(order: dict, out_dir: str,
                     charges: list[dict] | None = None, *,
                     bl_no_override: str | None = None,
                     shipped_date: str | None = None,
                     export_pdf: bool = True,
                     excel: object | None = None) -> dict:
    """生成工厂账单（EXCEL + 可选 PDF）。html 模式只产 PDF。"""
    bl = bl_no_override or order.get("bl_no") or "BL"
    po_tag = customer_po_tag(order)
    out_name = f"{safe_filename(str(bl))}{po_tag}+{safe_filename(doc_label('factory'))}.xlsx"
    out_xlsx = os.path.join(out_dir, out_name)
    result: dict[str, Any] = {"xlsx": out_xlsx, "pdf": None}

    from bol_forecast.rendering.renderer import renderer_mode
    if renderer_mode() == "html":
        from bol_forecast.rendering.fit_specs import apply_sizes
        from bol_forecast.rendering.html_render import render_doc
        ctx = build_factory_context(order, charges, bl_no_override=bl_no_override,
                                    shipped_date=shipped_date)
        apply_sizes("factory", ctx)
        out_pdf = os.path.join(out_dir, f"{safe_filename(str(bl))}{po_tag}+{safe_filename(doc_label('factory'))}.pdf")
        try:
            render_doc("factory", ctx, out_pdf)
            result["xlsx"] = None
            result["pdf"] = out_pdf
            result["total"] = ctx["total"]
            log.info("工厂账单 HTML 渲染完成: %s total=%s", out_pdf, ctx["total"])
        except Exception as e:
            log.warning("工厂账单 HTML 渲染失败: %s", e)
        return result

    cells, total = build_factory_cells(
        order, charges, bl_no_override=bl_no_override, shipped_date=shipped_date)
    result["total"] = total
    total_row = int(cells.get("_fac_total_row") or (FIRST_ROW + FAC_MAX_ROWS))
    charge_count = int(cells.get("_fac_charge_count") or 0)
    with XlsxWriter(template_path("factory"), out_xlsx, sheet=FAC_SHEET, excel=excel) as w:
        # 需求（2026-08-18）：费用行数不设上限——超出模板固定区（16~30 共 15 行）时，
        # 在其下方插入新行承接（汇总/账户区随行下移），Excel 仍可手动编辑。
        extra_rows = max(0, charge_count - FAC_MAX_ROWS)
        if extra_rows:
            w.insert_rows(FIRST_ROW + FAC_MAX_ROWS, extra_rows)   # 在 R31 起插入
        w.set_cells(cells)
        w.set_col_widths(FAC_COL_WIDTHS)   # 竖向 A4：缩窄列宽防右侧截断
        # 金额列自动缩字（防 ###，保证金额完整可读）
        for r in range(FIRST_ROW, FIRST_ROW + charge_count):
            try:
                w.fit_cell(f"G{r}", cells.get(f"G{r}"))
            except Exception as e:
                log.warning("工厂账单 G%s 缩字失败(继续): %s", r, e)
        try:
            w.fit_cell(f"G{total_row}", cells.get(f"G{total_row}"))
        except Exception as e:
            log.warning("工厂账单 G%d 缩字失败(继续): %s", total_row, e)
        # 文本溢出检测：货量/目的港/柜号/船名航次等长文本列（2026-08-17）
        for coord in ("H10", "H11", "H13", "B13"):
            try:
                w.fit_cell(coord, cells.get(coord))
            except Exception as e:
                log.warning("工厂账单 %s 文本溢出缩字失败(继续): %s", coord, e)
        # ★2026-08-20：WrapText 行高适配 —— 发货人/费用名称/备注多行文本
        for coord in ("B7", "B9", "B8", "H7", "H10", "H11", "H12", "H13",
                      "B11", "B12", "B13"):
            try:
                w.fit_text(coord, cells.get(coord), min_row_height=14.0)
            except Exception as e:
                log.warning("工厂账单 %s wrap行高适配失败(继续): %s", coord, e)
        for r in range(FIRST_ROW, FIRST_ROW + charge_count):
            for coord in (f"A{r}", f"I{r}"):
                try:
                    w.fit_text(coord, cells.get(coord), min_row_height=14.0)
                except Exception as e:
                    log.warning("工厂账单 %s wrap行高适配失败(继续): %s", coord, e)
        if export_pdf:
            out_pdf = os.path.join(
                out_dir, f"{safe_filename(str(bl))}{po_tag}+{safe_filename(doc_label('factory'))}.pdf")
            try:
                # 需求3：行数增多时整页缩放至单页（FitToPagesTall=1）
                max_r = max(51, total_row + 8)
                w.export_pdf(out_pdf, print_area=f"A1:K{max_r}",
                             orientation=1, single_page=True)
                result["pdf"] = out_pdf
            except Exception as e:
                # Excel COM 忙（0x800AC472）时 PDF 导出失败 —— 用 HTML 渲染器兜底
                log.warning("工厂账单 COM PDF 导出失败，改走 HTML 渲染兜底: %s", e)
                try:
                    from bol_forecast.rendering.fit_specs import apply_sizes
                    from bol_forecast.rendering.html_render import render_doc
                    ctx = build_factory_context(order, charges,
                                                bl_no_override=bl_no_override,
                                                shipped_date=shipped_date)
                    apply_sizes("factory", ctx)
                    render_doc("factory", ctx, out_pdf)
                    result["pdf"] = out_pdf
                except Exception as e2:
                    log.warning("工厂账单 HTML 兜底 PDF 也失败: %s", e2)
    log.info("工厂账单已生成: %s total=%s (pdf=%s) rows=%d", out_xlsx, total, result["pdf"], charge_count)
    return result
