# -*- coding: utf-8 -*-
"""开船提单生成：merged order -> 开船提单模版.xlsx 单元格。

写入策略（自我包含）：
  模板单元格原是对母表 `1. 海运散货补料模板.xls` 的外部公式引用
  （如 ='[1]BILL  DRAFT'!A2）。母表不在工作区，故直接用 COM 把这些
  公式格改写成计算好的字面值，LOGO/版式保留。

字段→单元格映射（按用户规则细化）：
  A2  Shipper            <- shipper_cn
  I2  B/L No.           <- bl_no（主页面可手输覆盖）
  A8  Consignee          <- consignee
  A14 Notify Party       <- notify_party (== consignee, FOB 铁律)
  A23 Vessel and Voyage  <- vessel
  D23 Port of Loading    <- pol
  A25 Port of Discharge  <- pod
  D25 Place of Delivery  <- pod
  F25 Final Destination  <- pod
  A28 Container No.      <- container_no
  C28 Seal No.           <- seal_no（默认 "//"）
  E28 Marks & Nos.       <- "{ctns}{unit}/{gw}KGS/{cbm}CBM"  e.g. 96CTNS/1175.3KGS/7.023CBM
  A30 固定               <- "CFS-CFS"（散货提单写死）
  A32 唛头               <- order_no（订单号）
  D33 No. of pkgs        <- ctns
  E33 包装单位           <- derive_pkg_unit(packing_kind, remark)：CTNS/PKGS/PLTS
  F33 Description        <- 品名（优先字段确认手动值 goods_name，回退 goods 首条）
  H33 Gross Weight(kgs)  <- gw
  J33 Measurement(M3)    <- cbm（数字过大时自动缩字，防 ###）
  E59 SAY(...)ONLY       <- "SAY ( {箱数英文+包装单位全称} ) ONLY"，合并原 F59/G59
  F59/G59                <- 清空（已并入 E59）
  K56 Shipped on board   <- shipped_date（== 轨迹 ATD；无 ATD 则留空，禁用申报日期）
"""
from __future__ import annotations

import logging
import os
from typing import Any

from bol_forecast.config import safe_filename, template_path, customer_po_tag
from bol_forecast.core.num2en import say_packages, unit_to_en
from .com_session import com_retry
from .writer_base import XlsxWriter

log = logging.getLogger(__name__)

BL_SHEET = "Sheet1"
# 打印区域：内容最右到 M 列，N 列以右为空白框，裁剪掉以收紧单页
BL_PRINT_AREA = "A1:M72"


# ---------------------------------------------------------------- 包装单位判定
def derive_pkg_unit(packing_kind: str | None, remark: str | None) -> str:
    """依据报关单『包装种类』栏 + 备注栏判定包装单位。

    - 英文缩写直接返回（CTNS/PKGS/PLTS/CARTONS/PACKAGES…）——
      字段确认栏位允许用户自由填写，UI 写入后不应被中文推断逻辑覆盖。
    - 纸质盒/箱           -> CTNS
    - 木质盒/箱           -> PKGS
    - 其他包装 + 备注含塑料托盘/托盘 -> PLTS
    - 其他 / 无法识别     -> PKGS（默认）
    """
    pk = (packing_kind or "").strip()
    rm = (remark or "").strip()
    upper = pk.upper()
    # 1) 英文缩写短路：CTNS / PKGS / PLTS / CARTON / PACKAGE / PALLET
    if upper in {"CTNS", "CARTONS", "CARTON"}:
        return "CTNS"
    if upper in {"PKGS", "PACKAGES", "PACKAGE"}:
        return "PKGS"
    if upper in {"PLTS", "PALLETS", "PALLET"}:
        return "PLTS"
    # 2) 中文关键字推断
    has_box = ("盒" in pk) or ("箱" in pk)
    if "纸" in pk and has_box:
        return "CTNS"
    if "木" in pk and has_box:
        return "PKGS"
    if "塑料托盘" in rm or "托盘" in rm or "托盘" in pk:
        return "PLTS"
    return "PKGS"


def _fmt_gw(gw: float) -> str:
    """1175.30 -> '1175.3'；1175 -> '1175'。去尾零。"""
    s = f"{float(gw):.2f}".rstrip("0").rstrip(".")
    return s if s else "0"


def _bl_description(order: dict, *, customer_po_on_bl: bool = True) -> str:
    """F33 品名多行展示：每个品名独立成行；客户 PO（勾选时）在末行另起新行。

    - 手动 goods_name（界面 textarea，每行一个品名）优先，按行拆分去空行、保序。
    - 否则取 goods_booking / goods_customs 列表，每个品名去重保序各成一行。
    - 客户 PO：customer_po_on_bl=True 且 PO 非空时，在品名末行追加 `PO:{po}` 独立一行。
      该勾选框**只影响提单呈现**，不影响话术变量（customer_po 仍按内容取值）与文件名
      （由 customer_po_tag 单独处理）。
    """
    manual = (order.get("goods_name") or "").strip()
    goods = order.get("goods_booking") or order.get("goods_customs") or []
    rows: list[str] = []
    if manual:
        # 界面手填：按换行拆分，每行一个品名，去空行保序
        for ln in manual.splitlines():
            s = ln.strip()
            if s:
                rows.append(s)
    elif goods:
        seen: set[str] = set()
        for g in goods:
            nm = (g.get("name_en") or g.get("name_cn") or "").strip()
            if nm and nm not in seen:
                seen.add(nm)
                rows.append(nm)
    # 客户 PO 末行新行（仅勾选显示）
    po = (order.get("customer_po") or "").strip()
    if po and customer_po_on_bl:
        rows.append(f"PO:{po}")
    return "\n".join(rows)


def _up(text):
    """提单字段统一转为大写（用户要求：提单所有填充字母大写）。"""
    if isinstance(text, str):
        return text.upper()
    return text


def build_bl_context(order: dict, *,
                     bl_no_override: str | None = None,
                     shipped_date: str | None = None,
                     seal_no: str | None = None,
                     customer_po_on_bl: bool = True) -> dict[str, Any]:
    """开船提单 业务逻辑唯一来源：order + overrides -> 命名键 dict。

    返回字段（与 bol_forecast/rendering/fit_specs.BL 对应；html 模板直接消费）：
      shipper_cn / bl_no / consignee / notify_party / vessel / pol / pod /
      container_no / seal / package_summary / marks / order_no / ctns / unit /
      goods_name / gw / cbm / say_line / shipped_date / freight_term / packing_kind
    """
    ctns = order.get("ctns") or 0
    gw = order.get("gw") or 0
    cbm = order.get("cbm") or 0
    unit = derive_pkg_unit(order.get("packing_kind"), order.get("remark"))
    unit_full = unit_to_en(unit)  # CARTONS / PACKAGES / PALLETS

    # E28: 件数+包装单位/重量KGS/体积CBM （公式组合，不上层大写）
    package_summary = f"{ctns}{unit}/{_fmt_gw(gw)}KGS/{float(cbm):.3f}CBM"

    # E59: SAY ( 箱数英文+包装单位全称 ) ONLY
    say_line = f"SAY ( {say_packages(int(ctns), unit)} ) ONLY"

    # 封条：字段确认可手动填写；默认 "//"
    seal = (seal_no or order.get("seal_no") or "//").strip() or "//"

    notify = order.get("notify_party") or order.get("consignee") or ""
    bl = bl_no_override or order.get("bl_no") or ""
    pod = order.get("pod") or ""
    pol = order.get("pol") or ""
    vessel = order.get("vessel") or ""

    return {
        "shipper_cn":      _up(order.get("shipper_cn") or ""),
        "bl_no":           _up(bl),
        "consignee":       _up(order.get("consignee") or ""),
        "notify_party":    _up(notify),
        "vessel":          _up(vessel),
        "pol":             _up(pol),
        "pod":             _up(pod),
        "container_no":    _up(order.get("container_no") or ""),
        "seal":            _up(seal),
        "package_summary": package_summary,
        "marks":           "",                                # 模板里 A29 留作后续追加
        "order_no":        _up(order.get("order_no") or ""),
        "ctns":            ctns,
        "unit":            _up(unit),
        "goods_name":      _up(_bl_description(order, customer_po_on_bl=customer_po_on_bl)),
        "gw":              gw,
        "cbm":             round(float(cbm), 3),
        "say_line":        say_line,
        "shipped_date":    _up(shipped_date or order.get("shipped_date") or order.get("atd") or ""),
        "freight_term":    "FREIGHT COLLECT",                 # 模板 A55 固定
        "packing_kind":    order.get("packing_kind") or "",
        # HTML 模板需要的额外渲染块
        "is_lcl":          True,                              # 散货提单（开船提单就是 LCL 模板）
        "package_summary_full": f"{int(ctns)} {unit_full}",   # 备用：英文箱数
    }


def bl_coords(ctx: dict[str, Any]) -> dict[str, Any]:
    """context -> {coord: value}（Excel COM 路径使用）。"""
    return {
        "A2":  ctx["shipper_cn"],
        "I2":  ctx["bl_no"],
        "A8":  ctx["consignee"],
        "A14": ctx["notify_party"],
        "A23": ctx["vessel"],
        "D23": ctx["pol"],
        "A25": ctx["pod"],
        "D25": ctx["pod"],
        "F25": ctx["pod"],
        "A28": ctx["container_no"],
        "C28": ctx["seal"],
        "E28": ctx["package_summary"],
        "A30": "CFS-CFS",
        "A32": ctx["order_no"],
        "D33": ctx["ctns"],
        "E33": ctx["unit"],
        "F33": ctx["goods_name"],
        "H33": ctx["gw"],
        "J33": ctx["cbm"],
        "E59": ctx["say_line"],
        "F59": "",
        "G59": "",
        "K56": ctx["shipped_date"],
    }


def build_bl_cells(order: dict, *,
                   bl_no_override: str | None = None,
                   shipped_date: str | None = None,
                   seal_no: str | None = None,
                   customer_po_on_bl: bool = True) -> dict[str, Any]:
    """薄封装：内部走 context + coords。COM 路径（与 build_*_cells 旧签名兼容）使用。"""
    ctx = build_bl_context(order, bl_no_override=bl_no_override,
                           shipped_date=shipped_date, seal_no=seal_no,
                           customer_po_on_bl=customer_po_on_bl)
    return bl_coords(ctx)


@com_retry()
def generate_bill_of_lading(order: dict, out_dir: str,
                            bl_no: str | None = None, *,
                             bl_no_override: str | None = None,
                             shipped_date: str | None = None,
                             seal_no: str | None = None,
                             customer_po_on_bl: bool = True,
                             export_pdf: bool = True,
                             excel: object | None = None) -> dict:
    """生成开船提单核对件（EXCEL + 可选 PDF），返回 {xlsx, pdf, pdf_path}。

    - bl_no / bl_no_override：B/L NO.，主页面可手输覆盖。
    - shipped_date：装船时间，主页面可编辑；缺省取轨迹 ATD，无 ATD 则留空。
    - seal_no：封条号，默认 "//"。
    - export_pdf：是否同步导出 PDF（仅当前 sheet，隐藏 sheet 保持隐藏）。
    - excel：可选复用的 COM Excel 实例（Req3 性能优化）。

    渲染器选择：见 bol_forecast.rendering.renderer.renderer_mode()。
    - html 模式：直接生成 PDF，无 xlsx
    - com  模式：xlsx + PDF（旧行为）
    """
    bl_no = bl_no_override or bl_no or order.get("bl_no") or "BL"
    po_tag = customer_po_tag(order)
    out_name = f"{safe_filename(str(bl_no))}{po_tag}+提单核对件.xlsx"
    out_xlsx = os.path.join(out_dir, out_name)
    ctx = build_bl_context(order, bl_no_override=bl_no_override,
                           shipped_date=shipped_date, seal_no=seal_no,
                           customer_po_on_bl=customer_po_on_bl)
    cells = bl_coords(ctx)
    result: dict[str, Any] = {"xlsx": out_xlsx, "pdf": None}

    # ---- 渲染器选择 ----
    from bol_forecast.rendering.renderer import renderer_mode
    if renderer_mode() == "html":
        from bol_forecast.rendering.fit_specs import apply_sizes
        from bol_forecast.rendering.html_render import render_doc
        apply_sizes("bl", ctx)
        out_pdf = os.path.join(out_dir, f"{safe_filename(str(bl_no))}{po_tag}+提单核对件.pdf")
        try:
            render_doc("bl", ctx, out_pdf)
            result["xlsx"] = None
            result["pdf"] = out_pdf
            log.info("开船提单 HTML 渲染完成: %s", out_pdf)
        except Exception as e:
            log.warning("HTML 渲染失败: %s", e)
        return result

    # ---- 原有 COM 路径 ----
    with XlsxWriter(template_path("bl"), out_xlsx, sheet=BL_SHEET, excel=excel) as w:
        w.set_cells(cells)
        # J33 体积列：数字过大时自动缩字（防 ###）
        try:
            w.fit_cell("J33", cells.get("J33"))
        except Exception as e:
            log.warning("J33 体积缩字失败(继续): %s", e)
        if export_pdf:
            out_pdf = os.path.join(out_dir, f"{safe_filename(str(bl_no))}{po_tag}+提单核对件.pdf")
            try:
                # 需求（2026-08-18）：提单核对件必须单页显示 → FitToPagesTall=1 整页缩放
                w.export_pdf(out_pdf, print_area=BL_PRINT_AREA, single_page=True)
                result["pdf"] = out_pdf
            except Exception as e:
                log.warning("PDF 导出失败: %s", e)
    log.info("开船提单已生成: %s (pdf=%s)", out_xlsx, result["pdf"])
    return result
