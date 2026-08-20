# -*- coding: utf-8 -*-
"""文本自适应字段表：与现 Excel COM fit_cell() 调用点对应。

每个文档（bl/telex/factory/invoice）列出一组 (ctx_key, base_pt, box_pt, bold)，
renderer 会在 ctx["sizes"] 里填入 {ctx_key: 拟合后字号(pt)}，模板用
`font-size:{{ ctx.sizes.key }}pt` 即可。

box_pt 是该字段在 A4 内的可视盒宽（pt），由 CSS 排版确定。
集中维护的好处：调 CSS 时同步调此表，文字"位置固定"与"大小自适应"解耦。
"""
from __future__ import annotations

from typing import TypedDict

from bol_forecast.rendering.textfit import fit_font_size


class FitField(TypedDict):
    base_pt: float
    box_pt: float
    bold: bool


# 盒宽估算（pt）按现 COM 缩列宽后的列宽 × Excel 列宽→像素系数计算。
# 提单 print_area A1:M72（约 12 列有效）→ 内容宽 559.27pt；列宽分配按 1 pt = 1.333 px。
# 工厂账单 print_area A1:K51 → 11 列分 559.27pt。
# INVOICE print_area A1:R45 → 18 列分 559.27pt，但右侧 1/3 大多空白。

# 注：以下 box_pt 是粗估初值，会在 compare_pdfs 迭代中根据实际渲染微调。
# 单元：pt
BL: dict[str, FitField] = {
    "shipper_cn":     {"base_pt": 11.0, "box_pt": 240.0, "bold": False},  # A2
    "bl_no":          {"base_pt": 14.0, "box_pt": 70.0,  "bold": True},   # I2
    "consignee":      {"base_pt": 11.0, "box_pt": 320.0, "bold": False},  # A8
    "notify_party":   {"base_pt": 11.0, "box_pt": 320.0, "bold": False},  # A14
    "vessel":         {"base_pt": 12.0, "box_pt": 130.0, "bold": False},  # A23
    "pol":            {"base_pt": 12.0, "box_pt": 130.0, "bold": False},  # D23
    "pod":            {"base_pt": 12.0, "box_pt": 130.0, "bold": False},  # A25/D25/F25
    "container_no":   {"base_pt": 11.0, "box_pt": 100.0, "bold": False},  # A28
    "package_summary": {"base_pt": 11.0, "box_pt": 100.0, "bold": False}, # E28
    "marks":          {"base_pt": 10.0, "box_pt": 240.0, "bold": False},  # A29
    "order_no":       {"base_pt": 11.0, "box_pt": 100.0, "bold": False},  # A32
    "ctns":           {"base_pt": 11.0, "box_pt": 50.0,  "bold": False},  # D33
    "unit":           {"base_pt": 11.0, "box_pt": 50.0,  "bold": False},  # E33
    "goods_name":     {"base_pt": 11.0, "box_pt": 200.0, "bold": False},  # F33
    "gw":             {"base_pt": 11.0, "box_pt": 50.0,  "bold": False},  # H33
    "cbm":            {"base_pt": 11.0, "box_pt": 50.0,  "bold": False},  # J33
    "say_line":       {"base_pt": 12.0, "box_pt": 280.0, "bold": True},   # E59
    "shipped_date":   {"base_pt": 11.0, "box_pt": 80.0,  "bold": False},  # K56
}

TELEX: dict[str, FitField] = {
    "shipper":    {"base_pt": 10.5, "box_pt": 280.0, "bold": False},  # B24
    "consignee":  {"base_pt": 10.5, "box_pt": 280.0, "bold": False},  # A8
    "bl_no":      {"base_pt": 14.0, "box_pt": 280.0, "bold": True},   # B21
    "container":  {"base_pt": 10.5, "box_pt": 200.0, "bold": False},  # B22
    "package":    {"base_pt": 10.5, "box_pt": 280.0, "bold": False},  # B23
    "pol":        {"base_pt": 10.5, "box_pt": 100.0, "bold": False},  # B26
    "pod":        {"base_pt": 10.5, "box_pt": 100.0, "bold": False},  # B28
    "vessel":     {"base_pt": 10.5, "box_pt": 100.0, "bold": False},  # B29
    "shipped_date": {"base_pt": 10.5, "box_pt": 100.0, "bold": False},  # B27
}

FACTORY: dict[str, FitField] = {
    "to_name":   {"base_pt": 12.0, "box_pt": 200.0, "bold": False},  # B7
    "attn":      {"base_pt": 11.0, "box_pt": 100.0, "bold": False},  # H7
    "from_name": {"base_pt": 11.0, "box_pt": 200.0, "bold": False},  # B8
    "shipper":   {"base_pt": 11.0, "box_pt": 200.0, "bold": False},  # B10
    "work_no":   {"base_pt": 11.0, "box_pt": 80.0,  "bold": False},  # H10
    "invoice_no":{"base_pt": 11.0, "box_pt": 80.0,  "bold": False},  # B11
    "bl_no":     {"base_pt": 11.0, "box_pt": 80.0,  "bold": False},  # H11
    "marks":     {"base_pt": 10.0, "box_pt": 100.0, "bold": False},  # H10 货量
    "pol":       {"base_pt": 11.0, "box_pt": 80.0,  "bold": False},  # B12 起运港
    "pod":       {"base_pt": 11.0, "box_pt": 80.0,  "bold": False},  # H12
    "shipped_date":{"base_pt": 11.0, "box_pt": 80.0,"bold": False},  # B12
    "vessel":    {"base_pt": 11.0, "box_pt": 100.0, "bold": False},  # B13
    "container": {"base_pt": 11.0, "box_pt": 100.0, "bold": False},  # H13
    # 费用行 G16..G31 统一缩字
    "charge_amount":  {"base_pt": 11.0, "box_pt": 50.0, "bold": False},
    "charge_total":   {"base_pt": 11.0, "box_pt": 50.0, "bold": True},
}

INVOICE: dict[str, FitField] = {
    "bill_no":     {"base_pt": 11.0, "box_pt": 130.0, "bold": False},  # G2
    "bill_date":   {"base_pt": 11.0, "box_pt": 130.0, "bold": False},  # G3
    "due_date":    {"base_pt": 11.0, "box_pt": 130.0, "bold": False},  # G4
    "bl_no":       {"base_pt": 11.0, "box_pt": 130.0, "bold": False},  # E8
    "container":   {"base_pt": 11.0, "box_pt": 130.0, "bold": False},  # E10
    "vessel":      {"base_pt": 11.0, "box_pt": 130.0, "bold": False},  # G10
    "pol":         {"base_pt": 11.0, "box_pt": 130.0, "bold": False},  # E12
    "pod":         {"base_pt": 11.0, "box_pt": 130.0, "bold": False},  # G12
    "ctns":        {"base_pt": 11.0, "box_pt": 80.0,  "bold": False},  # E14
    "goods_name":  {"base_pt": 11.0, "box_pt": 130.0, "bold": False},  # G14
    "gw":          {"base_pt": 11.0, "box_pt": 80.0,  "bold": False},  # E16
    "cbm":         {"base_pt": 11.0, "box_pt": 80.0,  "bold": False},  # G16
    "bill_to_name":{"base_pt": 13.0, "box_pt": 200.0, "bold": True},   # 客户抬头
    # 费用 G19..G25
    "charge_amount":  {"base_pt": 11.0, "box_pt": 80.0, "bold": False},
    "subtotal":       {"base_pt": 11.0, "box_pt": 100.0, "bold": True},
    "total":          {"base_pt": 11.0, "box_pt": 100.0, "bold": True},
}


def specs_for(doc_key: str) -> dict[str, FitField]:
    return {
        "bl": BL, "telex": TELEX, "factory": FACTORY, "invoice": INVOICE,
    }.get(doc_key, {})


def apply_sizes(doc_key: str, ctx: dict) -> dict:
    """计算 ctx["sizes"]，每个字段填入拟合后字号（pt）。原地修改 ctx，返回 ctx。"""
    spec = specs_for(doc_key)
    sizes: dict[str, float] = {}
    for key, fld in spec.items():
        text = ctx.get(key)
        if text is None or text == "":
            sizes[key] = fld["base_pt"]
            continue
        sizes[key] = fit_font_size(
            str(text), base_pt=fld["base_pt"], box_pt=fld["box_pt"],
            bold=fld["bold"], pad_pt=2.0,
        )
    ctx["sizes"] = sizes
    return ctx
