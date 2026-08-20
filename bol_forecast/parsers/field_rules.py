# -*- coding: utf-8 -*-
"""报关单 / 装货单 字段抽取规则表（数据驱动）。

设计原则：
  1. 页面按标题识别角色，绝不硬编码页码。
  2. 每字段可配多条正则，按顺序尝试，先命中先用。
  3. 正则全部失败时由调用方降级到坐标定位，再失败留空标 low。
"""
from __future__ import annotations

import re
import unicodedata

# ---------------- 页面角色 ----------------
PAGE_ROLE_PATTERNS = [
    ("CUSTOMS_DECL", r"中华人民共和国海关(出口|进口)货物报关单"),
    ("RELEASE_NOTICE", r"通关无纸化(出口|进口)放行通知书"),
    ("DECL_AGREEMENT", r"委托报关协议"),
    ("SHIPPING_ORDER", r"装\s*货\s*单|SHIPPING\s+ORDER|集装箱运输标准单"),
    ("PACKING_LIST", r"装箱单|PACKING\s+LIST"),
    ("COMM_INVOICE", r"商业发票|COMMERCIAL\s+INVOICE"),
]


def normalize_text(s: str) -> str:
    """全角转半角 + 统一空白，便于正则稳定匹配。"""
    if not s:
        return ""
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("\u3000", " ")
    return s


# ---------------- 后处理函数 ----------------
def _strip(v):
    return v.strip() if isinstance(v, str) else v


def _upper(v):
    return v.strip().upper() if isinstance(v, str) else v


def _to_int(v):
    try:
        return int(float(str(v).replace(",", "").strip()))
    except (TypeError, ValueError):
        return None


def _to_float(v):
    try:
        return float(str(v).replace(",", "").strip())
    except (TypeError, ValueError):
        return None


def _clean_company(v):
    """去掉统一社会信用代码括号等尾巴。"""
    if not isinstance(v, str):
        return v
    v = re.sub(r"[（(][0-9A-Z]{8,}[)）]", "", v)
    return v.strip(" :：\t")


def _strip_paren(v):
    """去掉首尾括号，用于指运港 (USA000) 之类。"""
    if not isinstance(v, str):
        return v
    return v.strip().strip("（）()").strip()


POST = {
    "strip": _strip,
    "upper": _upper,
    "int": _to_int,
    "float": _to_float,
    "company": _clean_company,
    "paren": _strip_paren,
}

# ---------------- 字段规则 ----------------
# (字段名, 页角色, [正则...], 后处理, 说明)
RULES = [
    # ---- 报关单主页 ----
    ("shipper_cn", "CUSTOMS_DECL", [
        r"境内发货人[^\n]*?\n\s*([^\n]{2,60}?)\s*(?:出境关别|$)",
        r"境内发货人[（(][^)）]*[)）]\s*([^\n]{2,60})",
        r"收发货人\s*\n?\s*([^\n]{2,60}?)(?:\s{2,}|$)",
    ], "company", "境内发货人/收发货人"),

    ("consignee_raw", "CUSTOMS_DECL", [
        r"境外收货人\s*\n?\s*([A-Za-z][^\n]{2,80}?)(?:\s{2,}|运输方式|$)",
    ], "strip", "境外收货人"),

    ("vessel_voyage_p1", "CUSTOMS_DECL", [
        r"运输工具名称及航次号\s*\n?\s*([^\n]{2,50}?)(?:\s{2,}|提运单号|$)",
        r"运输工具名称\s*\n?\s*([^\n]{2,50}?)(?:\s{2,}|提运单号|$)",
    ], "strip", "运输工具名称及航次号(代码形式)"),

    ("bl_no", "CUSTOMS_DECL", [
        r"提运单号\s*\n?\s*([A-Za-z0-9][A-Za-z0-9\-/]{4,24})",
    ], "upper", "提运单号=提单号"),

    ("pkg_count_p1", "CUSTOMS_DECL", [
        r"件数\s*(?:\n[^\n]*)?\n[^\n]*?\s(\d{1,6})\s+\d",       # 表格行内
        r"件数\s*[:：]?\s*(\d{1,6})",
    ], "int", "件数"),

    ("gross_kg_p1", "CUSTOMS_DECL", [
        r"毛重\s*[（(]\s*千克\s*[)）]\s*[:：]?\s*([\d,]+\.?\d*)",
    ], "float", "毛重(千克)"),

    ("net_kg_p1", "CUSTOMS_DECL", [
        r"净重\s*[（(]\s*千克\s*[)）]\s*[:：]?\s*([\d,]+\.?\d*)",
    ], "float", "净重(千克)"),

    ("trade_term", "CUSTOMS_DECL", [
        r"成交方式\s*[（(]?\d*[)）]?\s*\n?\s*([A-Z]{3})\b",
        r"\b(EXW|FOB|CIF|CFR|FCA|CPT|CIP|DAP|DDP|DDU)\b",
    ], "upper", "成交方式"),

    ("departure_port", "CUSTOMS_DECL", [
        r"离境口岸\s*[（(]?\d*[)）]?\s*\n?[^\n]*?([\u4e00-\u9fa5]{2,8})\s*$",
        r"离境口岸[^\n]*\n[^\n]*?\s([\u4e00-\u9fa5]{2,8})",
    ], "strip", "离境口岸(中文)"),

    ("dest_country", "CUSTOMS_DECL", [
        r"指运港\s*[（(]([A-Z]{3}\d{3})[)）]",
    ], "strip", "指运港代码"),

    ("contract_no", "CUSTOMS_DECL", [
        r"合同协议号\s*\n?\s*([A-Za-z0-9\-]{4,30})",
    ], "upper", "合同协议号"),

    ("customs_no", "CUSTOMS_DECL", [
        r"海关编号[:：]?\s*(\d{10,20})",
    ], "strip", "海关编号"),

    ("export_date", "CUSTOMS_DECL", [
        r"申报日期\s*\n?\s*(\d{8})",
    ], "strip", "申报日期 YYYYMMDD"),

    ("remark", "CUSTOMS_DECL", [
        r"备注[:：]?\s*(.+)",
    ], "strip", "备注栏（用于塑料托盘 PLTS 判定 + 柜号提取）"),

    # ---- 装货单页 ----
    ("vessel_p5", "SHIPPING_ORDER", [
        r"Ocean\s+Vessel[^\n]*\n\s*([A-Z][A-Z\s\.\-]{3,40}?)\s+[A-Z0-9]{2,6}\s+[A-Z]{3,}",
    ], "strip", "真实船名"),

    ("pol_p5", "SHIPPING_ORDER", [
        r"\b(YANTIAN|SHEKOU|SHENZHEN|NINGBO|SHANGHAI|QINGDAO|XIAMEN|GUANGZHOU|NANSHA|TIANJIN|DALIAN|FUZHOU)\b",
    ], "upper", "装货港"),
]


def compile_rules():
    out = []
    for name, role, pats, post, desc in RULES:
        out.append({
            "name": name,
            "role": role,
            "patterns": [re.compile(p, re.M) for p in pats],
            "post": POST.get(post, _strip),
            "desc": desc,
        })
    return out


COMPILED = compile_rules()

# ---------------- 常用港口中英映射 ----------------
PORT_CN2EN = {
    "盐田": "YANTIAN", "蛇口": "SHEKOU", "深圳": "SHENZHEN",
    "宁波": "NINGBO", "上海": "SHANGHAI", "青岛": "QINGDAO",
    "厦门": "XIAMEN", "广州": "GUANGZHOU", "南沙": "NANSHA",
    "天津": "TIANJIN", "大连": "DALIAN", "福州": "FUZHOU",
    "中山": "ZHONGSHAN", "佛山": "FOSHAN", "东莞": "DONGGUAN",
}

# 指运港代码 → 常用英文目的地（仅作默认建议，用户可改）
DEST_CODE2EN = {
    "USA000": "UNITED STATES",
    "USALB": "LONG BEACH, CA",
    "USLAX": "LOS ANGELES, CA",
    "USNYC": "NEW YORK, NY",
}


def cn_port_to_en(cn: str | None) -> str | None:
    if not cn:
        return None
    cn = cn.strip()
    for k, v in PORT_CN2EN.items():
        if k in cn:
            return v
    return None
