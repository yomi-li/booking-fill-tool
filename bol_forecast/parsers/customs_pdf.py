# -*- coding: utf-8 -*-
"""报关单 PDF 解析器。

三级降级策略：正则 → 标签下方坐标取值 → 留空标 low。
任何单字段失败都不会中断整体解析。
"""
from __future__ import annotations

import logging
import re
from typing import Any

import pdfplumber

from .field_rules import (COMPILED, DEST_CODE2EN, PAGE_ROLE_PATTERNS, POST,
                          cn_port_to_en, normalize_text)

log = logging.getLogger(__name__)

CNTR_RE = re.compile(r"\b([A-Z]{4}\d{7})\b")
CNTR_TYPE_RE = re.compile(r"\b(20|40|45)\s*[''\"]?\s*(GP|HQ|HC|RF|OT|FR)?\b", re.I)


# ---------------------------------------------------------------- 坐标抽取表
# 报关单主页（栅格表单）：标签行与值行错位约 11px。
# 每个 spec 对应一个「标签 → 下方值」的抽取；right boundary 由同行下一个
# 标签左边界动态截断，避免横向溢出到相邻列。
# post 复用 field_rules.POST；expect="number" 时非数字内容判空。
COORD_P1 = [
    {"name": "shipper_cn",   "label": r"境内发货人",        "post": "company"},
    {"name": "consignee_raw", "label": r"境外收货人",        "post": "strip"},
    {"name": "vessel_voyage_p1", "label": r"运输工具名称及航次号", "post": "strip"},
    {"name": "bl_no",        "label": r"提运单号",          "post": "upper"},
    {"name": "contract_no",  "label": r"合同协议号",        "post": "upper"},
    {"name": "departure_port", "label": r"离境口岸",        "post": "strip"},
    {"name": "trade_term",   "label": r"成交方式",          "post": "upper"},
    {"name": "export_date",  "label": r"申报日期",          "post": "strip"},
    {"name": "pkg_count_p1", "label": r"件数",              "post": "int", "expect": "number"},
    {"name": "gross_kg_p1",  "label": r"毛重\s*[（(]?\s*千克\s*[)）]?", "post": "float", "expect": "number"},
    {"name": "net_kg_p1",    "label": r"净重\s*[（(]?\s*千克\s*[)）]?", "post": "float", "expect": "number"},
    {"name": "packing_kind", "label": r"包装种类",          "post": "strip"},
]

# 装货单页（英文栅格，常见标签）。本样例无此页，定义以备真实装货单。
COORD_P5 = [
    {"name": "vessel_p5",    "label": r"Ocean\s*Vessel",    "post": "strip"},
    {"name": "voyage_p5",    "label": r"Voy\s*No",          "post": "strip"},
    {"name": "pol_p5",       "label": r"Port\s*of\s*Loading", "post": "upper"},
    {"name": "pod_p5",       "label": r"Port\s*of\s*Discharge", "post": "upper"},
    {"name": "measurement_p5", "label": r"Measurement",     "post": "float", "expect": "number"},
]


def _coord_extract(page, specs, fields, sources, source_name):
    """按 spec 列表逐字段调用 value_under_label，写入 fields。

    仅填充尚未被其他来源设值的字段（坐标法优先级低于已存在的精确值，
    但这里作为主力，通常先跑），避免覆盖。
    """
    for spec in specs:
        name = spec["name"]
        if fields.get(name) not in (None, ""):
            continue
        val = value_under_label(
            page,
            spec["label"],
            dy=spec.get("dy", (1, 30)),
            dx_pad=spec.get("dx_pad", 8),
            expect=spec.get("expect"),
        )
        if val is None:
            continue
        post = POST.get(spec.get("post", "strip"), POST["strip"])
        try:
            val = post(val)
        except Exception as e:
            log.warning("坐标后处理失败 %s: %s", name, e)
        if val not in (None, ""):
            fields[name] = val
            sources[name] = source_name



# ---------------------------------------------------------------- 页面识别
def classify_pages(pages_text: list[str]) -> dict[str, int]:
    """按页面标题区（前 300 字符）识别角色，避免正文提及关键词造成误判。

    例如「委托报关协议通用条款」正文里出现「装箱单」三字，
    若全文匹配就会被误判成装箱单页。
    """
    roles: dict[str, int] = {}
    for i, t in enumerate(pages_text):
        head = t[:300]
        for role, pat in PAGE_ROLE_PATTERNS:
            if role in roles:
                continue
            if re.search(pat, head, re.I):
                roles[role] = i
    return roles


# ---------------------------------------------------------------- 坐标兜底
def _find_label_bbox(words: list[dict], label_regex: str):
    """在 words 中寻找标签，支持跨 word 拼接（如 'Ocean' 'Vessel'）。"""
    rx = re.compile(label_regex, re.I)
    n = len(words)
    for i in range(n):
        acc = ""
        x0 = words[i]["x0"]
        x1 = words[i]["x1"]
        top = words[i]["top"]
        bottom = words[i]["bottom"]
        for j in range(i, min(i + 5, n)):
            w = words[j]
            if abs(w["top"] - top) > 4:
                break
            acc = (acc + " " + w["text"]).strip()
            x1 = max(x1, w["x1"])
            bottom = max(bottom, w["bottom"])
            if rx.fullmatch(acc) or rx.match(acc):
                return (x0, x1, top, bottom)
    return None


_NUM_RE = re.compile(r"^[\d,]+\.?\d*$")


def _right_boundary(words: list[dict], lab: tuple, page_width: float) -> float:
    """栅格表单关键：同一标签行上，下一个标签的左边界即当前值的右边界。

    不做截断的话，取值会横向溢出到相邻列（如「件数」取成「6 2442」）。
    """
    x0, x1, top, bottom = lab
    same_row = [w for w in words
                if abs(w["top"] - top) <= 4 and w["x0"] > x1 + 2]
    if not same_row:
        return page_width
    return min(w["x0"] for w in same_row) - 2


def value_under_label(page, label_regex: str, dy=(1, 30), dx_pad=8,
                      expect: str | None = None) -> str | None:
    """取标签正下方、且不越过右侧相邻列的文字。

    expect="number" 时，非数字内容判定为空（用于 Measurement 这类空栏，
    避免把表头第二行的中文标签误当成值）。
    """
    try:
        words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
    except Exception:
        return None
    lab = _find_label_bbox(words, label_regex)
    if not lab:
        return None
    x0, x1, _top, bottom = lab
    right = _right_boundary(words, lab, float(page.width))

    cand = [w for w in words
            if bottom + dy[0] <= w["top"] <= bottom + dy[1]
            and w["x1"] > x0 - dx_pad
            and w["x0"] < right]
    if not cand:
        return None
    cand.sort(key=lambda w: (round(w["top"], 1), w["x0"]))
    first_top = round(cand[0]["top"], 1)
    line = [w["text"] for w in cand if abs(round(w["top"], 1) - first_top) <= 3]
    val = " ".join(line).strip()
    if not val:
        return None

    if expect == "number":
        compact = val.replace(" ", "")
        if not _NUM_RE.match(compact):
            return None
        return compact
    return val


# ---------------------------------------------------------------- 主解析
def parse_customs_pdf(path: str) -> dict[str, Any]:
    """返回 {fields: {...}, pages: {role:index}, raw_text: {...}, warnings: []}"""
    result: dict[str, Any] = {
        "fields": {}, "pages": {}, "raw_text": {}, "warnings": [],
        "goods": [], "containers": [],
    }
    fields = result["fields"]

    try:
        pdf = pdfplumber.open(path)
    except Exception as e:
        result["warnings"].append(f"PDF 打开失败: {e}")
        return result

    with pdf:
        pages_text = []
        for pg in pdf.pages:
            try:
                pages_text.append(normalize_text(pg.extract_text() or ""))
            except Exception as e:
                log.warning("页面文本提取失败: %s", e)
                pages_text.append("")

        roles = classify_pages(pages_text)
        result["pages"] = roles
        for role, idx in roles.items():
            result["raw_text"][role] = pages_text[idx]

        # ---- 1) 坐标法优先（栅格表单主力） ----
        sources = {}
        if "CUSTOMS_DECL" in roles:
            _coord_extract(pdf.pages[roles["CUSTOMS_DECL"]], COORD_P1,
                           fields, sources, "报关单")
        if "SHIPPING_ORDER" in roles:
            _coord_extract(pdf.pages[roles["SHIPPING_ORDER"]], COORD_P5,
                           fields, sources, "装货单")
        result["sources"] = sources

        # ---- 2) 正则规则表兜底（仅填补坐标法未取到的字段） ----
        for rule in COMPILED:
            role = rule["role"]
            if role not in roles or fields.get(rule["name"]) not in (None, ""):
                continue
            text = pages_text[roles[role]]
            val = None
            for pat in rule["patterns"]:
                m = pat.search(text)
                if m:
                    val = m.group(1)
                    break
            if val is not None:
                try:
                    val = rule["post"](val)
                except Exception as e:
                    result["warnings"].append(f"{rule['name']} 后处理失败: {e}")
                    val = None
            if val not in (None, ""):
                fields[rule["name"]] = val
                sources[rule["name"]] = "正则兜底"

        # ---- 3) 装货单纯文本兜底 ----
        if "SHIPPING_ORDER" in roles:
            _parse_shipping_order_text(pages_text[roles["SHIPPING_ORDER"]], fields)

        # ---- 3) 集装箱号 ----
        all_text = "\n".join(pages_text)
        cntrs = list(dict.fromkeys(CNTR_RE.findall(all_text)))
        result["containers"] = cntrs

        # 柜号优先从报关单「装箱标箱数」备注段提取（用户明确要求）：
        # 取「装箱标箱数」之后所带的 四位字母+七位数字 柜号。
        decl_text = pages_text[roles["CUSTOMS_DECL"]] if "CUSTOMS_DECL" in roles else ""
        remark = fields.get("remark") or ""
        cand_text = "\n".join([remark, decl_text]) if remark else decl_text
        container_no = None
        for pat in (
            r"装箱标箱数.{0,120}?([A-Z]{4}\d{7})",
            r"集装箱标箱数.{0,120}?([A-Z]{4}\d{7})",
            r"标箱数.{0,120}?([A-Z]{4}\d{7})",
            r"集装箱号码[:：]?\s*([A-Z]{4}\d{7})",
        ):
            m = re.search(pat, cand_text)
            if m:
                container_no = m.group(1)
                break
        # 退而求其次：全文第一个柜号
        if not container_no and cntrs:
            container_no = cntrs[0]
        if container_no:
            fields["container_no"] = container_no
        fields["cntr_qty"] = len(cntrs)

        # 柜型（优先装货单）
        if "SHIPPING_ORDER" in roles:
            mt = re.search(r"([A-Z]{4}\d{7})\s+(20|40|45)\b",
                           pages_text[roles["SHIPPING_ORDER"]])
            if mt:
                fields["cntr_size"] = mt.group(2)

        # ---- 4) 商品明细 ----
        if "CUSTOMS_DECL" in roles:
            result["goods"] = _parse_goods(pages_text[roles["CUSTOMS_DECL"]])

    _post_derive(fields, result)
    return result


def _parse_shipping_order_text(text: str, fields: dict) -> None:
    """装货单纯文本兜底：船名/航次常常与 POL 同行出现。"""
    # 形如: ZIM MOUNT OLYMPUS 11E YANTIAN
    m = re.search(
        r"^([A-Z][A-Z\s\.\-]{3,40}?)\s+([A-Z0-9]{2,6})\s+"
        r"(YANTIAN|SHEKOU|SHENZHEN|NINGBO|SHANGHAI|QINGDAO|XIAMEN|NANSHA|GUANGZHOU)\s*$",
        text, re.M)
    if m:
        fields.setdefault("vessel_p5", m.group(1).strip())
        fields.setdefault("voyage_p5", m.group(2).strip())
        fields.setdefault("pol_p5", m.group(3).strip())

    m2 = re.search(r"Port\s+of\s+Discharge[^\n]*\n\s*([A-Za-z][A-Za-z\s,\.]{2,40}?)\s*$",
                   text, re.M | re.I)
    if m2 and not fields.get("pod_p5"):
        fields["pod_p5"] = m2.group(1).strip()


def _parse_goods(text: str) -> list[dict]:
    """抽取报关单商品明细行。"""
    goods = []
    anchor = re.search(r"项号\s+商品编号\s+商品名称及规格型号", text)
    if not anchor:
        return goods
    body = text[anchor.end():]
    for line in body.splitlines():
        line = line.strip()
        if not line:
            continue
        if line.startswith(("特殊关系", "报关人员", "申报单位", "兹申明")):
            break
        m = re.match(r"^(\d{1,2})\s+(\d{8,10})\s+([^\s]+)\s+(.*)$", line)
        if m:
            goods.append({
                "seq": int(m.group(1)),
                "hs_code": m.group(2),
                "name_cn": m.group(3),
                "rest": m.group(4).strip(),
            })
    return goods


def _post_derive(fields: dict, result: dict) -> None:
    """派生字段与候选值构建。"""
    # 规范字段合并（坐标/正则同名，去掉 _p1/_raw 后缀便于下游使用）
    if fields.get("pkg_count_p1") is not None:
        fields["pkg_count"] = fields["pkg_count_p1"]
    if fields.get("gross_kg_p1") is not None:
        fields["gross_kg"] = fields["gross_kg_p1"]
    if fields.get("net_kg_p1") is not None:
        fields["net_kg"] = fields["net_kg_p1"]
    if fields.get("consignee_raw"):
        fields["consignee"] = fields["consignee_raw"]

    # 船名航次双候选（不自动拍板，直接交界面勾选）
    candidates = []
    v5 = fields.get("vessel_p5")
    y5 = fields.get("voyage_p5")
    if v5:
        val = f"{v5} {y5}".strip() if y5 else v5
        candidates.append({"value": val, "source": "装货单(Ocean Vessel/Voy No)",
                           "recommend": True})
    v1 = fields.get("vessel_voyage_p1")
    if v1:
        candidates.append({"value": v1, "source": "报关单(运输工具名称及航次号)",
                           "recommend": False})
    result["vessel_candidates"] = candidates

    # 装货港：优先装货单英文，其次离境口岸中译英
    pol = fields.get("pol_p5") or cn_port_to_en(fields.get("departure_port"))
    if pol:
        fields["pol"] = pol

    # 卸货港
    pod = fields.get("pod_p5")
    if not pod:
        pod = DEST_CODE2EN.get((fields.get("dest_country") or "").upper())
    if pod:
        fields["pod"] = pod.upper()

    # 柜型描述 1X40HQ
    if fields.get("cntr_size"):
        qty = fields.get("cntr_qty") or 1
        fields["cntr_desc"] = f"{qty}X{fields['cntr_size']}HQ"

    # 申报日期 YYYYMMDD -> YYYY-MM-DD
    ed = fields.get("export_date")
    if ed and re.fullmatch(r"\d{8}", str(ed)):
        fields["export_date_iso"] = f"{ed[:4]}-{ed[4:6]}-{ed[6:]}"
