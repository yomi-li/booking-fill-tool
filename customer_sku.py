# -*- coding: utf-8 -*-
"""客户 SKU 产品库：JSON 持久化 + 增删改查。

库结构 customer_sku.json:
{
  "brands": ["Bozeman", "Coletti"],
  "skus": {
     "<SKU码>": {
        "sku": "<SKU码>",
        "cn_name": "中文品名", "en_name": "英文品名",
        "hs_code": "海关编码",
        "material_cn": "材质(中)", "material_en": "材质(英)",
        "brand": "品牌", "brand_type": "品牌类型",
        "model": "型号", "purpose": "用途", "electric_magnetic": "带电/磁",
        "qty_per_ctn": 每箱个数,
        "unit_price": 申报单价, "currency": "USD",
        "po_price": 采购单价, "po_currency": "CNY",
        "length": 长cm, "width": 宽cm, "height": 高cm,
        "image_path": "图片绝对路径或空", "product_link": "平台链接",
        "reference_id": "亚马逊内部编号(可选)"
     }, ...
  }
}
"""
import json
import os

BASE = os.path.dirname(os.path.abspath(__file__))
LIB_PATH = os.path.join(BASE, "customer_sku.json")
IMAGE_DIR = os.path.join(BASE, "sku_images")

SKU_FIELDS = [
    "sku", "cn_name", "en_name", "hs_code",
    "material_cn", "material_en", "brand", "brand_type",
    "model", "purpose", "electric_magnetic", "qty_per_ctn",
    "net_per_ctn", "gross_per_ctn", "total_cartons", "total_qty",
    "unit_price", "total_price", "currency", "po_price", "po_currency",
    "length", "width", "height",
    "image_path", "product_link", "reference_id",
]


def _default_lib():
    return {"brands": ["Bozeman", "Coletti"], "skus": {}}


def load_library():
    if not os.path.exists(LIB_PATH):
        return _default_lib()
    with open(LIB_PATH, "r", encoding="utf-8") as f:
        data = json.load(f)
    data.setdefault("brands", ["Bozeman", "Coletti"])
    data.setdefault("skus", {})
    return data


def save_library(lib):
    os.makedirs(os.path.dirname(LIB_PATH), exist_ok=True)
    with open(LIB_PATH, "w", encoding="utf-8") as f:
        json.dump(lib, f, ensure_ascii=False, indent=2)


def list_skus(keyword=""):
    lib = load_library()
    skus = list(lib["skus"].values())
    if keyword:
        kw = keyword.lower()
        skus = [
            s for s in skus
            if kw in (s.get("sku", "") or "").lower()
            or kw in (s.get("cn_name", "") or "").lower()
            or kw in (s.get("en_name", "") or "").lower()
            or kw in (s.get("brand", "") or "").lower()
        ]
    return {"brands": lib["brands"], "skus": skus}


def get_sku(sku):
    lib = load_library()
    return lib["skus"].get(sku)


def upsert_sku(sku, data: dict, merge: bool = True):
    """新增或更新一个 SKU。sku 为键且必须与 data['sku'] 一致。

    merge=True（默认，用于前端编辑/图片上传）：对于 data 中未提供的字段，
    保留库中已有值，避免编辑时把 Coletti 导入的 net_per_ctn/长宽高 等字段清零。
    merge=False（用于整条刷新，如 Coletti 导入）：完全以 data 为准。
    """
    lib = load_library()
    skus = lib["skus"]
    existing = skus.get(sku, {})
    rec = {}
    for f in SKU_FIELDS:
        if f == "sku":
            continue
        if f in data:
            val = data[f]
            rec[f] = "" if val is None else val
        elif merge and f in existing:
            rec[f] = existing[f]
        else:
            rec[f] = ""
    rec["sku"] = sku
    # 数值字段清洗
    for numf in ("qty_per_ctn", "net_per_ctn", "gross_per_ctn", "total_cartons",
                 "total_qty", "unit_price", "total_price", "po_price",
                 "length", "width", "height"):
        v = rec.get(numf)
        if v in (None, ""):
            rec[numf] = ""
        else:
            try:
                rec[numf] = float(v)
            except (TypeError, ValueError):
                rec[numf] = str(v)
    if "currency" not in rec or not rec["currency"]:
        rec["currency"] = "USD"
    if "po_currency" not in rec or not rec["po_currency"]:
        rec["po_currency"] = "CNY"
    skus[sku] = rec
    save_library(lib)
    return rec


def delete_sku(sku):
    lib = load_library()
    if sku in lib["skus"]:
        rec = lib["skus"].pop(sku)
        save_library(lib)
        # 同步清理该 SKU 孤立的产品图片文件
        img = rec.get("image_path")
        if img:
            try:
                bn = os.path.basename(img.replace("\\", "/"))
                p = os.path.join(IMAGE_DIR, bn)
                if os.path.exists(p):
                    os.remove(p)
            except Exception:
                pass
        return True
    return False


def match_skus(sku_codes):
    """给定一组 SKU 码，返回 [{sku, matched, record}]。"""
    lib = load_library()
    out = []
    for code in sku_codes:
        rec = lib["skus"].get(code)
        out.append({"sku": code, "matched": rec is not None, "record": rec})
    return out
