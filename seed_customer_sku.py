# -*- coding: utf-8 -*-
"""一次性：从 bozeman_library.json 规范化字段并补充 Packing List 中的新 SKU，生成 customer_sku.json。"""
import json
import os
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
SRC = r"C:\Users\admin\WorkBuddy\2026-07-30-16-24-30\bozeman_library.json"
OUT = os.path.join(BASE, "customer_sku.json")


def norm(rec, sku):
    return {
        "sku": sku,
        "cn_name": rec.get("chinese_name", ""),
        "en_name": rec.get("english_name", ""),
        "hs_code": rec.get("hs_code", ""),
        "material_cn": rec.get("material_cn", ""),
        "material_en": rec.get("material_en", ""),
        "brand": rec.get("brand", "Bozeman"),
        "brand_type": rec.get("brand_type", "境外品牌（贴牌生产）"),
        "model": rec.get("model", ""),
        "purpose": rec.get("use", ""),
        "electric_magnetic": rec.get("cargo_type", "普货"),
        "qty_per_ctn": rec.get("pieces_per_carton", ""),
        "unit_price": rec.get("unit_price", ""),
        "currency": "USD",
        "po_price": "",
        "po_currency": "CNY",
        "length": rec.get("length", ""),
        "width": rec.get("width", ""),
        "height": rec.get("height", ""),
        "image_path": "",
        "product_link": "",
        "reference_id": "",
    }


def blank(sku, **kw):
    base = {
        "sku": sku, "cn_name": "", "en_name": "", "hs_code": "",
        "material_cn": "不锈钢", "material_en": "stainless steel",
        "brand": "Bozeman", "brand_type": "境外品牌（贴牌生产）",
        "model": "", "purpose": "煮咖啡", "electric_magnetic": "普货",
        "qty_per_ctn": "", "unit_price": "", "currency": "USD",
        "po_price": "", "po_currency": "CNY",
        "length": "", "width": "", "height": "",
        "image_path": "", "product_link": "", "reference_id": "",
    }
    base.update(kw)
    return base


with open(SRC, "r", encoding="utf-8") as f:
    src = json.load(f)

skus = {}
for k, v in src.items():
    skus[k] = norm(v, k)

# ---- 实际 Packing List 中出现、但库里没有的 SKU（占位，待用户在界面补全）----
skus["Scoutmaster-BLK-24C"] = blank(
    "Scoutmaster-BLK-24C",
    cn_name="不锈钢咖啡壶24杯黑色",
    en_name="Stainless steel percolator 24 cup black",
    hs_code="7323.93.0045",
    model="24 cups",
    qty_per_ctn=12,
)
skus["Bozeman-SS-NDC-06C"] = blank(
    "Bozeman-SS-NDC-06C",
    cn_name="不锈钢咖啡壶原色6杯电磁",
    en_name="stainless steel percolator 06 cups induction",
    hs_code="7323.93.0000",
    model="6 cups",
    qty_per_ctn=12,
)
skus["Filters-BSKT-BWN-100PK"] = blank(
    "Filters-BSKT-BWN-100PK",
    cn_name="咖啡滤篮棕色100片装",
    en_name="Coffee filter basket brown 100PK",
    hs_code="",
    material_cn="纸", material_en="paper",
    model="100PK",
    purpose="过滤咖啡",
    qty_per_ctn=48,
)

lib = {"brands": ["Bozeman", "Coletti"], "skus": skus}
with open(OUT, "w", encoding="utf-8") as f:
    json.dump(lib, f, ensure_ascii=False, indent=2)

print(f"✅ 已生成 customer_sku.json，共 {len(skus)} 个 SKU")
print("新增（待补全）:", [k for k in ("Scoutmaster-BLK-24C", "Bozeman-SS-NDC-06C", "Filters-BSKT-BWN-100PK")])
