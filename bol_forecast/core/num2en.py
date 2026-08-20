# -*- coding: utf-8 -*-
"""数字 → 英文大写（纯本地实现，零依赖）。

用于提单 F59 的 SAY ... ONLY 栏位。
"""
from __future__ import annotations

_ONES = [
    "ZERO", "ONE", "TWO", "THREE", "FOUR", "FIVE", "SIX", "SEVEN", "EIGHT",
    "NINE", "TEN", "ELEVEN", "TWELVE", "THIRTEEN", "FOURTEEN", "FIFTEEN",
    "SIXTEEN", "SEVENTEEN", "EIGHTEEN", "NINETEEN",
]
_TENS = ["", "", "TWENTY", "THIRTY", "FORTY", "FIFTY", "SIXTY", "SEVENTY",
         "EIGHTY", "NINETY"]

# 包装单位中英/复数映射
_UNIT_EN = {
    "CTNS": "CARTONS", "CTN": "CARTONS", "CARTON": "CARTONS",
    "CARTONS": "CARTONS", "纸箱": "CARTONS", "箱": "CARTONS",
    "PKGS": "PACKAGES", "PKG": "PACKAGES", "PACKAGE": "PACKAGES",
    "PACKAGES": "PACKAGES", "件": "PACKAGES",
    "PLTS": "PALLETS", "PLT": "PALLETS", "PALLET": "PALLETS",
    "PALLETS": "PALLETS", "托": "PALLETS", "托盘": "PALLETS",
    "CASE": "CASES", "CASES": "CASES",
    "BAG": "BAGS", "BAGS": "BAGS", "袋": "BAGS",
    "ROLL": "ROLLS", "ROLLS": "ROLLS", "卷": "ROLLS",
    "SET": "SETS", "SETS": "SETS", "套": "SETS",
    "PCS": "PIECES", "PC": "PIECES", "PIECE": "PIECES",
    "PIECES": "PIECES", "个": "PIECES",
    "DRUM": "DRUMS", "DRUMS": "DRUMS", "桶": "DRUMS",
    "BALE": "BALES", "BALES": "BALES",
    "CRATE": "CRATES", "CRATES": "CRATES", "木箱": "CRATES",
    "BUNDLE": "BUNDLES", "BUNDLES": "BUNDLES", "捆": "BUNDLES",
}

_SINGULAR = {
    "CARTONS": "CARTON", "PACKAGES": "PACKAGE", "PALLETS": "PALLET",
    "CASES": "CASE", "BAGS": "BAG", "ROLLS": "ROLL", "SETS": "SET",
    "PIECES": "PIECE", "DRUMS": "DRUM", "BALES": "BALE",
    "CRATES": "CRATE", "BUNDLES": "BUNDLE",
}


def _below_1000(n: int) -> str:
    if n < 20:
        return _ONES[n]
    if n < 100:
        rest = n % 10
        return _TENS[n // 10] + ("-" + _ONES[rest] if rest else "")
    rest = n % 100
    return _ONES[n // 100] + " HUNDRED" + (" AND " + _below_1000(rest) if rest else "")


def int_to_en(n: int) -> str:
    """0 ~ 999,999,999 转英文大写。"""
    n = int(n)
    if n < 0:
        return "MINUS " + int_to_en(-n)
    if n == 0:
        return "ZERO"
    parts: list[str] = []
    for base, name in ((1_000_000_000, "BILLION"), (1_000_000, "MILLION"),
                       (1_000, "THOUSAND"), (1, "")):
        if n >= base:
            q, n = divmod(n, base)
            parts.append(_below_1000(q) + (" " + name if name else ""))
    return " ".join(p for p in parts if p).strip()


def unit_to_en(unit: str | None) -> str:
    """包装单位归一为英文复数形式。未知单位原样大写返回。"""
    u = (unit or "CTNS").strip().upper()
    return _UNIT_EN.get(u, u)


def say_packages(n: int, unit: str | None = "CTNS") -> str:
    """96, 'CTNS' -> 'NINETY-SIX CARTONS'

    模板 E59='SAY ( ' 与 G59=')ONLY' 已固定，此处只产出中间内容。
    """
    n = int(n or 0)
    u = unit_to_en(unit)
    if n == 1:
        u = _SINGULAR.get(u, u)
    return f"{int_to_en(n)} {u}"


def say_full(n: int, unit: str | None = "CTNS", hyphen: bool = False) -> str:
    """整句形式：SAY NINETY SIX CARTONS ONLY（config.bl.say_style=full 时使用）。"""
    body = say_packages(n, unit)
    if not hyphen:
        body = body.replace("-", " ")
    return f"SAY {body} ONLY"
