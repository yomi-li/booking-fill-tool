# -*- coding: utf-8 -*-
"""费用公式引擎 —— AST 白名单求值，绝不使用 eval/exec。

允许变量：CBM GW NW CTNS CNTR PCS PRICE QTY
允许函数：max min round ceil floor abs
允许运算：+ - * / // % ** 一元负号 以及比较/三元不参与（保持简单）
"""
from __future__ import annotations

import ast
import math
import operator as _op

__all__ = ["calc", "ALLOWED_VARS", "ALLOWED_FUNCS", "FormulaError", "compute_amount"]


class FormulaError(ValueError):
    """公式语法或变量不合法。"""


_BINOPS = {
    ast.Add: _op.add,
    ast.Sub: _op.sub,
    ast.Mult: _op.mul,
    ast.Div: _op.truediv,
    ast.FloorDiv: _op.floordiv,
    ast.Mod: _op.mod,
    ast.Pow: _op.pow,
}
_UNARYOPS = {ast.UAdd: _op.pos, ast.USub: _op.neg}

def _round(x, ndigits=0):
    """round 的第二参数必须是 int，而 AST 求值统一产出 float，此处强制转换。"""
    return round(float(x), int(ndigits))


ALLOWED_FUNCS = {
    "max": max,
    "min": min,
    "round": _round,
    "ceil": lambda x: float(math.ceil(float(x))),
    "floor": lambda x: float(math.floor(float(x))),
    "abs": abs,
}

ALLOWED_VARS = ("CBM", "GW", "NW", "CTNS", "CNTR", "PCS", "PRICE", "QTY")

_MAX_LEN = 200


def calc(expr: str, env: dict | None = None, ndigits: int = 2) -> float:
    """安全求值。env 为变量表，缺失变量按 0 处理。"""
    env = env or {}
    if not expr or not str(expr).strip():
        return 0.0
    src = str(expr).strip()
    if len(src) > _MAX_LEN:
        raise FormulaError("公式过长")

    try:
        tree = ast.parse(src, mode="eval")
    except SyntaxError as e:
        raise FormulaError(f"公式语法错误: {e.msg}") from e

    def ev(node):
        if isinstance(node, ast.Constant):
            if isinstance(node.value, bool) or not isinstance(node.value, (int, float)):
                raise FormulaError("只允许数字常量")
            return float(node.value)
        if isinstance(node, ast.Name):
            key = node.id.upper()
            if key not in ALLOWED_VARS:
                raise FormulaError(f"不允许的变量: {node.id}")
            v = env.get(key, env.get(node.id, 0))
            try:
                return float(v or 0)
            except (TypeError, ValueError):
                return 0.0
        if isinstance(node, ast.BinOp):
            fn = _BINOPS.get(type(node.op))
            if fn is None:
                raise FormulaError("不支持的运算符")
            right = ev(node.right)
            if type(node.op) in (ast.Div, ast.FloorDiv, ast.Mod) and right == 0:
                raise FormulaError("除数为零")
            return fn(ev(node.left), right)
        if isinstance(node, ast.UnaryOp):
            fn = _UNARYOPS.get(type(node.op))
            if fn is None:
                raise FormulaError("不支持的一元运算")
            return fn(ev(node.operand))
        if isinstance(node, ast.Call):
            if not isinstance(node.func, ast.Name):
                raise FormulaError("不支持的调用形式")
            name = node.func.id.lower()
            if name not in ALLOWED_FUNCS:
                raise FormulaError(f"不允许的函数: {node.func.id}")
            if node.keywords:
                raise FormulaError("函数不支持关键字参数")
            return float(ALLOWED_FUNCS[name](*[ev(a) for a in node.args]))
        raise FormulaError("公式含不支持的语法")

    return round(float(ev(tree.body)), ndigits)


def build_env(data) -> dict:
    """从解析结果构造公式变量表。data 支持 dict 或有属性的对象。"""
    def g(key, default=0):
        if isinstance(data, dict):
            return data.get(key, default)
        return getattr(data, key, default)

    return {
        "CBM": float(g("cbm") or 0),
        "GW": float(g("gross_kg") or 0),
        "NW": float(g("net_kg") or 0),
        "CTNS": float(g("ctns") or 0),
        "CNTR": float(g("cntr_qty") or 1),
        "PCS": float(g("pcs_total") or 0),
    }


def compute_amount(row: dict, env: dict) -> float:
    """按 calc_mode 统一算出金额，并套用保底/封顶。

    row: {calc_mode, unit_price, qty_var, qty, formula, min_amount, max_amount}
    """
    mode = (row.get("calc_mode") or "FIXED").upper()
    price = float(row.get("unit_price") or 0)

    if mode == "FIXED":
        qty = 1.0
        amount = price
    elif mode == "UNIT_PRICE":
        qv = (row.get("qty_var") or "").upper()
        if qv and qv in env:
            qty = float(env[qv])
        else:
            qty = float(row.get("qty") or 1)
        amount = price * qty
    elif mode == "FORMULA":
        qty = float(row.get("qty") or 1)
        amount = calc(row.get("formula") or "0", env)
    else:
        raise FormulaError(f"未知计算方式: {mode}")

    lo = row.get("min_amount")
    hi = row.get("max_amount")
    if lo not in (None, ""):
        amount = max(amount, float(lo))
    if hi not in (None, ""):
        amount = min(amount, float(hi))

    row["qty"] = round(qty, 4)
    row["amount"] = round(amount, 2)
    return row["amount"]
