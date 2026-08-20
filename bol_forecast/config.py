# -*- coding: utf-8 -*-
"""本模块（散货提单和预报生成）全局配置与路径常量。

注意：本文件刻意独立读取「本包内的 config.json」，绝不读取目标系统
「单证提取填充工具」的 config.json，避免配置串味。

所有路径一律绝对化（COM 只接受绝对路径）。DB 默认落在本包 data/
目录下；若目标系统通过环境变量 DATA_DIR 指定了持久卷，则落到该卷，
便于云端重部署后客户/话术/历史不丢失。
"""
from __future__ import annotations

import json
import os
from pathlib import Path

# 本包目录：booking-fill-tool/bol_forecast
BASE = Path(__file__).resolve().parent
# 目标系统根目录：booking-fill-tool
ROOT = BASE.parent

# 本模块专属配置（与目标的 config.json 隔离）
CONFIG_PATH = BASE / "config.json"


def _load() -> dict:
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


CFG: dict = _load()

# ---- 目录常量（全部绝对路径） ----
# 持久化目录：优先使用目标系统的 DATA_DIR 环境变量（云端持久卷），
# 否则落到本包 data/（本地便携版行为不变）。
DATA_DIR = Path(os.environ.get("DATA_DIR") or str(BASE / "data"))
TEMPLATES_DIR = ROOT / "bol_templates"
BACKUP_DIR = DATA_DIR / "backup"
JOBS_DIR = DATA_DIR / "jobs"
LOGS_DIR = DATA_DIR / "logs"
DB_PATH = DATA_DIR / "bol_forecast.db"
COM_PID_FILE = LOGS_DIR / "com_pids.txt"
TEMPLATE_FINGERPRINT = TEMPLATES_DIR / ".sha256"

for _d in (TEMPLATES_DIR, DATA_DIR, BACKUP_DIR, JOBS_DIR, LOGS_DIR):
    try:
        _d.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass


def template_path(doc_key: str) -> Path:
    """doc_key: bl | telex | factory | invoice"""
    name = CFG["template_files"][doc_key]
    return TEMPLATES_DIR / name


DOC_KEYS = ("bl", "telex", "factory", "invoice")


def doc_label(doc_key: str) -> str:
    return CFG["doc_labels"].get(doc_key, doc_key)


def expect_media(doc_key: str) -> int:
    return int(CFG["expect_media"].get(doc_key, 0))


# ---- Windows 文件名净化 ----
_ILLEGAL = '\\/:*?"<>|'


def safe_filename(name: str) -> str:
    out = "".join(("_" if c in _ILLEGAL else c) for c in str(name))
    return out.strip().strip(".") or "unnamed"


def customer_po_tag(order: dict) -> str:
    """从 order 取 customer_po，非空时返回 '_POxxx' 文件名后缀，否则空串。"""
    po = (order.get("customer_po") or "").strip()
    return f"_PO{safe_filename(po)}" if po else ""
