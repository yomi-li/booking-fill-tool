# -*- coding: utf-8 -*-
"""Windows 上 WeasyPrint 找不到 pango/glib 原生库。

开发机通过环境变量 WEASYPRINT_DLL_DIRECTORIES 注入路径。
本模块检测并自动注入「本机已有的」pango 路径：

  1. 用户手动设置 WEASYPRINT_DLL_DIRECTORIES：尊重之
  2. 否则：尝试 Tesseract-OCR 安装目录（C:\\Program Files\\Tesseract-OCR），
     这是 Tesseract-OCR 安装包自带的完整 pango/glib/harfbuzz/cairo 栈，
     是 WeasyPrint 官方推荐路径之外的零成本方案。
  3. 否则：不处理（用户须自行安装 MSYS2 pango 或装 GTK3-Runtime）。

生产环境（Linux Docker）走系统 pango，本模块 no-op。
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

_ENV_KEY = "WEASYPRINT_DLL_DIRECTORIES"
_TESSERACT_DEFAULT = r"C:\Program Files\Tesseract-OCR"


def _has_pango(d: Path) -> bool:
    """粗判：目录下是否存在 libpango/libgobject 这两个核心 DLL。"""
    if not d.is_dir():
        return False
    found = set()
    for f in d.iterdir():
        n = f.name.lower()
        if n.startswith("libpango-1.0-0"):
            found.add("pango")
        elif n.startswith("libgobject-2.0-0"):
            found.add("gobject")
    return "pango" in found and "gobject" in found


def ensure_pango_dll_dir() -> str | None:
    """Windows 上确保 WeasyPrint 能找到 pango DLL。返回生效的路径或 None。

    调用时机：html_render.py 模块导入或 render_doc() 入口处（import weasyprint 前）。
    Linux / macOS 上直接返回 None（走系统 pango）。
    """
    if sys.platform != "win32":
        return None
    cur = os.environ.get(_ENV_KEY, "").strip()
    if cur and any(_has_pango(Path(p)) for p in cur.split(os.pathsep) if p):
        return cur  # 用户已配置且有效，尊重之
    # 探测 Tesseract 目录
    t = Path(_TESSERACT_DEFAULT)
    if _has_pango(t):
        os.environ[_ENV_KEY] = str(t)
        return str(t)
    return None
