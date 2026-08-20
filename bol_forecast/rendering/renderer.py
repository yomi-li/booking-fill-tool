# -*- coding: utf-8 -*-
"""渲染器选择：env BOL_RENDERER > config.json "renderer" > 默认 html。

返回值: "html" | "com"
- "html": WeasyPrint HTML/CSS → PDF（新主路径，跨平台、Linux 部署无需 Office）
- "com":  Excel COM 复制模板 → 写值 → ExportAsFixedFormat（Windows+Office 机器上的回退/对版基线）

热切换：改环境变量或 config.json 后重启服务即生效，无需改代码。
"""
from __future__ import annotations

import os
import sys
from typing import Literal

from bol_forecast.config import CFG

RendererMode = Literal["html", "com"]

_ENV_KEY = "BOL_RENDERER"
_DEFAULT = "html"


def renderer_mode() -> RendererMode:
    """返回当前生效的渲染模式。非法值回退默认。"""
    env = os.environ.get(_ENV_KEY, "").strip().lower()
    if env in ("html", "com"):
        return env
    cfg_val = (CFG.get("renderer") or _DEFAULT).strip().lower()
    if cfg_val in ("html", "com"):
        return cfg_val
    return _DEFAULT  # type: ignore[return-value]


def is_html() -> bool:
    return renderer_mode() == "html"


def is_com() -> bool:
    return renderer_mode() == "com"
