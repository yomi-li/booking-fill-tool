# -*- coding: utf-8 -*-
"""HTML 模板渲染 + WeasyPrint → PDF。

设计要点：
- 模板与数据分离：模板 bol_forecast/html_templates/{doc_key}.html，
  渲染时把 context dict 喂给占位符 {{ ctx.field }}，再由 WeasyPrint 出 PDF。
- Jinja2 Environment 启动时只装 autoescape（PDF 不需要 HTML 转义，但保证安全）。
- 文本自适应：调用方先跑 textfit.apply_sizes(ctx, doc_key)，
  模板里相关字段用 <span style="font-size:{{ ctx.sizes.field }}pt">…</span>。
- 字体自适应：@font-face 把 Lato 嵌进 PDF；中文走系统字体回退栈。
- Linux/macOS 上系统 pango，Windows 上由 pango_runtime 自动注入。
"""
from __future__ import annotations

import logging
import os
from pathlib import Path
from typing import Any

from bol_forecast.config import ROOT

log = logging.getLogger(__name__)

# 模板与静态资源目录（绝对路径）
HTML_TEMPLATES_DIR = ROOT / "bol_forecast" / "html_templates"
STATIC_DIR = ROOT / "bol_forecast" / "static"


def _env():
    """延迟初始化 Jinja2 环境（避免无谓的 import-time 启动）。"""
    from jinja2 import Environment, FileSystemLoader, select_autoescape
    env = Environment(
        loader=FileSystemLoader(str(HTML_TEMPLATES_DIR)),
        autoescape=select_autoescape(["html", "xml"]),
        trim_blocks=True,
        lstrip_blocks=True,
    )
    return env


def render_doc(doc_key: str, ctx: dict[str, Any], out_pdf: str | os.PathLike) -> Path:
    """把 context 喂给 html_templates/{doc_key}.html，渲染为 PDF。

    ctx 至少包含业务字段（shipper/bl_no/...）+ 可选 sizes 字段（textfit 输出）。
    """
    out_pdf = Path(out_pdf)
    out_pdf.parent.mkdir(parents=True, exist_ok=True)

    # Windows 上确保 pango DLL 可被找到（Linux 上 no-op）
    from bol_forecast.rendering.pango_runtime import ensure_pango_dll_dir
    injected = ensure_pango_dll_dir()
    if injected:
        log.debug("WeasyPrint pango DLL 来源: %s", injected)

    env = _env()
    template = env.get_template(f"{doc_key}.html")
    html = template.render(ctx=ctx, static_url=str(STATIC_DIR))

    # WeasyPrint 函数内 import，便于在 Linux 上 import 整个 bol_forecast 不必强依赖 weasyprint
    import weasyprint
    wp = weasyprint.HTML(string=html, base_url=str(STATIC_DIR))
    wp.write_pdf(str(out_pdf))
    return out_pdf
