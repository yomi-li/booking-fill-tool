# -*- coding: utf-8 -*-
"""文本自适应：替代 Excel COM 的 fit_cell。

Excel 靠 ws.Range(coord).Text 反馈来判断溢出并循环缩字号。WeasyPrint 没有
即时反馈机制，因此我们用 PIL ImageFont 在 Python 端预测量：以 4x 渲染
精度取字宽（亚像素更准），从 base_pt 起按 0.5pt 步长向下直到宽度 ≤ 盒宽。

设计取舍：
- 不引入第三方 lib（fontTools 可做更精细度量但远重于 PIL；PIL 够准）
- 4x 字号渲染精度消除亚像素抖动
- 字符含 CJK 时切换到 CJK 字体度量（宋体 / Noto Sans CJK）
- 命中 min_pt 仍未放下时返回 min_pt（不抛错；CSS overflow:hidden 兜底）

性能：每字段一次测量；每个文档最多约 25 字段（factory 最多），
单文档测量约 50-200ms，HTML→PDF 链路的非主导开销。
"""
from __future__ import annotations

import logging
from functools import lru_cache
from pathlib import Path

from PIL import ImageFont

log = logging.getLogger(__name__)

# ---- 字体路径（与 bol_forecast/static/fonts/ 对齐） ----
_STATIC_FONTS = Path(__file__).resolve().parents[1] / "static" / "fonts"
_FONT_PATH_LATIN = _STATIC_FONTS / "Lato-Regular.ttf"
_FONT_PATH_LATIN_BOLD = _STATIC_FONTS / "Lato-Bold.ttf"
# Windows 自带；Linux 由 fonts-noto-cjk 提供（PDF 渲染走 Pango 直接找系统字体，不依赖此变量）
# 这里用于测量，所以 Linux 也得有一个；留 None 时回退到 latin 度量（结果偏大但仍可用）
_FONT_PATH_CJK = None
for _c in (r"C:\Windows\Fonts\msyh.ttc",
           r"C:\Windows\Fonts\msyh.ttf",
           r"C:\Windows\Fonts\simsun.ttc",
           r"C:\Windows\Fonts\simhei.ttf"):
    if Path(_c).exists():
        _FONT_PATH_CJK = _c
        break

# ---- 度量精度 ----
_RENDER_SCALE = 4       # 字号 × 4 渲染 → 像素级更准
_STEP_PT = 0.5
_MIN_PT = 8.0


@lru_cache(maxsize=64)
def _font(path: str, size_pt: float) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(path, round(size_pt * _RENDER_SCALE))


def _is_cjk(text: str) -> bool:
    return any("\u4e00" <= ch <= "\u9fff" for ch in text)


def _pick_font_path(text: str, bold: bool = False) -> str:
    if _is_cjk(text) and _FONT_PATH_CJK:
        return _FONT_PATH_CJK
    if bold and _FONT_PATH_LATIN_BOLD.exists():
        return str(_FONT_PATH_LATIN_BOLD)
    return str(_FONT_PATH_LATIN)


def text_width_pt(text: str, size_pt: float, bold: bool = False) -> float:
    """返回 text 在 size_pt 字号下的渲染宽度（pt）。多行取最大行。"""
    if not text:
        return 0.0
    path = _pick_font_path(text, bold=bold)
    f = _font(path, size_pt)
    lines = str(text).split("\n")
    return max(f.getlength(line) for line in lines) / _RENDER_SCALE


def fit_font_size(text: str, base_pt: float, box_pt: float,
                  bold: bool = False, pad_pt: float = 1.0) -> float:
    """从 base_pt 起按 _STEP_PT 步长向下，找到宽度 ≤ box_pt - pad_pt 的最大字号；
    下限 _MIN_PT；仍未放下则返回 _MIN_PT。
    """
    if not text or box_pt <= 0:
        return max(_MIN_PT, base_pt)
    target = box_pt - pad_pt
    size = base_pt
    while size > _MIN_PT and text_width_pt(text, size, bold=bold) > target:
        size -= _STEP_PT
    if size < _MIN_PT:
        size = _MIN_PT
    return size
