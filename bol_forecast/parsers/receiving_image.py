"""收货数据截图解析器。

从系统收货数据截图中提取：
  - 件数 (pkg_count)          —— 严格取「件数」关键字后的数字
  - 收货实重 (gross_weight)
  - 收费重 / 计费重 (chargeable_weight)
  - 收货体积 (volume_cbm)

提取策略（三级，按可靠度降序）：
  L1 行内锚定：同一行 `标签 ... 数字`，标签与数字之间不得跨越另一个标签
  L2 列对齐  ：标签在表头行、数值在其下方行的相同列区间（表格截图常见）
  L3 位置兜底：仅用于重量/体积；**件数绝不使用位置兜底**（避免误取收费重）

互不干扰保证：
  - 标签集合统一按「最长优先」切分，`收费重`/`计费重` 先于 `重量`/`实重` 命中；
  - 每个数字 token 只允许被一个字段占用（区间去重）；
  - `件数` 的候选值必须落在其标签的作用域内，跨标签的数字直接丢弃。

依赖：pytesseract + tesseract-ocr。OCR 不可用时返回空结果 + warning，不抛错。
"""
from __future__ import annotations

import logging
import os
import re
from typing import Any

from PIL import Image, ImageEnhance, ImageFilter

# pytesseract 为可选依赖：未安装时 OCR 降级为空结果，且不拖垮整个模块导入
def _get_pytesseract():
    try:
        import pytesseract
        return pytesseract
    except Exception:
        return None

log = logging.getLogger(__name__)

# 尝试定位 tesseract；常见 Windows 路径
_TESS_PATHS = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
]


def _init_tesseract() -> bool:
    """设置 tesseract 路径与中文训练数据目录。成功返回 True。"""
    pt = _get_pytesseract()
    if pt is None:
        return False
    # 指向项目级 tessdata 目录（含 eng + chi_sim），避免写入 Program Files
    # receiving_image.py 在 bol_forecast/parsers/，项目根在其上两级
    _project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    _tessdata_dir = os.path.join(_project_root, "tessdata")
    if os.path.isdir(_tessdata_dir) and os.path.isfile(os.path.join(_tessdata_dir, "chi_sim.traineddata")):
        os.environ["TESSDATA_PREFIX"] = _tessdata_dir
    cmd = getattr(pt.pytesseract, "tesseract_cmd", None)
    if cmd and os.path.isfile(cmd):
        return True
    for p in _TESS_PATHS:
        if os.path.isfile(p):
            pt.pytesseract.tesseract_cmd = p
            return True
    try:
        pt.get_tesseract_version()
        return True
    except Exception:
        return False


# ── 图像预处理管线 ──
def _preprocess_image(img: Image.Image) -> list[Image.Image]:
    """返回多份预处理后的图像，供不同 PSM 模式尝试。

    管线：
      1. 原图灰度 → 锐化 → 放大2x → Otsu 二值化 → 轻量去噪（主力）
      2. 原图灰度 → 强对比度(3x) → 放大2x → 自适应阈值（备选）
      3. 原图灰度 → 放大3x → 高对比度 → 二值化（高分辨率兜底）
    """
    base = img.convert("L")
    results: list[Image.Image] = []

    # ── 变体1：标准管线（最稳）──
    p1 = ImageEnhance.Sharpness(base).enhance(2.0)
    w, h = p1.size
    p1 = p1.resize((w * 2, h * 2), Image.LANCZOS)
    p1 = ImageEnhance.Contrast(p1).enhance(2.5)
    # Otsu 风格二值化：用均值作为动态阈值
    import statistics
    pixels = list(p1.getdata())
    try:
        thresh = statistics.mean(pixels) * 0.85
    except Exception:
        thresh = 160
    p1 = p1.point(lambda x: 255 if x > thresh else 0)
    p1 = p1.filter(ImageFilter.MedianFilter(size=3))
    results.append(p1)

    # ── 变体2：强对比度 + 自适应阈值（应对低对比截图）──
    p2 = ImageEnhance.Contrast(base).enhance(3.5)
    w, h = p2.size
    p2 = p2.resize((w * 2, h * 2), Image.LANCZOS)
    # 局部自适应：用邻域均值做阈值近似
    thresh2 = 140
    p2 = p2.point(lambda x: 255 if x > thresh2 else 0)
    results.append(p2)

    # ── 变体3：高分辨率放大（应对小字/低 DPI 截图）──
    p3 = base
    w, h = p3.size
    p3 = p3.resize((w * 3, h * 3), Image.LANCZOS)
    p3 = ImageEnhance.Contrast(p3).enhance(2.0)
    p3 = ImageEnhance.Sharpness(p3).enhance(1.5)
    # 用较高阈值保留清晰笔画
    p3 = p3.point(lambda x: 255 if x > 150 else 0)
    results.append(p3)

    return results


# ── 多策略 OCR 引擎 ─_
_OCR_CONFIGS = [
    "--psm 6 --oem 1",   # 均匀文本块（默认）
    "--psm 7 --oem 1",   # 单行文本（适合单行收货数据）
    "--psm 11 --oem 1",  # 稀疏文本（适合分散字段）
    "--psm 6 --oem 0",   # Legacy OEM 兜底
]


def _ocr_variants(pt, img_original: Image.Image) -> list[tuple[str, str]]:
    """对每份预处理图 × 每个 OCR 配置执行识别，返回 [(config, text)]。

    注意：不抛错，单个配置失败只记录日志并跳过。
    """
    variants: list[tuple[str, str]] = []
    preprocessed = _preprocess_image(img_original)
    # 也把原图加进去（有些情况下原图反而更好）
    all_images = [img_original.convert("L")] + preprocessed

    for img_idx, img in enumerate(all_images):
        for cfg_idx, config in enumerate(_OCR_CONFIGS):
            try:
                text = pt.image_to_string(img, lang="chi_sim+eng", config=config)
                if text and text.strip():
                    variants.append((f"img{img_idx}_cfg{cfg_idx}", text.strip()))
            except Exception as e:
                log.debug("OCR variant img%d cfg%s 失败: %s", img_idx, cfg_idx, e)
                continue

    # 去重：完全相同的文本只保留第一个（通常是最佳配置）
    seen: set[str] = set()
    unique: list[tuple[str, str]] = []
    for name, text in variants:
        if text not in seen:
            seen.add(text)
            unique.append((name, text))
    return unique


# ── 字段标签定义（顺序＝匹配优先级；同一字段内部亦按长→短） ──
# 注意：`收费重/计费重` 必须排在 `重量/实重/毛重` 之前，否则会被后者吃掉。
FIELD_LABELS: dict[str, list[str]] = {
    "chargeable_weight": [
        r"收费重量", r"计费重量", r"收费重", r"计费重", r"计costs?重",
        r"C\.?\s?W\b", r"Chargeable\s*Weight", r"Chargeable",
    ],
    "pkg_count": [
        r"总件数", r"实收件数", r"收货件数", r"件\s*数", r"件数",
        r"总箱数", r"箱\s*数", r"PKGS?\b", r"CTNS?\b", r"PCS\b", r"Packages?\b",
    ],
    "gross_weight": [
        r"收货实重", r"实际重量", r"实\s*重", r"毛\s*重", r"总重量", r"总\s*重",
        r"G\.?\s?W\b", r"Gross\s*Weight", r"重\s*量\b", r"Weight\b",
    ],
    "volume_cbm": [
        r"收货体积", r"实际体积", r"体\s*积", r"材\s*积", r"方\s*数",
        r"CBM\b", r"M\s?3\b", r"m³", r"Volume\b", r"立方",
    ],
}

# 数值 token：允许千分位逗号与小数
_RE_NUM = re.compile(r"(?<![\d.])(\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d{1,9}(?:\.\d{1,4})?)(?![\d])")

# 数字与标签之间允许出现的噪声（冒号/空格/单位/竖线/等号等），不允许出现汉字标签
_GAP_OK = re.compile(r"^[\s:：=\-—|｜/\\.,，、()（）\[\]]*$")

# 单位后缀，用于校验
_UNIT_AFTER = re.compile(r"^\s*(KGS?|公斤|千克|CBM|m³|M3|立方|件|箱|PKGS?|CTNS?)", re.I)

# 字段合理值域（防止把日期/单号当数值）
_RANGE = {
    "pkg_count": (1, 100000),
    "gross_weight": (0.01, 1_000_000),
    "chargeable_weight": (0.01, 1_000_000),
    "volume_cbm": (0.0001, 100_000),
}


def _to_float(s: str) -> float | None:
    try:
        return float(s.replace(",", ""))
    except (TypeError, ValueError):
        return None


def _build_label_index(text: str) -> list[tuple[int, int, str]]:
    """扫描全文，返回所有标签命中 [(start, end, field)]，按最长优先去重。"""
    hits: list[tuple[int, int, str]] = []
    for field, pats in FIELD_LABELS.items():
        for pat in pats:
            for m in re.finditer(pat, text, re.I):
                hits.append((m.start(), m.end(), field))
    # 最长优先 + 起点靠前优先；重叠区间只保留第一个
    hits.sort(key=lambda h: (h[0], -(h[1] - h[0])))
    picked: list[tuple[int, int, str]] = []
    for h in hits:
        if any(not (h[1] <= p[0] or h[0] >= p[1]) for p in picked):
            continue  # 与已选标签重叠 → 丢弃（避免 收费重/重量 互吃）
        picked.append(h)
    picked.sort(key=lambda h: h[0])
    return picked


def _in_range(field: str, val: float) -> bool:
    lo, hi = _RANGE[field]
    return lo <= val <= hi


def _extract_inline(line: str, labels: list[tuple[int, int, str]],
                    used: set[int]) -> dict[str, float]:
    """L1：同一行内，标签后取第一个数字；不得越过下一个标签。"""
    out: dict[str, float] = {}
    for i, (ls, le, field) in enumerate(labels):
        if field in out:
            continue
        # 作用域 = 本标签结束 → 下一个标签开始
        scope_end = labels[i + 1][0] if i + 1 < len(labels) else len(line)
        seg = line[le:scope_end]
        m = _RE_NUM.search(seg)
        if not m:
            continue
        # 标签与数字之间只允许噪声字符（不能夹着别的汉字词，防串行）
        if not _GAP_OK.match(seg[:m.start()]):
            continue
        abs_pos = le + m.start()
        if abs_pos in used:
            continue
        val = _to_float(m.group(1))
        if val is None or not _in_range(field, val):
            continue
        if field == "pkg_count" and abs(val - round(val)) > 1e-6:
            continue  # 件数必须是整数
        out[field] = val
        used.add(abs_pos)
    return out


def _ok_value(field: str, val: float | None) -> bool:
    if val is None or not _in_range(field, val):
        return False
    if field == "pkg_count" and abs(val - round(val)) > 1e-6:
        return False  # 件数必须是整数
    return True


def _extract_column(lines: list[str], used_fields: set[str]) -> dict[str, float]:
    """L2：表头行 + 数值行对齐（表格截图）。

    优先「序号配对」：表头标签数 == 数值行数字个数时，按出现顺序一一对应，
    这是表格 OCR 最稳的方式，天然杜绝件数取到收费重。
    数量不等时退回「列距离配对」，且每个数字只能被一个字段占用。
    """
    out: dict[str, float] = {}
    for idx, line in enumerate(lines):
        labels = _build_label_index(line)
        if not labels:
            continue
        if _RE_NUM.search(line):
            continue  # 行内含数字 → 属于 L1 的场景，避免重复/串行

        for j in range(idx + 1, min(idx + 4, len(lines))):
            nxt = lines[j]
            if _build_label_index(nxt):
                continue  # 仍是表头行，继续往下找数值行
            nums = [(m, _to_float(m.group(1))) for m in _RE_NUM.finditer(nxt)]
            nums = [(m, v) for m, v in nums if v is not None]
            if not nums:
                continue

            if len(nums) == len(labels):           # ① 序号配对
                for (ls, le, field), (m, v) in zip(labels, nums):
                    if field in out or field in used_fields or not _ok_value(field, v):
                        continue
                    out[field] = v
            else:                                   # ② 列距离配对（互斥）
                taken: set[int] = set()
                for ls, le, field in labels:
                    if field in out or field in used_fields:
                        continue
                    center = (ls + le) / 2
                    best: tuple[float, float, int] | None = None
                    for k, (m, v) in enumerate(nums):
                        if k in taken or not _ok_value(field, v):
                            continue
                        d = abs((m.start() + m.end()) / 2 - center)
                        if d > 14:      # 列偏移过大，不是同一列
                            continue
                        if best is None or d < best[0]:
                            best = (d, v, k)
                    if best is not None:
                        out[field] = best[1]
                        taken.add(best[2])
            break
    return out


def _fallback_by_position(text: str) -> dict[str, float]:
    """L3：位置兜底。**仅在一个关键字都没命中时启用**，且件数永不参与。

    只要有任意关键字命中，就说明 OCR 文本结构可读，缺失即真缺失 ——
    此时宁可留空也不猜，避免把收费重/体积的数字塞进别的字段。
    """
    out: dict[str, float] = {}
    nums = [v for v in (_to_float(m.group(1)) for m in _RE_NUM.finditer(text))
            if v is not None]
    decimals = [v for v in nums if abs(v - round(v)) > 1e-6]
    for v in decimals:
        if v > 10:
            out["gross_weight"] = round(v, 2)
            break
    for v in decimals:
        if 0.001 < v < 1000 and v != out.get("gross_weight"):
            out["volume_cbm"] = round(v, 2)
            break
    return out


def _compress_cjk_spaces(text: str) -> str:
    """压缩汉字/CJK 字符间的空格（tesseract 常在中文 OCR 输出字间空格）。

    反复执行直到稳定（处理「A B C」→「AB C」→「ABC」链式空格）。
    """
    import re as _re
    _pat = re.compile(r'([\u4e00-\u9fff\u3400-\u4dbf])\s+([\u4e00-\u9fff\u3400-\u4dbf])')
    prev = None
    while prev != text:
        prev = text
        text = _pat.sub(r'\1\2', text)
    return text


def extract_receiving_fields(text: str) -> dict[str, Any]:
    """从纯文本（OCR 结果或手工粘贴）抽取收货四字段。可独立单测。"""
    result: dict[str, Any] = {"warnings": []}
    if not text or not text.strip():
        result["warnings"].append("收货数据为空，未识别到任何内容")
        return result

    # 预处理：压缩 CJK 字间空格（tesseract 中文 OCR 常见输出问题）
    text = _compress_cjk_spaces(text)

    lines = [ln for ln in text.splitlines()]

    # ── L1 行内锚定 ──
    found: dict[str, float] = {}
    for line in lines:
        labels = _build_label_index(line)
        if not labels:
            continue
        used: set[int] = set()
        got = _extract_inline(line, labels, used)
        for k, v in got.items():
            found.setdefault(k, v)

    # ── L2 列对齐（补缺） ──
    missing = {f for f in FIELD_LABELS if f not in found}
    if missing:
        col = _extract_column(lines, set(found))
        for k, v in col.items():
            if k in missing:
                found.setdefault(k, v)

    # ── L3 位置兜底：仅当一个关键字都没命中时启用（件数不参与） ──
    if not found:
        for k, v in _fallback_by_position(text).items():
            found[k] = v
            result["warnings"].append(
                f"未匹配到任何字段关键字，{k} 按数值特征推断为 {v}，请人工核对")

    # ── 落值 + 规范化 ──
    if "pkg_count" in found:
        result["pkg_count"] = int(round(found["pkg_count"]))
    else:
        result["warnings"].append(
            "未识别到「件数」关键字后的数字，件数留空（不会误取收费重），请手动填写")
    for k, nd in (("gross_weight", 2), ("chargeable_weight", 2), ("volume_cbm", 2)):
        if k in found:
            result[k] = round(found[k], nd)

    # 交叉校验：件数与收费重不得为同一数值来源（理论上已被区间去重排除）
    if (result.get("pkg_count") is not None
            and result.get("chargeable_weight") is not None
            and float(result["pkg_count"]) == float(result["chargeable_weight"])):
        result["warnings"].append(
            "件数与收费重数值相同，可能识别串行，请核对截图")
    return result


def parse_receiving_text(text: str) -> dict[str, Any]:
    """解析手工粘贴的收货文本（无需 OCR）。"""
    r = extract_receiving_fields(text)
    r["_source"] = "text"
    r["raw_text"] = (text or "")[:2000]
    return r


def parse_receiving_image(image_path: str) -> dict[str, Any]:
    """解析收货数据截图（多策略 OCR + 最佳结果选取）。

    策略：
      1. 对截图做 3 种预处理 × 4 种 PSM/OEM 配置 = 最多 16 次 OCR
      2. 每次结果经 ``extract_receiving_fields`` 解析，统计命中的字段数
      3. 取命中字段最多的结果；并列时优先含「件数」的

    Returns
    -------
    dict
        ``{ pkg_count?, gross_weight?, chargeable_weight?, volume_cbm?,
            _source: "ocr", raw_text?, warnings[] }``
        字段缺失表示未能识别（不会用其它字段的数字顶替）。
    """
    result: dict[str, Any] = {"_source": "ocr", "warnings": []}

    if not _init_tesseract():
        msg = "tesseract 未安装或找不到（winget install UB-Mannheim.TesseractOCR）"
        result["warnings"].append(msg)
        log.warning(msg)
        return result

    if not os.path.isfile(image_path):
        msg = f"文件不存在: {image_path}"
        result["warnings"].append(msg)
        log.warning(msg)
        return result

    try:
        img = Image.open(image_path)
        pt = _get_pytesseract()
        if pt is None:
            raise RuntimeError("pytesseract 未安装")

        # ── 多策略 OCR ──
        variants = _ocr_variants(pt, img)
        if not variants:
            result["warnings"].append("所有 OCR 策略均未返回文本，图片可能为空或无法识别")
            return result

        # ── 选最佳：字段命中最多 > 含件数优先 > 文本最长 ──
        best_parsed = None
        best_score = (-1, -1, -1)  # (字段数, 有件数?, 文本长)
        best_raw = ""
        best_name = ""

        for vname, text in variants:
            parsed = extract_receiving_fields(text)
            n_fields = sum(1 for k in ("pkg_count", "gross_weight",
                                       "chargeable_weight", "volume_cbm")
                           if k in parsed)
            has_pkg = 1 if "pkg_count" in parsed else 0
            score = (n_fields, has_pkg, len(text))
            if score > best_score:
                best_score = score
                best_parsed = parsed
                best_raw = text[:2000]
                best_name = vname

        if best_parsed is None:
            result["warnings"].append("OCR 识别到文本但无法解析出任何有效字段")
            return result

        log.info("收货数据 OCR 最佳策略: %s (字段=%d, 件数=%d)",
                 best_name, best_score[0], best_score[1])
        warns = result["warnings"] + best_parsed.pop("warnings", [])
        result.update(best_parsed)
        result["raw_text"] = best_raw
        result["warnings"] = warns

        # 如果最佳结果仍缺关键字段，提示用户
        missing = [k for k in ("pkg_count", "gross_weight",
                               "chargeable_weight", "volume_cbm")
                   if k not in result]
        if missing and best_score[0] > 0:
            cn = {"pkg_count": "件数", "gross_weight": "实重",
                  "chargeable_weight": "收费重/计费重", "volume_cbm": "体积"}
            result["warnings"].append(
                f"部分字段未识别到（{'、'.join(cn[m] for m in missing)}），"
                "可尝试重新截取更清晰的图片或手动填写")

        log.info("收货数据解析完成: %s",
                 {k: v for k, v in result.items()
                  if k not in ("raw_text", "_source", "warnings")})
        return result

    except Exception as e:
        msg = f"OCR 失败: {e}"
        result["warnings"].append(msg)
        log.warning(msg)
        return result
