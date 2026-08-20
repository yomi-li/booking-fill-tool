"""轨迹信息解析器（截图 OCR / 粘贴文本）。

提取字段：
  - ATD (实际开船 / 发车日期)
  - ETA (预计到港日期)
  - 船名航次 (vessel_voyage)

ATD 取值优先级（高 → 低）：
  P1  「开船时间」后的具体时间          —— 业务铁律，最高优先
  P2  「X月X号已发车」中的 X月X号        —— 陆运/拖车口径，年份从上下文或当年推断
  P3  其它开船同义词（起运/起飞/发运/实际开船/ATD/Departure）
  P4  宽松窗口兜底（OCR 噪声场景）

ETA 取值优先级：
  P1  「预计到港时间」后的具体时间
  P2  其它到港同义词（预计到达/预计到站/ETA/Arrival）
  P3  宽松窗口兜底

输出日期一律规范为 YYYY-MM-DD；若原文带时分，额外给出 atd_time / eta_time。
任何一项识别不到即缺省不写该键，并在 warnings 中说明，绝不猜测填充。
"""
from __future__ import annotations

import datetime
import logging
import os
import re
from typing import Any

log = logging.getLogger(__name__)

# ── 日期片段 ──
# 2026-8-7 / 2026/08/07 / 2026.8.7 / 2026年8月7日
_D_FULL = r"(\d{4})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})\s*[日号]?"
# 8-7 / 8/7 / 8月7日 / 8月7号（无年份）
_D_MD = r"(\d{1,2})\s*[-/.月]\s*(\d{1,2})\s*[日号]?"
# 时分（可选秒）
_T_HM = r"(\d{1,2}\s*[:：]\s*\d{2}(?:\s*[:：]\s*\d{2})?)"

_RE_DATE_FULL = re.compile(_D_FULL)
_RE_DATE_SHORT = re.compile(r"(?:^|\D)(\d{1,2})[-/.](\d{1,2})(?=\D|$)")
_RE_YEAR = re.compile(r"(20\d{2})")

# ── 需求核心规则 ──
# R1：开船时间 / 预计到港时间 —— 关键字后紧跟的具体时间（允许中间有冒号、空格、换行）
_RE_ATD_ANCHOR = re.compile(
    r"开\s*船\s*时\s*间[^\d]{0,8}" + _D_FULL + r"(?:[^\d]{0,4}" + _T_HM + r")?")
_RE_ATD_ANCHOR_MD = re.compile(
    r"开\s*船\s*时\s*间[^\d]{0,8}" + _D_MD + r"(?:[^\d]{0,4}" + _T_HM + r")?")
_RE_ETA_ANCHOR = re.compile(
    r"预\s*计\s*到\s*港\s*时\s*间[^\d]{0,8}" + _D_FULL + r"(?:[^\d]{0,4}" + _T_HM + r")?")
_RE_ETA_ANCHOR_MD = re.compile(
    r"预\s*计\s*到\s*港\s*时\s*间[^\d]{0,8}" + _D_MD + r"(?:[^\d]{0,4}" + _T_HM + r")?")

# R2：「X月X号已发车」/「X月X日已发车」/「X月X号发车」/「已于X月X号发车」
_RE_DEPART_MD = re.compile(
    r"(\d{1,2})\s*月\s*(\d{1,2})\s*[日号][^\n。；;]{0,10}?(?:已\s*)?(?:发车|发运|起运|开船|出发)")

# 同义关键字（次级优先）
_ATD_KWS = [
    r"实际开船", r"开船", r"起运时间", r"起运", r"起飞时间", r"起飞",
    r"发车时间", r"发车", r"发运时间", r"发运", r"离港", r"出发",
    r"ATD", r"Actual\s*Departure", r"Departure",
]
_ETA_KWS = [
    r"预计到港", r"预计到达", r"预计到站", r"预计抵达", r"到港时间",
    r"到站时间", r"抵达时间", r"ETA", r"Estimated\s*Arrival", r"Arrival",
]

# ── 船名航次 ──
_RE_VESSEL = re.compile(
    r"船名(?:航次)?[：:\s]\s*([A-Z][A-Z0-9\s/|\.]+?\s*\d+\s*[A-Z]?)", re.I)
_RE_VESSEL_ALT = re.compile(
    r"(?:船名[：:]\s*)?([A-Z]{2,}[\s\.]?(?:I+|V+)?(?:\s*[/|]\s*\d{1,4}[E]?|\s+\d{1,4}[E]?))",
    re.I)
_RE_VESSEL_LOOSE = re.compile(r"\b([A-Z]{2,}(?:\s+[A-Z]+)*\s*[/|]\s*\d{1,4}[E]?)\b")
_RE_VESSEL_FALLBACK = re.compile(r"([A-Z]{2,}(?:\s+[A-Z]*)*)\s*[/|]\s*(\d{1,4}[E]?)")
_RE_VESSEL_ULTRA = re.compile(r"\b([A-Z]{2,})\s+[IV1]\s+(\d{1,4}[E]?)\b")

_VESSEL_STOP = re.compile(r"^(ETA|ATD|ETD|ETC|ARR|DEP|ETB|TBN|KGS|CBM)\b", re.I)
_VESSEL_PURE_NUM = re.compile(r"^[\d\s./|-]+$")


# ────────────────────────────── 工具 ──────────────────────────────
def _clean_int(s: str) -> int:
    return int(re.sub(r"[^\d]", "", s) or 0)


def _normalize_date(y: str | int, m: str | int, d: str | int) -> str | None:
    """标准化为 YYYY-MM-DD；非法日期返回 None。"""
    try:
        yi = int(y) if not isinstance(y, str) else _clean_int(y)
        mi = int(m) if not isinstance(m, str) else _clean_int(m)
        di = int(d) if not isinstance(d, str) else _clean_int(d)
        return datetime.date(yi, mi, di).isoformat()
    except (ValueError, TypeError):
        return None


def _norm_time(t: str | None) -> str | None:
    if not t:
        return None
    t = re.sub(r"\s", "", t).replace("：", ":")
    parts = t.split(":")
    try:
        h = int(parts[0]); mi = int(parts[1])
        if not (0 <= h <= 23 and 0 <= mi <= 59):
            return None
        return f"{h:02d}:{mi:02d}"
    except (ValueError, IndexError):
        return None


def _guess_year(text: str, month: int, day: int) -> int:
    """无年份场景推断年份：优先取文中出现的 20xx；否则用当前年。

    若推断出的日期比今天晚 6 个月以上，视为上一年（跨年轨迹回溯）。
    """
    m = _RE_YEAR.search(text)
    if m:
        return int(m.group(1))
    today = datetime.date.today()
    y = today.year
    try:
        cand = datetime.date(y, month, day)
    except ValueError:
        return y
    if (cand - today).days > 180:
        return y - 1
    return y


# ─────────────────────────── ATD / ETA ───────────────────────────
def _anchor_full(pat: re.Pattern, text: str) -> tuple[str, str | None] | None:
    m = pat.search(text)
    if not m:
        return None
    g = m.groups()
    iso = _normalize_date(g[0], g[1], g[2])
    if not iso:
        return None
    return iso, _norm_time(g[3] if len(g) > 3 else None)


def _anchor_md(pat: re.Pattern, text: str) -> tuple[str, str | None] | None:
    m = pat.search(text)
    if not m:
        return None
    g = m.groups()
    mo, dd = _clean_int(g[0]), _clean_int(g[1])
    iso = _normalize_date(_guess_year(text, mo, dd), mo, dd)
    if not iso:
        return None
    return iso, _norm_time(g[2] if len(g) > 2 else None)


def _by_keywords(text: str, kws: list[str]) -> tuple[str, str | None] | None:
    """关键字后 0~20 个非数字字符内的日期（先完整年月日，后月日）。"""
    for kw in kws:
        m = re.search(kw + r"[^\d]{0,20}" + _D_FULL + r"(?:[^\d]{0,4}" + _T_HM + r")?",
                      text, re.I)
        if m:
            g = m.groups()
            iso = _normalize_date(g[0], g[1], g[2])
            if iso:
                return iso, _norm_time(g[3] if len(g) > 3 else None)
    for kw in kws:
        m = re.search(kw + r"[^\d]{0,20}" + _D_MD + r"(?:[^\d]{0,4}" + _T_HM + r")?",
                      text, re.I)
        if m:
            g = m.groups()
            mo, dd = _clean_int(g[0]), _clean_int(g[1])
            iso = _normalize_date(_guess_year(text, mo, dd), mo, dd)
            if iso:
                return iso, _norm_time(g[2] if len(g) > 2 else None)
    return None


def _loose_window(text: str, kws: list[str]) -> str | None:
    """宽松兜底：日期前后 50 字符窗口内出现关键字即采纳（OCR 噪声场景）。"""
    for m in _RE_DATE_FULL.finditer(text):
        iso = _normalize_date(*m.groups())
        if not iso:
            continue
        w = text[max(0, m.start() - 50): m.end() + 50]
        if any(re.search(k, w, re.I) for k in kws):
            return iso
    for m in _RE_DATE_SHORT.finditer(text):
        mo, dd = _clean_int(m.group(1)), _clean_int(m.group(2))
        w = text[max(0, m.start() - 40): m.end() + 40]
        if any(re.search(k, w, re.I) for k in kws):
            iso = _normalize_date(_guess_year(text, mo, dd), mo, dd)
            if iso:
                return iso
    return None


def extract_atd(text: str) -> dict[str, Any]:
    """按业务优先级抽取 ATD。返回 {value?, time?, rule}。"""
    # P1 「开船时间」锚定
    hit = _anchor_full(_RE_ATD_ANCHOR, text) or _anchor_md(_RE_ATD_ANCHOR_MD, text)
    if hit:
        return {"value": hit[0], "time": hit[1], "rule": "开船时间"}
    # P2 「X月X号已发车」
    m = _RE_DEPART_MD.search(text)
    if m:
        mo, dd = _clean_int(m.group(1)), _clean_int(m.group(2))
        iso = _normalize_date(_guess_year(text, mo, dd), mo, dd)
        if iso:
            return {"value": iso, "time": None, "rule": "X月X号已发车"}
    # P3 同义关键字
    hit = _by_keywords(text, _ATD_KWS)
    if hit:
        return {"value": hit[0], "time": hit[1], "rule": "开船同义词"}
    # P4 宽松兜底
    iso = _loose_window(text, _ATD_KWS)
    if iso:
        return {"value": iso, "time": None, "rule": "宽松窗口"}
    return {"rule": None}


def extract_eta(text: str) -> dict[str, Any]:
    """按业务优先级抽取 ETA。返回 {value?, time?, rule}。"""
    hit = _anchor_full(_RE_ETA_ANCHOR, text) or _anchor_md(_RE_ETA_ANCHOR_MD, text)
    if hit:
        return {"value": hit[0], "time": hit[1], "rule": "预计到港时间"}
    hit = _by_keywords(text, _ETA_KWS)
    if hit:
        return {"value": hit[0], "time": hit[1], "rule": "到港同义词"}
    iso = _loose_window(text, _ETA_KWS)
    if iso:
        return {"value": iso, "time": None, "rule": "宽松窗口"}
    return {"rule": None}


# ─────────────────────────── 船名航次 ───────────────────────────
def _valid_vessel(v: str) -> bool:
    if not v or len(v) < 4:
        return False
    if _VESSEL_STOP.match(v) or _VESSEL_PURE_NUM.match(v):
        return False
    return True


def _find_vessel(text: str) -> str | None:
    for pat in (_RE_VESSEL, _RE_VESSEL_ULTRA, _RE_VESSEL_FALLBACK,
                _RE_VESSEL_ALT, _RE_VESSEL_LOOSE):
        for m in pat.finditer(text):
            g = m.groups()
            v = f"{g[0].strip()} {g[1].strip()}" if len(g) >= 2 and g[1] else g[0].strip()
            v = re.sub(r"\s+", " ", v).rstrip("|").strip()
            if _valid_vessel(v):
                return v
    return None


# ─────────────────────────── 主入口 ───────────────────────────
def parse_tracking(image_path_or_text: str, *, is_text: bool = False) -> dict[str, Any]:
    """解析轨迹信息（图片 OCR 或纯文本）。

    Returns
    -------
    dict
        ``{ atd?, atd_time?, atd_rule?, eta?, eta_time?, eta_rule?,
            vessel_voyage?, _source, raw_text?, warnings[] }``
    """
    result: dict[str, Any] = {"_source": "text" if is_text else "ocr", "warnings": []}

    if is_text:
        text = (image_path_or_text or "").strip()
        if not text:
            result["warnings"].append("轨迹文本为空")
            return result
    else:
        text = _ocr_image(image_path_or_text, result)
        if not text:
            return result

    result["raw_text"] = text[:2000]

    atd = extract_atd(text)
    if atd.get("value"):
        result["atd"] = atd["value"]
        result["atd_rule"] = atd["rule"]
        if atd.get("time"):
            result["atd_time"] = atd["time"]
    else:
        result["warnings"].append(
            "未识别到 ATD（未找到「开船时间」或「X月X号已发车」），请手动填写")

    eta = extract_eta(text)
    if eta.get("value"):
        result["eta"] = eta["value"]
        result["eta_rule"] = eta["rule"]
        if eta.get("time"):
            result["eta_time"] = eta["time"]
    else:
        result["warnings"].append("未识别到 ETA（未找到「预计到港时间」），请手动填写")

    vessel = _find_vessel(text)
    if vessel:
        result["vessel_voyage"] = vessel
    else:
        result["warnings"].append("未识别到船名航次，请手动填写")

    # 逻辑校验：ETA 早于 ATD 提示（不阻断）
    if result.get("atd") and result.get("eta") and result["eta"] < result["atd"]:
        result["warnings"].append(
            f"ETA({result['eta']}) 早于 ATD({result['atd']})，请核对轨迹")

    log.info("轨迹解析完成: atd=%s(%s) eta=%s(%s) vessel=%s",
             result.get("atd"), result.get("atd_rule"),
             result.get("eta"), result.get("eta_rule"), result.get("vessel_voyage"))
    return result


def _ocr_image(path: str, result: dict) -> str:
    """对图片做 OCR，返回文本。失败时写 warnings 并返回空串。"""
    if not os.path.isfile(path):
        result["warnings"].append(f"轨迹图片不存在: {path}")
        return ""
    try:
        import pytesseract
        from PIL import Image, ImageEnhance
    except Exception as e:
        result["warnings"].append(f"OCR 依赖缺失: {e}")
        return ""

    for p in (r"C:\Program Files\Tesseract-OCR\tesseract.exe",
              r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"):
        if os.path.isfile(p):
            pytesseract.pytesseract.tesseract_cmd = p
            break

    try:
        img = Image.open(path).convert("L")
        img = ImageEnhance.Contrast(img).enhance(2.0)
        text = pytesseract.image_to_string(img, lang="chi_sim+eng", config="--psm 6")
        text = (text or "").strip()
        if not text:
            result["warnings"].append("轨迹图片 OCR 未识别到文字，请改用文本粘贴")
        return text
    except Exception as e:
        msg = f"轨迹图片 OCR 失败: {e}"
        result["warnings"].append(msg)
        log.warning(msg)
        return ""
