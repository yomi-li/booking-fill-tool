# -*- coding: utf-8 -*-
"""把抽取/编辑后的数据填入固定模板并导出。

环境说明：本机 Excel COM（win32com）在当前 Python 3.13 / pywin32 组合下，
Worksheets/Sheets 的索引访问会返回损坏的代理对象（__call__.Range），
且工作簿 Close 被持久模态框阻塞（0x800ac472），无法稳定使用。
因此默认采用 openpyxl 填充。
关于 LOGO：openpyxl 在 load_workbook 时会自动保留模板内已有的浮动图片
（含其原始锚点），save 后不丢失——故模板自带 LOGO 则天然保留；当前桌面版
模板不含 LOGO（images=0），则不显示，无需额外处理。
产品图片则以「内嵌」方式（TwoCellAnchor 钉在 AD 列对应行）写入，
图片随行移动、与该产品行一一对齐，不再像浮动图那样错位。

每调用一次 fill_template(items, ...) 写入一票（多票请多次调用，分别生成文件）。
"""
import os
import io
import shutil
import tempfile

import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from openpyxl.utils import get_column_letter
from openpyxl.utils.units import pixels_to_EMU

SHEET_NAME = "客户货物运输委托书 "
# 1-based 列号（与模板表头严格对应）
COL = {
    "fba_box_no": 1, "cn_name": 2, "en_name": 3, "sku": 4, "hs_code": 5,
    "material_cn": 6, "material_en": 7, "brand": 8, "brand_type": 9,
    "model": 10, "purpose": 11, "electric_magnetic": 12,
    "total_cartons": 13, "net_per_ctn": 14, "gross_per_ctn": 15,
    "qty_per_ctn": 16, "total_qty": 17, "unit_price": 18,
    "total_price": 19, "currency": 20, "po_price": 21, "po_total": 22,
    "po_currency": 23, "length": 24, "width": 25, "height": 26,
    "reference_id": 27, "warehouse_code": 28,
    "product_link": 31,
}
IMG_COL = 30  # AD 列：图片

# 模板 LOGO 由 openpyxl 在 load_workbook 时自动保留（浮动图片 + 原锚点）。


def _img_dirs():
    """候选图片目录：优先 DATA_DIR（持久卷），其次项目目录。"""
    base = os.path.dirname(os.path.abspath(__file__))
    data = os.environ.get("DATA_DIR") or base
    out = []
    for d in (data, base):
        p = os.path.join(d, "sku_images")
        if p not in out:
            out.append(p)
    return out


def _resolve_image(image_path, sku=""):
    """图片路径跨平台兜底解析（解决换机/云端后绝对路径失效导致图片不显示）：
    1) 原路径存在 → 直接用；
    2) 按文件名在 sku_images/ 目录下找（云端/另一台电脑上绝对路径不存在的场景）；
    3) 按 SKU 码在 sku_images/ 目录下找（扩展名自动探测）。
    全部找不到返回 None（调用方跳过该行图片）。"""
    if not image_path:
        return None
    if os.path.exists(image_path):
        return image_path
    bn = os.path.basename(str(image_path).replace("\\", "/"))
    for img_dir in _img_dirs():
        if bn:
            cand = os.path.join(img_dir, bn)
            if os.path.exists(cand):
                return cand
        if sku:
            for ext in (".png", ".jpg", ".jpeg", ".webp"):
                cand = os.path.join(img_dir, f"{sku}{ext}")
                if os.path.exists(cand):
                    return cand
    return None


def _compress_image(src, max_px=240):
    """把源图等比缩放到 max_px 内并以压缩 PNG 输出到内存。
    返回 (BytesIO, w_px, h_px)；Pillow 不可用或处理失败时返回 (None,None,None)，
    调用方回退直接嵌入原图。"""
    try:
        from PIL import Image as PILImage
    except Exception:
        return None, None, None
    try:
        with PILImage.open(src) as im:
            im = im.convert("RGB")
            w, h = im.size
            scale = min(1.0, max_px / max(w, h))
            if scale < 1.0:
                im = im.resize((max(1, int(w * scale)), max(1, int(h * scale))),
                               PILImage.LANCZOS)
            buf = io.BytesIO()
            im.save(buf, format="PNG", optimize=True)
            w, h = im.size
            buf.seek(0)
            return buf, w, h
    except Exception:
        return None, None, None

def _has_content(it):
    return any(it.get(k) not in (None, "") for k in
               ("sku", "model", "fba_box_no", "cn_name", "en_name"))


def _copy_template_local(template_path: str) -> str:
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp.close()
    try:
        shutil.copy2(template_path, tmp.name)
    except Exception:
        import subprocess
        win_path = template_path.replace("/", "\\")
        subprocess.run(
            ["powershell", "-NoProfile", "-Command",
             f'Copy-Item -LiteralPath "{win_path}" -Destination "{tmp.name}" -Force'],
            check=True, capture_output=True,
        )
    return tmp.name


def _add_product_image(ws, row, image_path, sku="", target=95):
    """在 AD 列(row) 写入「内嵌」产品图片：用 TwoCellAnchor 把图片钉在该行
    单元格内，图片随行移动、与对应产品行一一对齐；并按图片高度撑开本行行高，
    保证图片完整落在这一行、不串到下一行。

    写入前先用 Pillow 把图片等比缩放到 240px 内并压缩（解决超大原图 11-12MB
    被完整嵌入导致 xlsx 巨大、Excel/WPS 解码失败显示空白的问题）；
    路径失效时按文件名/SKU 码兜底解析；Pillow 不可用或处理失败时回退原图。"""
    src = _resolve_image(image_path, sku)
    if not src:
        return
    try:
        buf, w_px, h_px = _compress_image(src)
        if buf is None:
            img = XLImage(src)
            w_px, h_px = img.width, img.height
        else:
            img = XLImage(buf)
            img.width, img.height = w_px, h_px
        if not w_px or not h_px:
            return
        # 等比缩放（以高度为主，保证一行能放下）
        scale = min(target / float(h_px), target / float(w_px), 1.0)
        disp_w = int(w_px * scale)
        disp_h = int(h_px * scale)
        img.width, img.height = disp_w, disp_h
        # 行高按图片高度(像素→点)撑开，+4 留白，保证整图落在该行内
        ws.row_dimensions[row].height = max(
            ws.row_dimensions[row].height or 15, disp_h * 0.75 + 4)
        # TwoCellAnchor：from=AD{row} 左上角，to=AD{row} 左上角+图片宽高(EMU)
        # 图片即“嵌入”在该单元格、随行移动，与产品行一一对应。
        marker_from = AnchorMarker(col=IMG_COL - 1, row=row - 1, colOff=0, rowOff=0)
        marker_to = AnchorMarker(
            col=IMG_COL - 1, row=row - 1,
            colOff=pixels_to_EMU(disp_w),
            rowOff=pixels_to_EMU(disp_h),
        )
        img.anchor = TwoCellAnchor(_from=marker_from, to=marker_to)
        ws.add_image(img)
    except Exception as e:
        print(f"[PRODUCT IMG ERR] {e}", flush=True)


def fill_template(items: list, template_path: str, output_path: str, currency_default="USD") -> str:
    local = _copy_template_local(template_path)
    try:
        wb = openpyxl.load_workbook(local)
        ws = wb[SHEET_NAME]

        # 清空数据区（保留表头、样式、其它 sheet）
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=31):
            for c in row:
                c.value = None

        # 1) 写数据矩阵
        r = 2
        for it in items:
            if not _has_content(it):
                continue
            for field, col in COL.items():
                v = it.get(field)
                if field == "currency":
                    v = (v or currency_default)
                if v in (None, ""):
                    v = None
                ws.cell(row=r, column=col, value=v)
            r += 1

        # 2) 产品图片（AD 列，按数据行，内嵌 + 压缩，跨平台路径兜底）
        r = 2
        for it in items:
            if not _has_content(it):
                continue
            _add_product_image(ws, r, it.get("image_path"), sku=it.get("sku", ""))
            r += 1

        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        wb.save(output_path)
        return output_path
    finally:
        try:
            os.remove(local)
        except Exception:
            pass
