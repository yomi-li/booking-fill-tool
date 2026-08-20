# -*- coding: utf-8 -*-
"""模板写入基类：复制模板 -> COM 打开 -> 写值 -> 保存。

设计：
  - 先 shutil.copy 模板到输出路径，再打开副本写入，避免 SaveAs 覆盖冲突，
    且 LOGO / 签章图片随副本完整保留（COM 写值不动图片）。
  - 外部公式链接单元格（如 ='[1]BILL  DRAFT'!A2）在 COM 打开时以
    UpdateLinks=0 禁止更新（母表缺失不报错），写入前先 ClearContents
    清掉公式再写字面值，实现「自我包含」输出，无需母表文件。
  - **铁律：绝不使用 openpyxl 重存整个模板文件**（曾用于预清除外部公式，
    但会改写全部 XML，存在破坏原始排版/单元格结构/框线的风险）。
    所有写入一律走 COM，模板逐字节复制到输出路径后原样保留。
  - 作为上下文管理器使用，确保异常也能退出 Excel。
"""
from __future__ import annotations

import logging
import os
import shutil
from typing import Any, Mapping

from .com_session import ExcelBusyError, launch_excel, open_workbook, quit_excel, excel_healthy

log = logging.getLogger(__name__)

# ---- Excel 瞬态忙错误重试 ----
# 0x80010001 = RPC_E_CALL_REJECTED（被呼叫方拒收）；0x800AC472 = Excel 正忙
# （打印/页面布局刷新期间）。二者均为瞬态，sleep 后重试即可恢复。
# ★2026-08-18：若多次重试仍持续被拒，说明实例已进入【永久忙态】（本机实证：
# 随机发生、60s 不恢复），继续重试无意义 → 抛 ExcelBusyError 交由 com_retry
# 杀进程重启后重试整单生成。
import time
try:
    import pywintypes
except ImportError:
    pywintypes = None

_REJECT_HRESULTS = (-2147418111, -2146777998)

def _is_reject(e) -> bool:
    return getattr(e, "hresult", None) in _REJECT_HRESULTS

def _safe_set(obj, attr, val, optional: bool = False, retries: int = 4):
    """带重试地设置 COM 属性，吸收 Excel 瞬态忙(0x800ac472/0x80010001)。

    PageSetup 等属性在 Excel 刷新页面布局时会偶发“正忙”并被拒绝，
    单次赋值失败后 sleep 重试通常即可成功。optional=True 时最终仍失败
    仅告警跳过（用于“不限制纵向页数”等可降级项）。
    """
    last = None
    for i in range(retries):
        try:
            setattr(obj, attr, val)
            return True
        except Exception as e:
            last = e
            hr = getattr(e, "hresult", None)
            if hr in _REJECT_HRESULTS and i < retries - 1:
                time.sleep(1.0)
                continue
            if optional:
                log.warning("设置 %s=%s 失败(可忽略): %s", attr, val, e)
                return False
            raise
    if optional:
        log.warning("设置 %s=%s 重试 %d 次仍失败(可忽略): %s", attr, val, retries, last)
        return False
    raise last


class XlsxWriter:
    def __init__(self, template_path: str, output_path: str,
                 sheet: str | None = None, visible: bool = False,
                 excel: object | None = None):
        self.template_path = template_path
        self.output_path = os.path.abspath(output_path)
        self.visible = visible
        self._own_excel = excel is None   # Req3：复用外部 Excel 实例时不负责退出
        self._excel = excel
        self._wb = None
        self._ws = None
        self._open()
        if sheet:
            self.select_sheet(sheet)

    def _open(self):
        os.makedirs(os.path.dirname(self.output_path), exist_ok=True)
        shutil.copy(self.template_path, self.output_path)
        # 注意：不再做任何 openpyxl 预处理。外部公式格交给
        # set_cell 的 ClearContents + 写值处理（UpdateLinks=0 打开不更新链接）。
        if self._excel is None:
            self._excel = launch_excel(self.visible)
        # ★2026-08-19：开工前最后做一次健康探测。实例可能在创建后 ~1-1.5s
        # 进入永久忙态；launch_excel 虽已探过，但到此处又经过若干毫秒，
        # 若已忙态则立即抛 ExcelBusyError 交 com_retry 重启，避免 open_workbook 才失败。
        if not excel_healthy():
            raise ExcelBusyError("开工前 Excel 实例已忙态（创建后进入永久忙窗）")
        self._wb = open_workbook(self._excel, self.output_path)
        try:
            self._ws = self._wb.ActiveSheet
        except ExcelBusyError:
            raise
        except Exception as e:
            if _is_reject(e):
                raise ExcelBusyError(f"读取 ActiveSheet 被拒（实例忙态）: {e}") from e
            raise

    # ---- 上下文管理 ----
    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        self.close(save=exc_type is None)
        return False

    # ---- 工作表选择 ----
    def select_sheet(self, name: str):
        try:
            self._ws = self._wb.Sheets(name)
        except ExcelBusyError:
            raise
        except Exception as e:
            if _is_reject(e):
                raise ExcelBusyError(f"切换工作表 {name} 被拒（实例忙态）: {e}") from e
            raise
        return self._ws

    @property
    def sheet(self):
        return self._ws

    # ---- 写值 ----
    @staticmethod
    def _write_value(rng, value: Any, retries: int = 6):
        """写 Range.Value，带 Excel 瞬态忙(0x800AC472/0x80010001/RPC拒绝)重试。

        批量写值 + 插入行后 Excel 常处于页面重算/刷新中，单次赋值易被拒。
        失败 sleep 后重试；★仍持续被拒（6 次 ≈ 5s）说明实例已永久忙态，
        立即抛 ExcelBusyError（BaseException）穿透所有 except-Exception，
        交由生成器 com_retry 杀进程重启重试，不再浪费 20s+ 干等。
        """
        last = None
        for i in range(retries):
            try:
                rng.Value = value
                return
            except Exception as e:
                last = e
                if _is_reject(e) and i < retries - 1:
                    time.sleep(0.8)
                    continue
                if _is_reject(e):
                    raise ExcelBusyError(
                        f"写值被拒 {retries} 次（Excel 永久忙态）: {e}") from e
                raise
        raise last

    def set_cell(self, coord: str, value: Any, sheet: str | None = None):
        ws = self._wb.Sheets(sheet) if sheet else self._ws
        if value is None:
            value = ""
        # COM 写入字符串时去除空字符与首尾空白
        if isinstance(value, str):
            value = value.replace("\x00", "").strip()
        try:
            rng = ws.Range(coord)
            # 外部公式格在母表缺失时可能处于错误/锁定状态，先清内容再写。
            # 合并单元格（如标题区）ClearContents 会报"无法对合并单元格执行此操作"，
            # 此时跳过清理直接写值即可（Value 赋值本身会覆盖公式）。
            try:
                rng.ClearContents()
            except Exception:
                pass
            self._write_value(rng, value)
        except ExcelBusyError:
            raise
        except Exception as e:
            if _is_reject(e):
                # ★2026-08-18：ws.Range() 访问器本身在被拒时也会抛 com_error，
                # 必须在此转成 ExcelBusyError 才能穿透 except-Exception、触发 com_retry 自愈。
                raise ExcelBusyError(f"写单元格 {coord} 被拒（实例忙态）: {e}") from e
            raise RuntimeError(f"写单元格 {coord} 失败: {e}") from e

    def fit_cell(self, coord: str, value: Any, sheet: str | None = None,
                 min_size: float = 8.0, step: float = 0.5,
                 check_overflow: bool = True):
        """写值 + 自动缩字/缩字防溢出：

        - 数值：出现 "###" 时逐步缩小字体（原有逻辑）。
        - 文本：当 check_overflow=True 时，写入后比较单元格显示文本 rng.Text
          与原始值（去除换行/空格标准化）。若不一致说明被截断/溢出，
          逐步缩小字体直至完整显示或到达 min_size。用于 VESSEL、BILL NUMBER、
          货量、目的港、柜号等长文本列。
        - 合并单元格作用于合并区域，字号统一调整。
        """
        ws = self._wb.Sheets(sheet) if sheet else self._ws
        orig = value
        if value is None:
            value = ""
        if isinstance(value, str):
            value = value.replace("\x00", "").strip()
        try:
            rng = ws.Range(coord)
            try:
                rng.ClearContents()
            except Exception:
                pass
            self._write_value(rng, value)
        except ExcelBusyError:
            raise
        except Exception as e:
            if _is_reject(e):
                raise ExcelBusyError(f"写单元格 {coord} 被拒（实例忙态）: {e}") from e
            raise RuntimeError(f"写单元格 {coord} 失败: {e}") from e

        # 数值：防 ### 溢出
        if isinstance(value, (int, float)):
            try:
                size = float(rng.Font.Size) or 11.0
                while size > min_size and "#" in str(rng.Text):
                    size -= step
                    rng.Font.Size = size
            except Exception:
                pass
            return

        # 文本：溢出检测与缩字
        if not check_overflow or not isinstance(orig, str) or not orig.strip():
            return
        try:
            def _norm(s):
                return str(s).replace("\r", "").replace("\n", "").replace(" ", "").replace("\x00", "").strip()
            target = _norm(orig)
            size = float(rng.Font.Size) or 11.0
            while size > min_size:
                displayed = _norm(rng.Text)
                if displayed == target:
                    break
                size -= step
                rng.Font.Size = size
        except Exception:
            pass

    def fit_text(self, coord: str, value: Any, min_size: float = 8.0,
                 step: float = 0.5, *, sheet: str | None = None,
                 min_row_height: float = 0.0):
        """文本适配：缩字防溢出 + 显式撑行高防遮盖（WrapText 模式下 Excel 不自适应）。

        - 先做现有 fit_cell 的缩字逻辑（单行/合并区适用）。
        - 再探测单元格是否 WrapText：若是，写入后根据实际行数
          显式 set_row_height，避免多行文本被下一行相邻单元格遮盖。
        - min_row_height：即使单行也保底的行高（pt），默认等于字体大小。
        """
        if value is None:
            value = ""
        if isinstance(value, str):
            value = value.replace("\x00", "").strip()
        # 先走原有缩字逻辑
        self.fit_cell(coord, value, check_overflow=True)
        # 再处理 WrapText 行高
        try:
            ws = self._wb.Sheets(sheet) if sheet else self._ws
            rng = ws.Range(coord)
            if not (rng.WrapText or False):
                return
            if not isinstance(value, str) or not value.strip():
                return
            # 估算需要多少行（按字符数/列宽经验比 14pt≈18字符）
            import math
            col_w = float(rng.Columns(1).ColumnWidth or 10.0)
            text_len = len(str(value).replace("\r", "").replace("\n", ""))
            chars_per_line = max(1, int(col_w * 1.6))  # 近似 14pt Arial
            lines = max(1, math.ceil(text_len / chars_per_line) + 1)  # +1 缓冲
            # 单行保底高度 ~14pt，每多一行 +~10pt（含行间距）
            row_h = max(min_row_height or 14.0, 14.0 + (lines - 1) * 10.0)
            self.set_row_height(rng.Row, row_h)
        except Exception as e:
            log.warning("fit_text %s 行高调整失败(继续): %s", coord, e)

    def set_row_height(self, row_index: int, height: float):
        """通过 COM 显式设置行高（pt）。WrapText 开启后 Excel 不自适应行高，
        必须在此显式设定以容纳多行文本，避免被相邻单元格遮盖。
        """
        try:
            self._ws.Rows(row_index).RowHeight = height
        except Exception as e:
            if _is_reject(e):
                raise ExcelBusyError(f"设置行高 {row_index}={height} 被拒（实例忙态）: {e}") from e
            log.warning("设置行高 %d=%s 失败(忽略): %s", row_index, height, e)

    def set_cells(self, mapping: Mapping[str, Any]):
        # 以下划线开头的键是元数据（如 _total_currency），不写入单元格
        for coord, val in mapping.items():
            if str(coord).startswith("_"):
                continue
            self.set_cell(coord, val)

    def set_number_format(self, coord: str, fmt: str, sheet: str | None = None):
        """通过 COM 设置单元格数字格式（如 '"$"#,##0.00' / '"€"#,##0.00'）。

        瞬态忙仅告警；★持续被拒抛 ExcelBusyError 触发自愈重试。
        """
        try:
            ws = self._wb.Sheets(sheet) if sheet else self._ws
            ws.Range(coord).NumberFormat = fmt
        except ExcelBusyError:
            raise
        except Exception as e:
            if _is_reject(e):
                raise ExcelBusyError(f"设置数字格式被拒（实例忙态）: {e}") from e
            log.warning("设置数字格式 %s=%s 失败(忽略): %s", coord, fmt, e)

    def insert_rows(self, row_index: int, count: int = 1, sheet: str | None = None) -> bool:
        """在指定行索引处插入 count 行（下方内容下移，模板格式随行带下）。

        需求（2026-08-18）：账单行数不设上限——费用明细超出模板固定区时，
        在其下方插入新行承接，避免覆盖 SUBTOTAL/账户区；Excel 中仍可手动编辑。
        xlShiftDown=-4121。失败返回 False（调用方回退到截断逻辑）。
        """
        if count <= 0:
            return True
        try:
            ws = self._wb.Sheets(sheet) if sheet else self._ws
            rng = ws.Range(
                ws.Cells(row_index, 1), ws.Cells(row_index + count - 1, 1))
            last = None
            for i in range(5):
                try:
                    rng.Insert(-4121)  # xlShiftDown
                    return True
                except Exception as e:
                    last = e
                    if _is_reject(e) and i < 4:
                        time.sleep(0.8)
                        continue
                    if _is_reject(e):
                        raise ExcelBusyError(
                            f"插入行被拒（实例忙态）: {e}") from e
                    raise
        except ExcelBusyError:
            raise
        except Exception as e:
            log.warning("插入行 R%d x%d 失败: %s", row_index, count, e)
            return False

    # ---- 列宽调整 ----
    def set_col_widths(self, widths: Mapping[str, float]):
        """通过 COM 设置列宽（不改动模板文件本身，仅影响输出副本）。

        widths: {"A": 6.5, "B": 7.0, ...} — 列字母 → 宽度（Excel 字符单位）。
        用 Range("A:A") 方式设置，兼容 early/late binding。
        """
        for col_letter, width in widths.items():
            try:
                self._ws.Range(f"{col_letter}:{col_letter}").ColumnWidth = width
            except Exception as e:
                if _is_reject(e):
                    raise ExcelBusyError(f"设置列宽被拒（实例忙态）: {e}") from e
                log.warning("设置列宽 %s=%s 失败: %s", col_letter, width, e)

    # ---- 收尾 ----
    def save(self):
        try:
            # ★2026-08-18：不再尝试切换 Calculation（本机 Excel 拒绝该属性，
            # 且失败会让实例进入永久忙态）。默认自动重算下 Save 时公式自然计算。
            self._wb.Save()
        except Exception as e:
            if _is_reject(e):
                raise ExcelBusyError(f"保存被拒（实例忙态）: {e}") from e
            log.warning("保存失败(忽略): %s", e)

    def export_pdf(self, pdf_path: str, sheet: str | None = None,
                   print_area: str | None = None,
                   orientation: int = 1,
                   single_page: bool = False) -> str:
        """导出当前（或指定）工作表为 PDF。隐藏 sheet 保持隐藏，不导出。

        xlTypePDF=0；铁律：PDF 导出仅按需展开当前 sheet。

        竖向排版约束（用户要求竖向 + 仅缩格子，不改格子位置/内容）：
        - orientation=1 Portrait（纵向 A4）；工厂/INVOICE 默认竖向。
        - FitToPagesWide=1：横向锁一页，所有列在一页宽度内完整显示（右侧不截断）。
        - FitToPagesTall=False：纵向允许自然分页（内容超长时多页）；
          single_page=True 时纵向也锁一页（整页缩放，用户要求提单单页）。
        - 列宽由调用方 set_col_widths() 预先缩窄至 A4 纵向可打印范围内。

        print_area: 指定打印区域（如 "A1:K51"），裁剪模板右侧 / 下方冗余空白。
        """
        ws = self._wb.Sheets(sheet) if sheet else self._ws
        pdf_path = os.path.abspath(pdf_path)
        os.makedirs(os.path.dirname(pdf_path), exist_ok=True)

        ps = ws.PageSetup
        if print_area:
            _safe_set(ps, "PrintArea", print_area, optional=True)

        # 先设置纸张方向（必须在 Zoom/FitToPages 之前，否则会被覆盖）
        _safe_set(ps, "PaperSize", 9, optional=True)             # xlPaperA4
        _safe_set(ps, "Orientation", orientation, optional=True)  # 1=Portrait, 2=Landscape

        # Orientation 设置后 Save 一次（部分 Excel 版本需落盘才生效）；
        # 失败属瞬态，导出用内存状态，忽略即可。
        try:
            self._wb.Save()
        except Exception as e:
            log.warning("保存失败(忽略): %s", e)

        # 用 _safe_set 包裹每个 PageSetup 赋值，吸收 Excel 瞬态忙(0x800ac472)。
        _safe_set(ps, "Zoom", False)                  # 关掉 Zoom 才能启用 FitToPages
        _safe_set(ps, "FitToPagesWide", 1)            # 横向锁一页，杜绝横向跨页/截断
        # 纵向：single_page=True → 锁一页（FitToPagesTall=1 整页缩放）；否则不强制
        _safe_set(ps, "FitToPagesTall", 1 if single_page else False, optional=True)
        _safe_set(ps, "Order", 1, optional=True)     # xlDownThenOver
        _safe_set(ps, "CenterHorizontally", True, optional=True)
        _safe_set(ps, "LeftMargin", 18, optional=True)    # 0.25 inch ≈ 0.64 cm
        _safe_set(ps, "RightMargin", 18, optional=True)
        _safe_set(ps, "TopMargin", 18, optional=True)
        _safe_set(ps, "BottomMargin", 18, optional=True)

        # ★2026-08-17：late binding(Dispatch) 下导出偶发 RPC_E_CALL_REJECTED
        # (0x80010001, 被呼叫方拒绝接收呼叫)，属瞬态忙错误，sleep 后重试（最多 5 次）。
        import time
        try:
            import pywintypes
        except ImportError:
            pywintypes = None
        for _attempt in range(5):
            try:
                ws.ExportAsFixedFormat(0, pdf_path)
                break
            except Exception as e:
                hresult = getattr(e, "hresult", None)
                # 覆盖两种瞬态忙：0x80010001(RPC_E_CALL_REJECTED) 与
                # 0x800AC472(Excel 正忙/打印锁定)，均 sleep 后重试。
                is_rejected = (pywintypes is not None and isinstance(e, pywintypes.com_error)
                               and hresult in (-2147418111, -2146777998))
                if is_rejected and _attempt < 4:
                    time.sleep(1.5)
                    continue
                if is_rejected:
                    # ★2026-08-18：5 次仍被拒 → 实例永久忙态，抛 ExcelBusyError
                    # 交由 com_retry 杀进程重启重试（不再走 HTML 兜底，保证 xlsx 落盘）。
                    raise ExcelBusyError(
                        f"PDF 导出被拒 5 次（实例忙态）: {e}") from e
                raise
        log.info("PDF 已导出: %s (orientation=%s)", pdf_path, orientation)
        return pdf_path
    def close(self, save: bool = True):
        try:
            if save:
                self.save()   # 忙态被拒时抛 ExcelBusyError，穿透到 com_retry
        finally:
            try:
                if self._wb is not None:
                    self._wb.Close(SaveChanges=False)
            except Exception:
                pass
            if self._own_excel:
                quit_excel(self._excel)
            self._wb = None
            self._excel = None
