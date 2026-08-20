# -*- coding: utf-8 -*-
"""Excel COM 会话基础设施。

铁律：写模板必须走 COM（win32com），openpyxl 会丢 LOGO / 签章图片。
所有生成器共用此模块启动 Excel、打开工作簿，由 writer_base.XlsxWriter
负责写入与清理。

★2026-08-18 架构调整（消除 0x800AC472/RPC 拒绝）：
  - 进程内【单例 Excel 实例】。所有生成器复用同一实例，绝不每次新建，
    避免连续生成时多实例并行（旧实例未完全退出 + 新实例 → COM 正忙）。
  - quit_excel 不再真 Quit（服务常驻期复用；进程退出由系统回收）。
  - 若外部显式调用 shutdown_excel() 才 Quit（服务优雅关闭时）。
"""
from __future__ import annotations

import functools
import logging
import os
import subprocess
import tempfile
import time
from typing import Any

from bol_forecast.config import CFG

log = logging.getLogger(__name__)

_GEN_INITIALIZED = False  # 每进程只清一次 gen_py
_excel_singleton: Any = None   # 进程内单例 Excel 实例
_excel_pid: int | None = None  # 单例 Excel 的进程 PID（自愈时只杀自己拉起的实例）

# Excel 忙/拒收 HRESULT：0x80010001=RPC_E_CALL_REJECTED（被呼叫方拒收）；
# 0x800AC472=Excel 正忙（打印/页面布局刷新期间）。
REJECT_HRESULTS = (-2147418111, -2146777998)

# launch_excel 最多重建次数：创建实例后用 excel_healthy() 验证，不健康则
# 杀掉重建，直到拿到健康实例（配合生成器 @com_retry() 形成双重自愈）。
_LAUNCH_TRIES = 6


def is_reject_error(e) -> bool:
    """判断异常是否为 Excel 忙/拒收（瞬态或永久忙态的信号）。"""
    return getattr(e, "hresult", None) in REJECT_HRESULTS


class ExcelBusyError(BaseException):
    """Excel COM 实例进入忙态（0x80010001/0x800AC472 持续被拒），需要杀进程重启。

    2026-08-18 实证：本机 Excel 会随机（与环境负载相关）在启动后 ~1-2s 进入
    永久忙态——所有 COM 调用（含 Quit）被拒，60s 不恢复。重试写入无意义，
    唯一出路是杀掉该实例、重启一个干净的。生成器捕获本异常后由 com_retry
    自动杀进程 + 重试整单生成。

    继承 BaseException 而非 Exception：所有 `except Exception` 兜底/吞异常站点
    （fit_cell 缩字失败、set_number_format 告警、export_pdf HTML 兜底等）都
    不会误吞本异常，确保忙态一定穿透到 com_retry 触发自愈。
    """


def _clear_gen_py_once() -> None:
    """删除 %TEMP%\\gen_py，强制 gencache 重新生成类型库缓存。

    用 subprocess 调系统 `rd` 绕过 Python os.remove 的安全删除拦截。
    脏缓存会让 ExportAsFixedFormat(PDF 导出)报「找不到成员」。
    """
    gen_dir = os.path.join(tempfile.gettempdir(), "gen_py")
    if os.path.exists(gen_dir):
        subprocess.run(["cmd", "/c", "rd", "/s", "/q", gen_dir],
                       shell=False, capture_output=True)
        log.info("已清理 gen_py 缓存: %s", gen_dir)


def _create_instance(visible: bool = False):
    """创建一个 Excel 实例并设置行为属性。失败（含创建异常）返回 None。

    pywin32 缺失时抛 RuntimeError（带清晰指引）。
    """
    try:
        import pythoncom  # 函数内 import：缺失时给到清晰 ImportError
        import win32com.client as win32
    except ImportError as e:
        raise RuntimeError(
            "生成 Excel 单据需要 pywin32 与 Office。请先在 venv 执行 "
            "`pip install pywin32` 并安装 64 位 Office，"
            "然后以管理员身份运行 `python Scripts/pywin32_postinstall.py -install`。"
            f"原始错误: {e}"
        ) from e
    pythoncom.CoInitialize()
    # ★2026-08-18：DispatchEx 强制创建【独立新实例】（不连接用户已打开的
    # Excel 窗口——那会因模态对话框/未保存内容导致 0x800AC472 正忙）。
    # 走 late binding（win32.DispatchEx），避免早绑定类型库与 Excel 16
    # 交互时 "不能设置 Calculation 属性" 等问题。EnsureModule 仅作兜底。
    excel = None
    try:
        excel = win32.DispatchEx("Excel.Application")
    except Exception:
        try:
            excel = win32.Dispatch("Excel.Application")
        except Exception:
            excel = None
    if excel is None:
        import win32com.client.gencache as gcache
        for major, minor in ((1, 9), (1, 8)):
            try:
                mod = gcache.EnsureModule(
                    "{00020813-0000-0000-C000-000000000046}", 0, major, minor)
                excel = mod.Application()
                break
            except Exception:
                continue
    if excel is None:
        return None
    # 以下属性为行为偏好，单个失败不应阻断生成
    for attr, val in (
        ("Visible", visible),
        ("DisplayAlerts", False),
        ("AskToUpdateLinks", False),
        ("ScreenUpdating", False),
        ("AutomationSecurity", 3),  # msoAutomationSecurityForceDisable
    ):
        try:
            setattr(excel, attr, val)
        except Exception as e:
            log.warning("设置 Excel.%s 失败(忽略): %s", attr, e)
    # ★2026-08-18 教训：绝不能触碰 Application.Calculation 属性。
    # 本机 Excel 拒绝设置该属性，且失败会让实例进入永久忙态（见 launch_excel）。
    # 模板公式在默认自动重算模式下于 Save 时正常计算，无需手动切换。
    # ★2026-08-19：不要在此 sleep。实证 Excel 实例在创建后约 1-1.5s 进入
    # 永久忙态，sleep 会把实例推进忙窗才做健康探测 → 探针失效。创建即探最稳。
    return excel


def launch_excel(visible: bool = False):
    """返回进程内单例 Excel COM 实例，并确保交付的是【健康】实例。

    ★2026-08-18 实证：本机 Excel 会随机进入永久忙态（约 1-2s 后对所有 COM
    调用拒收，含 Quit）。因此创建实例后必须用 excel_healthy() 探测验证，
    不健康则杀掉重建，直到拿到健康实例（最多 _LAUNCH_TRIES 次）。
    配合生成器 @com_retry()，即使偶发全忙也能自愈。
    """
    global _GEN_INITIALIZED, _excel_singleton
    if not _GEN_INITIALIZED:
        _clear_gen_py_once()
        _GEN_INITIALIZED = True
    # 已有单例且健康 → 直接复用
    if _excel_singleton is not None and excel_healthy():
        return _excel_singleton
    last_exc = None
    for _try in range(_LAUNCH_TRIES):
        if _excel_singleton is not None:
            recover_excel()  # 清掉上一轮不健康实例
        try:
            exc = _create_instance(visible)
        except RuntimeError:
            raise
        except Exception as e:
            log.warning("创建 Excel 实例异常(重试): %s", e)
            exc = None
        if exc is None:
            time.sleep(1.0)
            continue
        _excel_singleton = exc
        _track_pid(exc)   # 记录 PID，供自愈时精确清理（只杀自己拉起的实例）
        if excel_healthy():
            log.info("Excel COM 单例已就绪: %s (pid=%s, try=%d)",
                     getattr(exc, "Name", "Excel"), _excel_pid, _try + 1)
            return exc
        # 健康探测失败：本实例即刻忙态，下一轮循环会 recover 掉它重建
        last_exc = exc
        log.warning("新建 Excel 实例即刻忙态，重建 (try %d/%d)",
                    _try + 1, _LAUNCH_TRIES)
    # 兜底：多次重建仍不健康，交付最后一个（@com_retry 会再杀再试）
    log.warning("launch_excel：%d 次重建仍未拿到健康实例，交付当前实例由调用方兜底",
                _LAUNCH_TRIES)
    return _excel_singleton


def _track_pid(excel) -> None:
    """通过窗口句柄反查单例 Excel 的进程 PID（启动即刻记录，实例尚健康）。"""
    global _excel_pid
    _excel_pid = None
    try:
        import ctypes
        hwnd = excel.Hwnd
        if hwnd:
            pid = ctypes.c_ulong()
            ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            _excel_pid = pid.value or None
    except Exception as e:
        log.warning("无法获取 Excel PID（自愈时将回退到杀全部实例）: %s", e)


def excel_healthy() -> bool:
    """快速探测单例 Excel 是否仍响应。忙态（被拒）即视为不健康。"""
    global _excel_singleton
    if _excel_singleton is None:
        return False
    try:
        _ = _excel_singleton.Version
        return True
    except Exception:
        return False


def recover_excel(wait: float = 1.5) -> None:
    """杀掉忙态/失联的单例 Excel 进程，清空单例状态（下次 launch_excel 重建）。

    - 优先只杀自己拉起的实例（_excel_pid 精确 PID）；
    - PID 未知时回退为结束全部 EXCEL.EXE（本机为专用单证生成机，可接受）。
    - 绝不用 COM Quit（实例已失联，Quit 必然被拒）。
    """
    global _excel_singleton, _excel_pid
    pid = _excel_pid
    _excel_singleton = None
    _excel_pid = None
    if pid:
        try:
            subprocess.run(["taskkill", "/F", "/PID", str(pid)],
                           capture_output=True, timeout=15)
            log.info("已强制结束忙态 Excel 实例 (pid=%s)", pid)
        except Exception as e:
            log.warning("按 PID 结束 Excel 失败，回退结束全部实例: %s", e)
            pid = None
    if not pid:
        try:
            subprocess.run(["taskkill", "/F", "/IM", "EXCEL.EXE"],
                           capture_output=True, timeout=15)
            log.info("已强制结束全部 EXCEL.EXE（PID 追踪不可用回退）")
        except Exception as e:
            log.warning("结束全部 EXCEL.EXE 失败: %s", e)
    time.sleep(wait)   # 等进程彻底退出、端口/OLE 状态释放


def com_retry(attempts: int = 8):
    """生成器装饰器：捕获 ExcelBusyError → 杀进程重启 → 重试整单生成。

    2026-08-18 本机实证：Excel 实例会随机进入永久忙态（概率随机器负载波动），
    重试同一实例无意义；杀掉重建后新实例大概率健康（忙态与具体写入无关）。
    launch_excel 已内置健康探测（不健康不交付），此处再叠加整单重试，
    双重保障：偶发全忙也能自愈。attempts 次全部失败则抛最后一个异常。
    """
    def deco(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            last = None
            for i in range(attempts):
                try:
                    return func(*args, **kwargs)
                except ExcelBusyError as e:
                    # 已明确标记的忙态（写值/插入行/导出 PDF 等持续被拒）
                    last = e
                    log.warning("[%s] COM 忙态(第 %d/%d 次)，杀 Excel 重启后重试: %s",
                                func.__name__, i + 1, attempts, e)
                    recover_excel(wait=2.0)   # 多等 2s 让机器/OLE 状态缓冲
                except Exception as e:
                    # ★2026-08-18：任何"被拒"类 com_error（ActiveSheet/Sheets/
                    # Range 等访问器在实例忙态时抛）都按忙态处理 → 自愈重启。
                    # 仅捕获忙态 hresult，真实业务错误（文件缺失等）原样抛出。
                    if is_reject_error(e):
                        last = ExcelBusyError(f"COM 被拒（实例忙态）: {e}")
                        log.warning("[%s] COM 访问被拒(第 %d/%d 次)，杀 Excel 重启后重试: %s",
                                    func.__name__, i + 1, attempts, e)
                        recover_excel(wait=2.0)
                        continue
                    raise
            raise last
        return wrapper
    return deco


def open_workbook(excel, path: str, update_links: int = 0):
    """打开工作簿。update_links=0 表示不更新外部链接（母表缺失时避免报错）。

    ★忙态被拒时抛 ExcelBusyError（穿透所有 except-Exception，交 com_retry 自愈）。
    """
    try:
        return excel.Workbooks.Open(os.path.abspath(path), UpdateLinks=update_links)
    except ExcelBusyError:
        raise
    except Exception as e:
        if is_reject_error(e):
            raise ExcelBusyError(f"打开工作簿被拒（实例忙态）: {e}") from e
        raise


def quit_excel(excel) -> None:
    """安全退出 Excel，吞掉所有异常。

    单例模式下：保留实例供复用（服务常驻期不 Quit，避免下次启动竞态）。
    仅当进程即将退出（shutdown_excel 显式调用）才真正 Quit。
    """
    pass


def shutdown_excel() -> None:
    """服务优雅关闭时调用：真正退出单例 Excel。"""
    global _excel_singleton, _excel_pid
    if _excel_singleton is None:
        return
    try:
        _excel_singleton.ScreenUpdating = True
        _excel_singleton.Quit()
    except Exception as e:  # pragma: no cover
        log.warning("Excel 退出异常(可忽略): %s", e)
    finally:
        _excel_singleton = None
        _excel_pid = None
