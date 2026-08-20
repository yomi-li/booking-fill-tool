# -*- coding: utf-8 -*-
"""一键备份：把数据层（数据库 / 模板 / 配置 / 话术模板）打包为带时间戳的 zip。

仅备份"会变化且不可重建"的资产；jobs 产出目录不备份（可随时重新生成）。
备份落在 data/backup/，按时间命名，便于跨电脑迁移或回滚。
"""
from __future__ import annotations

import shutil
import zipfile
from datetime import datetime
from pathlib import Path

from bol_forecast.config import ROOT, DATA_DIR, TEMPLATES_DIR, BACKUP_DIR, DB_PATH, CONFIG_PATH

# 需要纳入备份的条目：(源路径, 在 zip 内的相对名)
def _manifest() -> list[tuple[Path, str]]:
    items: list[tuple[Path, str]] = []
    # 数据库（含 WAL / SHM 伴侣文件）
    for suf in ("", "-wal", "-shm"):
        p = Path(str(DB_PATH) + suf)
        if p.exists():
            items.append((p, f"data/{p.name}"))
    # 模板（保留目录结构）
    if TEMPLATES_DIR.exists():
        for f in TEMPLATES_DIR.rglob("*"):
            if f.is_file() and not f.name.startswith("."):
                items.append((f, f"templates/{f.name}"))
    # 配置 + 话术模板
    if CONFIG_PATH.exists():
        items.append((CONFIG_PATH, "config.json"))
    sp = DATA_DIR / "speech_templates.json"
    if sp.exists():
        items.append((sp, "data/speech_templates.json"))
    return items


def backup_now(tag: str = "") -> Path:
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    name = f"backup_{ts}{('_'+tag) if tag else ''}.zip"
    out = BACKUP_DIR / name
    items = _manifest()
    if not items:
        raise RuntimeError("无可备份内容（数据库/模板/配置均缺失）")
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for src, arc in items:
            z.write(src, arc)
    return out


if __name__ == "__main__":
    print("backup ->", backup_now())
