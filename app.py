# -*- coding: utf-8 -*-
"""本地网页应用后端：上传多份 PDF/XLS -> 规则转换(预览/可编辑) -> 导出填好模板。

零 AI 依赖：抽取完全由 rules.json 驱动的可配置规则完成；规则编辑器可自助维护。
"""
import json
import os
import base64
import io
import re
import zipfile
import datetime
import tempfile
import logging
from typing import List

logger = logging.getLogger(__name__)

from fastapi import FastAPI, UploadFile, File, Form, BackgroundTasks, Request
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

import extractor
import filler
import customer_sku
import amazon_pl

BASE = os.path.dirname(os.path.abspath(__file__))
# 可写数据目录：默认项目目录（本地/便携版行为不变）。
# 云端部署通过环境变量 DATA_DIR 指向持久卷（如 /data），
# 使 SKU 库 / 产品图片 / 规则 / 配置在重部署后不丢失。
DATA_DIR = os.environ.get("DATA_DIR") or BASE
os.makedirs(DATA_DIR, exist_ok=True)
CONFIG_PATH = os.path.join(DATA_DIR, "config.json")
RULES_PATH = os.path.join(DATA_DIR, "rules.json")
TEMPLATE_HTML = os.path.join(BASE, "templates", "index.html")
IMAGE_DIR = os.path.join(DATA_DIR, "sku_images")
os.makedirs(IMAGE_DIR, exist_ok=True)

app = FastAPI(title="单证提取填充工具")


def load_config():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def _resolve_tpl(p):
    """配置里的相对模板路径解析为项目根目录下的绝对路径；绝对/网络路径原样返回。"""
    if not p:
        return p
    if os.path.isabs(p) or "://" in p:
        return p
    return os.path.join(BASE, p)


def save_config(cfg):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def load_rules():
    with open(RULES_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def save_rules(rules):
    with open(RULES_PATH, "w", encoding="utf-8") as f:
        json.dump(rules, f, ensure_ascii=False, indent=2)


@app.get("/", response_class=HTMLResponse)
def index():
    with open(TEMPLATE_HTML, "r", encoding="utf-8") as f:
        html = f.read()
    return HTMLResponse(
        content=html,
        headers={
            "Cache-Control": "no-store, no-cache, must-revalidate, max-age=0",
            "Pragma": "no-cache",
            "Expires": "0",
        },
    )


@app.get("/api/config")
def get_config():
    return load_config()


@app.post("/api/save_config")
def post_save_config(
    template_path: str = Form(""),
    output_dir: str = Form(""),
    endpoint: str = Form(""),
    model: str = Form(""),
    api_key: str = Form("__UNCHANGED__"),
):
    cfg = load_config()
    if template_path:
        cfg["template_path"] = template_path
    if output_dir:
        cfg["output_dir"] = output_dir
    cfg.setdefault("llm", {})
    if endpoint:
        cfg["llm"]["endpoint"] = endpoint
    if model:
        cfg["llm"]["model"] = model
    if api_key != "__UNCHANGED__":
        cfg["llm"]["api_key"] = api_key
    save_config(cfg)
    return {"ok": True}


# --------------------------------------------------------------------------
# 规则管理 API（规则编辑器用）
# --------------------------------------------------------------------------

@app.get("/api/rules")
def get_rules():
    return load_rules()


@app.post("/api/rules")
async def post_rules(req: Request):
    """整体覆盖保存 rules.json（编辑器保存）。"""
    try:
        body = await req.json()
    except Exception:
        return JSONResponse(status_code=400, content={"error": "无效的 JSON"})
    if not isinstance(body, dict) or "rules" not in body:
        return JSONResponse(status_code=400, content={"error": "缺少 rules 字段"})
    save_rules(body)
    return {"ok": True, "count": len(body.get("rules", []))}


@app.post("/api/rules/test")
async def test_rule(
    file: UploadFile = File(...),
    rule_id: str = Form(""),
):
    """用某个规则(或自动匹配)对单个文档做抽取预览。"""
    rules = load_rules()
    data = await file.read()
    doc = extractor.read_doc(file.filename, data)
    dtype = extractor.detect_doc_type(doc)
    if rule_id:
        rule = next((r for r in rules["rules"] if r["id"] == rule_id), None)
    else:
        rule = extractor.find_rule(dtype, doc, rules["rules"])
    if not rule:
        return JSONResponse(status_code=404, content={
            "matched": False, "doc_type": dtype,
            "message": "未找到匹配规则，请在编辑器中新增/调整匹配条件。"})
    items = extractor.extract_with_rule(doc, rule, rules.get("settings", {}))
    return {
        "matched": True,
        "doc_type": dtype,
        "rule_id": rule["id"],
        "rule_name": rule["name"],
        "items": items,
    }


# --------------------------------------------------------------------------
# 转换 + 导出
# --------------------------------------------------------------------------

@app.post("/api/convert")
async def convert(files: List[UploadFile] = File(...)):
    """多文档规则抽取，返回 {shipment, items, method}。零 AI。"""
    rules = load_rules()
    docs = []
    for f in files:
        data = await f.read()
        docs.append({"filename": f.filename, "data": data})
    try:
        result = extractor.rule_extract(docs, rules)
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})
    result["used_key"] = False
    return result


@app.post("/api/export")
async def export(req: dict, background: BackgroundTasks):
    items = req.get("items", [])
    template_path = _resolve_tpl(req.get("template_path") or load_config().get("template_path", ""))
    output_name = req.get("output_name") or "客户货物运输托运书（系统下单用）-已填.xlsx"

    out = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    out.close()
    output_path = out.name
    try:
        filler.fill_template(items, template_path, output_path)
    except Exception as e:
        try:
            os.remove(output_path)
        except Exception:
            pass
        return JSONResponse(status_code=500, content={"error": str(e)})

    background.add_task(lambda p: _safe_remove(p), output_path)

    return FileResponse(
        output_path,
        filename=output_name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        background=background,
    )


def _safe_remove(p):
    try:
        os.remove(p)
    except Exception:
        pass


# --------------------------------------------------------------------------
# 客户 SKU 产品库（JSON 持久化 + 增删改查）
# --------------------------------------------------------------------------

@app.get("/api/customer_skus")
def get_customer_skus(keyword: str = ""):
    return customer_sku.list_skus(keyword)


@app.post("/api/customer_skus")
async def post_customer_sku(req: Request):
    try:
        body = await req.json()
    except Exception:
        return JSONResponse(status_code=400, content={"error": "无效的 JSON"})
    sku = (body.get("sku") or "").strip()
    if not sku:
        return JSONResponse(status_code=400, content={"error": "SKU 不能为空"})
    # op: 'add' = 新增（需查重报错）；'update' = 编辑（允许覆盖）
    op = (body.get("op") or "add").strip()
    lib = customer_sku.load_library()
    exists = sku in lib["skus"]
    if op == "add" and exists:
        return JSONResponse(status_code=409, content={
            "error": f"SKU『{sku}』已存在于产品库，请勿重复新增。如需修改请使用『编辑』功能。"})
    # 编辑时合并保留未在表单中的字段（如 net_per_ctn/长宽高/图片等）
    rec = customer_sku.upsert_sku(sku, body, merge=(op == "update"))
    return {"ok": True, "sku": sku, "record": rec}


@app.delete("/api/customer_skus/{sku}")
def delete_customer_sku(sku: str):
    ok = customer_sku.delete_sku(sku)
    return {"ok": ok}


@app.post("/api/customer_skus/{sku}/image")
async def upload_sku_image(sku: str, file: UploadFile = File(...)):
    ext = os.path.splitext(file.filename)[1] or ".png"
    safe = re.sub(r"[^\w.-]", "_", sku)
    dest = os.path.join(IMAGE_DIR, f"{safe}{ext}")
    data = await file.read()
    with open(dest, "wb") as f:
        f.write(data)
    abspath = os.path.abspath(dest)
    customer_sku.upsert_sku(sku, {"image_path": abspath})
    return {"ok": True, "image_path": abspath, "url": f"/sku_images/{os.path.basename(dest)}"}


# --------------------------------------------------------------------------
# Packing List 解析 + 托书生成
# --------------------------------------------------------------------------

@app.post("/api/parse_packing_list")
async def parse_packing_list(file: UploadFile = File(...)):
    data = await file.read()
    try:
        result = amazon_pl.parse_packing_list(data, file.filename)
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})
    return result


@app.post("/api/generate_bookings")
async def generate_bookings(req: Request):
    try:
        body = await req.json()
    except Exception:
        return JSONResponse(status_code=400, content={"error": "无效的 JSON"})
    tickets = body.get("tickets", [])
    # template_mode: "single"（默认，单票模板，每票一份文件）| "bulk"（批量下单模板）
    template_mode = (body.get("template_mode") or "single").strip().lower()
    cfg = load_config()
    if template_mode == "bulk":
        template_path = _resolve_tpl(body.get("bulk_template_path")
                                     or cfg.get("bulk_template_path", ""))
    else:
        template_path = _resolve_tpl(body.get("template_path") or cfg.get("template_path", ""))
    if not template_path:
        return JSONResponse(status_code=400, content={"error": "未配置模板路径"})
    if not tickets:
        return JSONResponse(status_code=400, content={"error": "没有可生成的票数据"})

    tmpdir = tempfile.mkdtemp(prefix="booking_")
    out_files = []
    try:
        if template_mode == "bulk":
            # 批量下单模板：所有票合成一个文件，序号 1、2、3…（同票共享序号）
            rows = []
            for idx, t in enumerate(tickets, 1):
                items = amazon_pl.build_items(t)
                if items:
                    rows.append({**t, "items": items})
            if not rows:
                return JSONResponse(status_code=500, content={"error": "没有可生成的票数据"})
            name = f"批量下单_{datetime.datetime.now():%Y%m%d_%H%M%S}.xlsx"
            out_path = os.path.join(tmpdir, name)
            try:
                import bulk_fill
                bulk_fill.fill_bulk_template(rows, template_path, out_path)
                out_files.append(out_path)
            except Exception as e:
                return JSONResponse(status_code=500,
                                    content={"error": f"批量下单模板生成失败: {e}"})
        else:
            for idx, t in enumerate(tickets, 1):
                items = amazon_pl.build_items(t)
                if not items:
                    continue
                wh = (t.get("warehouse") or f"ticket{idx}").strip()
                fba = (t.get("fba") or "").strip()
                name = f"托书_{wh}_{fba}.xlsx" if fba else f"托书_{wh}.xlsx"
                out_path = os.path.join(tmpdir, name)
                try:
                    filler.fill_template(items, template_path, out_path)
                    out_files.append(out_path)
                except Exception as e:
                    return JSONResponse(status_code=500,
                                        content={"error": f"票 {wh} 生成失败: {e}"})

        if not out_files:
            return JSONResponse(status_code=500, content={"error": "没有生成任何文件"})

        if template_mode == "bulk":
            # 批量模板：单文件直接下载，多文件同样打包
            if len(out_files) == 1:
                return FileResponse(
                    out_files[0],
                    filename=os.path.basename(out_files[0]),
                    media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        zip_path = os.path.join(tmpdir, "bookings.zip")
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for fp in out_files:
                zf.write(fp, os.path.basename(fp))
        return FileResponse(
            zip_path,
            filename=f"托书_{datetime.datetime.now():%Y%m%d_%H%M%S}.zip",
            media_type="application/zip",
            background=BackgroundTasks() if False else None,
        )
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})


# 静态服务 SKU 图片，供前端预览
app.mount("/sku_images", StaticFiles(directory=IMAGE_DIR), name="sku_images")


# ===== 集成模块：散货提单和预报生成 =====
# 复用本系统的 BasicAuth / 上传下载 / 错误格式；独立命名空间 /api/bol，互不冲突。
from bol_forecast.router import router as bol_router
from bol_forecast.config import JOBS_DIR as BOL_JOBS_DIR
from bol_forecast.data import db as bol_db

app.include_router(bol_router)

@app.get("/bol-forecast", response_class=HTMLResponse)
async def bol_forecast_page():
    p = os.path.join(BASE, "templates", "bol_forecast.html")
    with open(p, "r", encoding="utf-8") as f:
        return HTMLResponse(f.read())

@app.get("/bol-speech", response_class=HTMLResponse)
async def bol_speech_page():
    p = os.path.join(BASE, "templates", "bol_speech.html")
    with open(p, "r", encoding="utf-8") as f:
        return HTMLResponse(f.read())

# 生成产物（xlsx/pdf/zip）下载：router 返回 /bol-files/... URL
app.mount("/bol-files", StaticFiles(directory=str(BOL_JOBS_DIR)), name="bol_files")

@app.on_event("startup")
async def _bol_startup():
    bol_db.init_db()
    # ★COM 预热：服务启动时初始化 Excel 实例，避免用户第一次生成时
    # COM 进入"永久忙态"导致首次失败。com_retry 会自动重试，
    # 但预热后用户直接可用，无需第二次点击。
    try:
        from bol_forecast.generators.com_session import launch_excel
        launch_excel(visible=False)
        logger.info("COM 预热完成，Excel 实例已就绪")
    except Exception as e:
        logger.warning("COM 预热失败（用户第一次生成时仍会重试）: %s", e)


# --------------------------------------------------------------------------
# 可选访问鉴权（上线公网前必开）
# 设置环境变量 APP_PASSWORD 后，所有请求需带 Basic Auth（浏览器自动弹窗）。
# 未设置时完全透明，本地运行不受影响。
# 用原生 ASGI 中间件包裹，对 FileResponse（xlsx/zip 下载）零缓冲、零内存压力。
# --------------------------------------------------------------------------
class BasicAuthMiddleware:
    def __init__(self, app):
        self.app = app

    async def __call__(self, scope, receive, send):
        if scope.get("type") != "http" or not APP_PASSWORD:
            await self.app(scope, receive, send)
            return
        headers = dict(scope.get("headers", []))
        auth = headers.get(b"authorization", b"").decode(errors="ignore")
        expected = "Basic " + base64.b64encode(
            f"admin:{APP_PASSWORD}".encode()).decode()
        if auth != expected:
            await send({
                "type": "http.response.start",
                "status": 401,
                "headers": [
                    (b"www-authenticate",
                     b'Basic realm="booking-fill-tool"'),
                    (b"content-type", b"text/plain; charset=utf-8"),
                ],
            })
            await send({"type": "http.response.body",
                        "body": "401 Unauthorized".encode()})
            return
        await self.app(scope, receive, send)


APP_PASSWORD = os.environ.get("APP_PASSWORD", "")
app = BasicAuthMiddleware(app)


if __name__ == "__main__":
    import uvicorn
    # 固定访问地址：单一来源在 config.json（host/port）。
    # 127.0.0.1 是回环地址，任何电脑都指向本机自身，故在两台电脑上恒等生效。
    # 云平台通过环境变量 HOST/PORT 覆盖（标准做法），本地未设置时回退 config。
    _cfg = load_config()
    _host = os.environ.get("HOST", _cfg.get("host", "127.0.0.1"))
    _port = int(os.environ.get("PORT", _cfg.get("port", 8002)))
    uvicorn.run(app, host=_host, port=_port)
