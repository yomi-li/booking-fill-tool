#!/bin/sh
# 容器启动入口：把仓库内的"种子数据"首次拷贝到持久卷 DATA_DIR，
# 之后 SKU 库 / 图片 / 规则 / 配置 / bol_forecast 数据全部读写持久卷，
# 重部署不丢。
set -e

DATA_DIR="${DATA_DIR:-/data}"
mkdir -p "$DATA_DIR"

seed_copy() {
  # $1=源(镜像内) $2=目标(持久卷)
  if [ ! -e "$2" ]; then
    cp -r "$1" "$2" 2>/dev/null || true
  fi
}

seed_copy /app/customer_sku.json "$DATA_DIR/customer_sku.json"
seed_copy /app/rules.json        "$DATA_DIR/rules.json"
seed_copy /app/config.json       "$DATA_DIR/config.json"
seed_copy /app/sku_images        "$DATA_DIR/sku_images"

# bol_forecast：app.db 由 startup 事件自动建；首次空目录由 SQLite 写首次表
mkdir -p "$DATA_DIR/backup" "$DATA_DIR/jobs" "$DATA_DIR/logs"

# 模板 assets/ 与 html_templates/ 为只读资源，留在镜像内，不拷贝。

exec uvicorn app:app --host 0.0.0.0 --port "${PORT:-8000}"
