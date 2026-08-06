# booking-fill-tool — 云端部署镜像（Render / Railway / Fly.io 通用）
FROM python:3.11-slim

WORKDIR /app

# 先装依赖，利用 Docker 层缓存
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 再复制应用代码（sku_images / assets / customer_sku.json 等都会进来）
COPY . .

# 云平台通过环境变量 PORT 注入；默认 8000
ENV HOST=0.0.0.0 \
    PORT=8000

EXPOSE 8000

CMD ["sh", "-c", "uvicorn app:app --host 0.0.0.0 --port ${PORT:-8000}"]
