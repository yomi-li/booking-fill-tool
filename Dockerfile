# booking-fill-tool — 轻量服务器部署镜像（Render / Railway / Sealos 亦可用）
FROM python:3.11-slim

WORKDIR /app

# 先装依赖，利用 Docker 层缓存（使用国内镜像加速）
COPY requirements.txt .
RUN pip install --no-cache-dir -i https://pypi.tuna.tsinghua.edu.cn/simple -r requirements.txt

# 再复制应用代码（sku_images / assets / customer_sku.json 种子 等都会进来）
COPY . .

# 可写数据目录指向持久卷；模板 assets/ 只读留在镜像内
ENV HOST=0.0.0.0 \
    PORT=8000 \
    DATA_DIR=/data

VOLUME ["/data"]
EXPOSE 8000

COPY entrypoint.sh /entrypoint.sh
RUN chmod +x /entrypoint.sh

ENTRYPOINT ["/entrypoint.sh"]
