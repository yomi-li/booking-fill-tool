# booking-fill-tool — 轻量服务器部署镜像（Ubuntu 22.04 / Tencent Lighthouse）
#
# 渲染路径：WeasyPrint (HTML→PDF) 主路径。Excel COM (pywin32) 在 Linux
# 不可用，bol_forecast 在容器内只能走 BOL_RENDERER=html。

FROM python:3.11-slim

# 1) 系统依赖：WeasyPrint 运行时（pango/harfbuzz）+ CJK 字体 + 中英文
RUN apt-get update && apt-get install -y --no-install-recommends \
        libpango-1.0-0 \
        libpangoft2-1.0-0 \
        libharfbuzz-subset0 \
        libharfbuzz0b \
        libffi-dev \
        fonts-noto-cjk \
        fonts-liberation \
        fontconfig \
        curl \
        ca-certificates \
    && rm -rf /var/lib/apt/lists/* \
    && fc-cache -f

# 2) Python 依赖
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -i https://pypi.tuna.tsinghua.edu.cn/simple -r requirements.txt

# 3) 应用代码（bol_templates / bol_forecast / html_templates / static 一并进入）
COPY . .

# 4) 环境变量：HTTP 监听 8000（compose 映射到宿主机 8002）；DATA_DIR 走持久卷
ENV HOST=0.0.0.0 \
    PORT=8000 \
    DATA_DIR=/data \
    BOL_RENDERER=html \
    PYTHONUNBUFFERED=1

# 5) 持久卷：app.db / jobs / logs / 配置改写
VOLUME ["/data"]
EXPOSE 8000

# 6) 入口：首次启动把仓库种子数据拷贝到持久卷，之后读写都在持久卷
COPY entrypoint.sh /entrypoint.sh
RUN chmod +x /entrypoint.sh

ENTRYPOINT ["/entrypoint.sh"]
