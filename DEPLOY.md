# 部署到公网（换平台跑 Python 后端）

本工具是 FastAPI 后端，最适合部署到能跑常驻 Python 服务的平台：
**Render / Railway / Fly.io / 任意 VPS**。本文以 **Render** 为例（最省事、免费层可用），
但 `Dockerfile` 是通用的，Railway / Fly.io 同样能用。

> ⚠️ 仓库含客户 SKU 与托书数据，**务必设为私有仓库**再部署。

---

## 一、上线前必做：开访问鉴权

`app.py` 已内置可选 Basic Auth 中间件。部署时在平台的环境变量里加：

```
APP_PASSWORD=<一段够强的密码>
```

设置后，任何人访问都会弹出登录框（账号 `admin` / 密码即上述值）。
**不设则完全开放**——公网裸奔等于任何人都能上传、改你的 SKU 库、生成托书，切勿留空。

本地不设置 `APP_PASSWORD` 时该中间件完全透明，不影响你现在的双击使用。

## 二、部署步骤（Render，连 GitHub 自动构建）

1. 把整个 `booking-fill-tool/` 目录推到 **私有** GitHub 仓库。
2. 打开 https://render.com → New → Web Service → 连仓库。
3. 配置：
   - Runtime: **Docker**（用仓库里的 Dockerfile）
   - 或选 "Python" + Build Command `pip install -r requirements.txt` +
     Start Command `uvicorn app:app --host 0.0.0.0 --port $PORT`
   - Instance: Free（有休眠/月时长限制）或付费档（常驻）
4. Environment 变量：
   - `APP_PASSWORD` = 你的强密码（必填）
   - `PORT` 由 Render 自动注入，无需手填
5. 部署完成后，Render 给一个 `https://xxx.onrender.com` 公网地址。

Railway / Fly.io 同理：上传 Dockerfile 即可，`PORT` 由平台注入。

## 三、⚠️ 持久化坑（务必先看）

云平台的文件系统是 **临时** 的：

- 运行期你通过网页 **新增/编辑的 SKU、上传的产品图、改过的 rules.json**，
  在**重新部署或实例重启后会丢失**，回滚到仓库里的初始状态。
- 这是因为 `customer_sku.json` / `sku_images/` / `rules.json` 写在容器本地磁盘，
  平台不保证持久。

### 两种对策（请选择）

**方案 A — 接受临时（最简，推荐先用）**
把 Git 仓库当成 SKU 库的"唯一真相源"：要新增/修改产品，直接改仓库里的
`customer_sku.json` 和 `sku_images/` 然后重新部署。运行期网页只做"生成托书"这类
无状态操作。适合 SKU 库相对稳定的场景。

**方案 B — 接外部持久存储（最稳，需二次改造）**
- SKU 库 → 存 Supabase Postgres / Upstash 等
- 产品图 → 存对象存储（S3 / Cloudflare R2 / 腾讯云 COS）
- 需要我再改 `customer_sku.py` 与图片上传逻辑去对接。
告诉我你倾向哪家的存储，我来接。

## 四、本地验证（确保改动没破坏）

```bash
# 验证 app.py 能正常加载（含新中间件）
python -c "import app; print('app OK')"

# 本地照常双击 start_server.bat 或：
python -m uvicorn app:app --host 127.0.0.1 --port 8002
# 不设置 APP_PASSWORD → 行为与原先完全一致
```

## 五、回滚

部署出问题：在 Render 里 "Manual Deploy" 选上一个成功版本即可。
本地代码改动已通过 `git` 管理，回到旧 commit 即可。

---

### 小结
- 代码几乎不用改（已加 `PORT` 环境变量读取 + 可选 Basic Auth）。
- 上线关键两件事：**设 `APP_PASSWORD`** + **决定持久化方案（A 或 B）**。
- 数据隐私：仓库私有、公网加密码，这两点守住即可对外。
