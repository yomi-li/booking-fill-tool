@echo off
setlocal
set DIR=%~dp0
set PY=%DIR%python\python.exe
cd /d "%DIR%"

REM ============================================================
REM 固定访问地址：统一从 config.json 读取（单一事实来源）
REM 127.0.0.1 是回环地址，任何电脑都指向本机自身，故在两台
REM 电脑上恒等生效，不会出现"找不到地址"。
REM 如需改端口/地址，只改 config.json 即可，本脚本自动跟随。
REM ============================================================
for /f "tokens=*" %%h in ('%PY% -c "import json;print(json.load(open('config.json')).get('host','127.0.0.1'))"') do set HOST=%%h
for /f "tokens=*" %%p in ('%PY% -c "import json;print(json.load(open('config.json')).get('port',8002))"') do set PORT=%%p
for /f "tokens=*" %%u in ('%PY% -c "import json;print(json.load(open('config.json')).get('access_url','http://127.0.0.1:8002/'))"') do set ACCESS_URL=%%u

echo [booking-fill-tool] host=%HOST%  port=%PORT%
echo [booking-fill-tool] 固定访问地址: %ACCESS_URL%

REM 看门狗：端口已被占用 -> 服务可能已在运行，直接打开浏览器
%PY% -c "import socket;s=socket.socket();r=s.connect_ex(('127.0.0.1',%PORT%));s.close();raise SystemExit(0 if r!=0 else 1)" >nul 2>&1
if errorlevel 1 (
  echo [booking-fill-tool] 端口 %PORT% 已被占用（服务可能已在运行），直接打开浏览器。
  goto open_browser
)

REM 后台（最小化）启动 uvicorn 服务
start "booking-fill-tool-server" /min %PY% -m uvicorn app:app --host %HOST% --port %PORT%

REM 等待服务就绪再打开浏览器（约 3 秒）
ping -n 4 127.0.0.1 >nul

:open_browser
start "" %ACCESS_URL%
echo [booking-fill-tool] 已自动打开浏览器。服务窗口已最小化，请勿关闭（关闭即停止服务）。
echo [booking-fill-tool] 若浏览器未自动弹出，请手动访问: %ACCESS_URL%
pause >nul
endlocal
