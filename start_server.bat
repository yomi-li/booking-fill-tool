@echo off
chcp 65001 >nul
cd /d "C:\Users\admin\WorkBuddy\booking-fill-tool"
echo 启动单证流水线服务（端口 8002）...
echo 注意：Excel COM 生成需要在本机当前用户桌面会话中运行，请勿用后台服务方式启动。
set BOL_RENDERER=com
start "BOLForecast 8002" "C:\Users\admin\WorkBuddy\booking-fill-tool\python\python.exe" -m uvicorn app:app --host 127.0.0.1 --port 8002 --log-level info
echo 服务窗口已打开，请勿关闭该窗口。
pause
