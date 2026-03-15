@echo off
chcp 65001
echo ========================================================
echo 九灵AI 智能办公平台 - 局域网共享启动脚本
echo ========================================================
echo.
echo 正在启动服务，允许局域网设备访问...
echo.
echo 启动成功后，请在其他电脑/手机浏览器输入 Network URL 地址访问
echo 例如: http://192.168.1.x:8501
echo.
echo 按 Ctrl+C 可停止服务
echo.

python -m streamlit run app.py --server.address 0.0.0.0 --server.runOnSave true

pause