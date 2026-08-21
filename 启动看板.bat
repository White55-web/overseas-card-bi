@echo off
chcp 65001 >nul
cd /d "%~dp0"

:: 1. 写入防白屏与防卡顿配置
if not exist .streamlit mkdir .streamlit
(
    echo [server]
    echo headless = false
    echo enableCORS = false
    echo enableXsrfProtection = false
    echo enableWebsocketCompression = false
) > .streamlit\config.toml

:: 2. 直接启动 Streamlit 看板
start "看板服务" streamlit run dashboard.py

:: 3. 直接启动 cpolar 穿透（不套娃 cmd，绝对不会丢失）
start "公网穿透" cpolar http 8501

:: 4. 等待 2 秒唤起默认浏览器
timeout /t 2 /nobreak >nul
start http://localhost:8501