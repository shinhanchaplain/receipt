@echo off
cd /d "%~dp0"
chcp 65001 > nul
cls
echo ============================================================
echo   [ Starting Corporate Card Receipt Web Dashboard ]
echo ============================================================
echo.
echo Launching Streamlit web application...
echo Default URL: http://localhost:8501
echo (Press Ctrl + C in this window to stop the server)
echo.

python -m streamlit run app.py

pause
