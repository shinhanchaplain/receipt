@echo off
chcp 65001 > nul
cls
echo ============================================================
echo   [ Corporate Card Receipt Automation Performance Test ]
echo ============================================================
echo.

python --version > nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not in PATH.
    echo Please install Python 3.10 or higher.
    pause
    exit /b 1
)

echo [Step 1] Checking workspace files...
if not exist "test_performance.py" (
    echo [ERROR] test_performance.py not found.
    pause
    exit /b 1
)

echo [Step 2] Running full performance and integrity test...
echo.
python test_performance.py
set TEST_STATUS=%errorlevel%

echo.
echo ============================================================
if %TEST_STATUS% equ 0 (
    echo [SUCCESS] All performance tests passed!
    echo To launch the web dashboard, run 'run_app.bat'.
) else (
    echo [FAIL] Some tests failed. Please check the logs above.
)
echo ============================================================
echo.
pause
