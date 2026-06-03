@echo off
title Procurement App
color 0F

cd /d "%~dp0"

echo ========================================
echo         PROCUREMENT APP
echo ========================================
echo.

REM Активация виртуального окружения
call venv\Scripts\activate.bat

if errorlevel 1 (
    echo.
    echo ERROR: Failed to activate virtual environment.
    pause
    exit /b 1
)

echo Starting application...
echo.

python main.py

set EXIT_CODE=%ERRORLEVEL%

REM Если приложение завершилось с ошибкой - показываем сообщение
if not "%EXIT_CODE%"=="0" (
    echo.
    echo ========================================
    echo APPLICATION CRASHED
    echo Exit code: %EXIT_CODE%
    echo ========================================
    echo.
    pause
)

exit /b %EXIT_CODE%