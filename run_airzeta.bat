@echo off
title AirZeta Automation
cd /d "%~dp0"
echo ========================================
echo AirZeta Automation Start
echo ========================================
where python >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found.
    pause
    exit /b
)
python -m pip install openpyxl pandas xlrd lxml
python "%~dp0airzeta_automation.py"
pause