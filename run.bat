@echo off
chcp 65001 >nul 2>&1
echo ========================================
echo    Excel Processor - Run Script
echo ========================================
echo.

echo Checking dependencies...
py -3 -m pip show pandas >nul 2>&1
if errorlevel 1 (
    echo Installing dependencies...
    py -3 -m pip install pandas openpyxl tkinterdnd2 -q
)

echo Starting program...
py -3 main.py

if errorlevel 1 (
    echo.
    echo Program error. Please check the error message.
    pause
)
