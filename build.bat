@echo off
chcp 65001 >nul 2>&1
echo ========================================
echo    Excel Processor - Build Script
echo ========================================
echo.

echo [1/3] Checking Python 3...
py -3 --version >nul 2>&1
if errorlevel 1 (
    echo Error: Python 3 not found. Please install Python 3 first.
    pause
    exit /b 1
)

echo [2/3] Installing dependencies...
py -3 -m pip install pandas openpyxl tkinterdnd2 pyinstaller -q

echo [3/3] Building...
py -3 -m PyInstaller --onefile --windowed --name "ExcelProcessor" --clean main.py

echo.
echo ========================================
echo    Build completed!
echo    Output: dist\ExcelProcessor.exe
echo ========================================
pause
