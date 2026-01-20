@echo off
cd /d %~dp0
set "NOPAUSE="
if /I "%~1"=="--no-pause" set "NOPAUSE=1"
echo ========================================================
echo        AutoHyperlink Build Script
echo ========================================================

echo [1/3] Installing dependencies...
pip install pyinstaller watchdog xlrd==1.2.0 xlutils xlwt

echo [2/3] Building AutoHyperlink.exe...
pyinstaller --noconfirm --onefile --console ^
    --name "AutoHyperlink" ^
    --hidden-import=xlutils ^
    --hidden-import=watchdog ^
    --hidden-import=xlwt ^
    --hidden-import=xlrd ^
    "auto_hyperlink.py"

echo [3/3] Cleaning up...
if exist build rmdir /s /q build 2>nul
if exist build (
  timeout /t 2 /nobreak >nul
  rmdir /s /q build 2>nul
)
if exist AutoHyperlink.spec del AutoHyperlink.spec

echo.
echo ========================================================
echo Build Complete!
echo Executable location: %~dp0dist\AutoHyperlink.exe
echo ========================================================
echo Usage:
echo 1. Copy dist\AutoHyperlink.exe to the folder you want to monitor.
echo 2. Ensure the folder has year directories (e.g. 25, 26) and the Excel files.
echo 3. Double click to run.
echo ========================================================
if not defined NOPAUSE pause
