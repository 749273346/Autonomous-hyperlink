@echo off
cd /d %~dp0
set "NOPAUSE="
if /I "%~1"=="--no-pause" set "NOPAUSE=1"
echo ========================================================
echo        AutoHyperlink Build Script
echo ========================================================

echo [1/4] Installing dependencies...
pip install pyinstaller watchdog xlrd==1.2.0 xlutils xlwt pywin32

echo [2/4] Building AutoHyperlink.exe (Silent Mode)...
pyinstaller --noconfirm --onefile --noconsole ^
    --name "AutoHyperlink" ^
    --hidden-import=xlutils ^
    --hidden-import=watchdog ^
    --hidden-import=xlwt ^
    --hidden-import=xlrd ^
    "auto_hyperlink.py"

echo [3/4] Building FolderMonitor.exe (Silent Mode)...
pyinstaller --noconfirm --onefile --noconsole ^
    --name "FolderMonitor" ^
    --hidden-import=win32com.client ^
    --hidden-import=pythoncom ^
    "folder_monitor.py"

echo [4/4] Cleaning up...
if exist build rmdir /s /q build 2>nul
if exist build (
  timeout /t 2 /nobreak >nul
  rmdir /s /q build 2>nul
)
if exist AutoHyperlink.spec del AutoHyperlink.spec
if exist FolderMonitor.spec del FolderMonitor.spec

echo.
echo ========================================================
echo Build Complete!
echo Executables location: %~dp0dist\
echo   - AutoHyperlink.exe (Worker)
echo   - FolderMonitor.exe (Manager)
echo ========================================================
echo Usage:
echo 1. Copy BOTH files to the folder you want to monitor.
echo 2. Run FolderMonitor.exe ONCE.
echo 3. It will auto-start AutoHyperlink.exe when you open the folder.
echo ========================================================
if not defined NOPAUSE pause
