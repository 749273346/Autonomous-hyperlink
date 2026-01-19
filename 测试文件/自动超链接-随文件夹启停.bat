@echo off
setlocal
set "BASE=E:\QC-攻关小组\正在进行项目\自主超链接\Autonomous-hyperlink"
set "WATCH=%BASE%\测试文件"
set "PYW=%BASE%\.venv\Scripts\pythonw.exe"

if not exist "%PYW%" (
  set "PYW=%BASE%\.venv\Scripts\python.exe"
)

if exist "%WATCH%\folder_session_manager.py" (
  cd /d "%WATCH%"
  start "" "%PYW%" "%WATCH%\folder_session_manager.py"
)
endlocal

