@echo off
:: 强制结束资源管理器
taskkill /IM explorer.exe /F

:: 等待一秒确保进程结束
timeout /t 1 /nobreak >nul

:: 删除图标缓存文件
echo 正在清理图标缓存...
del /A /Q "%localappdata%\Microsoft\Windows\Explorer\iconcache*"

:: 重启资源管理器
start explorer.exe

echo 图标缓存刷新完成！
pause