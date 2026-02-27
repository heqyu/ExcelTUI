@echo off
REM Excel TUI 启动器 - 可设为 .xlsx/.xls 默认打开方式
cd /d "%~dp0"
python3 excel_tui.py "%~1"
if errorlevel 1 pause
