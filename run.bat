@echo off
chcp 65001 >nul
cd /d "%~dp0"

where uv >nul 2>nul
if errorlevel 1 (
    echo uv نصب نیست. لطفاً این دستور را در PowerShell اجرا کنید:
    echo powershell -c "irm https://astral.sh/uv/install.ps1 ^| iex"
    pause
    exit /b 1
)

start "" /b uv run --with PyQt6 --with openpyxl normalize_excel_gui.py
exit /b 0
