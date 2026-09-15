@echo off
echo ============================================
echo   Sincronizar Google Sheets -> BD
echo ============================================
echo.
cd /d C:\Datos\MCCI\SMCCI\backend
py sync_sheets.py
echo.
pause
