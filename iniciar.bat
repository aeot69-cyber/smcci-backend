@echo off
title SMCCI - Sistema de Consolidación MCCI
echo ========================================
echo   SMCCI - Iglesia MCCI
echo   Iniciando servidor...
echo   URL: http://127.0.0.1:5000
echo ========================================
cd /d "%~dp0"
py app.py
