@echo off
cd /d "%~dp0"
chcp 65001 >nul 2>&1
title Excel 空 Sheet 清理工具
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0ExcelSheetCleaner.ps1"
pause
