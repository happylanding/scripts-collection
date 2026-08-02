@echo off
cd /d "%~dp0"
chcp 65001 >nul
title Excel空Sheet清理工具
java -jar "excel-cleaner.jar"
pause
