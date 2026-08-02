@echo off
chcp 65001 >nul
title Excel空Sheet清理工具
java -jar "%~dp0excel-cleaner.jar" "%~dp0"
pause
