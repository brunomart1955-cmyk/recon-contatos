@echo off
chcp 65001 >nul
cd /d "%~dp0"
REM Arraste uma lista (.txt ou .csv) em cima deste arquivo, ou de dois cliques para usar lista.txt
if "%~1"=="" (
  node recon.js lista.txt
) else (
  node recon.js "%~1"
)
echo.
pause
