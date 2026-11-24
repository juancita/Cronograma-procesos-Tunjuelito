@echo off
chcp 65001 >nul
title Cronogramas de procesos SECOP II
cd /d "%~dp0"
if not exist "cronogramas.exe" (
  echo.
  echo   No se encuentra cronogramas.exe en esta carpeta.
  echo   Descomprima el ZIP completo y ejecute este archivo desde la misma carpeta.
  echo.
  pause
  exit /b 1
)
"%~dp0cronogramas.exe" %*
echo.
pause
