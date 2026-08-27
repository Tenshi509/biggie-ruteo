@echo off
title Biggie Paraguay - Ruteo de Distribucion
color 0C

cd /d "%~dp0"

echo =====================================================================
echo           BIGGIE MINIMARKET PARAGUAY - SISTEMA DE RUTEO
echo =====================================================================
echo.
echo  [+] Iniciando Sistema de Ruteo Local en segundo plano...
echo  [+] Abriendo ventana de aplicacion de escritorio nativa...
echo.

:: Iniciar servidor en segundo plano si no est? corriendo
powershell -Command "if (!(Test-NetConnection -ComputerName localhost -Port 3456 -InformationLevel Quiet)) { Start-Process node -ArgumentList 'server.js' -WindowStyle Hidden }"

:: Esperar 1.5 segundos a que el puerto responda
timeout /t 2 /nobreak >nul

:: Abrir en modo aplicaci?n de escritorio nativa (sin barras de navegador)
if exist "%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe" (
    start "" "%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe" --app="http://localhost:3456" --window-size=1440,900
    exit
)

if exist "%ProgramFiles%\Microsoft\Edge\Application\msedge.exe" (
    start "" "%ProgramFiles%\Microsoft\Edge\Application\msedge.exe" --app="http://localhost:3456" --window-size=1440,900
    exit
)

if exist "%ProgramFiles%\Google\Chrome\Application\chrome.exe" (
    start "" "%ProgramFiles%\Google\Chrome\Application\chrome.exe" --app="http://localhost:3456" --window-size=1440,900
    exit
)

if exist "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" (
    start "" "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" --app="http://localhost:3456" --window-size=1440,900
    exit
)

:: Fallback si no encuentra Edge/Chrome directamente
start msedge.exe --app="http://localhost:3456" --window-size=1440,900 || start "" http://localhost:3456
exit
