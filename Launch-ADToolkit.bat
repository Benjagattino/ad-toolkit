@echo off
title AD Toolkit — Launcher
echo Iniciando AD Toolkit...
echo.
powershell.exe -NoExit -ExecutionPolicy Bypass -Command ^
  "try { & '%~dp0AD-Toolkit-GUI.ps1' } catch { Write-Host ''; Write-Host 'ERROR: ' $_.Exception.Message -ForegroundColor Red; Write-Host 'En: ' $_.InvocationInfo.PositionMessage -ForegroundColor Yellow }"
echo.
pause
