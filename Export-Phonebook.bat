@echo off
chcp 65001 >nul 2>&1
title Export Phonebook
echo ============================================
echo   Export Phonebook from Active Directory
echo ============================================
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Export-Phonebook.ps1"
echo.
echo ============================================
echo   Done. Window closes in 3 seconds...
echo ============================================
timeout /t 3 /nobreak >nul
