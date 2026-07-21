@echo off
chcp 65001 >nul
title Синхронизация с GitHub

echo ========================================
echo  Синхронизация Phonebook из GitHub
echo ========================================
echo.

:: Пути
set "GITHUB_REPO=C:\Users\s2\Documents\Mino\phonebook-github"
set "WORK_DIR=C:\Users\s2\Documents\Mino\Phonebook-Web"

:: Проверка 존재ления каталогов
if not exist "%GITHUB_REPO%" (
    echo [ОШИБКА] Каталог GitHub не найден:
    echo   %GITHUB_REPO%
    echo.
    echo Сначала клонируйте репозиторий:
    echo   git clone https://github.com/dagmorport/phonebook-web.git "%GITHUB_REPO%"
    pause
    exit /b 1
)

if not exist "%WORK_DIR%" (
    echo [ОШИБКА] Рабочий каталог не найден:
    echo   %WORK_DIR%
    pause
    exit /b 1
)

:: Пulls из GitHub
echo [1/3] Обновление из GitHub...
cd /d "%GITHUB_REPO%"
git pull origin main
if %errorlevel% neq 0 (
    echo [ПРЕДУПРЕЖДЕНИЕ] Не удалось обновить из GitHub
)

:: Копирование index.html
echo [2/3] Копирование index.html...
copy /Y "%GITHUB_REPO%\index.html" "%WORK_DIR%\index.html" >nul
if %errorlevel% equ 0 (
    echo   OK: index.html скопирован
) else (
    echo   ОШИБКА: не удалось скопировать index.html
)

:: Копирование README.md (если нужен)
echo [3/3] Копирование README.md...
copy /Y "%GITHUB_REPO%\README.md" "%WORK_DIR%\README.md" >nul
if %errorlevel% equ 0 (
    echo   OK: README.md скопирован
) else (
    echo   ОШИБКА: не удалось скопировать README.md
)

echo.
echo ========================================
echo  Готово! 
echo ========================================
echo.
echo  Рабочий каталог: %WORK_DIR%
echo  Файлы обновлены: index.html, README.md
echo  Данные (employees.js) НЕ затронуты.
echo.
pause
