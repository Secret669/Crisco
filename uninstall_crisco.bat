@echo off
chcp 65001 >nul
title Crisco - Видалення програми
color 0C

echo.
echo ═══════════════════════════════════════════════════════════
echo   CRISCO - Видалення програми
echo ═══════════════════════════════════════════════════════════
echo.

REM Перевірка прав адміністратора
net session >nul 2>&1
if %errorLevel% == 0 (
    echo [OK] Запущено з правами адміністратора
) else (
    echo [!] Для видалення потрібні права адміністратора.
    echo     Клацніть правою кнопкою на цей файл та виберіть
    echo     "Запустити від імені адміністратора"
    echo.
    pause
    exit
)

echo.
echo УВАГА: Ця програма буде повністю видалена з вашого комп'ютера.
echo Всі документи в папці "Zaminy" також будуть видалені!
echo.
set /p confirm="Ви впевнені? (Y/N): "
if /i not "%confirm%"=="Y" (
    echo Видалення скасовано.
    pause
    exit
)

echo.
echo [1/4] Видалення ярлика з робочого столу...
if exist "%USERPROFILE%\Desktop\Crisco.lnk" del /F /Q "%USERPROFILE%\Desktop\Crisco.lnk"
echo [OK] Ярлик видалено

echo.
echo [2/4] Видалення ярлика з меню Пуск...
if exist "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Crisco" rmdir /S /Q "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Crisco"
echo [OK] Ярлик видалено

echo.
echo [3/4] Закриття програми (якщо запущена)...
taskkill /F /IM Crisco_Optimized.exe >nul 2>&1
echo [OK] Програма закрита

echo.
echo [4/4] Видалення файлів програми...
if exist "C:\Program Files\Crisco" (
    rmdir /S /Q "C:\Program Files\Crisco"
    echo [OK] Файли видалено
) else (
    echo [!] Папка програми не знайдена
)

echo.
echo ═══════════════════════════════════════════════════════════
echo   ✓ ВИДАЛЕННЯ ЗАВЕРШЕНО УСПІШНО!
echo ═══════════════════════════════════════════════════════════
echo.
echo Програма Crisco повністю видалена з вашого комп'ютера.
echo.
pause
