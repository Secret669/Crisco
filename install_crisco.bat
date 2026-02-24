@echo off
chcp 65001 >nul
title Crisco - Встановлення програми
color 0A

echo.
echo ═══════════════════════════════════════════════════════════
echo   CRISCO - Встановлення програми для ведення замін
echo ═══════════════════════════════════════════════════════════
echo.

REM Перевірка прав адміністратора
net session >nul 2>&1
if %errorLevel% == 0 (
    echo [OK] Запущено з правами адміністратора
) else (
    echo [!] Для встановлення потрібні права адміністратора.
    echo     Клацніть правою кнопкою на цей файл та виберіть
    echo     "Запустити від імені адміністратора"
    echo.
    pause
    exit
)

echo.
echo [1/5] Створення папки програми...
if not exist "C:\Program Files\Crisco" mkdir "C:\Program Files\Crisco"
echo [OK] Папка створена

echo.
echo [2/5] Копіювання файлів програми...
copy /Y "%~dp0Crisco_Portable\Crisco_Optimized.exe" "C:\Program Files\Crisco\" >nul
copy /Y "%~dp0Crisco_Portable\dataBase.mdb" "C:\Program Files\Crisco\" >nul
copy /Y "%~dp0Crisco_Portable\README.txt" "C:\Program Files\Crisco\" >nul
echo [OK] Файли скопійовано

echo.
echo [3/5] Створення папки для документів...
if not exist "C:\Program Files\Crisco\Zaminy" mkdir "C:\Program Files\Crisco\Zaminy"
echo [OK] Папка створена

echo.
echo [4/5] Створення ярлика на робочому столі...
powershell -Command "$WshShell = New-Object -comObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('%USERPROFILE%\Desktop\Crisco.lnk'); $Shortcut.TargetPath = 'C:\Program Files\Crisco\Crisco_Optimized.exe'; $Shortcut.WorkingDirectory = 'C:\Program Files\Crisco'; $Shortcut.Description = 'Програма для ведення замін'; $Shortcut.Save()"
echo [OK] Ярлик створено

echo.
echo [5/5] Створення ярлика в меню Пуск...
if not exist "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Crisco" mkdir "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Crisco"
powershell -Command "$WshShell = New-Object -comObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('%APPDATA%\Microsoft\Windows\Start Menu\Programs\Crisco\Crisco.lnk'); $Shortcut.TargetPath = 'C:\Program Files\Crisco\Crisco_Optimized.exe'; $Shortcut.WorkingDirectory = 'C:\Program Files\Crisco'; $Shortcut.Description = 'Програма для ведення замін'; $Shortcut.Save()"
echo [OK] Ярлик створено

echo.
echo ═══════════════════════════════════════════════════════════
echo   ✓ ВСТАНОВЛЕННЯ ЗАВЕРШЕНО УСПІШНО!
echo ═══════════════════════════════════════════════════════════
echo.
echo Програма встановлена в: C:\Program Files\Crisco\
echo Ярлик створено на робочому столі: Crisco
echo.
echo Тепер ви можете запустити програму з робочого столу
echo або через меню Пуск.
echo.
echo ═══════════════════════════════════════════════════════════
echo.

REM Запитати чи запустити програму
set /p launch="Запустити програму зараз? (Y/N): "
if /i "%launch%"=="Y" start "" "C:\Program Files\Crisco\Crisco_Optimized.exe"

echo.
echo Дякуємо за встановлення Crisco!
echo Натисніть будь-яку клавішу для виходу...
pause >nul
