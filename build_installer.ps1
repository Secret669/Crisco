# Скрипт для автоматичної збірки інсталятора Crisco
# PowerShell Script

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Crisco - Збірка інсталятора" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Крок 1: Перевірка PyInstaller
Write-Host "[1/4] Перевірка PyInstaller..." -ForegroundColor Yellow
$pyinstallerInstalled = $null
try {
    $pyinstallerInstalled = & pip show pyinstaller 2>$null
} catch {
    $pyinstallerInstalled = $null
}

if (-not $pyinstallerInstalled) {
    Write-Host "PyInstaller не встановлено. Встановлення..." -ForegroundColor Yellow
    pip install pyinstaller
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Помилка встановлення PyInstaller!" -ForegroundColor Red
        exit 1
    }
} else {
    Write-Host "PyInstaller вже встановлено." -ForegroundColor Green
}

# Крок 2: Збірка EXE через PyInstaller
Write-Host ""
Write-Host "[2/4] Збірка EXE файлу..." -ForegroundColor Yellow
Write-Host "Це може зайняти кілька хвилин..." -ForegroundColor Gray

# Видалення старих збірок
if (Test-Path "dist") {
    Remove-Item -Path "dist" -Recurse -Force
}
if (Test-Path "build") {
    Remove-Item -Path "build" -Recurse -Force
}

# Запуск PyInstaller
pyinstaller Crisco_Optimized.spec --clean

if ($LASTEXITCODE -ne 0) {
    Write-Host "Помилка при збірці EXE!" -ForegroundColor Red
    exit 1
}

if (-not (Test-Path "dist\Crisco_Optimized.exe")) {
    Write-Host "EXE файл не був створений!" -ForegroundColor Red
    exit 1
}

Write-Host "EXE файл успішно створено!" -ForegroundColor Green

# Крок 3: Перевірка Inno Setup
Write-Host ""
Write-Host "[3/4] Перевірка Inno Setup..." -ForegroundColor Yellow

$innoSetupPath = "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if (-not (Test-Path $innoSetupPath)) {
    Write-Host "Inno Setup не знайдено!" -ForegroundColor Red
    Write-Host "Будь ласка, встановіть Inno Setup з:" -ForegroundColor Yellow
    Write-Host "https://jrsoftware.org/isdl.php" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Після встановлення запустіть скрипт знову." -ForegroundColor Yellow
    
    # Пропонуємо відкрити сайт
    $openSite = Read-Host "Відкрити сайт Inno Setup в браузері? (Y/N)"
    if ($openSite -eq "Y" -or $openSite -eq "y") {
        Start-Process "https://jrsoftware.org/isdl.php"
    }
    
    Write-Host ""
    Write-Host "EXE файл створено і знаходиться в папці: dist\" -ForegroundColor Green
    Write-Host "Ви можете використовувати його без інсталятора." -ForegroundColor Yellow
    exit 0
}

# Крок 4: Створення інсталятора
Write-Host ""
Write-Host "[4/4] Створення інсталятора..." -ForegroundColor Yellow

& $innoSetupPath "installer_setup.iss"

if ($LASTEXITCODE -ne 0) {
    Write-Host "Помилка при створенні інсталятора!" -ForegroundColor Red
    exit 1
}

# Успішне завершення
Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "  Інсталятор успішно створено!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Розташування файлів:" -ForegroundColor Cyan
Write-Host "  - EXE файл: dist\Crisco_Optimized.exe" -ForegroundColor White
Write-Host "  - Інсталятор: installer_output\Crisco_Setup.exe" -ForegroundColor White
Write-Host ""
Write-Host "Тепер ви можете розповсюджувати Crisco_Setup.exe" -ForegroundColor Green
Write-Host "для встановлення програми на будь-якому комп'ютері!" -ForegroundColor Green
Write-Host ""

# Пропонуємо відкрити папку з інсталятором
$openFolder = Read-Host "Відкрити папку з інсталятором? (Y/N)"
if ($openFolder -eq "Y" -or $openFolder -eq "y") {
    explorer.exe "installer_output"
}
