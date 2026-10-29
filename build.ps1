# ================================================================
#  Сборка автономного EXE — ВАРИАНТ 3 (PowerShell)
#  Запуск:  powershell -ExecutionPolicy Bypass -File build.ps1
# ================================================================
$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot

# Имена файлов передаются как есть — PowerShell работает в Unicode,
# поэтому кириллица в путях не искажается.
$common = @(
    '--noconfirm', '--clean', '--onefile', '--windowed',
    '--name', 'Приемка_товаров',
    '--version-file', 'version_info.txt',
    '--add-data', 'app/templates;app/templates',
    '--add-data', 'data;data',
    '--add-data', 'Razdel1.xlsx;.',
    '--add-data', 'Ostalnoe.xlsx;.',
    '--add-data', 'CHANGELOG.md;.',
    'desktop.py'
)

Write-Host '[*] Ищу Python...'
$py = $null
foreach ($c in @('python', 'py')) {
    if (Get-Command $c -ErrorAction SilentlyContinue) { $py = $c; break }
}
if (-not $py) { Write-Host '[X] Python не найден в PATH.' -ForegroundColor Red; exit 1 }

Write-Host '[*] Проверяю наличие PyInstaller...'
& $py -m PyInstaller --version *> $null
if ($LASTEXITCODE -ne 0) {
    Write-Host '[!] PyInstaller не установлен. Устанавливаю...' -ForegroundColor Yellow
    & $py -m pip install pyinstaller
    if ($LASTEXITCODE -ne 0) { Write-Host '[X] Не удалось установить PyInstaller.' -ForegroundColor Red; exit 1 }
}

# Зависимости приложения (flask, openpyxl, python-docx и др.) должны быть
# установлены — иначе PyInstaller просто не включит недостающие модули в EXE.
Write-Host '[*] Проверяю зависимости (requirements.txt)...'
& $py -c "import flask, openpyxl, docx" *> $null
if ($LASTEXITCODE -ne 0) {
    Write-Host '[!] Недостающие зависимости. Устанавливаю из requirements.txt...' -ForegroundColor Yellow
    & $py -m pip install -r requirements.txt
    if ($LASTEXITCODE -ne 0) { Write-Host '[X] Не удалось установить зависимости.' -ForegroundColor Red; exit 1 }
}

Write-Host '[*] Запускаю сборку...'
& $py -m PyInstaller @common
if ($LASTEXITCODE -ne 0) {
    Write-Host ''
    Write-Host '[X] Сборка завершилась ошибкой (код $LASTEXITCODE). Текст — выше.' -ForegroundColor Red
    Write-Host '    Попробуйте вариант со спецификацией: pyinstaller build.spec --noconfirm'
    exit 1
}

Write-Host ''
Write-Host '[OK] Готово: dist\Приемка_товаров.exe' -ForegroundColor Green
