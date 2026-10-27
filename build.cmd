@echo off
rem ================================================================
rem  Сборка автономного EXE — ВАРИАНТ 2 (надёжная обработка кириллицы)
rem  Запуск: двойным щелчком либо  cmd /c build.cmd
rem ================================================================
chcp 65001 >nul
setlocal
cd /d "%~dp0"

where pyinstaller >nul 2>&1
if errorlevel 1 (
    echo [!] PyInstaller не найден. Установите его командой:
    echo        pip install pyinstaller
    pause
    exit /b 1
)

echo [*] Запускаю сборку через спецификацию build.spec ...
pyinstaller build.spec --noconfirm --clean
if errorlevel 1 (
    echo.
    echo [X] Сборка завершилась ошибкой. Текст ошибки — выше.
    echo     Другие варианты сборки описаны в README-BUILD.md
    pause
    exit /b 1
)

echo.
echo [OK] Готово: dist\Приемка_товаров.exe
pause
