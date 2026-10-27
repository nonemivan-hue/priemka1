@echo off
rem ================================================================
rem  SBORKA AVTONOMNOGO EXE -- VARIANT 1 (batch, ASCII-only).
rem  V etom skripte TOLKO komandyi BEZ kirillicy: on bolshe ne dayot
rem  oshibok is-not-recognized-as-an-internal-or-external-command,
rem  svyazannyih s kodovoy stranitsey konsoli.
rem  Trebuetsya: pip install pyinstaller
rem
rem  Druzhie variantyi sborki sm. README-BUILD.md:
rem    variant 2:  build.cmd   (spetsifikatsiya build.spec)
rem    variant 3:  powershell -ExecutionPolicy Bypass -File build.ps1
rem    variant 4:  pyinstaller build.spec --noconfirm
rem ================================================================
setlocal
cd /d "%~dp0"

where pyinstaller >nul 2>&1
if errorlevel 1 goto :nopyi

pyinstaller --noconfirm --clean --onefile --windowed ^
  --name "Priemka_tovarov" ^
  --version-file version_info.txt ^
  --add-data "app/templates;app/templates" ^
  --add-data "data;data" ^
  --add-data "Razdel1.xlsx;." ^
  --add-data "Ostalnoe.xlsx;." ^
  --add-data "Primer_formy.xlsx;." ^
  --add-data "CHANGELOG.md;." ^
  desktop.py
if errorlevel 1 goto :fail

echo.
echo Gotovo: dist\Priemka_tovarov.exe
echo Pri zhelanii exe mozhno pereimenovat v russkoe imya - na rabotu
echo ne vliyaet. Dannyie (papka data) sozdayutsya rjadom s exe pri
echo pervom zapuske.
goto :eof

:nopyi
echo [X] PyInstaller ne najden. Ustanovite:
echo        pip install pyinstaller
pause
exit /b 1

:fail
echo.
echo [X] SBORKA Zavershilas OSHIBKOY. Sm. tekst vyse.
echo Ispolzuite drugoy variant - sm. README-BUILD.md
pause
exit /b 1
