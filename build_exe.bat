@echo off
rem Сборка автономного EXE-файла программы (запуск без браузера и без Python)
rem Требуется: pip install pyinstaller
pyinstaller --noconfirm --onefile --version-file version_info.txt --name "Приемка_товаров" ^
  --add-data "app/templates;app/templates" ^
  --add-data "data;data" ^
  --add-data "Раздел1.xlsx;." --add-data "Остальное.xlsx;." ^
  desktop.py
echo.
echo Готово: файл dist\Приемка_товаров.exe запускается двойным щелчком,
echo окно приложения открывается без браузера.
pause
