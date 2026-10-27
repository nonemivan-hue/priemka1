# Сборка автономного EXE — 4 варианта

## Почему `build_exe.bat` выдавал ошибку

Текстовые файлы `.bat` сохранялись в кодировке **UTF-8**, а консоль Windows
(`cmd.exe`) читает их в **национальной кодовой странице** (для русской
Windows — CP866/OEM). Русские строки `echo ...` распаковывались как мусор и
cmd пытался **выполнить их как команды**:

```
'анавливаю...' is not recognized as an internal or external command
'r' is not recognized as an internal or external command
```

Сама сборка при этом даже не запускалась. Исправление: в bat-файле теперь
**только ASCII-команды** (никакой кириллицы вне комментариев, комментарии —
транслитом), переносы строк — обязательно Windows-стиля (CRLF).

---

## Вариант 1 (самый простой) — обновлённый `build_exe.bat`

Просто дважды щёлкнуть `build_exe.bat` (или запустить из cmd).
Требуется установленный PyInstaller:

```bat
pip install pyinstaller
```

Результат: `dist\Priemka_tovarov.exe` (exe можно переименовать по-русски —
на работу это не влияет).

## Вариант 2 — `build.cmd` + спецификация `build.spec`

```bat
cmd /c build.cmd
```

`build.cmd` сохраняет кириллические пути корректно (запускается через
`cmd /c`, содержит `chcp 65001`) и вызывает сборку по файлу `build.spec`.

## Вариант 3 — PowerShell (`build.ps1`) — самый надёжный для кириллицы

PowerShell работает в Unicode, поэтому русские имена файлов и параметры
не искажаются вообще:

```powershell
powershell -ExecutionPolicy Bypass -File build.ps1
```

Скрипт сам находит Python, при отсутствии устанавливает PyInstaller
(`python -m pip install pyinstaller`) и собирает `dist\Приемка_товаров.exe`.

## Вариант 4 — одна команда в консоли (без всяких bat)

```bat
pyinstaller build.spec --noconfirm
```

или напрямую:

```bat
pyinstaller --noconfirm --clean --onefile --windowed ^
  --name "Priemka_tovarov" --version-file version_info.txt ^
  --add-data "app/templates;app/templates" ^
  --add-data "data;data" ^
  --add-data "Razdel1.xlsx;." --add-data "Ostalnoe.xlsx;." ^
  --add-data "CHANGELOG.md;." desktop.py
```

---

## Что ещё изменено в проекте (для корректной работы EXE)

1. `_frozen_bootstrap.py` — новый модуль: при запуске собранного exe
   создаёт папку `data\` **рядом с exe** (первый запуск копирует `db.json`),
   добавляет временный каталог ресурсов в `sys.path` и переключает туда cwd
   (чтобы Flask нашёл шаблоны `app/templates`).
2. `desktop.py` — вызывает `_frozen_bootstrap.prepare()` до импорта приложения;
   ошибки запуска в собранном exe теперь показываются в диалоговом окне
   (и пишутся в `startup_error.log` рядом с exe), а не «исчезают бесследно».
3. `app/database.py`, `app/backup.py`, `app/internet_loaders.py`, `app/main.py` —
   каталог данных берётся из переменной окружения `PRIEMKA_DATA_DIR`
   (выставляется bootstrap'ом), в исходном коде — как раньше, `./data`.
4. `app/print_form.py` — шаблоны `Раздел1.xlsx` / `Остальное.xlsx` внутри exe
   ищутся и под кириллическими, и под ASCII-именами (`Razdel1.xlsx`,
   `Ostalnoe.xlsx`); ASCII-копии лежат в корне проекта и включаются в сборку
   всегда — так кириллица в аргументах командной строки больше не может
   сломать упаковку.
5. `CHANGELOG.md` упаковывается в exe — страница «О программе» работает
   и в собранной версии.

## Частые проблемы

| Симптом | Причина / решение |
|---|---|
| `'...' is not recognized...` в bat | Кириллица в bat-файле не в той кодировке. Используйте варианты 2–4 или новый ASCII-only `build_exe.bat`. |
| `--add-data: Wrong syntax SOURCE:DEST` | Так ругается PyInstaller, запущенный **на Linux/macOS**. На Windows разделитель — `;` (как в скриптах выше). Собирать нужно на Windows. |
| Окно exe мелькнуло и закрылось | Теперь ошибка показывается в диалоге / `startup_error.log`. Чаще всего: антивирус блокирует запуск, либо не установлен WebView2 (Windows 7 — см. requirements.txt: `pythonnet`). |
| Файлы `data\` пропали после обновления exe | Папка `data\` рядом с exe не перезаписывается — достаточно положить новый exe в ту же папку. |
