# -*- coding: utf-8 -*-
"""Служебный модуль для собранного EXE (используется desktop.py).

Вызывается из desktop.py при запуске PyInstaller-сборки (sys.frozen):
 - переносит каталог data/ рядом с exe, если его там нет (первый запуск);
 - переключает текущий каталог процесса на каталог ресурсов sys._MEIPASS,
   чтобы Flask нашёл шаблоны app/templates;
 - добавляет каталог ресурсов в sys.path.

Важно: функции вызываются ДО импорта модуля app.
"""
import os
import shutil
import sys


def prepare():
    """Подготовка окружения внутри собранного exe. Возвращает путь к exe."""
    exe_path = os.path.abspath(sys.executable)
    exe_dir = os.path.dirname(exe_path)
    base = getattr(sys, '_MEIPASS', exe_dir)  # каталог распакованных ресурсов

    # --- данные (db.json, backups, exports) держим РЯДОМ с exe ------------
    bundled_data = os.path.join(base, 'data')
    local_data = os.path.join(exe_dir, 'data')
    if os.path.isdir(bundled_data):
        if not os.path.exists(local_data):
            try:
                shutil.copytree(bundled_data, local_data)
            except OSError:
                pass
        # указываем модулям приложения, где хранить данные
        os.environ.setdefault('PRIEMKA_DATA_DIR', local_data)

    # --- рабочие директории Python ---------------------------------------
    if base not in sys.path:
        sys.path.insert(0, base)
    try:
        os.chdir(base)   # Flask ищет templates относительно cwd
    except OSError:
        pass

    return exe_path
