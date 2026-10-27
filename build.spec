# -*- coding: utf-8 -*-
"""Спецификация PyInstaller для сборки автономного EXE (альтернатива build_exe.bat).

Запуск из корня проекта:
    pyinstaller build.spec --noconfirm

Результат: dist/Приемка_товаров.exe
"""
import os

block_cipher = None


a = Analysis(
    ['desktop.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('app/templates', 'app/templates'),
        ('data', 'data'),
        # Шаблоны упаковываются и под кириллическими, и под ASCII-именами
        # (см. app/print_form.py:_res) — так сборка переживает любые
        # проблемы локальной кодовой страницы cmd при передаче аргументов.
        ('Раздел1.xlsx', '.'),
        ('Остальное.xlsx', '.'),
        ('Razdel1.xlsx', '.'),
        ('Ostalnoe.xlsx', '.'),
        ('CHANGELOG.md', '.'),
    ],
    hiddenimports=[
        # прокси-движки pywebview подбираются динамически — подключаем явно
        'webview.platforms.winforms',
        'webview.platforms.edgechromium',
        'webview.platforms.mshtml',
        'clr',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Приемка_товаров',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # без чёрного окна консоли
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version='version_info.txt',
    icon=None,              # при наличии файла app.ico заменить на icon='app.ico'
)
