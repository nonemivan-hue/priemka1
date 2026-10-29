# -*- coding: utf-8 -*-
"""Настольный запускчик программы приёмки товаров ГКУ КК УСЗН.

Запускает программу в отдельном окне приложения — БЕЗ браузера:

    python desktop.py

Интерфейс отображается во встроенном оконном компоненте (pywebview):
 - Windows 7/10/11 — WebView2 / MSHTML (входит в состав ОС);
 - Linux          — GTK WebKit / Qt WebEngine;
 - macOS          — WKWebView.

Локальный веб-сервер Flask поднимается во внутреннем потоке и доступен
только самому приложению (порт подбирается автоматически).
"""
import logging
import os
import socket
import sys
import threading

logging.getLogger('werkzeug').setLevel(logging.ERROR)

# --- подготовка окружения внутри собранного EXE (до импорта app) ----------
if getattr(sys, 'frozen', False):
    import _frozen_bootstrap
    _frozen_bootstrap.prepare()

from app.main import app  # noqa: E402
from version import APP_VERSION  # noqa: E402


DEFAULT_HOST = '127.0.0.1'
DEFAULT_PORT = 8765


def _free_port(start=DEFAULT_PORT, stop=DEFAULT_PORT + 100):
    """Возвращает первый свободный локальный порт."""
    for port in range(start, stop):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(0.2)
            if s.connect_ex(('127.0.0.1', port)) != 0:
                return port
    raise RuntimeError('Не найден свободный порт для запуска программы')


def _norm_host(value):
    s = (value or '').strip()
    return s if s else DEFAULT_HOST


def _norm_port(value, default=DEFAULT_PORT):
    s = str(value or '').strip()
    if not s:
        return default
    try:
        p = int(s)
        if 1 <= p <= 65535:
            return p
    except ValueError:
        pass
    return default


def _dialog_server_params(host, port):
    """Оконный диалог выбора IP/порта (tkinter входит в состав Python/EXE).

    Возвращает (host, port) или None, если пользователь отменил выбор
    (в этом случае используются значения по умолчанию).
    """
    try:
        import tkinter as tk
        from tkinter import ttk
    except Exception:
        return None
    result = {'r': None}
    root = tk.Tk()
    root.withdraw()
    root.title('Параметры веб-сервера')
    dlg = tk.Toplevel(root)
    dlg.title('Приёмка — адрес веб-сервера')
    dlg.resizable(False, False)
    frm = ttk.Frame(dlg, padding=12)
    frm.grid(sticky='nsew')
    ttk.Label(frm, text='IP-адрес сервера:').grid(row=0, column=0, sticky='w')
    e_host = ttk.Combobox(frm, width=24, values=[
        '127.0.0.1', '0.0.0.0', 'localhost'])
    e_host.set(host)
    e_host.grid(row=0, column=1, sticky='we', pady=(0, 6))
    ttk.Label(frm, text='Порт:').grid(row=1, column=0, sticky='w')
    e_port = ttk.Entry(frm, width=26)
    e_port.insert(0, str(port))
    e_port.grid(row=1, column=1, sticky='we', pady=(0, 6))
    ttk.Label(frm, foreground='#666', justify='left', text=(
        '127.0.0.1 — доступ только с этого компьютера\n'
        '0.0.0.0 — доступ из локальной сети по IP этого ПК')).grid(
        row=2, column=0, columnspan=2, sticky='w', pady=(0, 8))

    def ok(event=None):
        result['r'] = (_norm_host(e_host.get()), _norm_port(e_port.get()))
        dlg.destroy()

    def cancel():
        result['r'] = None
        dlg.destroy()

    btns = ttk.Frame(frm)
    btns.grid(row=3, column=0, columnspan=2, sticky='e')
    ttk.Button(btns, text='Запустить', command=ok).pack(side='right', padx=(6, 0))
    ttk.Button(btns, text='Отмена', command=cancel).pack(side='right')
    dlg.bind('<Return>', ok)
    dlg.bind('<Escape>', lambda e: cancel())
    e_port.focus_set()
    dlg.transient(root)
    dlg.grab_set()
    root.mainloop()
    return result['r']


def _ask_server_params():
    """Перед запуском предлагает пользователю выбрать IP-адрес и порт,
    на которых будет работать внутренний веб-сервер приложения.

    Порядок: аргументы командной строки (--host/--ip, --port) → запрос
    в консоли (если она есть) → оконный диалог (EXE без консоли) →
    значения по умолчанию (127.0.0.1:8765).
    """
    host, port = DEFAULT_HOST, DEFAULT_PORT
    args = sys.argv[1:]
    for i, a in enumerate(args):
        if a in ('--host', '--ip') and i + 1 < len(args):
            host = _norm_host(args[i + 1])
        elif a == '--port' and i + 1 < len(args):
            port = _norm_port(args[i + 1], port)
        elif a.startswith('--host='):
            host = _norm_host(a.split('=', 1)[1])
        elif a.startswith('--ip='):
            host = _norm_host(a.split('=', 1)[1])
        elif a.startswith('--port='):
            port = _norm_port(a.split('=', 1)[1], port)
    if ('--host' in args or '--ip' in args or '--port' in args or
            any(a.startswith(('--host=', '--ip=', '--port=')) for a in args)):
        return host, port  # явные параметры CLI — не спрашиваем

    if sys.stdin is not None and sys.stdin.isatty():
        print('Параметры веб-сервера приложения (Enter — по умолчанию):')
        print('  IP-адрес [по умолчанию %s; 0.0.0.0 — доступ из локальной сети]: '
              % DEFAULT_HOST, end='')
        host = _norm_host(input())
        print('  Порт     [по умолчанию %d]: ' % DEFAULT_PORT, end='')
        inp = input()
        new_port = _norm_port(inp, None)
        if new_port is None and str(inp).strip():
            print('  Некорректный порт "%s" — использован %d' % (inp, DEFAULT_PORT))
            new_port = DEFAULT_PORT
        port = new_port or DEFAULT_PORT
        return host, port

    params = _dialog_server_params(host, port)
    if params:
        return params
    return host, port


def main():
    try:
        import webview
    except ImportError:
        print('Для настольного режима установите зависимости:')
        print('    pip install -r requirements.txt')
        sys.exit(1)

    from app import main as app_main
    app_main.DESKTOP_MODE = True  # печать — сразу в файл/редактор, без браузера

    host, port = _ask_server_params()
    if host == '0.0.0.0':
        probe_host = '127.0.0.1'
    else:
        probe_host = host
    t = threading.Thread(
        target=lambda: app.run(host=host, port=port,
                               debug=False, use_reloader=False),
        daemon=True)
    t.start()

    # Ждём готовности внутреннего сервера
    for _ in range(100):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(0.2)
            if s.connect_ex((probe_host, port)) == 0:
                break
        import time
        time.sleep(0.1)

    webview.create_window(
        'Приёмка товаров — ГКУ КК УСЗН  v%s' % APP_VERSION,
        url='http://%s:%d' % (probe_host, port),
        width=1360, height=860, min_size=(1024, 640),
    )
    webview.start()  # блокирует до закрытия окна; приложение закрывается целиком


def _crash_dialog(exc):
    """Показывает ошибку в окне (в собранном EXE консоли нет)."""
    import traceback
    text = ''.join(traceback.format_exception(type(exc), exc, exc.__traceback__))
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(
            0, 'Приложение не запустилось:\n\n%s' % text,
            'Ошибка запуска — Приёмка товаров', 0x10)
    except Exception:
        try:
            log = os.path.join(os.path.dirname(os.path.abspath(sys.executable)),
                               'startup_error.log')
            with open(log, 'w', encoding='utf-8') as f:
                f.write(text)
        except Exception:
            pass


if __name__ == '__main__':
    try:
        main()
    except Exception as e:  # noqa: BLE001
        if getattr(sys, 'frozen', False):
            _crash_dialog(e)
        else:
            raise
