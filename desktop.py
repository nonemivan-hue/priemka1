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
import socket
import sys
import threading

logging.getLogger('werkzeug').setLevel(logging.ERROR)

from app.main import app  # noqa: E402
from version import APP_VERSION  # noqa: E402


def _free_port(start=8765, stop=8865):
    """Возвращает первый свободный локальный порт."""
    for port in range(start, stop):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(0.2)
            if s.connect_ex(('127.0.0.1', port)) != 0:
                return port
    raise RuntimeError('Не найден свободный порт для запуска программы')


def main():
    try:
        import webview
    except ImportError:
        print('Для настольного режима установите зависимости:')
        print('    pip install -r requirements.txt')
        sys.exit(1)

    from app import main as app_main
    app_main.DESKTOP_MODE = True  # печать — сразу в файл/редактор, без браузера

    port = _free_port()
    t = threading.Thread(
        target=lambda: app.run(host='127.0.0.1', port=port,
                               debug=False, use_reloader=False),
        daemon=True)
    t.start()

    # Ждём готовности внутреннего сервера
    for _ in range(100):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(0.2)
            if s.connect_ex(('127.0.0.1', port)) == 0:
                break
        import time
        time.sleep(0.1)

    webview.create_window(
        'Приёмка товаров — ГКУ КК УСЗН  v%s' % APP_VERSION,
        url='http://127.0.0.1:%d' % port,
        width=1360, height=860, min_size=(1024, 640),
    )
    webview.start()  # блокирует до закрытия окна; приложение закрывается целиком


if __name__ == '__main__':
    main()
