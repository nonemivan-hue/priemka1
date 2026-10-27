# -*- coding: utf-8 -*-
"""Запуск программы приёмки товаров ГКУ КК УСЗН.

По умолчанию программа запускается как настольное приложение (без браузера):
    python run.py            — окно приложения (pywebview)
    python run.py --server   — только внутренний сервер (доступ через браузер
                               по адресу http://localhost:5000, режим для сети)
"""
import sys


def run_server():
    from app.main import app
    print("Сервер запущен: http://localhost:5000")
    app.run(host='0.0.0.0', port=5000, debug=False)


def run_desktop():
    try:
        import desktop
    except ImportError:
        print('Не найден модуль desktop.py')
        sys.exit(1)
    try:
        import webview  # noqa: F401
    except ImportError:
        print('Не установлен pywebview — запуск в режиме сервера.')
        print('Для настольного режима: pip install -r requirements.txt')
        run_server()
        return
    desktop.main()


if __name__ == '__main__':
    if '--server' in sys.argv or '-s' in sys.argv:
        run_server()
    else:
        run_desktop()
