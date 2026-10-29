# -*- coding: utf-8 -*-
"""Запуск программы приёмки товаров ГКУ КК УСЗН.

По умолчанию программа запускается как настольное приложение (без браузера):
    python run.py            — окно приложения (pywebview); перед запуском
                               предлагается выбрать IP-адрес и порт веб-сервера
    python run.py --server   — только внутренний сервер (доступ через браузер),
                               при запуске из консоли также запрашиваются IP/порт
Аргументы (для любого режима):
    --host <IP> (--ip <IP>)  — адрес, на котором поднимается сервер
                                (127.0.0.1 — локально, 0.0.0.0 — из сети)
    --port <N>               — порт сервера
Если аргументы указаны явно — запрос не показывается.
"""
import sys


def _cli_host_port(default_host, default_port):
    """Разбирает --host/--ip/--port из командной строки."""
    host, port = None, None
    args = sys.argv[1:]
    for i, a in enumerate(args):
        if a in ('--host', '--ip') and i + 1 < len(args):
            host = args[i + 1].strip() or None
        elif a == '--port' and i + 1 < len(args):
            try:
                p = int(args[i + 1])
                if 1 <= p <= 65535:
                    port = p
            except ValueError:
                pass
        elif a.startswith('--host='):
            host = a.split('=', 1)[1].strip() or None
        elif a.startswith('--ip='):
            host = a.split('=', 1)[1].strip() or None
        elif a.startswith('--port='):
            try:
                p = int(a.split('=', 1)[1])
                if 1 <= p <= 65535:
                    port = p
            except ValueError:
                pass
    return host, port


def _ask_server_params(default_host='0.0.0.0', default_port=5000):
    """Запрашивает IP/порт в консоли (если есть tty); иначе — значения по умолчанию."""
    host, port = default_host, default_port
    if sys.stdin is not None and sys.stdin.isatty():
        print('Параметры веб-сервера (Enter — по умолчанию):')
        h = input('  IP-адрес [%s]: ' % default_host).strip()
        if h:
            host = h
        p = input('  Порт     [%d]: ' % default_port).strip()
        if p:
            try:
                pi = int(p)
                if 1 <= pi <= 65535:
                    port = pi
                else:
                    print('  Порт вне диапазона 1..65535 — использован %d' % default_port)
            except ValueError:
                print('  Некорректный порт — использован %d' % default_port)
    return host, port


def run_server():
    from app.main import app
    host, port = _cli_host_port(None, None)
    if host is None or port is None:
        ah, ap = _ask_server_params()
        host = host or ah
        port = port or ap
    print('Сервер запущен: http://%s:%d' % (host, port))
    app.run(host=host, port=port, debug=False)


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
