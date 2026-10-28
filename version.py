# -*- coding: utf-8 -*-
"""Версия программы приёмки товаров ГКУ КК УСЗН.

Единый источник версии для приложения, запускчиков и сборки EXE.
История изменений — в CHANGELOG.md.
"""
APP_VERSION = '1.8.0'
APP_NAME = 'Приемка товаров — ГКУ КК УСЗН'


def version_tuple():
    return tuple(int(x) for x in APP_VERSION.split('.'))


def version_string():
    return APP_VERSION
