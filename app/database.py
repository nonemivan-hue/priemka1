# -*- coding: utf-8 -*-
"""Слой хранения данных программы приёмки товаров ГКУ КК УСЗН.

Хранение: JSON-файл в каталоге data/ (простая переносимая база).
"""
import json
import os
import threading
from datetime import date, datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(BASE_DIR)
DATA_FILE = os.path.join(ROOT_DIR, 'data', 'db.json')

_lock = threading.Lock()

FIAS_ADDRESS_TEMPLATE = {
    "indeks": "",
    "strana": "Российская Федерация",
    "kray_oblast": "",
    "rayon": "",
    "gorod": "",
    "naselenny_punkt": "",
    "ulica": "",
    "dom": "",
    "korpus": "",
    "kvartira": "",
}


def _json_default(o):
    if isinstance(o, (date, datetime)):
        return o.strftime('%d.%m.%Y')
    raise TypeError(str(type(o)))


def empty_db():
    return {
        "constants": {
            "uchrezhdenie": {
                "учреждение_получатель": "",
                "обособленное_подразделение": "",
                "структурное_подразделение": "",
                "главный_администратор_доходов": "",
                "наименование_бюджета": "",
                "валюта": "Рубль",
                "адрес_грузополучателя": "",
                "inn": "",
                "kpp": "",
                "ogrn": "",
                "adres_zakazchika": dict(FIAS_ADDRESS_TEMPLATE),
                "mesto_postavki": dict(FIAS_ADDRESS_TEMPLATE),
                "rukovoditel": "",
            }
        },
        "catalogs": {
            "postavshiki": [],
            "vidy_dokumentov": [],
            "nomenklatura": [],
            "stavki_nds": ["без НДС", "0", "5", "7", "10", "20"],
            "strany_proiskhozhdeniya": [],
            "fias": [],
            "ed_izmereniya": [],
            "sotrudniki": [],
        },
        "documents": {
            "pri_kazy_comissii": [],
            "priemki": [],
        },
    }


def _load():
    if not os.path.exists(DATA_FILE):
        return empty_db()
    with open(DATA_FILE, 'r', encoding='utf-8') as f:
        return json.load(f)


def _save(db):
    os.makedirs(os.path.dirname(DATA_FILE), exist_ok=True)
    tmp = DATA_FILE + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(db, f, ensure_ascii=False, indent=2, default=_json_default)
    os.replace(tmp, DATA_FILE)


def get_db():
    with _lock:
        return _load()


def save_db(db):
    with _lock:
        _save(db)


def next_id(items):
    return max([i.get('id', 0) for i in items] + [0]) + 1


def address_to_string(addr):
    """Собирает строку адреса из элементов ФИАС-подобного адреса."""
    parts = []
    if addr.get('indeks'):
        parts.append(addr['indeks'])
    tail = []
    if addr.get('kray_oblast'):
        tail.append(addr['kray_oblast'] + ',')
    if addr.get('rayon'):
        tail.append('р-н ' + addr['rayon'] + ',')
    if addr.get('gorod'):
        tail.append('г. ' + addr['gorod'] + ',')
    elif addr.get('naselenny_punkt'):
        tail.append(addr['naselenny_punkt'] + ',')
    if addr.get('ulica'):
        tail.append('ул. ' + addr['ulica'] + ',')
    if addr.get('dom'):
        s = 'д. ' + addr['dom']
        if addr.get('korpus'):
            s += ', к. ' + addr['korpus']
        if addr.get('kvartira'):
            s += ', кв. ' + addr['kvartira']
        tail.append(s)
    body = ' '.join(tail).replace(' ,', ',').strip().rstrip(',')
    head = ' '.join(p for p in parts if p)
    if head and body:
        return head + ', ' + body
    return head or body


def init_defaults(db):
    """Заполняет базовые справочники значениями по умолчанию, если они пусты."""
    c = db['catalogs']
    if not c['vidy_dokumentov']:
        c['vidy_dokumentov'] = [
            {"id": 1, "наименование": "Приказ"},
            {"id": 2, "наименование": "Контракт"},
            {"id": 3, "наименование": "Договор"},
            {"id": 4, "наименование": "АКТ сдачи-приемки оказанных услуг"},
            {"id": 5, "наименование": "Накладная"},
            {"id": 6, "наименование": "УПД"},
            {"id": 7, "наименование": "Счёт на оплату"},
        ]
    if not c['strany_proiskhozhdeniya']:
        c['strany_proiskhozhdeniya'] = [{"id": 1, "kod": "643", "naimovanie": "Россия"}]
    return db
