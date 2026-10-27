# -*- coding: utf-8 -*-
"""Резервное копирование и восстановление базы данных программы.

Хранилище резервных копий: каталог data/backups/ (рядом с db.json).
Формат копии: JSON-файл с метаданными (версия программы, дата, статистика).

Структура файла копии:
{
  "meta": {
     "app": "Приемка товаров — ГКУ КК УСЗН",
     "program_version": "1.3.0",     — версия программы, создавшей копию
     "schema_version": 1,            — версия структуры данных
     "created": "дд.мм.гггг чч:мм:сс",
     "comment": "...",
     "stats": {"поставщики": n, "номенклатура": n, ...}
  },
  "data": { constants / catalogs / documents }   — снимок db.json
}
"""
import json
import os
import re
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(BASE_DIR)
DATA_DIR = os.environ.get('PRIEMKA_DATA_DIR') or os.path.join(ROOT_DIR, 'data')
BACKUP_DIR = os.path.join(DATA_DIR, 'backups')

SCHEMA_VERSION = 1
MAX_KEEP = 50          # максимум хранимых автоматических копий

_TS_RE = re.compile(r'^backup_(\d{8}_\d{6})')


def _now_stamp():
    return datetime.now().strftime('%Y%m%d_%H%M%S')


def _fname(comment='', stamp=None):
    stamp = stamp or _now_stamp()
    safe = re.sub(r'[^0-9A-Za-zА-Яа-яЁё_.-]', '_', comment or '')[:40]
    name = 'backup_%s' % stamp
    if safe:
        name += '_' + safe
    return name + '.json'


def stats_of(db):
    c = db.get('catalogs', {})
    d = db.get('documents', {})
    return {
        'поставщики': len(c.get('postavshiki', [])),
        'виды документов': len(c.get('vidy_dokumentov', [])),
        'номенклатура': len(c.get('nomenklatura', [])),
        'ставки НДС': len(c.get('stavki_nds', [])),
        'страны': len(c.get('strany_proiskhozhdeniya', [])),
        'адресный классификатор (ФИАС)': len(c.get('fias', [])),
        'ед. измерения': len(c.get('ed_izmereniya', [])),
        'сотрудники': len(c.get('sotrudniki', [])),
        'приказы комиссии': len(d.get('pri_kazy_comissii', [])),
        'приёмки': len(d.get('priemki', [])),
    }


def create_backup(db, comment='', program_version=''):
    """Создаёт файл резервной копии; возвращает путь к нему."""
    os.makedirs(BACKUP_DIR, exist_ok=True)
    payload = {
        'meta': {
            'app': 'Приемка товаров — ГКУ КК УСЗН',
            'program_version': program_version,
            'schema_version': SCHEMA_VERSION,
            'created': datetime.now().strftime('%d.%m.%Y %H:%M:%S'),
            'created_iso': datetime.now().isoformat(timespec='seconds'),
            'comment': comment or '',
            'stats': stats_of(db),
        },
        'data': db,
    }
    path = os.path.join(BACKUP_DIR, _fname(comment))
    tmp = path + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)
    prune_backups()
    return path


def list_backups():
    """Список копий (новые первыми): имя, дата, размер, версия, комментарий, статистика."""
    if not os.path.isdir(BACKUP_DIR):
        return []
    items = []
    for fn in os.listdir(BACKUP_DIR):
        if not (fn.startswith('backup_') and fn.endswith('.json')):
            continue
        p = os.path.join(BACKUP_DIR, fn)
        row = {'file': fn, 'path': p, 'size': os.path.getsize(p),
               'date': datetime.fromtimestamp(os.path.getmtime(p)).strftime('%d.%m.%Y %H:%M'),
               'program_version': '', 'comment': '', 'stats': {}, 'schema_version': None}
        try:
            with open(p, 'r', encoding='utf-8') as f:
                head = json.load(f)
            meta = head.get('meta', {})
            row['program_version'] = meta.get('program_version', '')
            row['comment'] = meta.get('comment', '')
            row['stats'] = meta.get('stats', {})
            row['schema_version'] = meta.get('schema_version')
            row['created'] = meta.get('created', row['date'])
        except Exception as e:
            row['error'] = str(e)
        items.append(row)
    items.sort(key=lambda r: r['file'], reverse=True)
    return items


def read_backup(path):
    """Читает файл копии; возвращает dict или бросает ValueError с текстом ошибки."""
    if not os.path.isfile(path):
        raise ValueError('Файл копии не найден')
    with open(path, 'r', encoding='utf-8') as f:
        payload = json.load(f)
    if not isinstance(payload, dict) or 'data' not in payload:
        raise ValueError('Файл не является резервной копией программы')
    return payload


def validate_data(data):
    """Проверяет минимальную структуру снимка базы."""
    if not isinstance(data, dict):
        return False
    for key in ('constants', 'catalogs', 'documents'):
        if key not in data:
            return False
    c, d = data['catalogs'], data['documents']
    if not isinstance(c, dict) or not isinstance(d, dict):
        return False
    for k in ('postavshiki', 'nomenklatura', 'sotrudniki'):
        if k not in c:
            return False
    for k in ('priemki', 'pri_kazy_comissii'):
        if k not in d:
            return False
    return True


def delete_backup(path):
    if os.path.isfile(path):
        os.remove(path)
        return True
    return False


def prune_backups(keep=MAX_KEEP):
    """Удаляет старые автоматические копии, оставляя последние `keep`."""
    items = list_backups()
    if len(items) <= keep:
        return
    for row in items[keep:]:
        try:
            delete_backup(row['path'])
        except Exception:
            pass


def auto_backup_if_needed(db, program_version=''):
    """Создаёт копию при запуске, если сегодня ещё не было копий."""
    today = datetime.now().strftime('%Y%m%d')
    for row in list_backups():
        if _TS_RE.match(row['file']) and today in row['file']:
            return None
    return create_backup(db, comment='авто', program_version=program_version)
