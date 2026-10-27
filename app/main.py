# -*- coding: utf-8 -*-
"""Веб-программа приёмки товаров ГКУ КК УСЗН (Flask).

Запуск:  python run.py (настольное приложение, без браузера)
         python run.py --server  — сетевой режим (http://localhost:5000)
"""
import io
import json
import os
import re
import sys
import uuid
from datetime import date, datetime

import openpyxl
from flask import (Flask, jsonify, redirect, render_template, request,
                   send_file, url_for)

from app import database as db_api
from app.database import (FIAS_ADDRESS_TEMPLATE, address_to_string, get_db,
                          init_defaults, next_id, save_db)
from app import internet_loaders as net
from app.print_form import build_priemka_xlsx
from app.print_form_61n import build_akt_priemki_61n_xlsx
from app.excel_export import export_prikaz_xlsx, export_priemka_xlsx
from app.report_volume import build_volume_report, export_volume_xlsx
from app.report_nomenclature import (build_nomenclature_report,
                                     export_nomenclature_xlsx)
from app import backup as bak
from version import APP_VERSION, APP_NAME

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.environ.get('PRIEMKA_DATA_DIR') or os.path.join(ROOT_DIR, 'data')
EXPORT_DIR = os.path.join(DATA_DIR, 'exports')

app = Flask(__name__)


@app.context_processor
def _inject_version():
    return {'app_version': APP_VERSION, 'app_name': APP_NAME}

# Флаг настольного режима: устанавливается запускчиком desktop.py (без браузера)
DESKTOP_MODE = False


def _norm_date(value):
    """Возвращает дату в формате дд.мм.гггг из ISO (yyyy-mm-dd) или как есть."""
    if not value:
        return ''
    s = str(value).strip()
    m = re.match(r'^(\d{4})-(\d{2})-(\d{2})$', s)  # значение type="date"
    if m:
        return '%s.%s.%s' % (m.group(3), m.group(2), m.group(1))
    return s


def _save_dialog(default_name):
    """Настольный режим (без браузера): предлагает пользователю путь для
    сохранения файла через стандартный диалог «Сохранить как…».

    Возвращает выбранный полный путь либо None (отмена / диалог недоступен —
    вызывающий код сохранит файл в папку программы data/exports)."""
    try:
        import webview  # не экспортирует FileDialog — проверяем через create
        from webview import FileDialog
    except Exception:
        return None
    try:
        win = webview.windows[0] if getattr(webview, 'windows', None) else None
        if win is None:
            return None
        res = win.create_file_dialog(
            FileDialog.SAVE, save_filename=default_name, extensions=['*.xlsx'])
    except Exception:
        return None
    if not res:
        return None  # пользователь отменил диалог
    path = res[0] if isinstance(res, (list, tuple)) else str(res)
    if not path.lower().endswith('.xlsx'):
        path += '.xlsx'
    return path

# Кнопка «Назад»: сервер запоминает предыдущую страницу в cookie _back
# (ставится после каждого перехода по ссылке) и отдаёт её в заголовке X-Back.
def _error_page(msg, code=500):
    """Понятная страница ошибки с кнопкой возврата (вместо пустого экрана)."""
    body = ('<html><head><meta charset="utf-8"><title>Ошибка</title></head>'
            '<body style="font-family:Segoe UI,Arial,sans-serif;padding:40px">'
            '<h2 style="color:#b00">&#9888; Ошибка</h2><p>%s</p>'
            '<p><a href="/" style="display:inline-block;padding:8px 18px;'
            'background:#2e7d32;color:#fff;border-radius:6px;'
            'text-decoration:none">&#127968; На главную</a>&nbsp;'
            '<a href="javascript:history.back()" style="display:inline-block;'
            'padding:8px 18px;background:#607d8b;color:#fff;border-radius:6px;'
            'text-decoration:none">&#8592; Назад</a></p></body></html>' % msg)
    return body, code


@app.after_request
def _back_header(resp):
    # Печатные формы в настольном режиме отдают JSON для fetch() из страницы —
    # не подставляем X-Back, чтобы cookie предыдущей страницы не перезаписывался.
    if DESKTOP_MODE and request.path.endswith(('/print', '/print61n')):
        return resp
    if resp.status_code == 200 and 'X-Back' not in resp.headers:
        back = request.cookies.get('_back')
        if back and back != request.path and not back.startswith('/static'):
            resp.headers['X-Back'] = back
    return resp


@app.context_processor
def _inject_back():
    ref = request.headers.get('Referer', '')
    back = None
    if ref:
        try:
            from urllib.parse import urlparse
            p = urlparse(ref).path or '/'
            if p != request.path:  # не ссылку на саму себя
                back = p
        except Exception:
            back = None
    if back is None:
        c = request.cookies.get('_back')
        if c and c != request.path and not c.startswith('/static'):
            back = c
    return {'back_url': back}


def _db():
    d = get_db()
    init_defaults(d)
    return d


# ---------------------------------------------------------------- страница ---
@app.before_request
def _remember_back():
    """Запоминаем предыдущую страницу (referer) — для кнопки «← Назад»,
    которая нужна в т.ч. в pywebview, где нет кнопки назад браузера."""
    ref = request.headers.get('Referer', '')
    if not ref:
        return
    try:
        from urllib.parse import urlparse
        p = urlparse(ref).path or '/'
    except Exception:
        return
    if p != request.path and not p.startswith('/static'):
        request._back_path = p


@app.after_request
def _back_cookie(resp):
    p = getattr(request, '_back_path', None)
    if p:
        resp.set_cookie('_back', p, max_age=3600, samesite='Lax')
    return resp

@app.route('/')
def index():
    d = _db()
    return render_template('index.html',
                           priemki=d['documents']['priemki'],
                           prikazy=d['documents']['pri_kazy_comissii'],
                           stats={
                               'postavshiki': len(d['catalogs']['postavshiki']),
                               'nomenklatura': len(d['catalogs']['nomenklatura']),
                               'sotrudniki': len(d['catalogs']['sotrudniki']),
                               'strany': len(d['catalogs']['strany_proiskhozhdeniya']),
                               'ed_izm': len(d['catalogs']['ed_izmereniya']),
                               'fias': len(d['catalogs']['fias']),
                           })


@app.route('/constants')
def constants_page():
    d = _db()
    uch = d['constants']['uchrezhdenie']
    sotrudniki = d['catalogs']['sotrudniki']
    return render_template('constants.html', uch=uch, sotrudniki=sotrudniki,
                           fias_keys=list(FIAS_ADDRESS_TEMPLATE.keys()))


@app.route('/constants/save', methods=['POST'])
def constants_save():
    d = _db()
    uch = d['constants']['uchrezhdenie']
    for k in ['учреждение_получатель', 'обособленное_подразделение',
              'структурное_подразделение', 'главный_администратор_доходов',
              'наименование_бюджета', 'валюта', 'адрес_грузополучателя',
              'inn', 'kpp', 'ogrn', 'rukovoditel']:
        if k in request.form:
            uch[k] = request.form[k].strip()
    for addr_key in ['adres_zakazchika', 'mesto_postavki']:
        addr = uch.get(addr_key) or dict(FIAS_ADDRESS_TEMPLATE)
        for f in FIAS_ADDRESS_TEMPLATE:
            name = '%s__%s' % (addr_key, f)
            if name in request.form:
                addr[f] = request.form[name].strip()
        uch[addr_key] = addr
        uch[addr_key + '_str'] = address_to_string(addr)
    save_db(d)
    return redirect(url_for('constants_page'))


CATALOG_TITLES = {
    'postavshiki': 'Поставщик',
    'vidy_dokumentov': 'Виды документов',
    'nomenklatura': 'Номенклатура',
    'stavki_nds': 'Ставки НДС',
    'strany_proiskhozhdeniya': 'Страна происхождения',
    'fias': 'Адресный классификатор (ФИАС)',
    'ed_izmereniya': 'Ед. измерения (ОКЕИ)',
    'sotrudniki': 'Сотрудники',
}

NET_CATALOGS = {'strany_proiskhozhdeniya', 'ed_izmereniya', 'fias'}


@app.route('/catalog/<name>')
def catalog_page(name):
    if name not in CATALOG_TITLES:
        abort404 = True
        return "Справочник не найден", 404
    d = _db()
    items = d['catalogs'][name]
    refs = None
    if name == 'nomenklatura':
        # справочники-источники для выпадающих списков полей номенклатуры
        refs = {
            'strany': d['catalogs']['strany_proiskhozhdeniya'],
            'ed_izm': d['catalogs']['ed_izmereniya'],
            'stavki': d['catalogs']['stavki_nds'],
        }
    return render_template('catalog.html', name=name,
                           title=CATALOG_TITLES[name], items=items,
                           is_net=name in NET_CATALOGS, refs=refs)


@app.route('/catalog/<name>/add', methods=['POST'])
def catalog_add(name):
    d = _db()
    data = json.loads(request.data or '{}')
    items = d['catalogs'][name]
    if name == 'stavki_nds':
        items.append(str(data.get('value', '')).strip())
    else:
        row = dict(data)
        row['id'] = next_id(items) if items and isinstance(items[0], dict) else len(items) + 1
        items.append(row)
    save_db(d)
    return jsonify({'ok': True})


@app.route('/catalog/<name>/update', methods=['POST'])
def catalog_update(name):
    d = _db()
    data = json.loads(request.data or '{}')
    idx = data.pop('_index', None)
    items = d['catalogs'][name]
    if idx is not None and 0 <= int(idx) < len(items):
        i = int(idx)
        if name == 'stavki_nds':
            items[i] = str(data.get('value', '')).strip()
        else:
            row = dict(data)
            # синхронизация связанных полей при выборе из справочников
            if name == 'nomenklatura':
                for s in d['catalogs']['strany_proiskhozhdeniya']:
                    if isinstance(s, dict) and str(s.get('kod', '')) == str(row.get('strana_kod', '')):
                        row['strana_name'] = s.get('naimovanie', '')
                        break
                for e in d['catalogs']['ed_izmereniya']:
                    if isinstance(e, dict) and str(e.get('kod', '')) == str(row.get('oei_kod', '')):
                        row['ed_izm'] = e.get('obozachenie') or e.get('naimovanie', '')
                        break
                for v in d['catalogs']['stavki_nds']:
                    sv = v if isinstance(v, str) else (v.get('value', '') if isinstance(v, dict) else '')
                    if str(sv).strip() == str(row.get('stavka_nds', '')).strip():
                        row['stavka_nds'] = str(sv).strip()
                        break
            row['id'] = items[i].get('id', i + 1) if isinstance(items[i], dict) else i + 1
            items[i] = row
    save_db(d)
    return jsonify({'ok': True})


@app.route('/catalog/<name>/delete', methods=['POST'])
def catalog_delete(name):
    d = _db()
    data = json.loads(request.data or '{}')
    items = d['catalogs'][name]
    # массовое удаление: список индексов indexes (или один _index)
    raw = data.get('indexes')
    if raw is None:
        raw = [data.get('_index', -1)]
    try:
        idxs = sorted({int(i) for i in raw}, reverse=True)
    except (TypeError, ValueError):
        return jsonify({'ok': False, 'msg': 'Некорректные индексы'}), 400
    removed = 0
    for idx in idxs:
        if 0 <= idx < len(items):
            del items[idx]
            removed += 1
    save_db(d)
    return jsonify({'ok': True, 'removed': removed})


@app.route('/catalog/<name>/load_net', methods=['POST'])
def catalog_load_net(name):
    d = _db()
    msg = ''
    if name == 'strany_proiskhozhdeniya':
        rows, msg = net.load_countries()
        d['catalogs'][name] = [dict(r, id=i + 1) for i, r in enumerate(rows)]
    elif name == 'ed_izmereniya':
        rows, msg = net.load_okoi()
        d['catalogs'][name] = [dict(r, id=i + 1) for i, r in enumerate(rows)]
    elif name == 'fias':
        rows, msg = net.load_fias_sample()
        d['catalogs'][name] = [dict(r, id=i + 1) for i, r in enumerate(rows)]
    else:
        return jsonify({'ok': False, 'msg': 'Для этого справочника загрузка из сети не предусмотрена'})
    save_db(d)
    return jsonify({'ok': True, 'msg': msg, 'count': len(d['catalogs'][name])})


# ------------------------------------------------------------ импорт xlsx ---
@app.route('/catalog/<name>/import', methods=['POST'])
def catalog_import(name):
    f = request.files.get('file')
    if not f:
        return jsonify({'ok': False, 'msg': 'Файл не выбран'}), 400
    wb = openpyxl.load_workbook(io.BytesIO(f.read()), data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    d = _db()
    added = 0
    if name == 'postavshiki':
        # шапка в 1 строке
        header, body = rows[0], rows[1:]
        keys = ['polnoe_naimovanie', 'kratkoe_naimovanie', 'adres', 'ogrn',
                'inn', 'kpp', 'dolzhnost', 'fio', 'dogovor', 'data_zakl', 'ikz']
        items = d['catalogs'][name]
        for r in body:
            if not r or not str(r[0] or '').strip():
                continue
            rec = {}
            for i, k in enumerate(keys):
                v = r[i] if i < len(r) else ''
                if isinstance(v, (date, datetime)):
                    v = v.strftime('%d.%m.%Y')
                rec[k] = str(v).strip() if v is not None else ''
            rec['id'] = next_id(items)
            items.append(rec)
            added += 1
    elif name == 'nomenklatura':
        # данные с 4-й строки (3 уровня шапки)
        items = d['catalogs'][name]
        strany = {s.get('naimovanie', '').lower(): s.get('kod', '')
                  for s in d['catalogs']['strany_proiskhozhdeniya'] if isinstance(s, dict)}
        for r in rows[3:]:
            if not r or not (str(r[0] or '').strip() or str(r[2] or '').strip()):
                continue
            kod_tovara, kratkoe, naimovanie = r[0], r[1], r[2]
            strana_kod, strana_name = r[3], r[4]
            oei_kod, ed_izm = r[5], r[6]
            stavka = r[7] if len(r) > 7 else ''
            deklaraciya = r[8] if len(r) > 8 else ''
            sertifikat = r[9] if len(r) > 9 else ''
            if isinstance(stavka, float):
                stavka = ('%g' % (stavka * 100)) + '%'
            rec = {
                'id': next_id(items),
                'kod_tovara': str(kod_tovara or '').strip(),
                'kratkoe_naimovanie': str(kratkoe or '').strip(),
                'naimovanie': str(naimovanie or '').strip(),
                'strana_kod': str(strana_kod or '').strip(),
                'strana_name': str(strana_name or '').strip(),
                'oei_kod': str(oei_kod or '').strip(),
                'ed_izm': str(ed_izm or '').strip(),
                'stavka_nds': str(stavka or '').strip(),
                'nom_deklaracii': str(deklaraciya or '').strip(),
                'nom_certifikata': str(sertifikat or '').strip(),
            }
            if rec['strana_name'] and not rec['strana_kod']:
                rec['strana_kod'] = strany.get(rec['strana_name'].lower(), '')
            items.append(rec)
            added += 1
    elif name == 'sotrudniki':
        items = d['catalogs'][name]
        body = rows[1:] if rows and str(rows[0][0] or '').lower().startswith(('долж', 'фио')) else rows
        for r in body:
            if not r or not any(str(x or '').strip() for x in r):
                continue
            a = str(r[0] or '').strip()
            b = str(r[1] or '').strip() if len(r) > 1 else ''
            low_a = a.lower()
            if low_a.startswith('долж') or low_a.startswith('фио'):
                continue
            # определяем порядок: ФИО обычно длиннее и содержит пробелы/отчества
            fio, dolzhnost = (b, a) if (len(a.split()) <= 2 and len(b.split()) >= 2) else (a, b)
            if not fio and not dolzhnost:
                continue
            items.append({'id': next_id(items), 'dolzhnost': dolzhnost, 'fio': fio})
            added += 1
    elif name == 'fias':
        items = d['catalogs'][name]
        header = [str(x or '').strip().lower() for x in rows[0]]
        def col(*names):
            for n in names:
                if n in header:
                    return header.index(n)
            return None
        ci = {
            'kray_oblast': col('край/область', 'регион', 'kray_oblast'),
            'rayon': col('район', 'rayon'),
            'gorod': col('город', 'gorod'),
            'naselenny_punkt': col('населенный пункт', 'naselenny_punkt'),
            'ulica': col('ул.', 'улица', 'ulica'),
            'dom': col('д.', 'дом', 'dom'),
            'indeks': col('индекс', 'indeks'),
        }
        for r in rows[1:]:
            rec = {'id': next_id(items) if items else 1}
            for k, idx in ci.items():
                rec[k] = str(r[idx]).strip() if idx is not None and idx < len(r) and r[idx] is not None else ''
            if any(rec.get(k) for k in ci):
                items.append(rec)
                added += 1
    else:
        return jsonify({'ok': False, 'msg': 'Импорт для этого справочника не поддерживается'}), 400
    save_db(d)
    return jsonify({'ok': True, 'added': added})


# ------------------------------------------------------------- документы ----
@app.route('/dokument/prikaz')
def prikazy_list():
    d = _db()
    return render_template('prikazy.html', items=d['documents']['pri_kazy_comissii'])


@app.route('/dokument/prikaz/new')
def prikaz_new():
    d = _db()
    return render_template('prikaz_form.html', item=None,
                           sotrudniki=d['catalogs']['sotrudniki'],
                           vidy=d['catalogs']['vidy_dokumentov'])


@app.route('/dokument/prikaz/<int:pid>')
def prikaz_edit(pid):
    d = _db()
    item = next((p for p in d['documents']['pri_kazy_comissii'] if p['id'] == pid), None)
    if item is None:
        return "Приказ не найден", 404
    return render_template('prikaz_form.html', item=item,
                           sotrudniki=d['catalogs']['sotrudniki'],
                           vidy=d['catalogs']['vidy_dokumentov'])


@app.route('/dokument/prikaz/save', methods=['POST'])
def prikaz_save():
    d = _db()
    items = d['documents']['pri_kazy_comissii']
    pid = request.form.get('id', type=int)
    predsedatel = request.form.get('predsedatel', '')
    members = request.form.getlist('chleny[]')
    item = {
        'id': pid or next_id(items),
        'nomer': request.form.get('nomer', ''),
        'data': _norm_date(request.form.get('data', '')),
        'nanimovanie': request.form.get('nanimovanie', ''),
        'predsedatel': predsedatel,
        'chleny': [m for m in members if m],
    }
    if pid:
        items[[i['id'] for i in items].index(pid)] = item
    else:
        items.append(item)
    save_db(d)
    return redirect(url_for('prikazy_list'))


@app.route('/dokument/prikaz/delete', methods=['POST'])
def prikaz_delete():
    d = _db()
    pid = int(json.loads(request.data)['id'])
    d['documents']['pri_kazy_comissii'] = [
        p for p in d['documents']['pri_kazy_comissii'] if p['id'] != pid]
    save_db(d)
    return jsonify({'ok': True})


@app.route('/dokument/prikaz/<int:pid>/excel')
def prikaz_excel(pid):
    """Выгрузка приказа комиссии в Excel (.xlsx)."""
    d = _db()
    item = next((p for p in d['documents']['pri_kazy_comissii'] if p['id'] == pid), None)
    if item is None:
        return 'Приказ не найден', 404
    os.makedirs(EXPORT_DIR, exist_ok=True)
    fname = 'Prikaz_%s_%s.xlsx' % (pid, uuid.uuid4().hex[:8])
    out = os.path.join(EXPORT_DIR, fname)
    try:
        export_prikaz_xlsx(item, out)
    except Exception as e:
        return 'Ошибка выгрузки в Excel: %s' % e, 500
    if DESKTOP_MODE:
        default_name = 'Приказ_%s.xlsx' % (item.get('nomer') or pid)
        target = _save_dialog(default_name)  # предложить пользователю путь
        if not target:
            target = os.path.join(EXPORT_DIR, default_name)
        try:
            import shutil
            shutil.copyfile(out, target)
        except Exception:
            target = out
        try:
            os.startfile(target)
        except Exception:
            pass
        return jsonify({'ok': True, 'file': target})
    return send_file(out, as_attachment=True,
                     download_name='Приказ_комиссии_%s.xlsx' % (item.get('nomer') or pid))


@app.route('/api/sotrudnik/<path:fio>')
def api_sotrudnik(fio):
    d = _db()
    for s in d['catalogs']['sotrudniki']:
        if s.get('fio') == fio:
            return jsonify(s)
    return jsonify({'fio': fio, 'dolzhnost': ''})


# ------------------------------------------------------------- приёмка ------
@app.route('/dokument/priemka')
def priemki_list():
    d = _db()
    return render_template('priemki.html', items=d['documents']['priemki'])


@app.route('/dokument/priemka/new')
def priemka_new():
    return _priemka_form(None)


@app.route('/dokument/priemka/<int:pid>')
def priemka_edit(pid):
    d = _db()
    item = next((p for p in d['documents']['priemki'] if p['id'] == pid), None)
    if item is None:
        return "Приёмка не найдена", 404
    return _priemka_form(item)


def _priemka_form(item):
    d = _db()
    return render_template('priemka_form.html', item=item,
                           db=d,
                           uch=d['constants']['uchrezhdenie'])


@app.route('/dokument/priemka/save', methods=['POST'])
def priemka_save():
    payload = json.loads(request.data)
    d = _db()
    items = d['documents']['priemki']
    pid = payload.get('id')
    payload['id'] = pid or next_id(items)
    # даты из календаря (type="date", ISO yyyy-mm-dd) — к формату дд.мм.гггг
    for k in ['дата', 'dok_osnovanie_data', 'osnovanie_data']:
        if k in payload:
            payload[k] = _norm_date(payload[k])
    # раздел 4: данные номенклатуры всегда берутся из справочника
    # (сотрудник вносит только количество и цену); ключ — код + наименование
    nom = {}
    for n in d['catalogs']['nomenklatura']:
        if isinstance(n, dict):
            nom[(str(n.get('kod_tovara', '')), str(n.get('naimovanie', '')))] = n
    nom_by_kod = {str(n.get('kod_tovara', '')): n for n in nom.values()}
    for i, poz in enumerate(payload.get('pozicii', [])):
        poz['nom'] = i + 1
        # поля «кол-во» и «цена» — не более 7 знаков (цифр)
        for knum in ('kolichestvo', 'cena'):
            val = str(poz.get(knum, '') or '').strip().replace(',', '.')
            if re.fullmatch(r'-?\d+(\.\d+)?', val):
                digits = val.lstrip('-').replace('.', '')
                if len(digits) > 7:
                    intp, _, frac = val.partition('.')
                    sign = '-' if intp.startswith('-') else ''
                    intp_d = intp.lstrip('-')
                    intp_d = intp_d[:7]                 # целая часть — до 7 цифр
                    frac = frac[:max(0, 7 - len(intp_d))]  # всего цифр <= 7
                    val = sign + intp_d + ('.' + frac if frac else '')
                    poz[knum] = val
        n = nom.get((str(poz.get('kod_tovara', '')), str(poz.get('naimovanie', '')))) \
            or nom_by_kod.get(str(poz.get('kod_tovara', '')))
        if n:
            for k in ['kod_tovara', 'kratkoe_naimovanie', 'naimovanie',
                      'strana_kod', 'strana_name', 'oei_kod',
                      'ed_izm', 'stavka_nds', 'nom_deklaracii', 'nom_certifikata']:
                poz[k] = n.get(k, '')
        poz['stavka_nds_text'] = poz.get('stavka_nds', '')
    if pid:
        items[[i['id'] for i in items].index(pid)] = payload
    else:
        items.append(payload)
    save_db(d)
    return jsonify({'ok': True, 'id': payload['id']})


@app.route('/dokument/priemka/delete', methods=['POST'])
def priemka_delete():
    d = _db()
    pid = int(json.loads(request.data)['id'])
    d['documents']['priemki'] = [p for p in d['documents']['priemki'] if p['id'] != pid]
    save_db(d)
    return jsonify({'ok': True})


@app.errorhandler(404)
def _err404(e):
    return _error_page("Страница или документ не найдены.", 404)


@app.errorhandler(500)
def _err500(e):
    return _error_page("Внутренняя ошибка программы: %s" % e, 500)


@app.route('/dokument/priemka/<int:pid>/print')
def priemka_print(pid):
    d = _db()
    item = next((p for p in d['documents']['priemki'] if p['id'] == pid), None)
    if item is None:
        return _error_page("Приёмка №%s не найдена." % pid, 404)
    os.makedirs(EXPORT_DIR, exist_ok=True)
    fname = 'priemka_%s_%s.xlsx' % (pid, uuid.uuid4().hex[:8])
    out = os.path.join(EXPORT_DIR, fname)
    try:
        build_priemka_xlsx(item, d['constants'], d['catalogs'], out)
    except Exception as e:
        return _error_page("Ошибка формирования печатной формы: %s" % e)
    # В настольном режиме (без браузера) предлагаем пользователю путь для
    # сохранения через диалог «Сохранить как…»; при отмене — папка программы.
    if DESKTOP_MODE:
        default_name = 'Приемка_форма0510452_%s.xlsx' % pid
        target = _save_dialog(default_name)  # предложить пользователю путь
        if not target:
            target = os.path.join(EXPORT_DIR, default_name)
        try:
            import shutil
            shutil.copyfile(out, target)
        except Exception:
            target = out
        try:
            os.startfile(target)  # Windows: открыть в Excel/LibreOffice
        except Exception:
            pass
        resp = jsonify({'ok': True, 'file': target})
        resp.headers['Cache-Control'] = 'no-store'
        return resp
    return send_file(out, as_attachment=True,
                     download_name='Приемка_форма0510452.xlsx')


@app.route('/dokument/priemka/<int:pid>/print61n')
def priemka_print_61n(pid):
    """Печатная форма «Акт приемки товаров, работ, услуг» (ОКУД 0510452),
    утверждённая Приказом Минфина России от 15.04.2021 № 61н."""
    d = _db()
    item = next((p for p in d['documents']['priemki'] if p['id'] == pid), None)
    if item is None:
        return _error_page("Приёмка №%s не найдена." % pid, 404)
    os.makedirs(EXPORT_DIR, exist_ok=True)
    fname = 'priemka_61n_%s_%s.xlsx' % (pid, uuid.uuid4().hex[:8])
    out = os.path.join(EXPORT_DIR, fname)
    try:
        build_akt_priemki_61n_xlsx(item, d['constants'], out)
    except Exception as e:
        return _error_page("Ошибка формирования акта по форме 61н: %s" % e)
    if DESKTOP_MODE:
        default_name = 'Акт_приемки_ОКУД0510452_61н_%s.xlsx' % pid
        target = _save_dialog(default_name)  # предложить пользователю путь
        if not target:
            target = os.path.join(EXPORT_DIR, default_name)
        try:
            import shutil
            shutil.copyfile(out, target)
        except Exception:
            target = out
        try:
            os.startfile(target)  # Windows: открыть в Excel/LibreOffice
        except Exception:
            pass
        resp = jsonify({'ok': True, 'file': target})
        resp.headers['Cache-Control'] = 'no-store'
        return resp
    return send_file(out, as_attachment=True,
                     download_name='Акт_приемки_ОКУД0510452_61н.xlsx')


@app.route('/dokument/priemka/<int:pid>/excel')
def priemka_excel(pid):
    """Выгрузка документа «Приёмка» в Excel (.xlsx) — реквизиты, позиции, итоги."""
    d = _db()
    item = next((p for p in d['documents']['priemki'] if p['id'] == pid), None)
    if item is None:
        return _error_page('Приёмка №%s не найдена.' % pid, 404)
    os.makedirs(EXPORT_DIR, exist_ok=True)
    fname = 'Priemka_excel_%s_%s.xlsx' % (pid, uuid.uuid4().hex[:8])
    out = os.path.join(EXPORT_DIR, fname)
    try:
        export_priemka_xlsx(item, d['constants'], out)
    except Exception as e:
        return 'Ошибка выгрузки в Excel: %s' % e, 500
    if DESKTOP_MODE:
        default_name = 'Приемка_Excel_%s.xlsx' % (item.get('номер') or pid)
        target = _save_dialog(default_name)  # предложить пользователю путь
        if not target:  # отмена диалога / диалог недоступен — папка программы
            target = os.path.join(EXPORT_DIR, default_name)
        try:
            import shutil
            shutil.copyfile(out, target)
        except Exception:
            target = out
        try:
            os.startfile(target)
        except Exception:
            pass
        return jsonify({'ok': True, 'file': target})
    return send_file(out, as_attachment=True,
                     download_name='Приемка_%s.xlsx' % (item.get('номер') or pid))


# ---------------------------------------------------------------- API -------
@app.route('/api/catalogs')
def api_catalogs():
    d = _db()
    return jsonify(d['catalogs'])


# ------------------------------------------------- отчёт по объёму ----------
@app.route('/otchet/obem')
def otchet_obem():
    """Отчёт по объёму принятых товаров (фильтры: период, поставщик)."""
    d = _db()
    date_from = request.args.get('date_from', '').strip()
    date_to = request.args.get('date_to', '').strip()
    postavshik = request.args.get('postavshik', '').strip()
    report = build_volume_report(d['documents']['priemki'],
                                 date_from=date_from, date_to=date_to,
                                 postavshik=postavshik)
    return render_template('otchet_obem.html', report=report,
                           date_from=date_from, date_to=date_to,
                           postavshik=postavshik,
                           postavshiki=d['catalogs']['postavshiki'])


@app.route('/otchet/obem/export')
def otchet_obem_export():
    """Выгрузка отчёта по объёму принятых товаров в XLSX."""
    d = _db()
    report = build_volume_report(
        d['documents']['priemki'],
        date_from=request.args.get('date_from', '').strip(),
        date_to=request.args.get('date_to', '').strip(),
        postavshik=request.args.get('postavshik', '').strip())
    os.makedirs(EXPORT_DIR, exist_ok=True)
    out = os.path.join(EXPORT_DIR, 'otchet_obem_%s.xlsx' % uuid.uuid4().hex[:8])
    export_volume_xlsx(report, out)
    if DESKTOP_MODE:
        saved = os.path.join(EXPORT_DIR, 'Отчет_обем_принятых_товаров.xlsx')
        try:
            import shutil
            shutil.copyfile(out, saved)
        except Exception:
            saved = out
        try:
            os.startfile(saved)  # Windows: открыть в редакторе таблиц
        except Exception:
            pass
        return jsonify({'ok': True, 'file': saved})
    return send_file(out, as_attachment=True,
                     download_name='Отчет_обем_принятых_товаров.xlsx')


# ------------------------ отчёт: количество номенклатуры по поставщикам ----
@app.route('/otchet/nomenklatura')
def otchet_nomenklatura():
    """Отчёт «Количество номенклатуры за период по поставщикам»."""
    d = _db()
    date_from = request.args.get('date_from', '').strip()
    date_to = request.args.get('date_to', '').strip()
    postavshik = request.args.get('postavshik', '').strip()
    group_by = request.args.get('group_by', 'dokument').strip()
    report = build_nomenclature_report(d['documents']['priemki'],
                                       date_from=date_from, date_to=date_to,
                                       postavshik=postavshik,
                                       group_by=group_by)
    return render_template('otchet_nomenklatura.html', report=report,
                           date_from=date_from, date_to=date_to,
                           postavshik=postavshik, group_by=group_by,
                           postavshiki=d['catalogs']['postavshiki'])


@app.route('/otchet/nomenklatura/export')
def otchet_nomenklatura_export():
    """Выгрузка отчёта о количестве номенклатуры в XLSX."""
    d = _db()
    report = build_nomenclature_report(
        d['documents']['priemki'],
        date_from=request.args.get('date_from', '').strip(),
        date_to=request.args.get('date_to', '').strip(),
        postavshik=request.args.get('postavshik', '').strip(),
        group_by=request.args.get('group_by', 'dokument').strip())
    os.makedirs(EXPORT_DIR, exist_ok=True)
    out = os.path.join(EXPORT_DIR, 'otchet_nomenklatura_%s.xlsx'
                       % uuid.uuid4().hex[:8])
    export_nomenclature_xlsx(report, out)
    if DESKTOP_MODE:
        saved = os.path.join(EXPORT_DIR,
                             'Отчет_количество_номенклатуры_по_поставщикам.xlsx')
        try:
            import shutil
            shutil.copyfile(out, saved)
        except Exception:
            saved = out
        try:
            os.startfile(saved)  # Windows: открыть в редакторе таблиц
        except Exception:
            pass
        return jsonify({'ok': True, 'file': saved})
    return send_file(out, as_attachment=True,
                     download_name='Отчет_количество_номенклатуры_по_поставщикам.xlsx')


@app.route('/api/constants')
def api_constants():
    d = _db()
    return jsonify(d['constants'])


# ------------------------------------------- резервное копирование ----------
@app.route('/about')
def about_page():
    """О программе и история версий."""
    changelog = ''
    if getattr(sys, 'frozen', False):
        p = os.path.join(getattr(sys, '_MEIPASS', ROOT_DIR), 'CHANGELOG.md')
    else:
        p = os.path.join(ROOT_DIR, 'CHANGELOG.md')
    if os.path.isfile(p):
        with open(p, 'r', encoding='utf-8') as f:
            changelog = f.read()
    return render_template('about.html', version=APP_VERSION,
                           app_name=APP_NAME, changelog=changelog)


@app.route('/backup')
def backup_page():
    """Окно резервного копирования и восстановления данных."""
    d = _db()
    copies = bak.list_backups()
    cur_stats = bak.stats_of(d)
    # автокопия при первом запуске за день
    try:
        if bak.auto_backup_if_needed(d, APP_VERSION):
            copies = bak.list_backups()
    except Exception:
        pass
    return render_template('backup.html', backups=copies,
                           cur_stats=cur_stats, version=APP_VERSION,
                           backup_dir=bak.BACKUP_DIR)


@app.route('/backup/create', methods=['POST'])
def backup_create():
    d = _db()
    if request.is_json:
        comment = request.json.get('comment', '')
    else:
        comment = request.form.get('comment', '')
    try:
        path = bak.create_backup(d, comment=comment, program_version=APP_VERSION)
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500
    if DESKTOP_MODE:
        try:
            os.startfile(os.path.dirname(path))  # открыть папку с копиями
        except Exception:
            pass
        return jsonify({'ok': True, 'file': path})
    # веб-режим: копия уже сохранена в data/backups, дополнительно отдаём файл
    return jsonify({'ok': True, 'file': path,
                    'download': url_for('backup_download', name=os.path.basename(path))})


@app.route('/backup/download/<path:name>')
def backup_download(name):
    """Скачивание файла копии на другой компьютер/носитель."""
    path = os.path.join(bak.BACKUP_DIR, os.path.basename(name))
    if not os.path.isfile(path):
        return 'Копия не найдена', 404
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))


@app.route('/backup/restore', methods=['POST'])
def backup_restore():
    """Восстановление из выбранной копии (с предварительной страховочной копией)."""
    name = (request.json.get('file') if request.is_json
            else request.form.get('file')) or ''
    name = os.path.basename(name)  # защита от path traversal
    path = os.path.join(bak.BACKUP_DIR, name)
    if not name or not os.path.isfile(path):
        return jsonify({'ok': False, 'error': 'Копия не найдена'}), 404
    try:
        payload = bak.read_backup(path)
    except Exception as e:
        return jsonify({'ok': False, 'error': 'Не удалось прочитать копию: %s' % e}), 400
    data = payload.get('data')
    if not bak.validate_data(data):
        return jsonify({'ok': False,
                        'error': 'Структура копии не соответствует базе программы'}), 400
    # страховочная копия текущего состояния перед восстановлением
    try:
        bak.create_backup(get_db(), comment='перед восстановлением',
                          program_version=APP_VERSION)
    except Exception:
        pass
    save_db(data)
    init_defaults(data)
    save_db(data)
    return jsonify({'ok': True,
                    'restored_from': name,
                    'version': payload.get('meta', {}).get('program_version', '')})


@app.route('/backup/upload', methods=['POST'])
def backup_upload():
    """Загрузка файла копии с диска и восстановление из него."""
    f = request.files.get('file')
    if not f or not f.filename.lower().endswith('.json'):
        return jsonify({'ok': False, 'error': 'Выберите файл копии .json'}), 400
    try:
        payload = json.load(io.StringIO(f.read().decode('utf-8')))
    except Exception as e:
        return jsonify({'ok': False, 'error': 'Файл повреждён или не является копией: %s' % e}), 400
    data = payload.get('data') if isinstance(payload, dict) else None
    if not bak.validate_data(data):
        return jsonify({'ok': False, 'error': 'Структура файла не соответствует базе программы'}), 400
    try:
        bak.create_backup(get_db(), comment='перед восстановлением (загрузка)',
                          program_version=APP_VERSION)
    except Exception:
        pass
    save_db(data)
    init_defaults(data)
    save_db(data)
    return jsonify({'ok': True,
                    'version': payload.get('meta', {}).get('program_version', '')})


@app.route('/backup/delete', methods=['POST'])
def backup_delete():
    name = os.path.basename((request.json.get('file') if request.is_json
                             else request.form.get('file')) or '')
    path = os.path.join(bak.BACKUP_DIR, name)
    ok = bak.delete_backup(path)
    if not ok:
        return jsonify({'ok': False, 'error': 'Файл не найден'}), 404
    return jsonify({'ok': True})


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False)
