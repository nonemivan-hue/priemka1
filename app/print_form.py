# -*- coding: utf-8 -*-
"""Печатная форма приёмки (форма 0510452).

«Приложение к распоряжению ректора № 3 от 22.01.2024г.», А4, альбомная.
Формируется по эталонным файлам Раздел1.xlsx (стр.1 — основной раздел)
и Остальное.xlsx (стр.2_3 — сведения о транспортировке/приёмке и пр.),
которые служат шаблонами: в них заполняются поля, выделенные жёлтым.
"""
import os
import re
import sys
from datetime import date, datetime

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import column_index_from_string

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _res(name):
    """Путь к ресурсу: внутри собранного EXE — каталог sys._MEIPASS.

    В сборке шаблоны лежат под ASCII-именами Razdel1.xlsx / Ostalnoe.xlsx
    (кириллические имена файлов при упаковке иногда ломаются).
    """
    if getattr(sys, 'frozen', False):
        base = getattr(sys, '_MEIPASS', ROOT_DIR)
        for alt in (name, {'Раздел1.xlsx': 'Razdel1.xlsx',
                           'Остальное.xlsx': 'Ostalnoe.xlsx'}.get(name, name)):
            p = os.path.join(base, alt)
            if os.path.isfile(p):
                return p
        return os.path.join(base, name)
    return os.path.join(ROOT_DIR, name)


TEMPLATE_STR1 = _res('Раздел1.xlsx')
TEMPLATE_STR23 = _res('Остальное.xlsx')

YELLOW = PatternFill('solid', fgColor='FFFFFF00')  # подсветка заполненных полей (жёлтый, как в шаблонах)


def _parse_date(value):
    if not value:
        return None
    if isinstance(value, (date, datetime)):
        return value
    m = re.match(r'(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})', str(value).strip())
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100:
            y += 2000
        try:
            return date(y, mo, d)
        except ValueError:
            return None
    return None


def _anchor(ws, coord):
    """Возвращает верхнюю левую ячейку объединённого диапазона, содержащего coord.

    В шаблонах формы многие поля — объединённые ячейки; запись в любую
    не-первую ячейку слияния вызывает "AttributeError: 'MergedCell' object
    attribute 'value' is read-only". Перенаправляем запись в первую ячейку.
    """
    from openpyxl.utils.cell import range_boundaries, get_column_letter
    min_col, min_row, max_col, max_row = range_boundaries(coord)
    # итерация от самых мелких диапазонов к крупным — устойчивый выбор якоря
    ranges = sorted(ws.merged_cells.ranges,
                    key=lambda r: (r.max_row - r.min_row) * (r.max_col - r.min_col))
    for rng in ranges:
        if (rng.min_row <= max_row and rng.max_row >= min_row and
                rng.min_col <= max_col and rng.max_col >= min_col):
            anchor = '%s%d' % (get_column_letter(rng.min_col), rng.min_row)
            # если координата лежит внутри слияния — это гарантированно его
            # верхняя левая ячейка; иначе (пересечение краёв разных слияний)
            # оставляем исходную координату, чтобы не перезаписать чужое поле
            c0, r0 = column_index_from_string(get_column_letter(min_col)), min_row
            if rng.min_row <= min_row <= rng.max_row and rng.min_col <= c0 <= rng.max_col:
                return anchor
    return coord


def _set(ws, coord, value, highlight=False):
    """Запись значения в ячейку. Заливки не используются (версия 1.7.0):
    печатная форма формируется программой целиком, жёлтая подсветка
    полей из файлов-примеров удалена."""
    cell = ws[_anchor(ws, coord)]
    cell.value = value
    return cell


# Ячейки, данные из которых убраны по требованию (версия 1.7.x):
#   стр.1: FF1, FJ1, FF2, FL2;  стр.2_3: FK2
CLEAR_STR1 = ('FF1', 'FJ1', 'FF2', 'FL2')
CLEAR_STR23 = ('FK2',)


def _clear_cells(ws, coords):
    """Очищает (обнуляет) указанные ячейки готовящейся формы."""
    for coord in coords:
        ws[_anchor(ws, coord)].value = None


# Столбцы с границами ячеек в таблице раздела 4 листа «стр.2_3»
# (позиции номенклатуры). Нужны для отрисовки сетки дополнительно
# вставленных строк (версия 1.7.2).
_TABLE4_MIN_COL = column_index_from_string('A')
_TABLE4_MAX_COL = column_index_from_string('FE')


def _fix_str23_table_lines(ws, itog_row=35):
    """Дорисовывает недостающие линии таблиц на листе «стр.2_3».

    В исходном шаблоне «Остальное.xlsx» у столбцов FE..FL границы не
    заданы вовсе — из-за этого правая рамка таблицы раздела 4
    (последняя колонка — «Отклонение по количеству (объему)», EX:FE)
    отсутствует, и в формируемом акте на стр.2_3 не хватает нарисованной
    линии справа. Здесь правой границей последней ячейки каждой строки
    таблицы (диапазон 29..itog_row) ставится та же линия, что у левых/
    нижних рамок этой строки; строку «Итого» замыкается линией medium.
    """
    from copy import copy as _copy

    def _style_of(b):
        return b.left.style or b.right.style or b.top.style or b.bottom.style

    for r in range(29, itog_row + 1):
        row_cells = [c for c in ws[r] if c.column <= _TABLE4_MAX_COL]
        last = max(row_cells, key=lambda c: c.column) if row_cells else None
        if last is None:
            continue
        b = last.border
        need = 'medium' if r == itog_row else (_style_of(b) or 'thin')
        if need and b.right.style != need:
            nb = _copy(b)
            nb.right = _copy(b.right)
            nb.right.style = need
            last.border = nb


def _remove_country_code_underline(ws):
    """Убирает подчёркивание цифрового кода страны происхождения товара.

    Колонка 4 таблицы раздела 4 — «Страна происхождения товара»:
    цифровой код (AH31/AH33/AH34...) и краткое наименование (AO...).
    В шаблоне значения в этих ячейках были с подчёркиванием; в
    сформированной форме подчёркивание не используется.
    """
    from openpyxl.styles import Font
    from copy import copy as _copy

    for row in ws.iter_rows(min_col=column_index_from_string('AH'),
                            max_col=column_index_from_string('AW')):
        for c in row:
            f = c.font
            if f is not None and f.underline:
                nf = _copy(f)
                nf.underline = None
                c.font = nf


def _clone_row_borders(ws, src_row, dst_row):
    """Копирует границы всех ячеек строки src_row в строку dst_row.

    Используется при добавлении дополнительных позиций (ниже строк
    33—34 шаблона), чтобы новые строки таблиц имели те же линии,
    что и строки-образцы.
    """
    from copy import copy as _copy
    for col in range(_TABLE4_MIN_COL, _TABLE4_MAX_COL + 1):
        s = ws.cell(row=src_row, column=col)
        d = ws.cell(row=dst_row, column=col)
        d.border = _copy(s.border)


def _split_date(dt):
    """Возвращает (день, месяц_рус_им., год2цифры) для ячеек формы."""
    months = ['', 'января', 'февраля', 'марта', 'апреля', 'мая', 'июня',
              'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря']
    if not dt:
        return '', '', ''
    return str(dt.day), months[dt.month], str(dt.year)[2:]


def nds_to_value(stavka):
    """Ставка НДС из справочника -> числовой коэффициент для расчёта."""
    s = str(stavka or '').strip()
    if not s or 'без' in s.lower():
        return 0
    m = re.search(r'(\d+(?:[.,]\d+)?)', s.replace('%', ''))
    if m:
        return float(m.group(1).replace(',', '.')) / 100.0
    return 0


def _clear_yellow(ws):
    """Снимает жёлтые заливки с ячеек листа.

    В файлах-примерах (Раздел1.xlsx / Остальное.xlsx) жёлтым выделены поля
    для ручного заполнения. Печатная форма заполняется программой целиком,
    поэтому подсветка удаляется: все ячейки становятся белыми."""
    from openpyxl.styles import PatternFill
    for row in ws.iter_rows():
        for c in row:
            f = c.fill
            if f is not None and f.patternType == 'solid' and f.fgColor is not None \
                    and str(f.fgColor.rgb or '').upper() in ('FFFFFF00', '00FFFF00'):
                c.fill = PatternFill(fill_type=None)


def build_priemka_xlsx(priemka, constants, catalogs, out_path):
    """Заполняет обе части формы и сохраняет в один workbook (2 листа)."""
    wb1 = openpyxl.load_workbook(TEMPLATE_STR1)
    ws1 = wb1.active
    _clear_yellow(ws1)
    wb23 = openpyxl.load_workbook(TEMPLATE_STR23)
    ws23 = wb23.active
    _clear_yellow(ws23)

    uch = constants['uchrezhdenie']
    doc = priemka

    dt_act = _parse_date(doc.get('дата')) or date.today()
    den, mesyac, god2 = _split_date(dt_act)

    # ---------- Раздел 1 (основной раздел) ---------------------------------
    _set(ws1, 'CQ4', uch.get('учреждение_получатель', ''))
    _set(ws1, 'EG4', 'И.Н. Харольская' if not uch.get('rukovoditel') else uch.get('rukovoditel'))
    # Утверждаем: дата документа-акта
    _set(ws1, 'CR6', den)
    _set(ws1, 'CW6', mesyac)
    _set(ws1, 'DO6', str(dt_act.year))
    # Номер акта
    _set(ws1, 'CI8', doc.get('номер', 'б/н'))
    _set(ws1, 'ER32', doc.get('номер', 'б/н'))
    _set(ws1, 'ER33', dt_act)
    # Дата акта (строка "от __ _______ 20__ г.")
    _set(ws1, 'BW11', den)
    _set(ws1, 'CB11', mesyac)
    _set(ws1, 'CT11', god2)

    # Константы учреждения
    _set(ws1, 'AS12', uch.get('учреждение_получатель', ''))
    _set(ws1, 'AS13', uch.get('обособленное_подразделение', ''))
    _set(ws1, 'AS14', uch.get('структурное_подразделение', ''))
    _set(ws1, 'AS15', uch.get('главный_администратор_доходов', ''))
    _set(ws1, 'AS16', uch.get('наименование_бюджета', ''))
    _set(ws1, 'AS17', uch.get('валюта', ''))
    _set(ws1, 'AS18', uch.get('адрес_грузополучателя', ''))
    _set(ws1, 'AS19', uch.get('учреждение_получатель', ''))
    _set(ws1, 'DL21', uch.get('inn', ''))
    _set(ws1, 'ER21', uch.get('kpp', ''))
    _set(ws1, 'ER19', uch.get('ogrn', ''))
    _set(ws1, 'AS23', doc.get('adres_zakazchika', uch.get('adres_zakazchika_str', '')))
    _set(ws1, 'AS24', doc.get('mesto_postavki', uch.get('mesto_postavki_str', '')))

    # Документ-основание о создании приёмочной комиссии (приказ)
    _set(ws1, 'AS25', doc.get('dok_osnovanie_vid', 'Приказ'))
    _set(ws1, 'ER25', doc.get('dok_osnovanie_nom', ''))
    _set(ws1, 'ER26', _parse_date(doc.get('dok_osnovanie_data')) or dt_act)

    # Основание приемки (вид, номер, дата, идентификатор)
    _set(ws1, 'AS27', doc.get('osnovanie_vid', 'Контракт'))
    _set(ws1, 'ER27', doc.get('osnovanie_nom', ''))
    _set(ws1, 'CV29', doc.get('ikz', ''))
    _set(ws1, 'ER29', _parse_date(doc.get('osnovanie_data')) or dt_act)

    # Документ об отгрузке
    _set(ws1, 'AS32', doc.get('dok_otgruzka', ''))

    # Сведения о поставщике (из справочника Поставщик по выбору)
    post = doc.get('postavshik', {}) or {}
    _set(ws1, 'AV40', post.get('polnoe_naimovanie', ''))
    _set(ws1, 'AV41', post.get('kratkoe_naimovanie', ''))
    _set(ws1, 'AV42', post.get('adres', ''))
    _set(ws1, 'AV43', post.get('ogrн', post.get('ogrn', '')))
    _set(ws1, 'AV44', post.get('inn', ''))
    _set(ws1, 'AV45', post.get('kpp', ''))

    # Наименование товара (ключ поиска номенклатуры) — ячейка FJ1 более
    # не заполняется и очищается вместе с FF1, FF2, FL2 (см. CLEAR_STR1).

    # ---------- Остальное (стр. 2—3) ---------------------------------------
    # Дата в правом верхнем углу (FK2) более не заполняется (см. CLEAR_STR23).

    # Удаляем данные из ячеек по требованию: стр.1 — FF1, FJ1, FF2, FL2;
    # стр.2_3 — FK2. Выполняется после всех записей, чтобы гарантированно
    # очистить и значения из шаблонов, и случайно записанные программой.
    _clear_cells(ws1, CLEAR_STR1)
    _clear_cells(ws23, CLEAR_STR23)

    # Раздел 2: транспортировка — если задана в документе
    transport = doc.get('transport', {}) or {}
    for col, key in [('G', 'otpravka_data'), ('S', 'otpravka_vremya'),
                     ('AC', 'pribytie_data'), ('AO', 'pribytie_vremya'),
                     ('AY', 'vydacha_data'), ('BK', 'vydacha_vremya'),
                     ('BU', 'vskrytie_data'), ('CG', 'vskrytie_vremya'),
                     ('CQ', 'dostavka_data'), ('DC', 'dostavka_vremya'),
                     ('DM', 'priemka_nachalo'), ('DY', 'priemka_konec'),
                     ('EK', 'mesto_sostavleniya')]:
        val = transport.get(key)
        if val:
            _set(ws23, '%s%d' % (col, 8), val)

    # Подписи: представитель поставщика
    pred = transport.get('predstavitel_postavshika', {}) or {}
    _set(ws23, 'AD18', pred.get('dolzhnost', ''))
    _set(ws23, 'BV18', pred.get('fio', ''))

    # Принято (должность/расшифровка — руководитель или ответственный)
    _set(ws23, 'AD23', uch.get('rukovoditel_dolzhnost', 'Руководитель'))
    _set(ws23, 'BV23', uch.get('rukovoditel', ''))

    # Раздел 4: позиции номенклатуры — добавляем строки копированием шаблона (33-34)
    template_rows = [33, 34]
    rows = doc.get('pozicii', []) or []
    import copy as _copy_mod
    from openpyxl.utils import get_column_letter as _gcl
    n = len(rows)
    needed = max(2, ((n + 1) // 2) * 2)  # чётное число строк блоками по 2
    if needed > 2:
        # ВНИМАНИЕ: дополнительные строки вставляем МЕТОДОМ СДВИВА
        # (insert_rows), иначе позиции перезаписывали бы строку «Итого»
        # и нижние разделы (35+). Границы новых строк копируются со
        # строк-образцов 33/34 — таблицы сохраняют все нарисованные
        # линии (версия 1.7.2).
        #
        # Важно: сдвиг выполняется ДО записи значений, потому что
        # openpyxl при insert_rows не переносит объединённые диапазоны
        # ниже вставляемой строки; все адреса строк «Итого», раздела 5,
        # комиссии и подписей (5.0) указаны уже со сдвигом.
        extra = needed - 2
        ws23.insert_rows(35, extra)
        for k in range(extra):
            _clone_row_borders(ws23, 33 if k % 2 == 0 else 34, 35 + k)
    shift = lambda r: r + max(0, needed - 2)  # смещение строк 35+ после вставки
    total_blocks = needed // 2
    for b in range(total_blocks):
        chunk = rows[b * 2:b * 2 + 2]
        if b == 0:
            _fill_section4_chunk(ws23, chunk, 33, catalogs)
        else:
            for i, poz in enumerate(chunk):
                r = 35 + (b - 1) * 2 + i
                _fill_extra_poz(ws23, r, poz)
    itog_row = 35 + max(0, needed - 2)  # строка «Итого» после вставки
    # Итого с НДС (EF/DD строки «Итого»)
    itog = 0.0
    for r in [33, 34] + list(range(35, itog_row)):
        v = ws23[_anchor(ws23, 'DD%d' % r)].value
        if isinstance(v, (int, float)):
            itog += v
    _set(ws23, 'EF%d' % itog_row, itog, highlight=False)

    # Приложение: наименование документа-основания (АКТ...)
    _set(ws23, 'Z%d' % shift(47), doc.get('dok_otgruzka', ''))
    _set(ws23, 'CV%d' % shift(48), doc.get('fail_pdf', ''))

    # Комиссия по приказу
    com = doc.get('komissiya_rows', []) or []
    slots = [(51, True), (53, False), (55, False), (57, False)]
    idx = 0
    if com:
        head = com[0]
        _set(ws23, 'AA%d' % shift(51), head.get('dolzhnost', ''))
        _set(ws23, 'BP%d' % shift(51), head.get('fio', ''))
        idx = 1
    for row_i in range(len(slots) - 1):
        if idx < len(com):
            m = com[idx]
            r = shift(slots[row_i + 1][0])
            _set(ws23, 'AA%d' % r, m.get('dolzhnost', ''))
            _set(ws23, 'BP%d' % r, m.get('fio', ''))
            idx += 1
    # Ответственный исполнитель
    resp = doc.get('otvetstvenny', {}) or {}
    _set(ws23, 'AA%d' % shift(59), resp.get('dolzhnost', ''))
    _set(ws23, 'BP%d' % shift(59), resp.get('fio', ''))
    _set(ws23, 'CO%d' % shift(59), resp.get('telefon', ''))
    _set(ws23, 'DL%d' % shift(59), resp.get('email', ''))

    # Даты подписей комиссией = дата акта
    _set(ws23, 'C%d' % shift(64), den)
    _set(ws23, 'H%d' % shift(64), mesyac)
    _set(ws23, 'Z%d' % shift(64), god2)
    _set(ws23, 'AA%d' % shift(66), uch.get('rukovoditel_dolzhnost', 'Руководитель'))
    _set(ws23, 'BP%d' % shift(66), uch.get('rukovoditel', ''))

    # Сохраняем обе части формы (два листа) в одном файле
    ws1.title = 'стр.1'
    ws23.title = 'стр.2_3'
    # Дорисовка недостающих линий таблиц и снятие подчёркивания кода
    # страны происхождения — после заполнения всех позиций (1.7.2).
    _fix_str23_table_lines(ws23, itog_row)
    _remove_country_code_underline(ws23)
    for ws in (ws1, ws23):
        ws.page_setup.orientation = 'landscape'
        ws.page_setup.paperSize = 9  # A4
        ws.sheet_view.showGridLines = False
    wb23.active = wb23.index(ws23)

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    # объединяем: копируем лист стр.2_3 из wb23 в wb1 через сохранение во временный поток
    import io as _io
    buf = _io.BytesIO()
    wb23.save(buf)
    buf.seek(0)
    wb23r = openpyxl.load_workbook(buf)
    src = wb23r['стр.2_3']
    tgt = wb1.create_sheet('стр.2_3')
    _copy_sheet(src, tgt)
    tgt.page_setup.orientation = 'landscape'
    tgt.page_setup.paperSize = 9
    ws1.page_setup.orientation = 'landscape'
    ws1.page_setup.paperSize = 9

    wb1.save(out_path)
    return out_path


def _copy_sheet(src, tgt):
    from openpyxl.utils import get_column_letter
    for row in src.iter_rows():
        for cell in row:
            if cell.__class__.__name__ == 'MergedCell':
                continue
            nc = tgt.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                from copy import copy as _copy
                nc.font = _copy(cell.font)
                nc.border = _copy(cell.border)
                nc.fill = _copy(cell.fill)
                nc.number_format = cell.number_format
                nc.alignment = _copy(cell.alignment)
                nc.protection = _copy(cell.protection)
    for rng in list(src.merged_cells.ranges):
        try:
            tgt.merge_cells(str(rng))
        except Exception:
            pass
    for letter, dim in src.column_dimensions.items():
        tgt.column_dimensions[letter].width = dim.width
        tgt.column_dimensions[letter].hidden = dim.hidden
    for rd, dim in src.row_dimensions.items():
        tgt.row_dimensions[rd].height = dim.height
        tgt.row_dimensions[rd].hidden = dim.hidden
    tgt.sheet_format = src.sheet_format
    tgt.freeze_panes = src.freeze_panes


def _fill_extra_poz(ws, r, poz):
    """Заполняет дополнительно вставленную строку раздела 4 (ниже шаблона)."""
    cols = {
        'A': 'nom', 'F': 'kod_tovara', 'O': 'naimovanie',
        'AH': 'strana_kod', 'AO': 'strana_name', 'AX': 'oei_kod',
        'BD': 'ed_izm', 'BL': 'kolichestvo', 'BT': 'cena',
        'CL': 'stavka_nds_text', 'DN': 'nom_deklaracii', 'DZ': 'nom_certifikata',
    }
    for col, key in cols.items():
        v = poz.get(key, '')
        if key == 'nom' and not isinstance(v, int):
            continue
        _set(ws, '%s%d' % (col, r), v)
    kolichestvo = float(poz.get('kolichestvo') or 0)
    cena = float(poz.get('cena') or 0)
    bez = kolichestvo * cena
    koef = nds_to_value(poz.get('stavka_nds_text'))
    summa_nds = round(bez * koef, 2)
    _set(ws, 'CC%d' % r, bez)
    _set(ws, 'CT%d' % r, summa_nds)
    _set(ws, 'DD%d' % r, round(bez + summa_nds, 2))


def _fill_section4_chunk(ws, chunk, first_row, catalogs):
    """Заполняет блок строк 33-34 раздела 4 (позиции номенклатуры)."""
    cols = {
        'A': 'nom', 'F': 'kod_tovara', 'O': 'naimovanie',
        'AH': 'strana_kod', 'AO': 'strana_name', 'AX': 'oei_kod',
        'BD': 'ed_izm', 'BL': 'kolichestvo', 'BT': 'cena',
        'CL': 'stavka_nds_text', 'DN': 'nom_deklaracii', 'DZ': 'nom_certifikata',
    }
    for i, poz in enumerate(chunk):
        r = first_row + i
        for col, key in cols.items():
            if col == 'A':
                continue  # номер позиции проставим целым числом ниже
            _set(ws, '%s%d' % (col, r), poz.get(key, ''))
        _set(ws, 'A%d' % r, int(poz.get('nom') or (r - first_row + 1)))
        cc = column_index_from_string('CC')
        ct = column_index_from_string('CT')
        dd = column_index_from_string('DD')
        kolichestvo = float(poz.get('kolichestvo') or 0)
        cena = float(poz.get('cena') or 0)
        bez_nds = kolichestvo * cena
        koef = nds_to_value(poz.get('stavka_nds_text'))
        summa_nds = round(bez_nds * koef, 2)
        with_nds = round(bez_nds + summa_nds, 2)
        _set(ws, 'CC%d' % r, bez_nds)
        _set(ws, 'CT%d' % r, summa_nds)
        _set(ws, 'DD%d' % r, with_nds)
    # очищаем неиспользованные строки блока
    for i in range(len(chunk), 2):
        r = first_row + i
        for col in list(cols.keys()) + ['CC', 'CT', 'DD']:
            cell = ws[_anchor(ws, '%s%d' % (col, r))]
            if col in ('CC', 'DD'):
                cell.value = '=IF(BL%d=0,BT%d,BL%d*BT%d)' % (r, r, r, r) if r == 34 else None
            elif col == 'CT':
                cell.value = '=IF(CL%d="","",CC%d*CL%d)' % (r, r, r) if r == 34 else None
            else:
                if r != 33:
                    cell.value = None
