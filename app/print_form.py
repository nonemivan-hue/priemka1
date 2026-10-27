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

# Эталонный образец печатной формы (Приказ Минфина России от 15.04.2021
# № 61н, ОКУД 0510452 — ред. приказов Минфина от 30.10.2023 № 174н и др.).
# Если он присутствует рядом с программой, форма формируется на его основе:
# расположение блоков, заголовки строк/колонок и ширина сетки точно
# соответствуют файлу «Пример формы.XLSX» из удалённого репозитория.
# В собранном EXE файл упаковывается под ASCII-именем Primer_formy.xlsx.
TEMPLATE_EXAMPLE = _res('Пример формы.XLSX')
if not os.path.isfile(TEMPLATE_EXAMPLE):
    alt = _res('Primer_formy.xlsx')
    if os.path.isfile(alt):
        TEMPLATE_EXAMPLE = alt

EXAMPLE_AVAILABLE = os.path.isfile(TEMPLATE_EXAMPLE)


def _example_wb():
    """Открывает свежую копию эталонного образца (или None, если файла нет)."""
    if not EXAMPLE_AVAILABLE:
        return None
    try:
        return openpyxl.load_workbook(TEMPLATE_EXAMPLE)
    except Exception:
        return None

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


def _strip_example_leftovers(ws):
    """Удаляет из эталонного образца остатки примера заполнения.

    В файле «Пример формы.XLSX» часть ячеек содержит демонстрационные
    значения и устаревшие внешние формулы ([1]стр.1!, [2]Информация! и
    т.п.), которые программа перезаполняет сама. Обнуляются поля ввода
    обеих раскладок — рабочей (Раздел1/Остальное) и эталонной (образец);
    служебные ячейки, заполняемые программой по своим координатам, при
    этом не затрагиваются.
    """
    str1_fields = ['CQ4', 'EG4', 'CA9', 'CI8', 'CR6', 'CW6', 'DO6', 'DO7',
                   'BW11', 'CB11', 'CT11', 'BW12', 'CB12', 'CT12',
                   'AS12', 'AS13', 'AS14', 'AS15', 'AS16', 'AS17', 'AS18',
                   'AS19', 'AS20', 'AS23', 'AS24', 'AS25', 'AS27', 'AS28',
                   'AS29', 'AS31', 'AS36',
                   'DL21', 'DP24', 'ER21', 'ER19', 'ER22', 'ER24', 'ER25', 'ER26',
                   'ER25', 'ER26', 'ER27', 'ER29', 'ER30', 'ER31', 'ER32', 'ER33',
                   'ER36', 'ER37', 'CV29', 'DB33', 'DB35',
                   'AV40', 'AV41', 'AV42', 'AV43', 'AV44', 'AV45',
                   'CA44', 'CA45', 'CA46', 'CA47', 'CA48', 'CA49', 'CA50', 'CA51']
    # стр.2_3: подписи, приложение, строки позиций раздела 4 (33..34)
    str23_fields = ['AD18', 'BV18', 'AD23', 'BV23',
                    'AA51', 'BP51', 'AA53', 'BP53', 'AA55', 'BP55',
                    'AA57', 'BP57', 'AA59', 'BP59', 'CO59', 'DL59',
                    'AA66', 'BP66',
                    'AB79', 'BQ79', 'AB81', 'BQ81', 'AB83', 'BQ83',
                    'AB85', 'BQ85', 'AB87', 'BQ87', 'CP87', 'DM87',
                    'AB89', 'BQ89', 'AB94', 'BQ94', 'AB99', 'BQ99',
                    'AB104', 'BQ104',
                    'Z47', 'CV48', 'O74', 'AA75', 'CW75',
                    'O76', 'AA77', 'CW77']
    if ws.title == 'стр.1':
        coords = str1_fields
    elif ws.title == 'стр.2_3':
        coords = str23_fields
    else:
        return
    for coord in coords:
        try:
            ws[_anchor(ws, coord)].value = None
        except Exception:
            pass
    if ws.title != 'стр.2_3':
        return
    # строки позиций раздела 4 (33—34): демо-значения и внешние формулы
    for r in (33, 34):
        for col in ('A', 'F', 'O', 'AH', 'AO', 'AX', 'BD', 'BL', 'BT',
                    'CC', 'CL', 'CT', 'DD', 'DN', 'DZ'):
            try:
                c = ws[_anchor(ws, '%s%d' % (col, r))]
            except Exception:
                continue
            v = c.value
            if isinstance(v, str) and (v.startswith('=') or v.strip() in ('1', '2')):
                c.value = None
            elif v is not None and not isinstance(v, str):
                c.value = None


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


def _clear_template_prompts(ws):
    """Удаляет из эталонного образца подписи-подсказки в полях ввода.

    В чистом бланке «Пример формы.XLSX» строки вида «(должность)»,
    «(подпись)», «(расшифровка подписи)», «20», «"», « г.» — это
    подсказки для ручного заполнения; программа заполняет эти поля
    данными документа, поэтому подсказки снимаются (только точные
    совпадения, чтобы не задеть заголовки таблиц)."""
    prompts = {'(должность)', '(подпись)', '(расшифровка подписи)',
               '(номер контактного телефона)', '(наименование документа)',
               '(имя файла.pdf)', '"', '“', '”', '20', '20 ', ' г.', 'г.'}
    # Строки с заголовками таблиц (не трогать ни при каких условиях):
    # в них встречаются точные совпадения «1», «Итого» и т.п.
    protected_rows = {30, 31, 32} if ws.title == 'стр.2_3' else set()
    for row in ws.iter_rows():
        for c in row:
            if c.row in protected_rows:
                continue
            v = c.value
            if isinstance(v, str) and v.strip() in prompts:
                c.value = None
    # В строках-образцах позиций раздела 4 (33—34) снимаем нумерацию
    # столбцов «1»..«13» — это разметка примера, а не данные бланка.
    if ws.title == 'стр.2_3':
        for r in (33, 34):
            for c in ws[r]:
                if isinstance(c.value, str) and c.value.strip().isdigit():
                    c.value = None
                elif isinstance(c.value, int) and 1 <= c.value <= 15:
                    c.value = None


def build_priemka_xlsx(priemka, constants, catalogs, out_path):
    """Заполняет обе части формы и сохраняет в один workbook (2 листа).

    Если рядом с программой лежит эталонный файл «Пример формы.XLSX»
    (Приказ Минфина № 61н), форма строится на его основе — расположение
    блоков, заголовки и ширина сетки точно соответствуют образцу.
    Иначе используются рабочие шаблоны Раздел1.xlsx / Остальное.xlsx.
    """
    use_example = EXAMPLE_AVAILABLE
    if use_example:
        try:
            # Эталонный образец содержит ОБА листа в одном workbook —
            # работаем с ним целиком (важно: при загрузке двух копий
            # файла лист «стр.2_3» из второй копии не попал бы в
            # сохраняемый wb1 и позиции терялись).
            wb1 = openpyxl.load_workbook(TEMPLATE_EXAMPLE)
            ws1 = wb1['стр.1']
            _strip_example_leftovers(ws1)
            _clear_template_prompts(ws1)
            _clear_yellow(ws1)
            ws23 = wb1['стр.2_3']
            _strip_example_leftovers(ws23)
            _clear_template_prompts(ws23)
            _clear_yellow(ws23)
            wb23 = wb1  # для совместимости с общим кодом сохранения
        except Exception:
            # образец повреждён/недоступен — откат к рабочим шаблонам
            use_example = False
    if not use_example:
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
    if use_example:
        _set(ws1, 'CQ4', uch.get('учреждение_получатель', ''))
        _set(ws1, 'EG4', uch.get('rukovoditel', '') or 'И.Н. Харольская')
        # Утверждаем: дата документа-акта (образец: «20» DL7, год DO7)
        _set(ws1, 'DL7', '20')
        _set(ws1, 'DO7', god2)           # образец: год DO7, «г.» DR7
    else:
        _set(ws1, 'CQ4', uch.get('учреждение_получатель', ''))
        _set(ws1, 'EG4', uch.get('rukovoditel', '') or 'И.Н. Харольская')
        _set(ws1, 'CR6', den)
        _set(ws1, 'CW6', mesyac)
        _set(ws1, 'DO6', str(dt_act.year))
    # Номер акта: в образце поле — ячейка CA9 («АКТ № ___»),
    # ER32 входит в слияние ER31:FE32 и не является полем ввода.
    if use_example:
        _set(ws1, 'CA9', doc.get('номер', 'б/н'))
    else:
        _set(ws1, 'CI8', doc.get('номер', 'б/н'))
        _set(ws1, 'ER32', doc.get('номер', 'б/н'))
        _set(ws1, 'ER33', dt_act)
    # Дата акта (строка "от __ _______ 20__ г.")
    if use_example:
        _set(ws1, 'BW12', den)
        _set(ws1, 'CB12', mesyac)
        _set(ws1, 'CT12', god2)
    else:
        _set(ws1, 'BW11', den)
        _set(ws1, 'CB11', mesyac)
        _set(ws1, 'CT11', god2)

    # Константы учреждения
    if use_example:
        # Раскладка эталонного бланка (ред. № 174н): строки AS13..AS20,
        # код Сводного реестра — ER13/ER14, Глава по БК — ER16, ОКТМО — ER17,
        # ОКЕИ — ER18, ОГРН — ER22, ИНН — DP24, КПП — ER24, лицевой счёт — ER25.
        _set(ws1, 'AS13', uch.get('учреждение_получатель', ''))
        _set(ws1, 'ER13', uch.get('kod_svodnogo_reestra', ''))
        _set(ws1, 'AS14', uch.get('обособленное_подразделение', ''))
        _set(ws1, 'ER14', uch.get('kod_svodnogo_reestra_obos', ''))
        _set(ws1, 'AS15', uch.get('структурное_подразделение', ''))
        _set(ws1, 'AS16', uch.get('главный_администратор_доходов', ''))
        _set(ws1, 'ER16', uch.get('glava_bk', ''))
        _set(ws1, 'AS17', uch.get('наименование_бюджета', ''))
        _set(ws1, 'ER17', uch.get('oktmo', ''))
        _set(ws1, 'AS18', uch.get('валюта', ''))
        _set(ws1, 'ER18', uch.get('valuta_code', ''))
        _set(ws1, 'AS19', uch.get('адрес_грузополучателя', ''))
        _set(ws1, 'AS20', uch.get('учреждение_получатель', ''))
        _set(ws1, 'ER20', uch.get('platezhny_dokument_nom', ''))
        _set(ws1, 'ER21', uch.get('platezhny_dokument_data', ''))
        _set(ws1, 'ER22', uch.get('ogrn', ''))
        _set(ws1, 'DP24', uch.get('inn', ''))
        _set(ws1, 'ER24', uch.get('kpp', ''))
        _set(ws1, 'ER25', uch.get('litsevoy_schet', ''))
        _set(ws1, 'ER26', uch.get('razdel_litsevogo_scheta', ''))
        _set(ws1, 'AS27', doc.get('adres_zakazchika', uch.get('adres_zakazchika_str', '')))
        _set(ws1, 'AS28', doc.get('mesto_postavki', uch.get('mesto_postavki_str', '')))
        # Документ-основание о создании приёмочной комиссии (приказ)
        _set(ws1, 'AS29', doc.get('dok_osnovanie_vid', 'Приказ'))
        _set(ws1, 'ER29', doc.get('dok_osnovanie_nom', ''))
        _set(ws1, 'ER30', _parse_date(doc.get('dok_osnovanie_data')) or dt_act)
        # Основание приемки (вид, номер, дата)
        _set(ws1, 'AS31', doc.get('osnovanie_vid', 'Контракт'))
        _set(ws1, 'ER31', doc.get('osnovanie_nom', ''))
        _set(ws1, 'ER33', _parse_date(doc.get('osnovanie_data')) or dt_act)
        # ИКЗ (идентификатор государственного контракта) — DB35
        _set(ws1, 'DB35', doc.get('ikz', ''))
        # Документ об отгрузке: наименование AS36, номер ER36, дата ER37
        _set(ws1, 'AS36', doc.get('dok_otgruzka', ''))
        _set(ws1, 'ER36', doc.get('dok_otgruzka_nom', ''))
        _set(ws1, 'ER37', _parse_date(doc.get('dok_otgruzka_data')))
        # Сведения о поставщике (таблица раздела 1, колонка «Сведения о
        # поставщике» — CA44..CA49)
        post = doc.get('postavshik', {}) or {}
        _set(ws1, 'CA44', post.get('polnoe_naimovanie', ''))
        _set(ws1, 'CA45', post.get('kratkoe_naimovanie', ''))
        _set(ws1, 'CA46', post.get('adres', ''))
        _set(ws1, 'CA47', post.get('ogrн', post.get('ogrn', '')))
        _set(ws1, 'CA48', post.get('inn', ''))
        _set(ws1, 'CA49', post.get('kpp', ''))
        _set(ws1, 'CA50', post.get('litsevoy_schet', ''))
        _set(ws1, 'CA51', post.get('razdel_litsevogo_scheta', ''))
    else:
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
    # Номера позиций: берём из данных, при отсутствии — сквозная нумерация.
    for i, poz in enumerate(rows):
        if not isinstance(poz.get('nom'), int) or poz.get('nom') in (None, '', 0):
            poz['nom'] = i + 1
    # Строка «Итого» в макетах: в эталонном бланке метка — CT35, в рабочем
    # шаблоне — EF35; обе лежат на строке 35 (вставка выполняется ниже).
    try:
        for _r in range(29, 46):
            for _c in ws23[_r]:
                if isinstance(_c.value, str) and _c.value.strip().lower().startswith('итого'):
                    itog_label_row = _c.row
                    break
            else:
                continue
            break
    except Exception:
        pass
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
        # openpyxl при insert_rows НЕ переносит объединённые диапазоны
        # ниже вставляемой строки. Поэтому сдвиг выполняем вручную:
        # 1) запоминаем все значения/стили/слияния начиная со строки
        #    «Итого»; 2) чистим эти строки; 3) записываем блок на extra
        #    строк ниже, восстанавливая слияния со сдвигом. Строки
        #    35..35+extra-1 становятся новыми рядами раздела 4 — их
        #    оформление копируется со строк-образцов 33/34 (версия 1.7.2:
        #    таблицы сохраняют все нарисованные линии).
        from openpyxl.utils import get_column_letter as _gcl2
        src_from = itog_label_row
        maxr = ws23.max_row
        maxc = ws23.max_column
        saved_merges = [(rng.min_col, rng.min_row, rng.max_col, rng.max_row)
                        for rng in ws23.merged_cells.ranges
                        if rng.min_row >= src_from]
        block = {}
        for r in range(src_from, maxr + 1):
            for c in range(1, maxc + 1):
                cell = ws23.cell(row=r, column=c)
                if cell.value is not None or cell.has_style:
                    block[(r, c)] = (cell.value,
                                     _copy_mod.copy(cell._style) if cell.has_style else None)
        # очистка нижнего блока и его слияний
        for mc, mr_, xc, xr in saved_merges:
            try:
                ws23.unmerge_cells('%s%d:%s%d' % (_gcl2(mc), mr_, _gcl2(xc), xr))
            except Exception:
                pass
        for r in range(src_from, maxr + 1):
            for c in range(1, maxc + 1):
                cell = ws23.cell(row=r, column=c)
                cell.value = None
                cell.style = 'Normal'
        # запись блока со сдвигом вниз
        for (r, c), (val, stl) in block.items():
            cell = ws23.cell(row=r + extra, column=c)
            cell.value = val
            if stl is not None:
                cell._style = stl
        for mc, mr_, xc, xr in saved_merges:
            ws23.merge_cells('%s%d:%s%d' % (_gcl2(mc), mr_ + extra, _gcl2(xc), xr + extra))
        # оформление новых строк позиций — копированием строк-образцов
        for k in range(extra):
            src_r = 33 if k % 2 == 0 else 34
            dst_r = 35 + k
            for c in range(1, maxc + 1):
                s = ws23.cell(row=src_r, column=c)
                d = ws23.cell(row=dst_r, column=c)
                if s.has_style:
                    d._style = _copy_mod.copy(s._style)
    shift = lambda r: r + max(0, needed - 2)  # смещение строк «Итого»+ после вставки
    total_blocks = needed // 2
    for b in range(total_blocks):
        chunk = rows[b * 2:b * 2 + 2]
        if b == 0:
            _fill_section4_chunk(ws23, chunk, 33, catalogs)
        else:
            for i, poz in enumerate(chunk):
                r = 35 + (b - 1) * 2 + i
                _fill_extra_poz(ws23, r, poz)
    itog_row = itog_label_row + max(0, needed - 2)  # строка «Итого» после вставки
    # Итого с НДС: сумма считается НЕ по ячейкам листа, а напрямую по
    # реально заполненным позициям (1.7.4) — так исключается захват
    # остаточных значений из строк-образцов или смещённых строк «Итого».
    # В эталонном образце итог пишется в DD «Итого» (CT — подпись),
    # в рабочем шаблоне — EF.
    itog = 0.0
    for poz in rows:
        kolichestvo = float(poz.get('kolichestvo') or 0)
        cena = float(poz.get('cena') or 0)
        bez = kolichestvo * cena
        koef = nds_to_value(poz.get('stavka_nds_text'))
        itog += round(bez + round(bez * koef, 2), 2)
    itog = round(itog, 2)
    if use_example:
        _set(ws23, 'DD%d' % itog_row, itog, highlight=False)
    else:
        _set(ws23, 'EF%d' % itog_row, itog, highlight=False)

    # Приложение: наименование документа-основания (АКТ...) — в обоих
    # макетах блоки приложения находятся ниже области вставки строк,
    # поэтому смещаются на extra.
    if use_example:
        _set(ws23, 'AA%d' % shift(74), doc.get('dok_otgruzka', ''))
        _set(ws23, 'CW%d' % shift(75), doc.get('fail_pdf', ''))
    else:
        _set(ws23, 'Z%d' % shift(47), doc.get('dok_otgruzka', ''))
        _set(ws23, 'CV%d' % shift(48), doc.get('fail_pdf', ''))

    # Комиссия по приказу: в эталонном образце строки подписей идут НИЖЕ
    # блока вставки дополнительных позиций и смещаются на extra; в рабочем
    # шаблоне «Остальное.xlsx» блоки комиссии/приложения (47—69) находятся
    # ДО строки вставки (35), а подпись руководителя (66) — ПОСЛЕ неё,
    # поэтому смещается только pred_r.
    com = doc.get('komissiya_rows', []) or []
    if use_example:
        head_r, memb_rs, resp_r, pred_r = 79, [81, 83, 85], 87, 94
        d_col, f_col = 'AB', 'BQ'
        tel_col, mail_col, tr_col, mr_col, yr_col = 'CP', 'DM', 'F', 'R', 'X'
    else:
        head_r, memb_rs, resp_r, pred_r = 51, [53, 55, 57], 59, 66
        d_col, f_col = 'AA', 'BP'
        tel_col, mail_col, tr_col, mr_col, yr_col = 'CO', 'DL', 'C', 'H', 'Z'
    shf = (lambda r: r + max(0, needed - 2)) if use_example else (lambda r: r)
    idx = 0
    if com:
        head = com[0]
        hr = shf(head_r)
        _set(ws23, '%s%d' % (d_col, hr), head.get('dolzhnost', ''))
        _set(ws23, '%s%d' % (f_col, hr), head.get('fio', ''))
        idx = 1
    for row_i in range(len(memb_rs)):
        if idx < len(com):
            m = com[idx]
            r = shf(memb_rs[row_i])
            _set(ws23, '%s%d' % (d_col, r), m.get('dolzhnost', ''))
            _set(ws23, '%s%d' % (f_col, r), m.get('fio', ''))
            idx += 1
    # Ответственный исполнитель
    resp = doc.get('otvetstvenny', {}) or {}
    rr = shf(resp_r)
    _set(ws23, '%s%d' % (d_col, rr), resp.get('dolzhnost', ''))
    _set(ws23, '%s%d' % (f_col, rr), resp.get('fio', ''))
    _set(ws23, '%s%d' % (tel_col, rr), resp.get('telefon', ''))
    _set(ws23, '%s%d' % (mail_col, rr), resp.get('email', ''))

    # Даты подписей комиссией = дата акта; подпись «принято» — руководитель.
    # Строка pred_r в обоих макетах лежит ниже области вставки — смещаем.
    pr = shift(pred_r)
    dr = shf(pred_r - 2) if use_example else pred_r - 2
    _set(ws23, '%s%d' % (tr_col, dr), den)
    _set(ws23, '%s%d' % (mr_col, dr), mesyac)
    _set(ws23, '%s%d' % (yr_col, dr), god2)
    _set(ws23, '%s%d' % (d_col, pr), uch.get('rukovoditel_dolzhnost', 'Руководитель'))
    _set(ws23, '%s%d' % (f_col, pr), uch.get('rukovoditel', ''))

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
    if use_example:
        # Эталонный образец уже содержит оба листа («стр.1» и «стр.2_3»)
        # в одном workbook — копирование второго листа не требуется, иначе
        # появляется дубликат «стр.2_31».
        wb1.save(out_path)
        return out_path
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
