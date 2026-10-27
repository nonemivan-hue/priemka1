# -*- coding: utf-8 -*-
"""Печатная форма «Акт о приеме товаров, работ, услуг» (ОКУД 0510452).

Форма утверждена Приказом Минфина России от 15.04.2021 № 61н
(приложение № 3 к порядку применения универсальных передаточных
документов). Структура и расположение блоков соответствуют образцу
«Пример формы.XLSX»: заголовок с датой составления и номером, блоки
«Получатель» и «Поставщик», основание приемки, сведения о
транспортировке, вскрытии тары и сроках приемки, раздел 4 — таблица
показателей приемки (13 колонок) со строками «Итого» и «Всего»
(сумма прописью), раздел 5 — заключение по результатам приемки,
подписи председателя и членов комиссии, ответственного исполнителя
и представителя поставщика.

Существующая форма учреждения (Раздел1.xlsx / Остальное.xlsx, маршрут
/dokument/priemka/<id>/print) сохранена без изменений; данный модуль
реализует альтернативный вариант — /dokument/priemka/<id>/print61n.
"""
import os
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

from app.print_form import _parse_date, nds_to_value

# Оформление: как в образце «Пример формы.XLSX» — сплошная сетка из тонких
# рамок, Times New Roman 12, перенос текста, выравнивание сверху/слева.
THIN = Side(style='thin', color='000000')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
FONT = Font(name='Times New Roman', size=12)
FONT_B = Font(name='Times New Roman', size=12, bold=True)
CELL_AL = Alignment(horizontal='left', vertical='top', wrap_text=True)
CEN_AL = Alignment(horizontal='center', vertical='center', wrap_text=True)

# Таблица показателей приёмки (раздел 4) — 13 колонок, как в образце
POZ_HEADERS = ['Номер\nпозиции', 'Наименование\nтовара (работы, услуги)',
               'Код\nОКПД2', 'Код\nед. изм.', 'Кол-во\n(объем)',
               'Цена,\nруб./шт.', 'Сумма без\nНДС, руб.', 'Сумма\nНДС, руб.',
               'Стоимость с\nНДС, руб.', 'Страна\nпроисхождения',
               'Номер декларации\nо соответствии',
               'Номер сертификата /\nдокумента качества', 'При-\nмечание']
POZ_N = len(POZ_HEADERS)          # 13
LAST_COL = POZ_N                  # форма занимает колонки A..M
NCOLS = POZ_N                     # совместимость со старым кодом

# Ширины колонок A..J (метки/реквизиты) + K..M (таблица позиций)
WIDTHS = [13, 9, 10, 20, 12, 9, 10, 11, 12, 13, 14, 16, 10]


def _triplet(n, feminine=False):
    words = []
    h, r = divmod(n, 100)
    if h:
        words.append(_HUND[h])
    if 10 <= r < 20:
        words.append(_TEENS[r - 10])
    else:
        t, u = divmod(r, 10)
        if t:
            words.append(_TENS[t])
        if u:
            words.append((_ONES_F if feminine else _ONES)[u])
    return ' '.join(words)


def _plural(n, one, few, many):
    n10, n100 = n % 10, n % 100
    if 11 <= n100 <= 14:
        return many
    if n10 == 1:
        return one
    if 2 <= n10 <= 4:
        return few
    return many


_ONES = ['', 'один', 'два', 'три', 'четыре', 'пять', 'шесть', 'семь',
         'восемь', 'девять']
_ONES_F = ['', 'одна', 'две', 'три', 'четыре', 'пять', 'шесть', 'семь',
           'восемь', 'девять']
_TEENS = ['десять', 'одиннадцать', 'двенадцать', 'тринадцать', 'четырнадцать',
          'пятнадцать', 'шестнадцать', 'семнадцать', 'восемнадцать',
          'девятнадцать']
_TENS = ['', '', 'двадцать', 'тридцать', 'сорок', 'пятьдесят', 'шестьдесят',
         'семьдесят', 'восемьдесят', 'девяносто']
_HUND = ['', 'сто', 'двести', 'триста', 'четыреста', 'пятьсот', 'шестьсот',
         'семьсот', 'восемьсот', 'девятьсот']


def num_to_words_rubles(amount):
    """Сумма в рублях прописью: «... рублей 00 копеек».

    При нулевой сумме возвращает «Ноль рублей 00 копеек» (в т.ч. если
    сумма равна нулю, но есть ненулевые копейки — «N копеек»).
    """
    amount = round(float(amount or 0), 2)
    rub = int(amount)
    kop = int(round((amount - rub) * 100))
    if rub == 0 and kop == 0:
        return 'Ноль рублей 00 копеек'
    if rub == 0:
        return '%d %s' % (kop, _plural(kop, 'копейка', 'копейки', 'копеек'))
    parts = []
    millions = rub // 1_000_000
    thousands = (rub // 1000) % 1000
    rest = rub % 1000
    if millions:
        parts += [_triplet(millions),
                  _plural(millions, 'миллион', 'миллиона', 'миллионов')]
    if thousands:
        parts += [_triplet(thousands, feminine=True),
                  _plural(thousands, 'тысяча', 'тысячи', 'тысяч')]
    if rest or not (millions or thousands):
        parts.append(_triplet(rest))
    parts.append(_plural(rub, 'рубль', 'рубля', 'рублей'))
    return '%s %02d копеек' % (' '.join(parts).strip().capitalize(), kop)


# Совместимость со старым кодом модуля (прежние имена стилей)
BOLD = FONT_B
SMALL = Font(name='Times New Roman', size=9)
CENTER = CEN_AL
LEFT = CELL_AL


def _fmt_dt(value):
    dt = _parse_date(value)
    return dt.strftime('%d.%m.%Y') if dt else (str(value) if value else '')


def _num(v):
    try:
        f = float(v or 0)
    except (TypeError, ValueError):
        return 0.0
    return round(f, 2)


def _put(ws, row, c1, c2, value='', bold=False, align=None, height=None):
    """Объединяет диапазон c1..c2 в строке row, пишет текст, ставит рамки."""
    if c2 > c1:
        ws.merge_cells(start_row=row, start_column=c1, end_row=row,
                       end_column=c2)
    cell = ws.cell(row=row, column=c1, value=value)
    for c in range(c1, c2 + 1):
        cc = ws.cell(row=row, column=c)
        cc.border = BORDER
        cc.font = FONT_B if bold else FONT
        cc.alignment = align or CELL_AL
    if height:
        cur = ws.row_dimensions[row].height or 0
        ws.row_dimensions[row].height = max(cur, height)
    return cell


def _pair_row(ws, row, label, value, lab_w=4, height=None):
    """Строка «метка (A..lab_w) | значение (lab_w+1..LAST_COL)»."""
    _put(ws, row, 1, lab_w, label, bold=True, height=height)
    _put(ws, row, lab_w + 1, LAST_COL, value, height=height)


def _section_header(ws, row, title, height=None):
    _put(ws, row, 1, LAST_COL, title, bold=True, align=CEN_AL, height=height)


def _adres_str(d, fallback=''):
    if isinstance(d, dict):
        parts = [d.get('indeks', ''), d.get('strana', ''),
                 d.get('kray_oblast', ''), d.get('gorod', ''),
                 d.get('ulica', ''), d.get('dom', '')]
        return ', '.join(filter(None, parts)) or fallback
    return str(d) if d else fallback


def build_akt_priemki_61n_xlsx(priemka, constants, out_path):
    """Формирует акт приёмки по форме ОКУД 0510452 (Приказ № 61н).

    Раскладка листа соответствует файлу-образцу «Пример формы.XLSX».
    """
    doc = priemka or {}
    uch = (constants or {}).get('uchrezhdenie', {}) or {}

    wb = Workbook()
    ws = wb.active
    ws.title = 'Акт приемки 0510452'
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = 9  # A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_view.showGridLines = False
    for i, w in enumerate(WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    dt_act = _parse_date(doc.get('дата')) or date.today()
    nomer = doc.get('номер') or 'б/н'

    post = doc.get('postavshik', {}) or {}
    if not isinstance(post, dict):
        post = {'polnoe_naimovanie': str(post)}

    row = 1
    # ---------------------------------------------------- Заголовок
    _put(ws, row, 1, LAST_COL, 'АКТ о приеме товаров, работ, услуг',
         bold=True, align=CEN_AL, height=22)
    row += 1
    _put(ws, row, 1, LAST_COL,
         '(ОКУД 0510452, форма утверждена Приказом Минфина России '
         'от 15.04.2021 № 61н)', align=CEN_AL)
    ws.cell(row=row, column=1).font = Font(name='Times New Roman', size=10)
    row += 1
    _put(ws, row, 1, 3, 'Дата составления', bold=True)
    _put(ws, row, 4, 6, dt_act.strftime('%d.%m.%Y'), align=CEN_AL)
    _put(ws, row, 7, 8, 'Номер документа', bold=True)
    _put(ws, row, 9, LAST_COL, nomer, align=CEN_AL)
    row += 2

    # ------------------------------------------------------ Получатель
    _section_header(ws, row, 'Получатель')
    row += 1
    zak_adres = (doc.get('adres_zakazchika') or
                 uch.get('adres_zakazchika_str') or
                 _adres_str(uch.get('adres_zakazchika'),
                            uch.get('адрес_грузополучателя', '')))
    _put(ws, row, 1, 3, 'Наименование', bold=True)
    _put(ws, row, 4, 8, uch.get('учреждение_получатель', ''), height=30)
    _put(ws, row, 9, LAST_COL,
         'Обособленное подразделение: %s' %
         uch.get('обособленное_подразделение', ''), height=30)
    row += 1
    _put(ws, row, 1, 3, 'Адрес', bold=True)
    _put(ws, row, 4, 8, zak_adres, height=26)
    _put(ws, row, 9, LAST_COL,
         'Место поставки (приемки): %s' %
         (doc.get('mesto_postavki') or uch.get('mesto_postavki_str') or
          _adres_str(uch.get('mesto_postavki'))), height=26)
    row += 1
    _put(ws, row, 1, 3, 'ИНН / КПП / ОГРН', bold=True)
    _put(ws, row, 4, LAST_COL,
         '%s / %s / %s' % (uch.get('inn', ''), uch.get('kpp', ''),
                           uch.get('ogrn', '')))
    row += 1
    _put(ws, row, 1, 3, 'Лицевой счет', bold=True)
    _put(ws, row, 4, LAST_COL, uch.get('лицевой_счет', uch.get('lsc', '')))
    row += 2

    # ------------------------------------------------------- Поставщик
    _section_header(ws, row, 'Поставщик')
    row += 1
    _put(ws, row, 1, 3, 'Наименование', bold=True)
    _put(ws, row, 4, LAST_COL,
         '%s%s' % (post.get('polnoe_naimovanie', ''),
                   (' (%s)' % post.get('kratkoe_naimovanie'))
                   if post.get('kratkoe_naimovanie') else ''), height=30)
    row += 1
    _put(ws, row, 1, 3, 'Адрес', bold=True)
    _put(ws, row, 4, LAST_COL, post.get('adres', ''), height=26)
    row += 1
    _put(ws, row, 1, 3, 'ИНН / КПП / ОГРН', bold=True)
    _put(ws, row, 4, LAST_COL,
         '%s / %s / %s' % (post.get('inn', ''), post.get('kpp', ''),
                           post.get('ogrн', post.get('ogrn', ''))))
    row += 2

    # ------------------------------------------ Основание приемки
    _section_header(ws, row,
                    'Основание приемки (документ, на основании которого '
                    'производится приемка)')
    row += 1
    _put(ws, row, 1, 3, 'Вид документа', bold=True)
    _put(ws, row, 4, 6, doc.get('osnovanie_vid', ''))
    _put(ws, row, 7, LAST_COL,
         'Документ о создании приёмочной комиссии: %s № %s от %s' %
         (doc.get('dok_osnovanie_vid', 'Приказ'),
          doc.get('dok_osnovanie_nom', ''),
          _fmt_dt(doc.get('dok_osnovanie_data'))))
    row += 1
    _put(ws, row, 1, 3, 'Номер документа', bold=True)
    _put(ws, row, 4, 6, doc.get('osnovanie_nom', ''))
    _put(ws, row, 7, LAST_COL,
         'Идентификатор госконтракта (ИКЗ): %s' % doc.get('ikz', ''))
    row += 1
    _put(ws, row, 1, 3, 'Дата документа', bold=True)
    _put(ws, row, 4, 6, _fmt_dt(doc.get('osnovanie_data')))
    _put(ws, row, 7, LAST_COL,
         'Документ об отгрузке: %s' % doc.get('dok_otgruzka', ''))
    row += 2

    # ------------------------ Сведения о транспортировке и сроках приёмки
    _section_header(ws, row,
                    'Сведения о транспортировке, вскрытии тары и сроках приемки')
    row += 1
    tr = doc.get('transport', {}) or {}
    transport_lines = [
        ('Вид транспорта', tr.get('vid_transporta', '')),
        ('Дата/время отправки (отгрузки)',
         '%s %s' % (_fmt_dt(tr.get('otpravka_data')),
                    tr.get('otpravka_vremya', ''))),
        ('Дата/время прибытия',
         '%s %s' % (_fmt_dt(tr.get('pribytie_data')),
                    tr.get('pribytie_vremya', ''))),
        ('Дата/время выдачи грузополучателю',
         '%s %s' % (_fmt_dt(tr.get('vydacha_data')),
                    tr.get('vydacha_vremya', ''))),
        ('Дата/время вскрытия тары',
         '%s %s' % (_fmt_dt(tr.get('vskrytie_data')),
                    tr.get('vskrytie_vremya', ''))),
        ('Начало приемки', _fmt_dt(tr.get('priemka_nachalo'))),
        ('Окончание приемки', _fmt_dt(tr.get('priemka_konec'))),
        ('Место составления акта',
         tr.get('mesto_sostavleniya') or uch.get('mesto_postavki_str', '')),
    ]
    has_tr = any(str(v).strip() for _, v in transport_lines)
    for label, value in transport_lines:
        _pair_row(ws, row, label, value if has_tr else '')
        row += 1
    pred = tr.get('predstavitel_postavshika', {}) or {}
    _pair_row(ws, row, 'Представитель поставщика',
              '%s, %s' % (pred.get('dolzhnost', ''), pred.get('fio', ''))
              if pred else '')
    row += 2

    # ------------------------------------------------ 4. Показатели приёмки
    _section_header(ws, row,
                    '4. Показатели приемки (количественные и качественные)')
    row += 1
    hdr_row = row
    for ci, h in enumerate(POZ_HEADERS, start=1):
        c = ws.cell(row=hdr_row, column=ci, value=h)
        c.font = FONT_B
        c.alignment = CEN_AL
        c.border = BORDER
    ws.row_dimensions[hdr_row].height = 40
    row += 1

    pozicii = doc.get('pozicii', []) or []
    it_bez = it_nds = it_full = 0.0
    for i, poz in enumerate(pozicii, start=1):
        kolichestvo = _num(poz.get('kolichestvo'))
        cena = _num(poz.get('cena'))
        bez = round(kolichestvo * cena, 2)
        koef = nds_to_value(poz.get('stavka_nds_text') or poz.get('stavka_nds'))
        nds = round(bez * koef, 2)
        full = round(bez + nds, 2)
        it_bez += bez
        it_nds += nds
        it_full += full
        naim = poz.get('naimovanie') or poz.get('kratkoe_naimovanie', '')
        strana = ''
        if poz.get('strana_name') or poz.get('strana_kod'):
            strana = '%s (%s)' % (poz.get('strana_name', ''),
                                  poz.get('strana_kod', ''))
        ed = poz.get('ed_izm', '')
        ed_val = ('%s (%s)' % (ed, poz.get('oei_kod'))
                  if ed and poz.get('oei_kod') else ed)
        vals = [poz.get('nom', i), naim, poz.get('kod_tovara', ''),
                ed_val, kolichestvo, cena, bez, nds, full, strana,
                poz.get('nom_deklaracii', ''), poz.get('nom_certifikata', ''),
                'соответствует']
        for ci, v in enumerate(vals, start=1):
            c = ws.cell(row=row, column=ci, value=v)
            c.font = FONT
            c.border = BORDER
            c.alignment = CELL_AL if ci in (2, 10, 11, 12) else CEN_AL
            if ci in (6, 7, 8, 9):
                c.number_format = '#,##0.00'
        ws.row_dimensions[row].height = 30
        row += 1
    if not pozicii:
        _put(ws, row, 1, LAST_COL, 'Позиции не указаны', align=CEN_AL)
        row += 1

    # Итого
    _put(ws, row, 1, 6, 'Итого', bold=True, align=CEN_AL)
    for ci, v in [(7, round(it_bez, 2)), (8, round(it_nds, 2)),
                  (9, round(it_full, 2))]:
        c = ws.cell(row=row, column=ci, value=v)
        c.font = FONT_B
        c.number_format = '#,##0.00'
        c.alignment = CEN_AL
    for ci in (10, 11, 12, 13):
        cc = ws.cell(row=row, column=ci)
        cc.border = BORDER
        cc.font = FONT
    row += 1
    # Всего — сумма прописью
    _put(ws, row, 1, 2, 'Всего', bold=True, align=CEN_AL)
    if pozicii:
        total_text = '%s. В том числе НДС: %s.' % (
            num_to_words_rubles(round(it_full, 2)),
            num_to_words_rubles(round(it_nds, 2)) if it_nds else 'без НДС')
    else:
        total_text = ''
    _put(ws, row, 3, LAST_COL, total_text, height=26)
    row += 2

    # ----------------------------------------- 5. Заключение по приёмке
    _section_header(ws, row, '5. Заключение по результатам приемки')
    row += 1
    default_zak = ('Фактические данные о количестве (объеме) и качестве '
                   'принятых товаров (работ, услуг) соответствуют документу, '
                   'на основании которого произведена приемка. Отклонения по '
                   'количеству и качеству не выявлены. Товары (работы, услуги) '
                   'приняты.')
    _put(ws, row, 1, LAST_COL, doc.get('zakuchenie', default_zak), height=40)
    row += 2

    # --------------------------------------------------------- Подписи
    _section_header(ws, row, 'Подписи')
    row += 1
    com = doc.get('komissiya_rows', []) or []
    if com:
        head = com[0]
        _put(ws, row, 1, 2, 'Председатель комиссии', bold=True)
        _put(ws, row, 3, 5, head.get('dolzhnost', ''))
        _put(ws, row, 6, 8, head.get('fio', ''))
        _put(ws, row, 9, LAST_COL, 'подпись                      дата')
        row += 1
        for m in com[1:]:
            _put(ws, row, 1, 2, 'Член комиссии', bold=True)
            _put(ws, row, 3, 5, m.get('dolzhnost', ''))
            _put(ws, row, 6, 8, m.get('fio', ''))
            _put(ws, row, 9, LAST_COL, 'подпись                      дата')
            row += 1
    else:
        _put(ws, row, 1, 2, 'Председатель комиссии', bold=True)
        _put(ws, row, 3, LAST_COL,
             '_______________________________         ________________       '
             '«___» ____________ 20___ г.')
        row += 1
        _put(ws, row, 1, 2, 'Члены комиссии', bold=True)
        _put(ws, row, 3, LAST_COL,
             '_______________________________         ________________       '
             '«___» ____________ 20___ г.')
        row += 1

    resp = doc.get('otvetstvenny', {}) or {}
    _put(ws, row, 1, 2, 'Ответственный исполнитель', bold=True)
    _put(ws, row, 3, 5, resp.get('dolzhnost', ''))
    _put(ws, row, 6, 8, resp.get('fio', ''))
    _put(ws, row, 9, LAST_COL, 'подпись                      дата')
    row += 1
    _put(ws, row, 1, 2, 'Поставщик (представитель)', bold=True)
    _put(ws, row, 3, 8,
         '%s, %s' % (pred.get('dolzhnost', ''), pred.get('fio', ''))
         if pred else '')
    _put(ws, row, 9, LAST_COL, 'подпись                      дата')

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    wb.save(out_path)
    return out_path
