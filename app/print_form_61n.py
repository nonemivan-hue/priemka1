# -*- coding: utf-8 -*-
"""Печатная форма «Акт приемки товаров, работ, услуг» (ОКУД 0510452).

Форма утверждена Приказом Минфина России от 15.04.2021 № 61н
(«Об утверждении универсальных передаточных документов»).
Формируется программой целиком (без файлов-шаблонов) в отдельный
Excel-файл; структура соответствует форме 0510452:

  * заголовок, реквизиты сторон (получатель / поставщик);
  * раздел 1 — данные о документе-основании и документе об отгрузке;
  * раздел 2 — сведения о приёмке (транспортировка, вскрытие, сроки);
  * таблица показателей приёмки по позициям (количество, цена, суммы, НДС);
  * раздел с отклонениями;
  * подписи комиссии (председатель/члены по приказу), ответственного
    исполнителя и руководителя.

Существующая форма учреждения (Раздел1.xlsx / Остальное.xlsx, маршрут
/dokument/priemka/<id>/print) сохранена без изменений; данный модуль
добавляет альтернативный вариант — /dokument/priemka/<id>/print61n.
"""
import os
from datetime import date, datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from app.print_form import _parse_date, nds_to_value

THIN = Side(style='thin', color='000000')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BOLD = Font(bold=True)
SMALL = Font(size=9)
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT = Alignment(horizontal='left', vertical='top', wrap_text=True)
HEAD_FILL = PatternFill('solid', fgColor='D9E1F2')

NCOLS = 13  # ширина таблицы показателей


def _merge(ws, row, c1, c2, value, bold=False, align=None, border=True):
    ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
    cell = ws.cell(row=row, column=c1, value=value)
    if bold:
        cell.font = BOLD
    cell.alignment = align or LEFT
    if border:
        for c in range(c1, c2 + 1):
            ws.cell(row=row, column=c).border = BORDER
    return cell


def _label_value(ws, row, label, value, lab_w=(1, 4), val_w=(5, NCOLS)):
    _merge(ws, row, lab_w[0], lab_w[1], label, bold=True, border=True)
    _merge(ws, row, val_w[0], val_w[1], value, border=True)


def _fmt_dt(value):
    dt = _parse_date(value)
    return dt.strftime('%d.%m.%Y') if dt else (str(value) if value else '')


def _num(v):
    try:
        f = float(v or 0)
    except (TypeError, ValueError):
        return 0.0
    return round(f, 2)


POZ_HEADERS = ['№ п/п', 'Наименование\n(код товара)', 'Страна\nпроисхождения',
               'Ед. изм.', 'Кол-во', 'Цена,\nруб.', 'Стоимость без\nНДС, руб.',
               'Ставка\nНДС', 'Сумма\nНДС, руб.', 'Стоимость с\nНДС, руб.',
               'Номер деклар.\nо соотв.', 'Номер сертификата /\ndок-та кач-ва',
               'Примечание']
POZ_WIDTHS = [6, 34, 14, 8, 9, 10, 13, 8, 12, 13, 14, 16, 14]


def build_akt_priemki_61n_xlsx(priemka, constants, out_path):
    """Формирует акт приёмки по форме ОКУД 0510452 (Приказ Минфина № 61н)."""
    doc = priemka
    uch = constants.get('uchrezhdenie', {}) or {}
    wb = Workbook()
    ws = wb.active
    ws.title = 'Акт 0510452'
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = 9  # A4
    ws.sheet_view.showGridLines = False
    for i, w in enumerate(POZ_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    dt_act = _parse_date(doc.get('дата')) or date.today()
    nomer = doc.get('номер') or 'б/н'

    row = 1
    # ------------------------- Заголовок ----------------------------------
    _merge(ws, row, 1, NCOLS,
           'АКТ приемки товаров, работ, услуг', bold=True, align=CENTER, border=False)
    ws.cell(row=row, column=1).font = Font(bold=True, size=14)
    row += 1
    _merge(ws, row, 1, NCOLS,
           '(ОКУД 0510452) — форма утверждена Приказом Минфина России '
           'от 15.04.2021 № 61н', align=CENTER, border=False)
    ws.cell(row=row, column=1).font = SMALL
    row += 1
    _merge(ws, row, 1, NCOLS,
           'Дата составления: %s        Номер документа: %s' %
           (dt_act.strftime('%d.%m.%Y'), nomer), align=CENTER, border=False)
    row += 2

    # ---------------------- Реквизиты сторон ------------------------------
    _merge(ws, row, 1, NCOLS, 'Получатель', bold=True, align=CENTER)
    for c in range(1, NCOLS + 1):
        ws.cell(row=row, column=c).fill = HEAD_FILL
    row += 1
    _label_value(ws, row, 'Наименование организации (учреждение-получатель)',
                 uch.get('учреждение_получатель', '')); row += 1
    _label_value(ws, row, 'Обособленное / структурное подразделение',
                 ', '.join(filter(None, [uch.get('обособленное_подразделение', ''),
                                         uch.get('структурное_подразделение', '')]))); row += 1
    _label_value(ws, row, 'ИНН / КПП / ОГРН',
                 '%s / %s / %s' % (uch.get('inn', ''), uch.get('kpp', ''),
                                   uch.get('ogrn', ''))); row += 1
    _label_value(ws, row, 'Адрес (местонахождение)',
                 doc.get('adres_zakazchika', uch.get('адрес_грузополучателя', ''))); row += 1

    post = doc.get('postavshik', {}) or {}
    if not isinstance(post, dict):
        post = {'polnoe_naimovanie': str(post)}
    _merge(ws, row, 1, NCOLS, 'Поставщик', bold=True, align=CENTER)
    for c in range(1, NCOLS + 1):
        ws.cell(row=row, column=c).fill = HEAD_FILL
    row += 1
    _label_value(ws, row, 'Наименование (полное / краткое)',
                 '%s (%s)' % (post.get('polnoe_naimovanie', ''),
                              post.get('kratkoe_naimovanie', ''))
                 if post.get('kratkoe_naimovanie') else post.get('polnoe_naimovanie', '')); row += 1
    _label_value(ws, row, 'Адрес поставщика', post.get('adres', '')); row += 1
    _label_value(ws, row, 'ИНН / КПП / ОГРН',
                 '%s / %s / %s' % (post.get('inn', ''), post.get('kpp', ''),
                                   post.get('ogrн', post.get('ogrn', '')))); row += 1
    row += 1

    # ---------------- Раздел 1: документ-основание ------------------------
    _merge(ws, row, 1, NCOLS,
           'Раздел 1. Данные о документе, на основании которого производится приемка',
           bold=True, align=CENTER)
    for c in range(1, NCOLS + 1):
        ws.cell(row=row, column=c).fill = HEAD_FILL
    row += 1
    _label_value(ws, row, 'Вид документа-основания', doc.get('osnovanie_vid', '')); row += 1
    _label_value(ws, row, 'Номер документа-основания', doc.get('osnovanie_nom', '')); row += 1
    _label_value(ws, row, 'Дата документа-основания',
                 _fmt_dt(doc.get('osnovanie_data'))); row += 1
    _label_value(ws, row, 'Идентификатор госконтракта (ИКЗ)', doc.get('ikz', '')); row += 1
    _label_value(ws, row, 'Документ об отгрузке и приемке',
                 doc.get('dok_otgruzka', '')); row += 1
    _label_value(ws, row, 'Место поставки (приемки)',
                 doc.get('mesto_postavki', uch.get('mesto_postavki_str', ''))); row += 1
    _label_value(ws, row, 'Документ о создании приёмочной комиссии',
                 '%s № %s от %s' % (doc.get('dok_osnovanie_vid', 'Приказ'),
                                    doc.get('dok_osnovanie_nom', ''),
                                    _fmt_dt(doc.get('dok_osnovanie_data')))); row += 2

    # ------------- Раздел 2: сведения о приёмке / транспортировке ---------
    _merge(ws, row, 1, NCOLS,
           'Раздел 2. Сведения о приёмке (транспортировке, вскрытии тары)',
           bold=True, align=CENTER)
    for c in range(1, NCOLS + 1):
        ws.cell(row=row, column=c).fill = HEAD_FILL
    row += 1
    tr = doc.get('transport', {}) or {}
    tr_rows = [
        ('Дата/время отправки (отгрузки)', '%s %s' % (_fmt_dt(tr.get('otpravka_data')),
                                                      tr.get('otpravka_vremya', ''))),
        ('Дата/время прибытия', '%s %s' % (_fmt_dt(tr.get('pribytie_data')),
                                           tr.get('pribytie_vremya', ''))),
        ('Дата/время выдачи грузополучателю', '%s %s' % (_fmt_dt(tr.get('vydacha_data')),
                                                         tr.get('vydacha_vremya', ''))),
        ('Дата/время вскрытия тары', '%s %s' % (_fmt_dt(tr.get('vskrytie_data')),
                                                tr.get('vskrytie_vremya', ''))),
        ('Начало приёмки', _fmt_dt(tr.get('priemka_nachalo'))),
        ('Окончание приёмки', _fmt_dt(tr.get('priemka_konec'))),
        ('Место составлении акта', tr.get('mesto_sostavleniya', '')),
    ]
    has_tr = any(v.strip() for _, v in tr_rows)
    for label, value in tr_rows:
        _label_value(ws, row, label, value if has_tr else '—'); row += 1
    pred = tr.get('predstavitel_postavshika', {}) or {}
    _label_value(ws, row, 'Представитель поставщика',
                 '%s, %s' % (pred.get('dolzhnost', ''), pred.get('fio', ''))
                 if pred else ''); row += 2

    # ------------------- Таблица показателей приёмки -----------------------
    _merge(ws, row, 1, NCOLS,
           'Показатели приемки (количественные и качественные)',
           bold=True, align=CENTER)
    for c in range(1, NCOLS + 1):
        ws.cell(row=row, column=c).fill = HEAD_FILL
    row += 1
    hdr_row = row
    for ci, h in enumerate(POZ_HEADERS, start=1):
        c = ws.cell(row=hdr_row, column=ci, value=h)
        c.font = BOLD
        c.alignment = CENTER
        c.border = BORDER
    ws.row_dimensions[hdr_row].height = 30
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
        kod = poz.get('kod_tovara', '')
        strana = '%s (%s)' % (poz.get('strana_name', ''), poz.get('strana_kod', '')) \
            if (poz.get('strana_name') or poz.get('strana_kod')) else ''
        vals = [poz.get('nom', i),
                '%s%s' % (naim, (' (код: %s)' % kod) if kod else ''),
                strana, poz.get('ed_izm', ''), kolichestvo, cena, bez,
                poz.get('stavka_nds_text') or poz.get('stavka_nds', ''),
                nds, full, poz.get('nom_deklaracii', ''),
                poz.get('nom_certifikata', ''), 'соответствует']
        for ci, v in enumerate(vals, start=1):
            c = ws.cell(row=row, column=ci, value=v)
            c.border = BORDER
            c.alignment = LEFT if ci in (2, 3) else CENTER
            if ci in (6, 7, 9, 10):
                c.number_format = '#,##0.00'
                c.alignment = Alignment(horizontal='right', vertical='top')
            if ci == 5:
                c.number_format = '#,##0.###'
        row += 1
    if not pozicii:
        _merge(ws, row, 1, NCOLS, 'Позиции не указаны', align=CENTER)
        row += 1

    # Итого
    _merge(ws, row, 1, 6, 'Итого', bold=True)
    for ci, v in [(7, round(it_bez, 2)), (9, round(it_nds, 2)), (10, round(it_full, 2))]:
        c = ws.cell(row=row, column=ci, value=v)
        c.font = BOLD
        c.number_format = '#,##0.00'
    for ci in range(1, NCOLS + 1):
        ws.cell(row=row, column=ci).border = BORDER
    row += 2

    # --------------------- Отклонения при приёмке --------------------------
    _merge(ws, row, 1, NCOLS,
           'Выявленные отклонения по количеству (объему) и качеству',
           bold=True, align=CENTER)
    for c in range(1, NCOLS + 1):
        ws.cell(row=row, column=c).fill = HEAD_FILL
    row += 1
    _merge(ws, row, 1, NCOLS,
           'Отклонений по количеству и качеству не выявлено. '
           'Товары (работы, услуги) приняты в количестве и качестве, '
           'соответствующем документу-основанию.', border=True)
    ws.row_dimensions[row].height = 30
    row += 2

    # --------------------------- Подписи -----------------------------------
    _merge(ws, row, 1, NCOLS, 'Подписи', bold=True, align=CENTER)
    for c in range(1, NCOLS + 1):
        ws.cell(row=row, column=c).fill = HEAD_FILL
    row += 1
    com = doc.get('komissiya_rows', []) or []
    if com:
        head = com[0]
        _label_value(ws, row, 'Председатель приёмочной комиссии',
                     '%s _______________ %s' % (head.get('dolzhnost', ''),
                                                head.get('fio', '')))
        row += 1
        chleny = ', '.join('%s (%s)' % (m.get('fio', ''), m.get('dolzhnost', ''))
                           for m in com[1:])
        _merge(ws, row, 1, 4, 'Члены комиссии', bold=True)
        _merge(ws, row, 5, NCOLS, chleny)
        ws.row_dimensions[row].height = 30
        row += 1
        _merge(ws, row, 1, NCOLS,
               'Подписи членов комиссии: ' +
               '          '.join('_______________ %s' % m.get('fio', '')
                                 for m in com[1:]) if len(com) > 1 else '',
               border=True)
        row += 1
    else:
        _label_value(ws, row, 'Члены приёмочной комиссии',
                     'комиссия не назначена (приказ не выбран)'); row += 1

    resp = doc.get('otvetstvenny', {}) or {}
    _label_value(ws, row, 'Ответственный исполнитель',
                 '%s, %s   тел.: %s' % (resp.get('dolzhnost', ''), resp.get('fio', ''),
                                        resp.get('telefon', ''))); row += 1
    ruk_d = uch.get('rukovoditel_dolzhnost', 'Руководитель')
    ruk = uch.get('rukovoditel', '')
    _label_value(ws, row, 'Утверждён руководителем',
                 '%s _______________ %s    «__» ________ 20__ г.' % (ruk_d, ruk))
    row += 1
    _merge(ws, row + 1, 1, NCOLS,
           'Электронная подпись / файл документа: %s' % doc.get('fail_pdf', ''),
           border=False)
    ws.cell(row=row + 1, column=1).font = SMALL

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    wb.save(out_path)
    return out_path
