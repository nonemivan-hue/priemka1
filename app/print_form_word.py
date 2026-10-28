# -*- coding: utf-8 -*-
"""Выгрузка документа «Приёмка» в Word (.docx).

Раскладка и тексты соответствуют образцу «Акт приемки оказанных услуг,
результатов отдельного этапа исполнения контракта — копия.rtf»:

    ГКУ КК – управление социальной защиты населения в г-к Геленджике
    Акт приемки поставленных товаров, выполненных работ, оказанных услуг...
    г. Геленджик | «DD» месяца YYYY год
    Приемочная комиссия в составе: председатель / члены комиссии
    При исполнении государственного контракта № ... от ... осуществила приемку...
    <Поставщик> (далее – Поставщик) предоставлены результаты...
    Таблица: N п/п | Наименование | Количество (объем)
    Экспертиза -> соответствует; Отчет о транзакциях; Заключение 1..4;
    Подписи: Председатель комиссии / Члены комиссии.

Примечания к образцу (что откуда подставляется):
  * шапка учреждения — constants.uchrezhdenie.учреждение_получатель;
  * город — из адреса заказчика (constants.adres_zakazchika.gorod);
  * дата акта — item.дата;
  * председатель/члены комиссии — приказ item.prikaz_id (prikazy), при
    отсутствии приказа — ответственный исполнитель + руководитель;
  * номер и дата контракта — item.osnovanie_nom / osnovanie_data;
  * поставщик — item.postavshik.kratkoe_naimovanie;
  * позиции — item.pozicii (наименование, количество+ед.изм., цена*кол-во);
  * суммы прописью — num_to_words_rubles (раздел 3 заключения).
"""

import re

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

from app.print_form_61n import num_to_words_rubles

MONTHS = {1: 'января', 2: 'февраля', 3: 'марта', 4: 'апреля', 5: 'мая',
          6: 'июня', 7: 'июля', 8: 'августа', 9: 'сентября', 10: 'октября',
          11: 'ноября', 12: 'декабря'}


def _fmt_money(v):
    try:
        f = float(v)
        if abs(f - round(f)) < 1e-9:
            return '{:,.0f}'.format(f).replace(',', ' ')
        return '{:,.2f}'.format(f).replace(',', ' ')
    except Exception:
        return str(v)


def _split_date(s):
    """'31.05.2026' -> ('31', 'мая', '2026')."""
    try:
        dd, mm, yy = str(s).strip().split('.')
        month = MONTHS.get(int(mm), '')
        return dd.strip(), month, yy.strip()
    except Exception:
        return '', '', ''


def _short_org(full):
    """Свернуть длинное наименование учреждения к виду образца."""
    f = (full or '').lower()
    if 'управление социальной защиты' in f and 'геленджик' in f:
        return ('Государственное казенное Учреждение Краснодарского края – '
                'управление социальной защиты населения в городе-курорте '
                'Геленджике')
    return full or ''


def build_akt_priemki_docx(item, constants, out_path):
    uch = (constants or {}).get('uchrezhdenie', {}) or {}
    az = uch.get('adres_zakazchika', {}) or {}
    gorod = (az.get('gorod') or 'Геленджик').strip()
    if not gorod.startswith('г'): 
        gorod = 'г. ' + gorod

    dd, month, yyyy = _split_date(item.get('дата', ''))

    doc = Document()
    st = doc.styles['Normal']
    st.font.name = 'Times New Roman'
    st.font.size = Pt(12)
    for sec in doc.sections:
        sec.top_margin = Cm(2)
        sec.bottom_margin = Cm(2)
        sec.left_margin = Cm(2)
        sec.right_margin = Cm(1.5)

    def p(text='', align=None, bold=False, italic=False, size=12,
          space_after=4, indent_first=None, indent_left=None):
        par = doc.add_paragraph()
        par.paragraph_format.space_after = Pt(space_after)
        if align is not None:
            par.alignment = align
        if indent_first is not None:
            par.paragraph_format.first_line_indent = Cm(indent_first)
        if indent_left is not None:
            par.paragraph_format.left_indent = Cm(indent_left)
        run = par.add_run(text)
        run.bold = bold
        run.italic = italic
        run.font.size = Pt(size)
        return par

    CTR = WD_ALIGN_PARAGRAPH.CENTER
    JUS = WD_ALIGN_PARAGRAPH.JUSTIFY

    # ---- Шапка: учреждение --------------------------------------------
    org = _short_org(uch.get('учреждение_получатель', ''))
    parts = [ln.strip() for ln in org.split('\n') if ln.strip()]
    if len(parts) == 1 and len(org) > 70:
        parts = ['Государственное казенное',
                 'Учреждение Краснодарского края',
                 '– управление социальной защиты',
                 'населения в городе-курорте Геленджике']
    for ln in parts[:4]:
        p(ln, align=CTR, bold=True, space_after=0)
    p('', space_after=2)

    # ---- Заголовок ------------------------------------------------------
    p('Акт приемки', align=CTR, bold=True, size=14, space_after=0)
    p('поставленных товаров, выполненных работ, оказанных услуг, результатов '
      'отдельного этапа исполнения контракта', align=CTR, space_after=0)
    p('(приемка осуществляется приемочной комиссией)', align=CTR,
      italic=True, space_after=8)

    # ---- Город и дата ---------------------------------------------------
    dnum = '«%s»' % dd if dd else ''
    dtxt = ('%s %s %s год' % (dnum, month, yyyy)).strip()
    par = doc.add_paragraph()
    par.paragraph_format.space_after = Pt(8)
    r = par.add_run(gorod)
    r.add_tab()
    par.add_run('\t\t')
    par.add_run(dtxt)

    # ---- Комиссия -------------------------------------------------------
    prikazy = (constants or {}).get('_prikazy') or []
    prikaz = next((pr for pr in prikazy
                   if pr.get('id') == item.get('prikaz_id')), None)
    pred, chleny = [], []
    if prikaz:
        sost = prikaz.get('sostav') or []
        for m in sost:
            fn = (m.get('фамилия_имя_отчество') or m.get('fio')
                  or ' '.join(filter(None, [m.get('фамилия'),
                                            (m.get('имя') or '')[:1] + '.',
                                            (m.get('отчество') or '')[:1] + '.'])))
            fio = str(fn).strip()
            if m.get('фамилия') and not (m.get('фамилия_имя_отчество') or m.get('fio')):
                # инициалы вида «С.А.» без пробелов — как в образце
                fio = '%s %s' % (m['фамилия'],
                                 ''.join(x[:1] + '.' for x in
                                         (m.get('имя'), m.get('отчество')) if x))
            dolj = m.get('должность') or ''
            role = (m.get('роль') or '').lower()
            entry = '%s%s' % (fio, ' – ' + dolj if dolj else '')
            if 'председ' in role:
                pred.append(entry)
            else:
                chleny.append(entry)
    if not pred:
        otv = item.get('otvetstvenny', {}) or {}
        fio = otv.get('fio') or ''
        dolj = otv.get('dolzhnost') or ''
        head = uch.get('rukovoditel') or ''
        if fio:
            pred.append('%s%s' % (fio, ' – ' + dolj if dolj else ''))
        elif head:
            pred.append(head)
        if fio and not pred:
            pred.append(fio)

    p('Приемочная комиссия в составе:', space_after=2)
    if pred:
        p('        председатель комиссии %s;' % pred[0].rstrip('.'), indent_left=0.5,
          space_after=2)
    if chleny:
        p('члены комиссии:', space_after=0)
        for c in chleny:
            p('- %s;' % c.rstrip('.'), indent_left=1.0, space_after=0)
        p('', space_after=2)

    # ---- Основание (контракт) -------------------------------------------
    vid = item.get('osnovanie_vid') or 'государственного контракта'
    nom = item.get('osnovanie_nom') or ''
    o_dd, o_mm, o_yyyy = _split_date(item.get('osnovanie_data', ''))
    contr_txt = ''
    if nom or o_dd:
        contr_txt = '№ %s от %s %s %s года' % (nom, o_dd, o_mm, o_yyyy)
        contr_txt = contr_txt.replace('  ', ' ').strip('/ от')
        if not nom:
            contr_txt = 'от %s %s %s года' % (o_dd, o_mm, o_yyyy)
    prepos = {'контракт': 'государственного контракта',
              'допсоглашение': 'дополнительного соглашения',
              'заявка': 'заявки'}.get(vid.lower(), vid.lower())
    sent = ('При исполнении %s %s, осуществила приемку поставленного товара/'
            'выполненной работы/оказанной услуги (далее соответственно - '
            'контракт, результаты, предусмотренные контрактом).'
            % (prepos, contr_txt)).replace('  ', ' ')
    p(sent, align=JUS, indent_first=1.25, space_after=6)

    post = ((item.get('postavshik') or {}).get('kratkoe_naimovanie')
            or (item.get('postavshik') or {}).get('polnoe_naimovanie')
            or 'Поставщик')
    p('%s (далее – Поставщик) предоставлены результаты, предусмотренные '
      'контрактом, в приведенных ниже объемах:' % post,
      align=JUS, indent_first=1.25, space_after=6)

    # ---- Таблица позиций --------------------------------------------------
    pozicii = item.get('pozicii') or []
    tbl = doc.add_table(rows=1, cols=3)
    tbl.style = 'Table Grid'
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = tbl.rows[0].cells
    for i, h in enumerate(('N\nп/п', 'Наименование', 'Количество (объем)')):
        para = hdr[i].paragraphs[0]
        para.alignment = CTR
        run = para.add_run(h)
        run.bold = True
    total = 0.0
    for n, pz in enumerate(pozicii, 1):
        cells = tbl.add_row().cells
        cells[0].paragraphs[0].alignment = CTR
        cells[0].text = str(n)
        cells[1].text = pz.get('naimovanie') or ''
        kolichestvo = pz.get('kolichestvo')
        ed = pz.get('ed_izm') or ''
        qty = ('%s %s' % (kolichestvo, ed)).strip()
        cells[2].text = qty
        try:
            total += float(kolichestvo or 0) * float(pz.get('cena') or 0)
        except Exception:
            pass
    tbl.columns[0].width = Cm(1.5)
    tbl.columns[1].width = Cm(11)
    tbl.columns[2].width = Cm(4)
    p('', space_after=4)

    # ---- Экспертиза и отчет ----------------------------------------------
    p('Для проверки предоставленных Поставщиком результатов, предусмотренных '
      'контрактом, в части их соответствия условиям контракта Заказчиком '
      'проведена экспертиза, по результатам которой установлено, что '
      'результаты исполнения соответствуют условиям контракта.',
      align=JUS, indent_first=1.25, space_after=4)
    otdok = item.get('dok_otgruzka') or ''
    rep = ('Отчет о транзакциях.' if not otdok
           else 'Приемочной комиссии представлены на рассмотрение: %s.'
                % otdok.rstrip('.'))
    p(rep, align=JUS, indent_first=1.25, space_after=4)
    p('Рассмотрев представленные отчетные документы и материалы, приемочная '
      'комиссия вынесла следующее заключение:', align=JUS,
      indent_first=1.25, space_after=4)

    words = num_to_words_rubles(round(total, 2))
    # num_to_words_rubles возвращает «... рублей NN копеек» — для формата
    # образца: «97 400 (Девяносто семь тысяч четыреста) рублей, 00 копеек»
    kop = ('%.2f' % total).split('.')[1]
    w = words.strip()
    for tail in (' рублей 00 копеек', ' рубля 00 копеек', ' рубль 00 копеек'):
        if w.endswith(tail):
            w = w[:-len(tail)].rstrip()
            break
    else:
        m = re.search(r'\s(\d{2}) копеек$', w)
        if m:
            kop = m.group(1)
            w = re.sub(r' (\S+ )?(рублей|рубля|рубль) \d{2} копеек$', '', w).strip()
    p('1. Результаты, предусмотренные контрактом, соответствуют условиям '
      'контракта. Отчетная документация оформлена надлежащим образом, '
      'представлена в полном объеме и в установленные сроки.',
      align=JUS, indent_first=1.25, space_after=2)
    p('2. Считать принятыми результаты, предусмотренные контрактом.',
      align=JUS, indent_first=1.25, space_after=2)
    p('3. Цена контракта составляет: %s (%s) рублей, %s копеек.'
      % (_fmt_money(total), w, kop),
      align=JUS, indent_first=1.25, space_after=2)
    p('Стоимость результатов, предусмотренных контрактом, подлежащая оплате '
      'Заказчиком, составляет: %s (%s) рублей, %s копеек.'
      % (_fmt_money(total), w, kop),
      align=JUS, indent_first=1.25, space_after=2)
    p('4. Настоящий акт в соответствии с условиями контракта является '
      'основанием для проведения расчетов Заказчика с Исполнителем.',
      align=JUS, indent_first=1.25, space_after=8)

    # ---- Подписи ----------------------------------------------------------
    p('', space_after=2)
    sign_lines = []
    if pred:
        sign_lines.append(('Председатель комиссии:', pred[0].rstrip('.')))
    for i, c in enumerate(chleny):
        label = 'Члены комиссии:' if i == 0 else ''
        sign_lines.append((label, c.rstrip('.')))
    if not chleny and not pred:
        otv = item.get('otvetstvenny', {}) or {}
        if otv.get('fio'):
            sign_lines.append(('Ответственный исполнитель:', otv['fio']))
    for label, who in sign_lines:
        par = doc.add_paragraph()
        par.paragraph_format.space_after = Pt(6)
        par.add_run(label)
        par.add_run('\t' + '_' * 30 + '\t')
        par.add_run(who + ';')

    doc.save(out_path)
    return out_path
