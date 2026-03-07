"""
Генератор титульного листа по ГОСТ.
Подставляет данные студента в шаблон, поддерживает рамку.
"""

import os

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from .formatter import (
    GOST, prepare_template,
    _move_sectpr_to_end, _set_margins,
)

FONT = "Times New Roman"


# ================================================================
# ВСПОМОГАТЕЛЬНЫЕ
# ================================================================

def _para(doc, text, size=14, bold=False, align='center',
          space_before=0, space_after=0):
    """Добавляет параграф с форматированием"""
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    try:
        run._element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
    except:
        pass

    aligns = {
        'center': WD_ALIGN_PARAGRAPH.CENTER,
        'left': WD_ALIGN_PARAGRAPH.LEFT,
        'right': WD_ALIGN_PARAGRAPH.RIGHT,
        'justify': WD_ALIGN_PARAGRAPH.JUSTIFY,
    }
    p.paragraph_format.alignment = aligns.get(align, WD_ALIGN_PARAGRAPH.CENTER)
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.15
    p.paragraph_format.first_line_indent = Cm(0)
    return p


def _empty(doc, count=1):
    """Пустые строки"""
    for _ in range(count):
        _para(doc, '', size=14)


def _remove_borders(table):
    """Убирает все границы таблицы"""
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)

    for old in tblPr.findall(qn('w:tblBorders')):
        tblPr.remove(old)

    borders = OxmlElement('w:tblBorders')
    for name in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        b = OxmlElement(f'w:{name}')
        b.set(qn('w:val'), 'none')
        b.set(qn('w:sz'), '0')
        b.set(qn('w:space'), '0')
        b.set(qn('w:color'), 'auto')
        borders.append(b)
    tblPr.append(borders)


def _cell_bottom_border(cell):
    """Нижняя граница ячейки (линия подчёркивания)"""
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn('w:tcBorders')):
        tcPr.remove(old)

    borders = OxmlElement('w:tcBorders')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '4')
    bottom.set(qn('w:space'), '0')
    bottom.set(qn('w:color'), '000000')
    borders.append(bottom)
    tcPr.append(borders)


def _cell(cell, text, size=14, bold=False, align='left', underline=False):
    """Заполняет ячейку таблицы"""
    cell.text = ''
    p = cell.paragraphs[0]
    run = p.add_run(text)
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    if underline:
        _cell_bottom_border(cell)
    try:
        run._element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
    except:
        pass

    aligns = {
        'center': WD_ALIGN_PARAGRAPH.CENTER,
        'right': WD_ALIGN_PARAGRAPH.RIGHT,
        'left': WD_ALIGN_PARAGRAPH.LEFT,
    }
    p.paragraph_format.alignment = aligns.get(align, WD_ALIGN_PARAGRAPH.LEFT)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.first_line_indent = Cm(0)


# ================================================================
# ГЛАВНАЯ ФУНКЦИЯ
# ================================================================

def create_title_page(data, output_path, template_path=None):
    """
    Генерирует титульный лист.

    data:
        institution     — название учебного заведения
        work_title      — "ОТЧЕТ ПО ПРАКТИЧЕСКОЙ РАБОТЕ" и т.д.
        work_number     — номер работы
        specialty_code  — "09.02.07"
        specialty_name  — "Информационные системы и программирование"
        subject         — "Web-программирование"
        group           — "9ПР-3.23П"
        student_id      — "194кр"
        student_name    — "П.Ю. Пьянникова"
        teacher_name    — "А.Г. Королев"
        city            — "Красноярск"
        year            — "2026"
    """

    # ─── Создание документа ───
    if template_path and os.path.exists(template_path):
        doc = prepare_template(template_path)
        print(f"  🖼️ Титульный лист С рамкой")
    else:
        doc = Document()
        print(f"  📄 Титульный лист без рамки")

    _set_margins(doc)

    # ═══════════════════════════════════════════
    # 1. ШАПКА — Министерство + Колледж
    # ═══════════════════════════════════════════
    institution = data.get('institution', '').strip()
    if not institution:
        institution = (
            'краевое государственное бюджетное профессиональное '
            'образовательное учреждение\n'
            '«Красноярский колледж радиоэлектроники '
            'и информационных технологий»'
        )

    _para(doc, 'Министерство образования Красноярского края',
          size=14, space_after=4)

    # Разбиваем institution по \n
    for line in institution.split('\n'):
        _para(doc, line.strip(), size=14, space_after=2)

    # ═══════════════════════════════════════════
    # 2. ПУСТОЕ МЕСТО
    # ═══════════════════════════════════════════
    _empty(doc, 5)

    # ═══════════════════════════════════════════
    # 3. НАЗВАНИЕ РАБОТЫ
    # ═══════════════════════════════════════════
    work_title = data.get('work_title', 'ОТЧЕТ ПО ПРАКТИЧЕСКОЙ РАБОТЕ')
    work_number = data.get('work_number', '').strip()

    title = work_title.upper()
    if work_number:
        title += f' №{work_number}'

    _para(doc, title, size=14, bold=True, space_after=8)

    # Специальность
    spec_code = data.get('specialty_code', '').strip()
    spec_name = data.get('specialty_name', '').strip()
    if spec_code or spec_name:
        spec_text = f'{spec_code} {spec_name}'.strip()
        _para(doc, spec_text, size=14, space_after=2)
        _para(doc, 'код и наименование специальности', size=10, space_after=6)

    # Дисциплина
    subject = data.get('subject', '').strip()
    if subject:
        _para(doc, subject, size=14, space_after=2)
        _para(doc, 'наименование дисциплины', size=10)

    # ═══════════════════════════════════════════
    # 4. ПУСТОЕ МЕСТО
    # ═══════════════════════════════════════════
    _empty(doc, 3)

    # ═══════════════════════════════════════════
    # 5. ТАБЛИЦА: Студент / Преподаватель
    # ═══════════════════════════════════════════
    group = data.get('group', '').strip()
    student_id = data.get('student_id', '').strip()
    student_name = data.get('student_name', '').strip()
    teacher_name = data.get('teacher_name', '').strip()

    # Собираем "группа, зач.книжка"
    parts = [p for p in [group, student_id] if p]
    group_id = ', '.join(parts)

    table = doc.add_table(rows=5, cols=4)
    _remove_borders(table)

    # Ширины колонок
    widths = [Cm(3), Cm(5.5), Cm(3.5), Cm(4.5)]
    for row in table.rows:
        for i, w in enumerate(widths):
            row.cells[i].width = w

    # Строка 0: Студент | группа, зач | (подпись) | ФИО
    _cell(table.cell(0, 0), 'Студент', size=14)
    _cell(table.cell(0, 1), group_id, size=14, align='center', underline=True)
    _cell(table.cell(0, 2), '', size=14, underline=True)
    _cell(table.cell(0, 3), student_name, size=14, align='right', underline=True)

    # Строка 1: подписи-пояснения
    _cell(table.cell(1, 0), '', size=9)
    _cell(table.cell(1, 1), 'номер группы, зачетной книжки', size=9, align='center')
    _cell(table.cell(1, 2), 'подпись, дата', size=9, align='center')
    _cell(table.cell(1, 3), 'инициалы, фамилия', size=9, align='center')

    # Строка 2: пустая
    for c in range(4):
        _cell(table.cell(2, c), '', size=9)

    # Строка 3: Преподаватель
    _cell(table.cell(3, 0), 'Преподаватель', size=14)
    _cell(table.cell(3, 1), '', size=14, underline=True)
    _cell(table.cell(3, 2), '', size=14, underline=True)
    _cell(table.cell(3, 3), teacher_name, size=14, align='right', underline=True)

    # Строка 4: подписи-пояснения
    _cell(table.cell(4, 0), '', size=9)
    _cell(table.cell(4, 1), '', size=9)
    _cell(table.cell(4, 2), 'подпись, дата', size=9, align='center')
    _cell(table.cell(4, 3), 'инициалы, фамилия', size=9, align='center')

    # ═══════════════════════════════════════════
    # 6. ГОРОД, ГОД
    # ═══════════════════════════════════════════
    _empty(doc, 4)

    city = data.get('city', 'Красноярск').strip()
    year = data.get('year', '2026').strip()
    _para(doc, f'{city}, {year}', size=14)

    # ═══════════════════════════════════════════
    # СОХРАНЕНИЕ
    # ═══════════════════════════════════════════
    _move_sectpr_to_end(doc)
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)

    print(f"  ✅ Титульный лист: {output_path}")
    return output_path