"""
Форматтер курсовых работ по правилам учебного заведения.

Правила:
  • Шрифт: Times New Roman 14pt, межстрочный интервал 1,5
  • Поля: левое 3 см, правое 1,5 см, верх/низ 2 см
  • Отступ первой строки: 1,25 см
  • Выравнивание основного текста: по ширине

Заголовки 1-го уровня (special_heading / toc_heading / chapter_heading):
  — всегда начинаются с новой страницы (pageBreakBefore)

  special_heading  (ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ, АННОТАЦИЯ, ПРИЛОЖЕНИЕ …):
    • по центру, ВСЕ ЗАГЛАВНЫЕ, жирный, Heading 1

  toc_heading  (СОДЕРЖАНИЕ / ОГЛАВЛЕНИЕ):
    • по центру, ВСЕ ЗАГЛАВНЫЕ, жирный + поле TOC

  chapter_heading  («1. Название раздела»):
    • по ширине, красная строка 1,25 см, жирный, Heading 1

Заголовки 2-го уровня (section_heading — «1.1 …»):
  • по ширине, красная строка 1,25 см, жирный, Heading 2

Заголовки 3-го уровня (subsection_heading — «1.1.1 …»):
  • по ширине, красная строка 1,25 см, жирный (курсив), Heading 3

Список (list_item): тире «— », по ширине, красная строка 1,25 см

Таблицы: стиль Table Grid, 12pt
Подписи к рисункам: 12pt, по центру, «Рисунок N – Название»
"""

import os
import re

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ================================================================
# КОНСТАНТЫ
# ================================================================

KURSOVAYA = {
    "font":          "Times New Roman",
    "size":          Pt(14),
    "size_table":    Pt(12),
    "size_caption":  Pt(12),
    "line_spacing":  1.5,
    "indent":        Cm(1.25),
    "margin_left":   Cm(3.0),
    "margin_right":  Cm(1.5),
    "margin_top":    Cm(2.0),
    "margin_bottom": Cm(2.0),
}


# ================================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ================================================================

def _set_font(run, bold: bool = False, italic: bool = False, size=None):
    """Устанавливает шрифт TNR нужного размера для run."""
    run.font.name = KURSOVAYA["font"]
    run.font.size = size or KURSOVAYA["size"]
    run.font.bold = bold
    run.font.italic = italic
    try:
        rPr = run._element.get_or_add_rPr()
        fonts_el = rPr.find(qn("w:rFonts"))
        if fonts_el is None:
            fonts_el = OxmlElement("w:rFonts")
            rPr.insert(0, fonts_el)
        fonts_el.set(qn("w:eastAsia"), KURSOVAYA["font"])
        fonts_el.set(qn("w:cs"), KURSOVAYA["font"])
    except Exception:
        pass


def _para_fmt(para,
              alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
              line_spacing: float = 1.5,
              first_line_indent=None,
              space_before: float = 0,
              space_after: float = 0):
    """Устанавливает форматирование абзаца."""
    fmt = para.paragraph_format
    fmt.alignment = alignment
    fmt.line_spacing = line_spacing
    fmt.first_line_indent = first_line_indent
    fmt.space_before = Pt(space_before)
    fmt.space_after = Pt(space_after)


def _add_page_break_before(para):
    """Добавляет pageBreakBefore=true через pPr XML."""
    pPr = para._element.get_or_add_pPr()
    # Удаляем старый тег если есть, чтобы не дублировать
    for old in pPr.findall(qn("w:pageBreakBefore")):
        pPr.remove(old)
    pb = OxmlElement("w:pageBreakBefore")
    pb.set(qn("w:val"), "1")
    pPr.append(pb)


def _clear_heading_style_numbering(para):
    """
    Убирает автоматическую нумерацию из стиля Heading 1/2/3,
    чтобы слово «ВВЕДЕНИЕ» не превращалось в «1 ВВЕДЕНИЕ».
    """
    pPr = para._element.get_or_add_pPr()
    numPr = pPr.find(qn("w:numPr"))
    if numPr is not None:
        pPr.remove(numPr)


def _insert_toc_field(doc: Document):
    """
    Вставляет поле Table of Contents (TOC) после текущей позиции.
    Поле обновляется при первом открытии документа в Word (Ctrl+A, F9).
    Параметры: заголовки Heading 1..3, гиперссылки, без номеров стилей.
    """
    para = doc.add_paragraph()
    _para_fmt(para,
              alignment=WD_ALIGN_PARAGRAPH.LEFT,
              line_spacing=KURSOVAYA["line_spacing"],
              first_line_indent=Cm(0))

    def _run_field(r_elem, fld_type: str, text: str = ""):
        fc = OxmlElement("w:fldChar")
        fc.set(qn("w:fldCharType"), fld_type)
        r_elem.append(fc)
        if text:
            t = OxmlElement("w:t")
            t.set(qn("xml:space"), "preserve")
            t.text = text
            r_elem.append(t)

    # begin
    r_begin = para.add_run()
    _run_field(r_begin._r, "begin")

    # instrText
    r_instr = para.add_run()
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = ' TOC \\o "1-3" \\h \\z \\u '
    r_instr._r.append(instr)

    # separate + placeholder text
    r_sep = para.add_run()
    _run_field(r_sep._r, "separate")

    r_text = para.add_run("Нажмите Ctrl+A, затем F9 для обновления оглавления")
    r_text.font.name = KURSOVAYA["font"]
    r_text.font.size = KURSOVAYA["size"]
    r_text.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

    # end
    r_end = para.add_run()
    _run_field(r_end._r, "end")

    return para


def _gost_image_width(image_path: str, max_w_cm: float = 16.0, max_h_cm: float = 18.0):
    """Возвращает ширину вставки с сохранением пропорций."""
    try:
        from PIL import Image
        with Image.open(image_path) as im:
            px_w, px_h = im.size
            dpi_raw = im.info.get("dpi")
            dpi_x = float(dpi_raw[0]) if dpi_raw and dpi_raw[0] else 96.0
            dpi_y = float(dpi_raw[1]) if dpi_raw and len(dpi_raw) > 1 and dpi_raw[1] else dpi_x
        w_cm = (px_w / dpi_x) * 2.54
        h_cm = (px_h / dpi_y) * 2.54
        if w_cm > 0 and h_cm > 0:
            scale = min(max_w_cm / w_cm, max_h_cm / h_cm, 1.0)
            return Cm(w_cm * scale)
    except Exception:
        pass
    return Cm(min(max_w_cm, 16.0))


# ================================================================
# ГЛАВНАЯ ФУНКЦИЯ
# ================================================================

def create_kursovaya_document(elements: list, output_path: str) -> str:
    """
    Создаёт новый .docx с правильным оформлением курсовой работы.

    Args:
        elements:     список dict из kursovaya_classifier
        output_path:  путь для сохранения результата

    Returns:
        output_path
    """
    print(f"\n{'='*60}")
    print("📚 ФОРМАТИРОВАНИЕ КУРСОВОЙ РАБОТЫ")
    print(f"{'='*60}")

    doc = Document()

    # ── Поля страницы ──────────────────────────────────────────────
    for sec in doc.sections:
        sec.left_margin  = KURSOVAYA["margin_left"]
        sec.right_margin = KURSOVAYA["margin_right"]
        sec.top_margin   = KURSOVAYA["margin_top"]
        sec.bottom_margin = KURSOVAYA["margin_bottom"]

    # Счётчик — первый заголовок 1-го уровня не нуждается в разрыве
    # (иначе документ начнётся с пустой страницы)
    _first_lvl1 = True
    _written = 0

    for elem in elements:
        etype = elem.get("type", "paragraph")
        text  = (elem.get("text") or "").strip()

        # ── TABLE ─────────────────────────────────────────────────
        if etype == "table":
            rows_data = elem.get("rows") or []
            if not rows_data:
                continue
            ncol  = max(len(r) for r in rows_data)
            nrows = len(rows_data)
            tbl = doc.add_table(rows=nrows, cols=ncol)
            try:
                tbl.style = "Table Grid"
            except (KeyError, ValueError):
                pass
            for ri, row_cells in enumerate(rows_data):
                for ci in range(ncol):
                    txt = row_cells[ci] if ci < len(row_cells) else ""
                    cell = tbl.cell(ri, ci)
                    cell.text = txt
                    for p in cell.paragraphs:
                        _para_fmt(p,
                                  alignment=WD_ALIGN_PARAGRAPH.LEFT,
                                  line_spacing=KURSOVAYA["line_spacing"],
                                  first_line_indent=Cm(0))
                        for r in p.runs:
                            _set_font(r, size=KURSOVAYA["size_table"])
            _written += 1
            continue

        # ── IMAGE ─────────────────────────────────────────────────
        if etype == "image":
            for img_path in elem.get("image_paths", []):
                if not os.path.exists(img_path):
                    continue
                para = doc.add_paragraph()
                _para_fmt(para,
                          alignment=WD_ALIGN_PARAGRAPH.CENTER,
                          line_spacing=KURSOVAYA["line_spacing"],
                          first_line_indent=Cm(0),
                          space_before=6, space_after=6)
                run = para.add_run()
                try:
                    w = _gost_image_width(img_path)
                    run.add_picture(img_path, width=w)
                    _written += 1
                except Exception as exc:
                    print(f"  ⚠️  Изображение: {exc}")
            continue

        # ── FIGURE CAPTION ────────────────────────────────────────
        if etype == "figure_caption":
            clean = re.sub(
                r"^рисунок\s+\d+\s*[–\-—.:]\s*", "", text, flags=re.IGNORECASE
            )
            fig_num = elem.get("figure_num", 1)
            para = doc.add_paragraph()
            _para_fmt(para,
                      alignment=WD_ALIGN_PARAGRAPH.CENTER,
                      line_spacing=1.0,
                      first_line_indent=Cm(0),
                      space_before=6, space_after=12)
            run = para.add_run(f"Рисунок {fig_num} – {clean}")
            _set_font(run, size=KURSOVAYA["size_caption"])
            _written += 1
            continue

        # Все остальные типы требуют ненулевого текста
        if not text:
            continue

        # ── TOC HEADING (СОДЕРЖАНИЕ / ОГЛАВЛЕНИЕ) ────────────────
        if etype == "toc_heading":
            para = doc.add_paragraph(style="Heading 1")
            _clear_heading_style_numbering(para)
            if not _first_lvl1:
                _add_page_break_before(para)
            _first_lvl1 = False
            _para_fmt(para,
                      alignment=WD_ALIGN_PARAGRAPH.CENTER,
                      line_spacing=KURSOVAYA["line_spacing"],
                      first_line_indent=Cm(0),
                      space_before=0, space_after=12)
            run = para.add_run(text.upper())
            _set_font(run, bold=True)
            # Вставляем поле TOC
            _insert_toc_field(doc)
            _written += 1
            continue

        # ── SPECIAL HEADING (ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ, ПРИЛОЖЕНИЕ …) ─
        if etype == "special_heading":
            para = doc.add_paragraph(style="Heading 1")
            _clear_heading_style_numbering(para)
            if not _first_lvl1:
                _add_page_break_before(para)
            _first_lvl1 = False
            _para_fmt(para,
                      alignment=WD_ALIGN_PARAGRAPH.CENTER,
                      line_spacing=KURSOVAYA["line_spacing"],
                      first_line_indent=Cm(0),
                      space_before=0, space_after=12)
            run = para.add_run(text.upper())
            _set_font(run, bold=True)
            _written += 1
            continue

        # ── CHAPTER HEADING («1. Название раздела») ──────────────
        if etype == "chapter_heading":
            para = doc.add_paragraph(style="Heading 1")
            _clear_heading_style_numbering(para)
            if not _first_lvl1:
                _add_page_break_before(para)
            _first_lvl1 = False
            _para_fmt(para,
                      alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                      line_spacing=KURSOVAYA["line_spacing"],
                      first_line_indent=KURSOVAYA["indent"],
                      space_before=0, space_after=12)
            run = para.add_run(text)
            _set_font(run, bold=True)
            _written += 1
            continue

        # ── SECTION HEADING («1.1 Название») ─────────────────────
        if etype == "section_heading":
            para = doc.add_paragraph(style="Heading 2")
            _clear_heading_style_numbering(para)
            _para_fmt(para,
                      alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                      line_spacing=KURSOVAYA["line_spacing"],
                      first_line_indent=KURSOVAYA["indent"],
                      space_before=12, space_after=6)
            run = para.add_run(text)
            _set_font(run, bold=True)
            _written += 1
            continue

        # ── SUBSECTION HEADING («1.1.1 Название») ────────────────
        if etype == "subsection_heading":
            para = doc.add_paragraph(style="Heading 3")
            _clear_heading_style_numbering(para)
            _para_fmt(para,
                      alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                      line_spacing=KURSOVAYA["line_spacing"],
                      first_line_indent=KURSOVAYA["indent"],
                      space_before=6, space_after=6)
            run = para.add_run(text)
            _set_font(run, bold=True)
            _written += 1
            continue

        # ── LIST ITEM ─────────────────────────────────────────────
        if etype == "list_item":
            clean = re.sub(r"^[—\-–•●○]\s*", "", text).strip()
            # Сохраняем нумерованные пункты как есть, маркированные — с тире
            if re.match(r"^\d+\.", text):
                display = text
            else:
                display = "— " + clean
            para = doc.add_paragraph()
            _para_fmt(para,
                      alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                      line_spacing=KURSOVAYA["line_spacing"],
                      first_line_indent=KURSOVAYA["indent"],
                      space_before=0, space_after=0)
            run = para.add_run(display)
            _set_font(run)
            _written += 1
            continue

        # ── PARAGRAPH (основной текст) ────────────────────────────
        para = doc.add_paragraph()
        _para_fmt(para,
                  alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                  line_spacing=KURSOVAYA["line_spacing"],
                  first_line_indent=KURSOVAYA["indent"],
                  space_before=0, space_after=0)
        run = para.add_run(text)
        _set_font(run)
        _written += 1

    # ── Сохранение ────────────────────────────────────────────────
    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    doc.save(output_path)

    print(f"\n  ✅ Записано элементов: {_written}")
    print(f"  📂 Файл: {output_path} ({os.path.getsize(output_path):,} байт)")

    return output_path
