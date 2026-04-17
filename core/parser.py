"""
Парсер документов .docx
Извлекает параграфы, изображения, свойства форматирования
"""

import os
import re
import zipfile
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph


def extract_images_from_docx(docx_path, output_folder):
    """Извлекает все изображения из .docx в папку"""
    images = {}
    os.makedirs(output_folder, exist_ok=True)
    with zipfile.ZipFile(docx_path, 'r') as z:
        for fi in z.filelist:
            if fi.filename.startswith('word/media/'):
                name = os.path.basename(fi.filename)
                path = os.path.join(output_folder, name)
                with z.open(fi.filename) as src, open(path, 'wb') as dst:
                    dst.write(src.read())
                images[name] = path
    return images


def get_paragraph_images(para):
    """Находит ID изображений внутри параграфа"""
    rIds = []
    for drawing in para._p.findall('.//' + qn('w:drawing')):
        for blip in drawing.findall('.//' + qn('a:blip')):
            embed = blip.get(qn('r:embed'))
            if embed and embed not in rIds:
                rIds.append(embed)
    return rIds


def resolve_image_path(doc, rId, extracted_images):
    """Преобразует rId в путь к файлу изображения"""
    try:
        rel = doc.part.rels[rId]
        name = os.path.basename(rel.target_ref)
        return extracted_images.get(name)
    except:
        return None


def paragraph_to_dict(para, doc, extracted_images, index):
    """
    Словарь свойств параграфа или None, если пустой (без текста и без картинки).
    """
    text = para.text.strip()

    image_rIds = get_paragraph_images(para)
    has_image = len(image_rIds) > 0
    image_paths = []
    for rId in image_rIds:
        path = resolve_image_path(doc, rId, extracted_images)
        if path:
            image_paths.append(path)

    if not text and not has_image:
        return None

    numId = None
    level = 0
    pPr = para._p.pPr
    if pPr is not None and pPr.numPr is not None:
        if pPr.numPr.numId is not None:
            numId = int(pPr.numPr.numId.val)
        if pPr.numPr.ilvl is not None:
            level = int(pPr.numPr.ilvl.val)

    bold = False
    if para.runs:
        non_empty = [r for r in para.runs if r.text.strip()]
        if non_empty:
            bold = all(r.bold for r in non_empty)

    centered = False
    try:
        centered = para.paragraph_format.alignment == WD_ALIGN_PARAGRAPH.CENTER
    except Exception:
        pass

    font_name = None
    font_size_pt = None
    if para.runs and len(para.runs) > 0:
        run = para.runs[0]
        if run.font:
            font_name = run.font.name
            if run.font.size:
                font_size_pt = run.font.size.pt

    line_spacing = None
    try:
        ls = para.paragraph_format.line_spacing
        if ls:
            line_spacing = float(ls)
    except Exception:
        pass

    first_indent_cm = None
    try:
        fi = para.paragraph_format.first_line_indent
        if fi:
            first_indent_cm = round(fi / 360000, 2)
    except Exception:
        pass

    return {
        "index": index,
        "text": text,
        "numId": numId,
        "level": level,
        "bold": bold,
        "centered": centered,
        "has_numbering": numId is not None,
        "has_image": has_image,
        "image_paths": image_paths,
        "font_name": font_name,
        "font_size_pt": font_size_pt,
        "line_spacing": line_spacing,
        "first_indent_cm": first_indent_cm,
    }


def collect_paragraphs(doc, extracted_images):
    """
    Собирает параграфы тела документа (как doc.paragraphs — без ячеек таблиц).

    Возвращает список словарей:
    - index, text, numId, level, bold, centered
    - has_numbering, has_image, image_paths
    - font_name, font_size_pt, line_spacing, first_indent_cm
    """
    paragraphs = []
    for i, para in enumerate(doc.paragraphs):
        d = paragraph_to_dict(para, doc, extracted_images, i)
        if d:
            paragraphs.append(d)
    return paragraphs


def _table_to_rows(table: Table) -> list[list[str]]:
    """Текст ячеек по строкам (упрощённо, без учёта сложного merge)."""
    rows_out: list[list[str]] = []
    for row in table.rows:
        cells = [
            cell.text.replace("\r", " ").replace("\n", " ").strip()
            for cell in row.cells
        ]
        rows_out.append(cells)
    return rows_out


def collect_ordered_blocks(doc, extracted_images):
    """
    Порядок блоков тела документа: параграфы и таблицы как в оригинале.

    Возвращает:
      blocks — список ("p", idx) | ("t", rows), где idx — индекс в par_list;
      par_list — список словарей параграфов (только непустые), для нумерации и classify.
    """
    blocks: list = []
    par_list: list = []
    body_para_index = 0
    for item in doc.iter_inner_content():
        if isinstance(item, Paragraph):
            d = paragraph_to_dict(item, doc, extracted_images, body_para_index)
            body_para_index += 1
            if d is None:
                continue
            par_list.append(d)
            blocks.append(("p", len(par_list) - 1))
        elif isinstance(item, Table):
            rows = _table_to_rows(item)
            if rows:
                blocks.append(("t", rows))
    return blocks, par_list


def assign_numbers_globally(paragraphs):
    """
    Группирует нумерованные элементы по numId+level.
    Возвращает:
      numbering: {index: номер}
      is_list_item: {index: True/False}
    """
    groups = {}
    for i, p in enumerate(paragraphs):
        if p["has_numbering"]:
            key = (p["numId"], p["level"])
            if key not in groups:
                groups[key] = []
            groups[key].append(i)

    numbering = {}
    is_list_item = {}

    for key, indices in groups.items():
        is_multi = len(indices) >= 2
        for num, idx in enumerate(indices, 1):
            numbering[idx] = num
            is_list_item[idx] = is_multi

    return numbering, is_list_item


def check_page_margins(doc):
    """Проверяет поля страницы на соответствие ГОСТ"""
    errors = []
    expected = {
        "Левое поле": 3.0,
        "Правое поле": 1.5,
        "Верхнее поле": 2.0,
        "Нижнее поле": 2.0,
    }

    try:
        section = doc.sections[0]
        actuals = {
            "Левое поле": section.left_margin,
            "Правое поле": section.right_margin,
            "Верхнее поле": section.top_margin,
            "Нижнее поле": section.bottom_margin,
        }

        for name, actual in actuals.items():
            if actual:
                actual_cm = round(actual / 360000, 1)
                expected_cm = expected[name]
                if abs(actual_cm - expected_cm) > 0.2:
                    errors.append({
                        'type': 'margins',
                        'description': f'{name}: {actual_cm} см → должно быть {expected_cm} см',
                    })
    except:
        pass

    return errors