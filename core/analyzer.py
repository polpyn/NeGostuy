"""
Главный модуль — связывает парсер, классификатор, форматтер.
Вызывается из Django views.
"""

import os
import time
import shutil
import tempfile
from collections import Counter

from docx import Document

from .parser import (
    extract_images_from_docx,
    collect_paragraphs,
    assign_numbers_globally,
    check_page_margins,
)
from .classifier import classify_element_simple, classify_all
from .formatter import create_gost_document


# Параметры ГОСТ для проверки
GOST_CHECK = {
    "font": "Times New Roman",
    "font_size": 14,
    "line_spacing": 1.5,
    "indent_cm": 1.25,
}


def check_gost_compliance(p, elem_type):
    """
    Проверяет один элемент на соответствие ГОСТ.
    Возвращает: (status, errors, warnings)
    """
    errors = []
    warnings = []

    # Шрифт
    if p.get("font_name") and p["font_name"] != GOST_CHECK["font"]:
        errors.append(
            f'Шрифт: "{p["font_name"]}" → должен быть "{GOST_CHECK["font"]}"'
        )

    # Размер шрифта
    if p.get("font_size_pt") and p["font_size_pt"] != GOST_CHECK["font_size"]:
        size = p["font_size_pt"]
        errors.append(
            f'Размер шрифта: {size}pt → должен быть {GOST_CHECK["font_size"]}pt'
        )

    # Межстрочный интервал
    if p.get("line_spacing") and p["line_spacing"] != GOST_CHECK["line_spacing"]:
        ls = p["line_spacing"]
        warnings.append(
            f'Межстрочный интервал: {ls} → должен быть {GOST_CHECK["line_spacing"]}'
        )

    # Отступ первой строки
    if elem_type == "text" and p.get("first_indent_cm") is not None:
        indent = p["first_indent_cm"]
        if abs(indent - GOST_CHECK["indent_cm"]) > 0.1:
            errors.append(
                f'Отступ первой строки: {indent} см → должен быть {GOST_CHECK["indent_cm"]} см'
            )

    # Выравнивание
    if elem_type == "text" and p.get("centered"):
        warnings.append('Основной текст выровнен по центру → должен быть по ширине')

    if elem_type == "heading" and not p.get("centered") and not p.get("bold"):
        warnings.append('Заголовок не выделен (не жирный и не по центру)')

    # Статус
    if errors:
        return "error", errors, warnings
    elif warnings:
        return "warning", errors, warnings
    else:
        return "correct", errors, warnings


def analyze_document(file_path, template_path=None):
    """
    Главная функция — полный анализ и форматирование документа.

    Args:
        file_path: путь к исходному .docx
        template_path: путь к рамке .docx (необязательно)

    Returns:
        dict с результатами
    """

    temp_dir = tempfile.mkdtemp(prefix='negostuy_')

    try:
        # === ШАГ 1: ПАРСИНГ ===
        extracted_images = extract_images_from_docx(file_path, temp_dir)
        doc = Document(file_path)
        paragraphs = collect_paragraphs(doc, extracted_images)

        # === ШАГ 2: КЛАССИФИКАЦИЯ И ПРОВЕРКА ===
        elements_detail = []
        type_counts = Counter()
        total_errors = 0
        total_warnings = 0
        errors_list = []

        for p in paragraphs:
            elem_type = classify_element_simple(p)
            type_counts[elem_type] += 1

            elem_status, errors, warnings = check_gost_compliance(p, elem_type)
            total_errors += len(errors)
            total_warnings += len(warnings)

            text_preview = p["text"][:80]
            if len(p["text"]) > 80:
                text_preview += "..."

            elements_detail.append({
                "type": elem_type,
                "text": text_preview,
                "status": elem_status,
                "errors": errors,
                "warnings": warnings,
            })

            for err in errors:
                errors_list.append({
                    "type": _get_error_category(err),
                    "description": err,
                })

        # Поля страницы
        margin_errors = check_page_margins(doc)
        total_errors += len(margin_errors)
        for me in margin_errors:
            errors_list.append(me)
            elements_detail.insert(0, {
                "type": "text",
                "text": me["description"],
                "status": "error",
                "errors": [me["description"]],
                "warnings": [],
            })

        # Оценка
        if total_errors == 0 and total_warnings == 0:
            grade = "✅ ИДЕАЛЬНО — Полное соответствие ГОСТ"
        elif total_errors == 0:
            grade = f"⚠️ ХОРОШО — Есть {total_warnings} предупреждений"
        elif total_errors <= 5:
            grade = f"📝 НЕПЛОХО — {total_errors} ошибок"
        elif total_errors <= 15:
            grade = f"⚠️ УДОВЛЕТВОРИТЕЛЬНО — {total_errors} ошибок"
        else:
            grade = f"❌ ТРЕБУЕТ ИСПРАВЛЕНИЯ — {total_errors} ошибок"

        # === ШАГ 3: ФОРМАТИРОВАНИЕ С РАМКОЙ ===
        numbering, is_list_item = assign_numbers_globally(paragraphs)
        classified_elements = classify_all(paragraphs, numbering, is_list_item)

        output_dir = os.path.join(os.path.dirname(file_path), '..', 'results')
        os.makedirs(output_dir, exist_ok=True)

        base_name = os.path.splitext(os.path.basename(file_path))[0]
        output_path = os.path.join(output_dir, f"{base_name}_GOST.docx")

        # ПЕРЕДАЁМ РАМКУ В ФОРМАТТЕР
        create_gost_document(
            classified_elements,
            output_path,
            template_path=template_path
        )

        if template_path:
            print(f"✅ Документ создан С РАМКОЙ: {output_path}")
        else:
            print(f"✅ Документ создан без рамки: {output_path}")

        images_count = sum(1 for e in classified_elements if e["type"] == "image")

        return {
            'total_elements': len(paragraphs),
            'errors_count': total_errors,
            'warnings_count': total_warnings,
            'headings_count': type_counts.get('heading', 0),
            'paragraphs_count': type_counts.get('text', 0),
            'lists_count': type_counts.get('list_item', 0),
            'images_count': images_count,
            'grade': grade,
            'elements_detail': elements_detail,
            'report': {
                'errors': errors_list,
                'summary': dict(type_counts),
            },
            'output_path': output_path,
        }

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def _get_error_category(error_text):
    """Определяет категорию ошибки по тексту"""
    t = error_text.lower()
    if 'шрифт' in t and 'размер' not in t:
        return 'font'
    elif 'размер' in t:
        return 'font_size'
    elif 'выравн' in t or 'центр' in t:
        return 'alignment'
    elif 'отступ' in t:
        return 'indent'
    elif 'интервал' in t:
        return 'spacing'
    elif 'поле' in t:
        return 'margins'
    return 'other'