"""
Классификатор элементов документа.
Определяет тип каждого параграфа:
heading, text, list_item, image, figure_caption и т.д.
"""

import re


def is_figure_caption(text):
    """Проверяет, является ли текст подписью к рисунку"""
    return bool(re.match(r'^рисунок\s+\d+', text.lower()))


def classify_element_simple(p):
    """
    Простая классификация для отчёта (heading / text / list_item / figure_caption).
    Используется для проверки ГОСТ.
    """
    text = p["text"]
    tl = text.lower().strip()

    if text and is_figure_caption(text):
        return "figure_caption"

    caps_headings = [
        'введение', 'заключение', 'выводы',
        'список литературы', 'список использованных источников',
        'содержание', 'аннотация', 'abstract', 'оглавление'
    ]
    if tl in caps_headings:
        return "heading"

    title_headings = [
        'ход работы', 'теоретическая часть', 'практическая часть',
        'цель работы', 'задание', 'вывод', 'теория',
    ]
    if tl in title_headings:
        return "heading"

    title_patterns = [
        r'^практическая\s+работа', r'^лабораторная\s+работа',
        r'^курсовая\s+работа', r'^контрольная\s+работа',
        r'^отчет\s+по\s+практической', r'^отчёт\s+по\s+практической',
        r'^отчет\s+по\s+лабораторной', r'^отчёт\s+по\s+лабораторной',
    ]
    if any(re.match(pat, tl) for pat in title_patterns):
        return "heading"

    if text.isupper() and 2 < len(text) < 150 and not text.endswith('.'):
        return "heading"

    if re.match(r'^\d+\.\d+\.\d+\.?\s+', text):
        return "heading"
    if re.match(r'^\d+\.\d+\.?\s+', text):
        return "heading"

    if p.get("bold") and p.get("centered") and len(text) < 100:
        return "heading"

    if p.get("has_numbering"):
        return "list_item"

    if re.match(r'^[•\-–—\*●○]\s*', text):
        return "list_item"

    if re.match(r'^\d+\.\s+', text):
        return "list_item"

    return "text"


def classify_all_from_blocks(blocks, paragraphs, numbering, is_list_item):
    """
    Классификация с сохранением порядка таблиц в теле документа.
    blocks: список ("p", idx) | ("t", rows).
    """
    elements = []
    figure_counter = 0

    for kind, payload in blocks:
        if kind == "t":
            elements.append({"type": "table", "rows": payload})
            continue
        i = payload
        p = paragraphs[i]
        figure_counter = _classify_one_paragraph(
            elements, p, i, numbering, is_list_item, figure_counter
        )

    return elements


def _classify_one_paragraph(elements, p, i, numbering, is_list_item, figure_counter):
    """Добавляет в elements элементы для одного параграфа; возвращает новый figure_counter."""
    text = p["text"]
    tl = text.lower() if text else ""

    if p.get("has_image"):
        elements.append({
            "type": "image",
            "text": "",
            "image_paths": p.get("image_paths", [])
        })
        if text and is_figure_caption(text):
            figure_counter += 1
            elements.append({
                "type": "figure_caption",
                "text": text,
                "figure_num": figure_counter
            })
        elif text:
            elements.append({"type": "paragraph", "text": text})
        return figure_counter

    if not text:
        return figure_counter

    if is_figure_caption(text):
        figure_counter += 1
        elements.append({
            "type": "figure_caption",
            "text": text,
            "figure_num": figure_counter
        })
        return figure_counter

    caps_headings = [
        'введение', 'заключение', 'выводы',
        'список литературы', 'список использованных источников',
        'содержание', 'аннотация', 'abstract', 'оглавление'
    ]
    if tl in caps_headings:
        elements.append({"type": "heading_caps", "text": text})
        return figure_counter

    title_headings = [
        'ход работы', 'теоретическая часть', 'практическая часть',
        'цель работы', 'задание', 'вывод', 'теория',
    ]
    title_patterns = [
        r'^практическая\s+работа', r'^лабораторная\s+работа',
        r'^контрольная\s+работа', r'^курсовая\s+работа',
        r'^отчет\s+по\s+практической', r'^отчёт\s+по\s+практической',
        r'^отчет\s+по\s+лабораторной', r'^отчёт\s+по\s+лабораторной',
    ]

    if tl in title_headings:
        elements.append({"type": "heading_title", "text": text})
        return figure_counter

    if any(re.match(pat, tl) for pat in title_patterns):
        elements.append({"type": "heading_title", "text": text})
        return figure_counter

    if text.isupper() and 2 < len(text) < 150 and not text.endswith('.'):
        elements.append({"type": "heading_caps", "text": text})
        return figure_counter

    if re.match(r'^\d+\.\d+\.\d+\.?\s+', text):
        elements.append({"type": "heading_subsection", "text": text})
        return figure_counter
    if re.match(r'^\d+\.\d+\.?\s+', text):
        elements.append({"type": "heading_section", "text": text})
        return figure_counter

    if p.get("has_numbering"):
        num = numbering.get(i, 1)
        is_list = is_list_item.get(i, False)

        if is_list:
            elements.append({
                "type": "numbered_item",
                "text": text,
                "number": num
            })
        else:
            if len(text) < 80 and not text.endswith('.'):
                lvl = p.get("level", 0)
                if lvl == 0:
                    elements.append({"type": "heading_chapter", "text": text})
                elif lvl == 1:
                    elements.append({"type": "heading_section", "text": text})
                else:
                    elements.append({"type": "heading_subsection", "text": text})
            else:
                elements.append({
                    "type": "numbered_item",
                    "text": text,
                    "number": num
                })
        return figure_counter

    match_num = re.match(r'^(\d+)\.\s+(.+)', text)
    if match_num:
        num = int(match_num.group(1))
        rest = match_num.group(2)

        if len(rest) > 30:
            elements.append({
                "type": "numbered_item",
                "text": rest,
                "number": num
            })
            return figure_counter

        if rest[0].isupper() and not rest.endswith('.'):
            elements.append({"type": "heading_chapter", "text": text})
            return figure_counter

        elements.append({
            "type": "numbered_item",
            "text": rest,
            "number": num
        })
        return figure_counter

    if re.match(r'^[•\-–—\*●○]\s*', text):
        elements.append({"type": "list_item", "level": 1, "text": text})
        return figure_counter

    elements.append({"type": "paragraph", "text": text})
    return figure_counter


def classify_all(paragraphs, numbering, is_list_item):
    """Классификация только параграфов (без таблиц), обратная совместимость."""
    blocks = [("p", i) for i in range(len(paragraphs))]
    return classify_all_from_blocks(blocks, paragraphs, numbering, is_list_item)
