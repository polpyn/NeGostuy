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
    Простая классификация для отчёта (heading / text / list_item).
    Используется для проверки ГОСТ.
    """
    text = p["text"]
    tl = text.lower().strip()

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


def classify_all(paragraphs, numbering, is_list_item):
    """
    Полная классификация для форматирования документа.
    Возвращает список элементов с детальными типами:
    heading_caps, heading_title, heading_chapter, heading_section,
    heading_subsection, paragraph, numbered_item, list_item,
    image, figure_caption
    """
    elements = []
    figure_counter = 0

    for i, p in enumerate(paragraphs):
        text = p["text"]
        tl = text.lower() if text else ""

        # === ИЗОБРАЖЕНИЕ ===
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
            continue

        if not text:
            continue

        # === ПОДПИСЬ К РИСУНКУ ===
        if is_figure_caption(text):
            figure_counter += 1
            elements.append({
                "type": "figure_caption",
                "text": text,
                "figure_num": figure_counter
            })
            continue

        # === ЗАГОЛОВКИ КАПСОМ ===
        caps_headings = [
            'введение', 'заключение', 'выводы',
            'список литературы', 'список использованных источников',
            'содержание', 'аннотация', 'abstract', 'оглавление'
        ]
        if tl in caps_headings:
            elements.append({"type": "heading_caps", "text": text})
            continue

        # === ЗАГОЛОВКИ БЕЗ КАПСА ===
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
            continue

        if any(re.match(pat, tl) for pat in title_patterns):
            elements.append({"type": "heading_title", "text": text})
            continue

        # === КАПС В ТЕКСТЕ ===
        if text.isupper() and 2 < len(text) < 150 and not text.endswith('.'):
            elements.append({"type": "heading_caps", "text": text})
            continue

        # === НУМЕРАЦИЯ В ТЕКСТЕ: 1.1. или 1.1.1. ===
        if re.match(r'^\d+\.\d+\.\d+\.?\s+', text):
            elements.append({"type": "heading_subsection", "text": text})
            continue
        if re.match(r'^\d+\.\d+\.?\s+', text):
            elements.append({"type": "heading_section", "text": text})
            continue

        # === НУМЕРОВАННЫЙ ЭЛЕМЕНТ (Word numId) ===
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
            continue

        # === НУМЕРАЦИЯ В ТЕКСТЕ: "2. Настроила..." ===
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
                continue

            if rest[0].isupper() and not rest.endswith('.'):
                elements.append({"type": "heading_chapter", "text": text})
                continue

            elements.append({
                "type": "numbered_item",
                "text": rest,
                "number": num
            })
            continue

        # === МАРКЕРЫ В ТЕКСТЕ ===
        if re.match(r'^[•\-–—\*●○]\s*', text):
            elements.append({"type": "list_item", "level": 1, "text": text})
            continue

        # === ОБЫЧНЫЙ ПАРАГРАФ ===
        elements.append({"type": "paragraph", "text": text})

    return elements