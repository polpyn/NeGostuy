"""
Классификатор элементов курсовой работы.

Типы элементов:
  special_heading  — ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ, АННОТАЦИЯ, ПРИЛОЖЕНИЕ ...
                     (новая страница, по центру, ЗАГЛАВНЫЕ, жирный)
  toc_heading      — СОДЕРЖАНИЕ / ОГЛАВЛЕНИЕ
                     (новая страница + поле TOC Word)
  chapter_heading  — нумерованный раздел: «1. Название»
                     (новая страница, по ширине, жирный)
  section_heading  — подраздел: «1.1 Название»
  subsection_heading — пункт: «1.1.1 Название»
  list_item        — маркированный или нумерованный пункт
  figure_caption   — подпись к рисунку
  paragraph        — основной текст
  table            — таблица (rows)
  image            — рисунок (image_paths)
"""

import re


# ----------------------------------------------------------------
# Ключевые слова специальных заголовков 1-го уровня
# ----------------------------------------------------------------
SPECIAL_KEYWORDS = {
    "введение",
    "заключение",
    "аннотация",
    "abstract",
    "список литературы",
    "список использованных источников",
    "список источников",
    "список использованной литературы",
    "библиографический список",
    "перечень сокращений",
    "сокращения",
    "обозначения",
    "определения",
    "нормативные ссылки",
    "термины и определения",
}

TOC_KEYWORDS = {"содержание", "оглавление"}

# «Приложение А» / «Приложение Б» — короткий заголовок (не начало предложения)
# Текст должен быть <= 80 символов и соответствовать шаблону
APPENDIX_RE = re.compile(r"^приложени[еяё](\s+[а-яёА-ЯЁa-zA-Z0-9][\w\-]*(\s+[\w\-]*){0,3})?$", re.IGNORECASE)

# «1. Название» — нумерованный раздел с точкой (1-2 цифры перед точкой)
CHAPTER_DOT_RE = re.compile(r"^\d{1,2}\.\s+\S")
# «2 Название» — нумерованный раздел без точки, следующее слово с заглавной буквы
CHAPTER_SPACE_RE = re.compile(r"^\d{1,2}\s+[А-ЯЁA-Z]")

# 1.1 или 1.1. + текст (без третьего уровня)
SECTION_RE = re.compile(r"^\d+\.\d+\.?\s+\S")
# 1.1.1 или 1.1.1. + текст
SUBSECTION_RE = re.compile(r"^\d+\.\d+\.\d+\.?\s+\S")

FIGURE_CAPTION_RE = re.compile(r"^рисунок\s+\d+", re.IGNORECASE)
LIST_DASH_RE = re.compile(r"^[—\-–•●○]\s+")
NUMBERED_ITEM_RE = re.compile(r"^\d+\.\s+")


def _is_section_heading_text(text: str) -> bool:
    """1.1 / 1.1. — без третьего уровня."""
    return bool(SECTION_RE.match(text)) and not bool(SUBSECTION_RE.match(text))


def classify_kursovaya_element(p_data: dict) -> str:
    """
    Определяет тип одного абзаца для курсовой работы.

    p_data — словарь из parser.paragraph_to_dict, содержит:
      text, bold, centered, has_numbering, level,
      font_size_pt, line_spacing, first_indent_cm,
      has_image, image_paths, style_name (опционально)
    """
    text: str = (p_data.get("text") or "").strip()

    if not text and not p_data.get("has_image"):
        return "empty"

    if p_data.get("has_image"):
        return "image"

    tl = text.lower().strip()

    # ── Специальные заголовки (точное совпадение по ключевым словам) ───
    if tl in TOC_KEYWORDS:
        return "toc_heading"

    if tl in SPECIAL_KEYWORDS:
        return "special_heading"

    # «Приложение А» / «Приложение Б» — короткий заголовок (≤ 80 символов,
    # не обычное предложение типа «Приложение представляет собой...»)
    if APPENDIX_RE.match(tl) and len(text) <= 80:
        return "special_heading"

    # Текст полностью заглавными → специальный заголовок
    if (text.isupper()
            and 3 < len(text) < 120
            and not text.rstrip().endswith(";")
            and not SUBSECTION_RE.match(text)
            and not SECTION_RE.match(text)):
        tl_clean = text.lower()
        if tl_clean in TOC_KEYWORDS:
            return "toc_heading"
        return "special_heading"

    # ── Подпись к рисунку ──────────────────────────────────────────────
    if FIGURE_CAPTION_RE.match(text):
        return "figure_caption"

    # ── Иерархия нумерованных заголовков ──────────────────────────────
    if SUBSECTION_RE.match(text):
        return "subsection_heading"

    if _is_section_heading_text(text):
        return "section_heading"

    # «1. Название» или «2 Название» — нумерованный раздел.
    # Требования: 1-2 цифры, следующее слово с заглавной буквы,
    # длина ≤ 120 символов, НЕ заканчивается на ; или , (признак пункта).
    _ends_as_list = text.rstrip().endswith((";", ","))
    if not _ends_as_list:
        if CHAPTER_DOT_RE.match(text) or CHAPTER_SPACE_RE.match(text):
            rest_match = re.match(r"^\d+[\.\s]\s*(.+)$", text)
            rest = rest_match.group(1).strip() if rest_match else ""
            if rest and rest[0].isupper() and len(rest) <= 120:
                return "chapter_heading"

    # ── Маркированный список ───────────────────────────────────────────
    if LIST_DASH_RE.match(text):
        return "list_item"

    # ── Нумерованный пункт ────────────────────────────────────────────
    if NUMBERED_ITEM_RE.match(text):
        return "list_item"

    # ── Элементы Word-нумерации ────────────────────────────────────────
    if p_data.get("has_numbering"):
        return "list_item"

    return "paragraph"


def classify_kursovaya_blocks(blocks: list, paragraphs: list) -> list:
    """
    Классифицирует все блоки (параграфы и таблицы) для курсовой работы.

    blocks  — список ("p", idx) | ("t", rows)
    paragraphs — список dicts из parser

    Возвращает список elements-dict совместимых с kursovaya_formatter.
    """
    elements: list[dict] = []
    figure_counter = 0

    for kind, payload in blocks:
        if kind == "t":
            elements.append({"type": "table", "rows": payload, "text": ""})
            continue

        p = paragraphs[payload]
        etype = classify_kursovaya_element(p)

        if etype == "empty":
            continue

        if etype == "image":
            elements.append({
                "type": "image",
                "text": "",
                "image_paths": p.get("image_paths", []),
            })
            # Если вслед за изображением идёт подпись — она будет следующим элементом
            continue

        if etype == "figure_caption":
            figure_counter += 1
            elements.append({
                "type": "figure_caption",
                "text": p["text"],
                "figure_num": figure_counter,
            })
            continue

        elements.append({
            "type": etype,
            "text": p["text"],
        })

    return elements
