"""
LLM-проверка затекстовой библиографии (раздел «БИБЛИОГРАФИЧЕСКОЕ ОПИСАНИЕ»).

Цель: НЕ править документ, а вернуть список ошибок для отчёта.
"""

from __future__ import annotations

import json
import re
from typing import Any


_BIB_HEADS = {
    "библиографическое описание",
    "библиографический список",
    "список литературы",
    "список использованных источников",
    "список источников",
    "список использованной литературы",
}


def extract_bibliography_entries(paragraph_texts: list[str]) -> list[tuple[int, str]]:
    """
    Возвращает (index, text) для абзацев, входящих в затекстовую библиографию.
    Индекс — глобальный индекс абзаца (как в python-docx doc.paragraphs).
    """
    entries: list[tuple[int, str]] = []
    mode = False
    for i, raw in enumerate(paragraph_texts):
        t = (raw or "").strip()
        if not t:
            continue
        low = t.lower().strip()
        if low in _BIB_HEADS:
            mode = True
            continue
        if not mode:
            continue
        # Выходим при явном крупном заголовке следующего раздела
        if low.startswith("приложени") or low in {"введение", "заключение", "содержание", "оглавление"}:
            break
        if re.match(r"^\d{1,2}(\.|\\s)\\s+\\S", t) and not re.match(r"^\d+\\.", t):
            # Похоже на «1. Техническое ...» или «2 Специальная ...»
            break
        entries.append((i, t))
    return entries


def _build_prompt(entries: list[tuple[int, str]]) -> str:
    """
    Проверяем по примерам пользователя:
    - Затекстовая ссылка: "47. Автор. ... Город, год. С. 256–300."
    - Может быть вариант со '//' для статьи: "... // ... М.: Наука, 1983. С. 109."
    """
    lines = [f"[{idx}] {text}" for idx, text in entries]
    block = "\n".join(lines)
    return (
        "Ты проверяешь затекстовые библиографические ссылки на русском языке.\n"
        "Тебе дают список строк (каждая начинается с [INDEX]).\n\n"
        "ПРАВИЛА (по примерам):\n"
        "- Строка должна начинаться с номера: \"N.\" (например \"47.\")\n"
        "- Должен быть год (4 цифры).\n"
        "- Должно быть указание страниц: \"С. N\" или \"С. N–M\".\n"
        "- Если есть \"//\" (статья/раздел), после него должен быть источник и выходные данные.\n"
        "- НЕ придумывай недостающие данные. Только отмечай, чего не хватает/что выглядит неверно.\n\n"
        "Верни ТОЛЬКО JSON без markdown в формате:\n"
        "{\n"
        "  \"items\": [\n"
        "    {\"index\": 139, \"ok\": true},\n"
        "    {\"index\": 140, \"ok\": false, \"issues\": [\"нет 'С.' с страницами\", \"нет года\"]}\n"
        "  ]\n"
        "}\n\n"
        "СТРОГО: issues — короткие строки (до 120 символов), без предложений на 5 строк.\n\n"
        "СТРОКИ ДЛЯ ПРОВЕРКИ:\n"
        + block
    )


def _parse_items(text: str) -> list[dict[str, Any]]:
    raw = (text or "").strip()
    if not raw:
        return []
    raw = re.sub(r"^```(?:json)?\\s*", "", raw, flags=re.IGNORECASE)
    raw = re.sub(r"\\s*```$", "", raw)
    try:
        data = json.loads(raw)
    except Exception:
        # Попытка вытащить первый JSON объект
        start = raw.find("{")
        end = raw.rfind("}")
        if start >= 0 and end > start:
            try:
                data = json.loads(raw[start:end + 1])
            except Exception:
                return []
        else:
            return []
    items = data.get("items") if isinstance(data, dict) else None
    if not isinstance(items, list):
        return []
    out = []
    for it in items:
        if not isinstance(it, dict):
            continue
        try:
            idx = int(it.get("index"))
        except Exception:
            continue
        ok = bool(it.get("ok", False))
        issues = it.get("issues") or []
        if not isinstance(issues, list):
            issues = []
        issues_s = [str(x).strip()[:200] for x in issues if str(x).strip()]
        out.append({"index": idx, "ok": ok, "issues": issues_s})
    return out


def check_bibliography_entries_with_ai(entries: list[tuple[int, str]]) -> tuple[list[dict[str, Any]], str]:
    """
    Возвращает (items, err). items: [{"index": i, "issues":[...]}] только для ok=false.
    """
    if not entries:
        return [], ""
    try:
        from .ai_postprocess import _call_llm  # reuse providers/keys from settings
    except Exception as exc:
        return [], f"LLM unavailable: {exc}"

    prompt = _build_prompt(entries[:40])
    actions, err, _meta = _call_llm(prompt)
    # _call_llm возвращает list[dict] (actions) для своего формата; здесь используем raw inner text не доступно.
    # Поэтому используем fallback: если провайдер вернул не actions, но err пуст — считаем, что не поддержано.
    if err:
        return [], err
    # Некоторые провайдеры вернут пустой список actions; попробуем распарсить из meta preview через публичный путь нельзя.
    # Поэтому договоримся: используем только OpenRouter/Gemini с JSON-ответом или Ollama с json, где _call_llm
    # уже умеет вытащить list[dict]. Ожидаем, что это будут dict'ы items.
    # Если модель вернула {"items":[...]} — _call_llm вернёт это как список dict? Не гарантировано.
    # Поддержим оба варианта:
    if isinstance(actions, list) and actions and any("items" in a for a in actions if isinstance(a, dict)):
        # Если вернули outer dict как action
        for a in actions:
            if isinstance(a, dict) and isinstance(a.get("items"), list):
                parsed = _parse_items(json.dumps(a, ensure_ascii=False))
                break
        else:
            parsed = []
    else:
        # Иногда модель может вернуть list items напрямую
        parsed = []
        for a in actions:
            if isinstance(a, dict) and "index" in a:
                parsed.append(a)
        # нормализуем в наш формат
        normed = []
        for it in parsed:
            try:
                idx = int(it.get("index"))
            except Exception:
                continue
            ok = bool(it.get("ok", False))
            issues = it.get("issues") or []
            if not isinstance(issues, list):
                issues = []
            issues_s = [str(x).strip()[:200] for x in issues if str(x).strip()]
            normed.append({"index": idx, "ok": ok, "issues": issues_s})
        parsed = normed

    bad = [it for it in parsed if isinstance(it, dict) and not bool(it.get("ok")) and it.get("issues")]
    return bad, ""

