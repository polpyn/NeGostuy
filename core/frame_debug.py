"""
Отладка подстановок в рамку (шапки/низ, document.xml).

Включение записи снимков при патче:
  set NEGOSTUY_FRAME_DEBUG=1
  (Windows PowerShell: $env:NEGOSTUY_FRAME_DEBUG="1")

Или путь к папке:
  set NEGOSTUY_FRAME_DEBUG=D:\\tmp\\negostuy_debug

В debug_frame/<timestamp>/ пишутся *_BEFORE.xml, *_AFTER.xml, meta и summary.
"""

from __future__ import annotations

import os
import re
import zipfile
from pathlib import Path


def project_root() -> Path:
    return Path(__file__).resolve().parent.parent


def resolve_frame_debug_dir() -> str | None:
    raw = os.environ.get("NEGOSTUY_FRAME_DEBUG", "").strip()
    if not raw:
        return None
    root = project_root()
    low = raw.lower()
    if low in ("1", "true", "yes", "on"):
        d = root / "debug_frame"
        d.mkdir(parents=True, exist_ok=True)
        return str(d)
    p = Path(raw)
    p.mkdir(parents=True, exist_ok=True)
    return str(p)


def scan_docx_xml(docx_path: str) -> dict:
    """Сводка по плейсхолдерам и полям во всех word/*.xml."""
    out: dict = {"path": docx_path, "parts": []}
    with zipfile.ZipFile(docx_path, "r") as z:
        for name in sorted(z.namelist()):
            if not name.startswith("word/") or not name.endswith(".xml"):
                continue
            if "/_rels/" in name:
                continue
            raw = z.read(name)
            try:
                text = raw.decode("utf-8")
            except UnicodeDecodeError:
                text = raw.decode("utf-8", "replace")
            flags = []
            if "{{" in text:
                flags.append("has_double_brace")
            if "nomer" in text.lower():
                flags.append("has_nomer_substr")
            if "{{nomer}}" in text:
                flags.append("literal_nomer")
            if "{{list}}" in text:
                flags.append("literal_list")
            if "fldSimple" in text and "PAGE" in text:
                flags.append("fldSimple_PAGE")
            if "w:sz w:val=" in text:
                m = re.findall(r'<w:sz w:val="(\d+)"', text)
                if m:
                    flags.append(f"sz_vals_sample={m[:8]}")
            if "w:szCs w:val=" in text:
                m = re.findall(r'<w:szCs w:val="(\d+)"', text)
                if m:
                    flags.append(f"szCs_vals_sample={m[:8]}")
            if flags:
                out["parts"].append({"file": name, "flags": flags, "size": len(text)})
    return out


def format_scan_report(scan: dict) -> str:
    lines = [f"Файл: {scan['path']}", ""]
    if not scan["parts"]:
        lines.append("(Нет совпадений по типичным маркерам в word/*.xml)")
        return "\n".join(lines)
    for p in scan["parts"]:
        lines.append(f"--- {p['file']} ({p['size']} bytes) ---")
        for f in p["flags"]:
            lines.append(f"  • {f}")
    return "\n".join(lines)


def write_patch_debug_session(
    debug_parent: str,
    docx_path: str,
    zachet_number: str | None,
    before: dict[str, str],
) -> str | None:
    """Пишет каталог с BEFORE; AFTER нужно дописать после сохранения docx."""
    import datetime

    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    session = os.path.join(debug_parent, stamp)
    os.makedirs(session, exist_ok=True)

    meta = os.path.join(session, "00_meta.txt")
    with open(meta, "w", encoding="utf-8") as f:
        f.write(f"zachet_number = {zachet_number!r}\n")
        f.write(f"docx_path = {docx_path}\n")
        f.write(f"NEGOSTUY_FRAME_DEBUG = {os.environ.get('NEGOSTUY_FRAME_DEBUG', '')!r}\n")
        f.write(f"CWD = {os.getcwd()!r}\n")

    for name, text in sorted(before.items()):
        safe = name.replace("/", "__")
        with open(os.path.join(session, f"{safe}_BEFORE.xml"), "w", encoding="utf-8") as f:
            f.write(text)

    return session


def write_after_and_summary(session: str, docx_path: str, before: dict[str, str]) -> None:
    """После patch_docx_headers: дописать AFTER и summary."""
    summary_path = os.path.join(session, "01_summary.txt")
    lines: list[str] = []

    with zipfile.ZipFile(docx_path, "r") as z:
        for name in sorted(before.keys()):
            safe = name.replace("/", "__")
            after = z.read(name).decode("utf-8")
            with open(os.path.join(session, f"{safe}_AFTER.xml"), "w", encoding="utf-8") as f:
                f.write(after)
            b, a = before[name], after
            lines.append(f"\n{'='*60}\n{name}\n{'='*60}")
            lines.append(f"  BEFORE: {{nomer}} literal = {'{{nomer}}' in b}")
            lines.append(f"  AFTER:  {{nomer}} literal  = {'{{nomer}}' in a}")
            lines.append(f"  BEFORE: {{list}} literal = {'{{list}}' in b}")
            lines.append(f"  AFTER:  {{list}} literal   = {'{{list}}' in a}")
            lines.append(f"  BEFORE: substring 'nomer' (ci) = {'nomer' in b.lower()}")
            lines.append(f"  AFTER:  substring 'nomer' (ci) = {'nomer' in b.lower()}")
            lines.append(f"  AFTER:  fldSimple PAGE     = {'fldSimple' in a and 'PAGE' in a}")
            # фрагменты, если плейсхолдер остался
            if "{{nomer}}" in a or "{{list}}" in a or "{{" in a:
                i = a.find("{{")
                if i >= 0:
                    frag = a[max(0, i - 60) : i + 80]
                    lines.append("  FRAGMENT AFTER @ pos %d: %r" % (i, frag))

    scan = scan_docx_xml(docx_path)
    lines.append("\n\n=== Полное сканирование результата ===\n")
    lines.append(format_scan_report(scan))

    with open(summary_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    readme = os.path.join(project_root(), "debug_frame", "README_HOWTO.txt")
    os.makedirs(os.path.dirname(readme), exist_ok=True)
    if not os.path.isfile(readme):
        with open(readme, "w", encoding="utf-8") as f:
            f.write(
                """Как включить отладку рамки
========================

PowerShell (перед py manage.py runserver):
  $env:NEGOSTUY_FRAME_DEBUG="1"

CMD:
  set NEGOSTUY_FRAME_DEBUG=1

После «Проверить и исправить» смотрите папку:
  debug_frame/<дата_время>/
   00_meta.txt       — что пришло на сервер (zachet)
    *_BEFORE.xml      — XML до патча
    *_AFTER.xml       — XML после патча
    01_summary.txt    — сравнение и скан

Без сервера — просканировать любой .docx:
  py debug_frame_scan.py путь\\к\\файлу.docx
  py debug_frame_scan.py путь\\к\\файлу.docx выходная_папка
"""
            )
